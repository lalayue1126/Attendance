/**
 * ══════════════════════════════════════════════════════════════
 *  NFC / QR 打刻システム — Google Apps Script バックエンド
 *
 *  対になるフロントエンド : index.html (GitHub Pages)
 *
 *  【最初にやること】
 *   1. CONFIG.SHEET_ID にスプレッドシートIDを入れる
 *   2. メニューから setup() を1回実行する
 *   3. ウェブアプリとしてデプロイ
 *        実行するユーザー   : 自分
 *        アクセスできるユーザー: 全員
 *   4. 発行された /exec URL を index.html の CONFIG.API_URL に貼る
 * ══════════════════════════════════════════════════════════════
 */

const CONFIG = {
  SHEET_ID: 'ここにスプレッドシートIDを入れる',

  TZ: 'Asia/Singapore',

  // 勤務日の境界。4 なら 00:00〜03:59 の打刻は前日の勤務として扱う
  DAY_BOUNDARY_HOUR: 4,

  // 同一種別の連続打刻をこの秒数内は無視する（誤タップ・二重読み対策）
  DEDUP_WINDOW_SEC: 60,

  // FIXED 拠点の既定半径（locations の radius_m が空のとき使用）
  DEFAULT_RADIUS_M: 100,

  // 逆ジオコーディング（PORTABLE のみ）。不要なら false
  REVERSE_GEOCODE: true
};

const SHEETS = {
  EMPLOYEES: 'employees',
  DEVICES: 'devices',
  LOCATIONS: 'locations',
  EVENTS: 'punch_events',
  CORRECTIONS: 'corrections'
};

const TYPE_LABEL = {
  IN: '出勤', OUT: '退勤', BREAK_START: '休憩開始', BREAK_END: '休憩終了'
};


/* ════════════════════════════════════════════
   1. エントリポイント
   ════════════════════════════════════════════ */

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return json({ ok: false, error: 'BAD_REQUEST', message: 'リクエストが空です。' });
    }

    const req = JSON.parse(e.postData.contents);

    switch (req.action) {
      case 'register': return json(handleRegister(req));
      case 'state':    return json(handleState(req));
      case 'punch':    return json(handlePunch(req));
      default:
        return json({ ok: false, error: 'BAD_ACTION', message: '不明な操作です。' });
    }

  } catch (err) {
    // 例外を握りつぶさず記録する。原因調査で必ず要る。
    logError(err);
    return json({
      ok: false, error: 'SERVER',
      message: 'サーバー側でエラーが発生しました。管理者に連絡してください。'
    });
  }
}

/**
 * doGet は打刻には使わないが、デプロイ確認用に生かしておく。
 * ブラウザで /exec を開いて {"ok":true,...} が出れば公開設定は正しい。
 */
function doGet() {
  return json({ ok: true, service: 'punch-api', time: new Date().toISOString() });
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


/* ════════════════════════════════════════════
   2. 端末登録
   ════════════════════════════════════════════ */

function handleRegister(req) {
  const code = String(req.employee_code || '').trim().toUpperCase();
  const pin  = String(req.pin || '').trim();

  if (!code || !pin) {
    return { ok: false, error: 'NOT_FOUND' };
  }

  const emp = findEmployee(code);
  if (!emp || !truthy(emp.is_active)) {
    return { ok: false, error: 'NOT_FOUND' };
  }

  if (hashPin(pin, emp.pin_salt) !== String(emp.pin_hash)) {
    return { ok: false, error: 'BAD_PIN' };
  }

  const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
  const now = new Date();

  sheet(SHEETS.DEVICES).appendRow([
    token, emp.employee_code, now, now,
    String(req.label || '').slice(0, 60), true
  ]);

  return { ok: true, token: token, name: emp.name };
}


/* ════════════════════════════════════════════
   3. 打刻画面の初期表示
   ════════════════════════════════════════════ */

function handleState(req) {
  const emp = authenticate(req.token);
  if (!emp) return { ok: false, error: 'AUTH' };

  const site = findLocation(req.l);
  if (!site) return { ok: false, error: 'UNKNOWN_LOC' };

  const last = getLastPunchToday(emp.employee_code);

  return {
    ok: true,
    name: emp.name,
    loc_name: site.name,
    loc_type: site.type,
    suggest: suggestNext(last),
    last: last ? {
      type: last.punch_type,
      type_label: TYPE_LABEL[last.punch_type] || last.punch_type,
      at_label: Utilities.formatDate(last.punched_at, CONFIG.TZ, 'HH:mm')
    } : null
  };
}

/**
 * 次に押すべき打刻種別を推定する。
 * あくまで「おすすめ」であり、確定はしない。
 * 自動確定にすると打刻漏れ1回で以降すべてが反転してしまう。
 */
function suggestNext(last) {
  if (!last) return 'IN';
  return last.punch_type === 'OUT' ? 'IN' : 'OUT';
}


/* ════════════════════════════════════════════
   4. 打刻の記録
   ════════════════════════════════════════════ */

function handlePunch(req) {
  const lock = LockService.getScriptLock();

  // 同時打刻で行が壊れるのを防ぐ。取れなければ諦めて再試行を促す。
  try {
    lock.waitLock(20000);
  } catch (e) {
    return { ok: false, error: 'BUSY', message: '混み合っています。もう一度お試しください。' };
  }

  try {
    const emp = authenticate(req.token);
    if (!emp) return { ok: false, error: 'AUTH' };

    const site = findLocation(req.l);
    if (!site) return { ok: false, error: 'UNKNOWN_LOC' };

    const type = String(req.type || '').toUpperCase();
    if (!TYPE_LABEL[type]) {
      return { ok: false, error: 'BAD_TYPE', message: '打刻の種別が不正です。' };
    }

    // ── 冪等性チェック : オフライン再送で二重に入らないようにする
    const uuidKey = 'uuid:' + req.client_uuid;
    const cache = CacheService.getScriptCache();
    if (req.client_uuid && (cache.get(uuidKey) || existsUuid(req.client_uuid))) {
      return { ok: true, dedup: true, punched_at: new Date().toISOString(),
               loc_name: site.name, geo_status: 'OK' };
    }

    // ── 誤タップ対策 : 直近の同一種別を無視
    const last = getLastPunchToday(emp.employee_code);
    if (last && last.punch_type === type &&
        (new Date() - last.punched_at) < CONFIG.DEDUP_WINDOW_SEC * 1000) {
      return { ok: false, error: 'TOO_SOON' };
    }

    // ── 位置情報の判定 : ここで固定 / 持ち運びが分岐する
    const geo = req.geo || { status: 'TIMEOUT' };
    let geoStatus = String(geo.status || 'TIMEOUT');
    let distance = '';
    let address  = '';

    if (geoStatus === 'OK' && isNum(geo.lat) && isNum(geo.lng)) {

      if (site.type === 'FIXED' && isNum(site.lat) && isNum(site.lng)) {
        const radius = isNum(site.radius_m) ? Number(site.radius_m) : CONFIG.DEFAULT_RADIUS_M;
        distance = haversine(geo.lat, geo.lng, Number(site.lat), Number(site.lng));

        // GPS 誤差を許容範囲に加算する。
        // これをやらないと屋内 Wi-Fi 測位で正当な打刻が大量に弾かれる。
        const tolerance = radius + Math.min(Number(geo.accuracy) || 0, 300);
        geoStatus = distance <= tolerance ? 'OK' : 'OUT_OF_RANGE';

      } else {
        // PORTABLE は検証せず、どこで打刻したかを住所として残す
        address = CONFIG.REVERSE_GEOCODE ? reverseGeocode(geo.lat, geo.lng) : '';
      }
    }

    // ── 追記 : 時刻はサーバー側を正とする
    const now = new Date();
    sheet(SHEETS.EVENTS).appendRow([
      Utilities.getUuid(),
      emp.employee_code,
      emp.name,
      site.loc_id,
      site.name,
      now,
      type,
      businessDate(now),
      isNum(geo.lat) ? geo.lat : '',
      isNum(geo.lng) ? geo.lng : '',
      isNum(geo.accuracy) ? Math.round(Number(geo.accuracy)) : '',
      distance,
      geoStatus,
      address,
      req.client_uuid || '',
      req.client_time || '',
      String(req.ua || '').slice(0, 60),
      false
    ]);

    if (req.client_uuid) cache.put(uuidKey, '1', 21600); // 6時間

    return {
      ok: true,
      type: type,
      punched_at: now.toISOString(),
      loc_name: site.name,
      geo_status: geoStatus,
      distance: distance,
      address: address
    };

  } finally {
    lock.releaseLock();
  }
}


/* ════════════════════════════════════════════
   5. 認証とデータ取得
   ════════════════════════════════════════════ */

function authenticate(token) {
  if (!token) return null;

  const t = readTable(SHEETS.DEVICES);
  for (let i = 0; i < t.rows.length; i++) {
    const d = t.rows[i];
    if (String(d.token) === String(token) && truthy(d.is_active)) {
      // 最終利用日時を更新（休眠端末の把握に使う）
      t.sheet.getRange(i + 2, t.col.last_used_at + 1).setValue(new Date());
      return findEmployee(d.employee_code);
    }
  }
  return null;
}

function findEmployee(code) {
  const t = readTable(SHEETS.EMPLOYEES);
  const key = String(code).trim().toUpperCase();
  for (const e of t.rows) {
    if (String(e.employee_code).trim().toUpperCase() === key) return e;
  }
  return null;
}

function findLocation(locId) {
  if (!locId) return null;
  const t = readTable(SHEETS.LOCATIONS);
  const key = String(locId).trim();
  for (const s of t.rows) {
    if (String(s.loc_id).trim() === key && truthy(s.is_active)) return s;
  }
  return null;
}

/**
 * 本人の当日分の最終打刻を返す。
 * 末尾から遡るので、行数が増えても速度は落ちない。
 */
function getLastPunchToday(code) {
  const sh = sheet(SHEETS.EVENTS);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;

  const t = readHeader(SHEETS.EVENTS);
  const today = businessDate(new Date());
  const from = Math.max(2, lastRow - 500);
  const values = sh.getRange(from, 1, lastRow - from + 1, sh.getLastColumn()).getValues();

  for (let i = values.length - 1; i >= 0; i--) {
    const r = values[i];
    if (String(r[t.col.employee_code]) !== String(code)) continue;
    if (truthy(r[t.col.is_voided])) continue;
    if (String(r[t.col.business_date]) !== today) continue;
    return {
      punch_type: String(r[t.col.punch_type]),
      punched_at: new Date(r[t.col.punched_at])
    };
  }
  return null;
}

function existsUuid(uuid) {
  const sh = sheet(SHEETS.EVENTS);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return false;

  const t = readHeader(SHEETS.EVENTS);
  const from = Math.max(2, lastRow - 200);
  const col = sh.getRange(from, t.col.client_uuid + 1, lastRow - from + 1, 1).getValues();
  return col.some(r => String(r[0]) === String(uuid));
}


/* ════════════════════════════════════════════
   6. ユーティリティ
   ════════════════════════════════════════════ */

function sheet(name) {
  const sh = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(name);
  if (!sh) throw new Error('シートが見つかりません: ' + name + ' → setup() を実行してください');
  return sh;
}

/** ヘッダー行から列名→インデックスの対応を作る（列順を変えても壊れないようにする） */
function readHeader(name) {
  const sh = sheet(name);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const col = {};
  headers.forEach((h, i) => col[String(h).trim()] = i);
  return { sheet: sh, headers: headers, col: col };
}

function readTable(name) {
  const t = readHeader(name);
  const lastRow = t.sheet.getLastRow();
  const rows = [];
  if (lastRow >= 2) {
    const values = t.sheet.getRange(2, 1, lastRow - 1, t.headers.length).getValues();
    for (const v of values) {
      const o = {};
      t.headers.forEach((h, i) => o[String(h).trim()] = v[i]);
      rows.push(o);
    }
  }
  return { sheet: t.sheet, col: t.col, headers: t.headers, rows: rows };
}

/** チェックボックスの true と文字列の "TRUE" の両方を受け付ける */
function truthy(v) {
  if (v === true) return true;
  const s = String(v).trim().toUpperCase();
  return s === 'TRUE' || s === 'YES' || s === '1' || s === 'Y';
}

function isNum(v) {
  return v !== null && v !== undefined && v !== '' && !isNaN(Number(v));
}

/** 勤務日。境界時刻より前の打刻は前日扱いにする（夜勤・日跨ぎ対応） */
function businessDate(d) {
  const shifted = new Date(d.getTime() - CONFIG.DAY_BOUNDARY_HOUR * 3600 * 1000);
  return Utilities.formatDate(shifted, CONFIG.TZ, 'yyyy-MM-dd');
}

/** 2地点間の距離（メートル） */
function haversine(lat1, lng1, lat2, lng2) {
  const R = 6371000, rad = d => d * Math.PI / 180;
  const dLat = rad(lat2 - lat1), dLng = rad(lng2 - lng1);
  const a = Math.sin(dLat / 2) ** 2 +
            Math.cos(rad(lat1)) * Math.cos(rad(lat2)) * Math.sin(dLng / 2) ** 2;
  return Math.round(R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a)));
}

/** 逆ジオコーディング。失敗しても打刻は止めない */
function reverseGeocode(lat, lng) {
  const key = 'geo:' + Number(lat).toFixed(4) + ',' + Number(lng).toFixed(4);
  const cache = CacheService.getScriptCache();
  const hit = cache.get(key);
  if (hit !== null) return hit;

  try {
    const res = Maps.newGeocoder().setLanguage('ja').reverseGeocode(lat, lng);
    const addr = (res.results && res.results[0]) ? res.results[0].formatted_address : '';
    cache.put(key, addr, 21600);
    return addr;
  } catch (e) {
    return '';
  }
}

/** PIN は平文で保存しない。ソルト＋ペッパー付き SHA-256 */
function hashPin(pin, salt) {
  const pepper = PropertiesService.getScriptProperties().getProperty('PEPPER') || '';
  const raw = String(salt) + '|' + String(pin) + '|' + pepper;
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return bytes.map(b => ((b & 0xFF) + 0x100).toString(16).slice(1)).join('');
}

function logError(err) {
  console.error(err.stack || err.message || err);
  try {
    const sh = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName('errors');
    if (sh) sh.appendRow([new Date(), String(err.message || err), String(err.stack || '')]);
  } catch (e) { /* ログ失敗で本処理を止めない */ }
}


/* ════════════════════════════════════════════
   7. 初期セットアップ（エディタから手動実行）
   ════════════════════════════════════════════ */

/**
 * 最初に1回だけ実行する。
 * 必要なシートとヘッダー、ペッパーを作成する。既存シートは壊さない。
 */
function setup() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  ss.setSpreadsheetTimeZone(CONFIG.TZ);

  const defs = {
    [SHEETS.EMPLOYEES]: ['employee_code', 'name', 'email', 'pin_salt', 'pin_hash', 'is_active'],
    [SHEETS.DEVICES]:   ['token', 'employee_code', 'registered_at', 'last_used_at', 'label', 'is_active'],
    [SHEETS.LOCATIONS]: ['loc_id', 'name', 'type', 'lat', 'lng', 'radius_m', 'is_active'],
    [SHEETS.EVENTS]:    ['event_id', 'employee_code', 'employee_name', 'loc_id', 'loc_name',
                         'punched_at', 'punch_type', 'business_date', 'lat', 'lng', 'accuracy_m',
                         'distance_m', 'geo_status', 'address', 'client_uuid', 'client_time',
                         'user_agent', 'is_voided'],
    [SHEETS.CORRECTIONS]: ['correction_id', 'original_event_id', 'employee_code', 'field',
                           'old_value', 'new_value', 'reason', 'requested_by', 'approved_by', 'corrected_at'],
    'errors': ['at', 'message', 'stack']
  };

  Object.keys(defs).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(defs[name]);
      sh.getRange(1, 1, 1, defs[name].length).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  });

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('PEPPER')) {
    props.setProperty('PEPPER', Utilities.getUuid() + Utilities.getUuid());
  }

  // 見本の場所を1件ずつ入れておく（座標は実地で書き換える）
  const loc = ss.getSheetByName(SHEETS.LOCATIONS);
  if (loc.getLastRow() === 1) {
    loc.appendRow(['HQ01', '本社入口', 'FIXED', 1.3521, 103.8198, 100, true]);
    loc.appendRow(['MOBILE_A01', '携帯タグ A', 'PORTABLE', '', '', '', true]);
  }

  Logger.log('セットアップが完了しました。次に addEmployees() を実行してください。');
}

/**
 * 社員を一括登録し、初期PINを発行する。
 * 下の list を書き換えてから実行し、ログに出るPINを本人に配布する。
 * PIN はハッシュ化して保存されるため、この場で控えないと二度と表示できない。
 */
function addEmployees() {
  const list = [
    { code: 'E001', name: '山田 太郎', email: '' },
    { code: 'E002', name: '鈴木 花子', email: '' }
  ];

  const sh = sheet(SHEETS.EMPLOYEES);
  const out = [];

  list.forEach(p => {
    if (findEmployee(p.code)) {
      out.push(p.code + ' : 既に登録済み（スキップ）');
      return;
    }
    const pin  = String(Math.floor(100000 + Math.random() * 900000)); // 6桁
    const salt = Utilities.getUuid();
    sh.appendRow([p.code.toUpperCase(), p.name, p.email || '', salt, hashPin(pin, salt), true]);
    out.push(p.code + ' / ' + p.name + ' → PIN: ' + pin);
  });

  Logger.log('─── 発行結果（この画面を閉じると二度と見られません）───\n' + out.join('\n'));
}

/** PIN を再発行する。忘れた人が出たら実行する。 */
function resetPin() {
  const code = 'E001';   // ← 対象の社員コードに書き換える

  const t = readTable(SHEETS.EMPLOYEES);
  const idx = t.rows.findIndex(e =>
    String(e.employee_code).trim().toUpperCase() === code.trim().toUpperCase());
  if (idx < 0) return Logger.log('見つかりません: ' + code);

  const pin  = String(Math.floor(100000 + Math.random() * 900000));
  const salt = Utilities.getUuid();
  t.sheet.getRange(idx + 2, t.col.pin_salt + 1).setValue(salt);
  t.sheet.getRange(idx + 2, t.col.pin_hash + 1).setValue(hashPin(pin, salt));

  Logger.log(code + ' の新しい PIN: ' + pin);
}

/** 端末を失効させる。紛失・機種変更時に実行する。 */
function deactivateDevices() {
  const code = 'E001';   // ← 対象の社員コードに書き換える

  const t = readTable(SHEETS.DEVICES);
  let n = 0;
  t.rows.forEach((d, i) => {
    if (String(d.employee_code).trim().toUpperCase() === code.trim().toUpperCase()
        && truthy(d.is_active)) {
      t.sheet.getRange(i + 2, t.col.is_active + 1).setValue(false);
      n++;
    }
  });
  Logger.log(code + ' の端末を ' + n + ' 件失効させました。本人に再登録を案内してください。');
}

/** 打刻を取り消す（物理削除はしない）。event_id を指定して実行する。 */
function voidEvent() {
  const eventId = '';          // ← event_id を貼る
  const reason  = '打刻ミス';   // ← 理由

  if (!eventId) return Logger.log('event_id を指定してください。');

  const t = readTable(SHEETS.EVENTS);
  const idx = t.rows.findIndex(r => String(r.event_id) === eventId);
  if (idx < 0) return Logger.log('見つかりません: ' + eventId);

  t.sheet.getRange(idx + 2, t.col.is_voided + 1).setValue(true);
  sheet(SHEETS.CORRECTIONS).appendRow([
    Utilities.getUuid(), eventId, t.rows[idx].employee_code, 'is_voided',
    'FALSE', 'TRUE', reason, Session.getActiveUser().getEmail(), '', new Date()
  ]);
  Logger.log('取り消しました: ' + eventId);
}
