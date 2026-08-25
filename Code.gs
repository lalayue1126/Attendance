/**
 * ══════════════════════════════════════════════════════════════
 *  NFC / QR Time Clock System — Google Apps Script backend
 *
 *  Paired frontend: index.html (GitHub Pages)
 *
 *  【First-time setup】
 *   1. Put the spreadsheet ID into CONFIG.SHEET_ID
 *   2. Run setup() once from the editor
 *   3. Deploy as a web app
 *        Execute as: Me
 *        Who has access: Anyone
 *   4. Paste the issued /exec URL into CONFIG.API_URL in index.html
 *   5. Enter each date's scheduled hours (start time + hours) into the
 *      schedules sheet by hand. employee_code and practice_loc_id are both
 *      optional — leave employee_code blank for "everyone" and
 *      practice_loc_id blank for "any location". This lets different staff
 *      working at the same location have different hours, and the same
 *      employee have different hours depending on which location they
 *      check in at. The most specific row (employee + location) wins.
 *   6. Register practice locations in the locations sheet (type is
 *      either FIXED or PORTABLE)
 *      → The "Practice Location" buttons on the punch screen are built
 *        from this automatically. FIXED locations are GPS-checked;
 *        PORTABLE locations are not (only the address is recorded).
 *   7. Run installWeeklyTrigger() once
 *      → Every Monday at 3am, last week's (Mon–Sun) worked hours per
 *        employee are automatically summarized into weekly_summary
 *   8. Set a password in setReportPassword() and run it once
 *      → This password gates report.html, the admin-only page that builds
 *        the attendance report (by location / employee / date range) and
 *        writes it into the "report" sheet.
 *
 *  【Only 2 links are distributed to staff】
 *   Fixed locations : index.html?l=FIXED     (one shared link for all fixed sites)
 *   Portable        : index.html?l=PORTABLE
 *   l is not a specific location ID — it's one of these two mode names.
 *   Either link then lets staff pick the actual place from the
 *   "Practice Location" buttons, sourced from locations rows matching
 *   that type.
 *
 *  【Admin-only report page】
 *   report.html (also on GitHub Pages) — password-protected, builds a
 *   pivot report (employees × dates) and writes it to the "report" sheet.
 * ══════════════════════════════════════════════════════════════
 */

const CONFIG = {
  SHEET_ID: '1tCaUQjiepw5tsJOqvk8uuFCDQsXrW3A50Vsmn4JrVGQ',

  TZ: 'Asia/Singapore',

  // Business-day boundary. With 4, punches between 00:00–03:59 count as the previous day.
  DAY_BOUNDARY_HOUR: 4,

  // Ignore a repeat of the same punch type within this many seconds (mis-taps / double reads).
  DEDUP_WINDOW_SEC: 60,

  // Default radius for FIXED locations (used when locations.radius_m is blank).
  DEFAULT_RADIUS_M: 100,

  // Reverse geocoding (PORTABLE only). Set to false if not needed.
  REVERSE_GEOCODE: true,

  // Fallback schedule used only when the schedules sheet has no row for that day
  // (a safety net for the normal workflow of entering hours per date by hand).
  DEFAULT_SCHEDULED_START: '09:00',
  DEFAULT_SCHEDULED_HOURS: 8
};

const SHEETS = {
  EMPLOYEES: 'employees',
  DEVICES: 'devices',
  LOCATIONS: 'locations',
  EVENTS: 'punch_events',
  CORRECTIONS: 'corrections',
  SCHEDULES: 'schedules',
  WEEKLY_SUMMARY: 'weekly_summary'
};

const TYPE_LABEL = {
  IN: 'Check In', OUT: 'Leave Early'
};

const MODE_LABEL = {
  FIXED: 'Fixed', PORTABLE: 'Portable'
};


/* ════════════════════════════════════════════
   1. Entry points
   ════════════════════════════════════════════ */

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return json({ ok: false, error: 'BAD_REQUEST', message: 'The request was empty.' });
    }

    const req = JSON.parse(e.postData.contents);

    switch (req.action) {
      case 'register':    return json(handleRegister(req));
      case 'state':       return json(handleState(req));
      case 'punch':       return json(handlePunch(req));
      case 'report_meta': return json(handleReportMeta(req));
      case 'report':      return json(handleReport(req));
      default:
        return json({ ok: false, error: 'BAD_ACTION', message: 'Unknown action.' });
    }

  } catch (err) {
    // Never swallow the exception silently — logging is essential for debugging.
    logError(err);
    return json({
      ok: false, error: 'SERVER',
      message: 'A server error occurred. Please contact your administrator.'
    });
  }
}

/**
 * doGet isn't used for punching, but is kept for deployment verification.
 * Opening /exec in a browser and seeing {"ok":true,...} confirms the deployment is public.
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
   2. Device registration
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

  appendRowByHeader(SHEETS.DEVICES, {
    token: token,
    employee_code: emp.employee_code,
    registered_at: now,
    last_used_at: now,
    label: String(req.label || '').slice(0, 60),
    is_active: true
  });

  return { ok: true, token: token, name: emp.name };
}


/* ════════════════════════════════════════════
   3. Initial state for the punch screen
   ════════════════════════════════════════════ */

function handleState(req) {
  const emp = authenticate(req.token);
  if (!emp) return { ok: false, error: 'AUTH' };

  const mode = validateMode(req.l);
  if (!mode) return { ok: false, error: 'BAD_MODE' };

  const last = getLastPunchToday(emp.employee_code);

  return {
    ok: true,
    name: emp.name,
    loc_mode: mode,
    loc_mode_label: MODE_LABEL[mode],
    suggest: suggestNext(last),
    practice_locations: listLocationsByType(mode),
    last: last ? {
      type: last.punch_type,
      type_label: TYPE_LABEL[last.punch_type] || last.punch_type,
      at_label: Utilities.formatDate(last.punched_at, CONFIG.TZ, 'HH:mm')
    } : null
  };
}

/**
 * Guesses which punch type should come next.
 * This is only a suggestion, never auto-confirmed —
 * auto-confirming would flip every subsequent punch after a single missed one.
 */
function suggestNext(last) {
  if (!last) return 'IN';
  return last.punch_type === 'OUT' ? 'IN' : 'OUT';
}


/* ════════════════════════════════════════════
   4. Recording a punch
   ════════════════════════════════════════════ */

function handlePunch(req) {
  const lock = LockService.getScriptLock();

  // Prevents concurrent punches from corrupting a row. Give up and ask the client to retry if the lock can't be acquired.
  try {
    lock.waitLock(20000);
  } catch (e) {
    return { ok: false, error: 'BUSY', message: 'The system is busy. Please try again.' };
  }

  try {
    const emp = authenticate(req.token);
    if (!emp) return { ok: false, error: 'AUTH' };

    const mode = validateMode(req.l);
    if (!mode) return { ok: false, error: 'BAD_MODE' };

    const type = String(req.type || '').toUpperCase();
    if (!TYPE_LABEL[type]) {
      return { ok: false, error: 'BAD_TYPE', message: 'Invalid punch type.' };
    }

    const practiceLoc = findLocationByType(req.practice_loc_id, mode);
    if (!practiceLoc) {
      return { ok: false, error: 'BAD_PRACTICE_LOC', message: 'Please select a practice location.' };
    }

    // ── Idempotency check: prevents an offline retry from being recorded twice
    const uuidKey = 'uuid:' + req.client_uuid;
    const cache = CacheService.getScriptCache();
    if (req.client_uuid && (cache.get(uuidKey) || existsUuid(req.client_uuid))) {
      return { ok: true, dedup: true, punched_at: new Date().toISOString(),
               loc_name: MODE_LABEL[mode], geo_status: 'OK' };
    }

    // ── Mis-tap guard: ignore an immediate repeat of the same punch type
    const last = getLastPunchToday(emp.employee_code);
    if (last && last.punch_type === type &&
        (new Date() - last.punched_at) < CONFIG.DEDUP_WINDOW_SEC * 1000) {
      return { ok: false, error: 'TOO_SOON' };
    }

    // ── Location check: FIXED is verified against the coordinates of the
    // selected practice location; PORTABLE isn't verified, only the address is recorded.
    const geo = req.geo || { status: 'TIMEOUT' };
    let geoStatus = String(geo.status || 'TIMEOUT');
    let distance = '';
    let address  = '';

    if (geoStatus === 'OK' && isNum(geo.lat) && isNum(geo.lng)) {

      if (mode === 'FIXED' && isNum(practiceLoc.lat) && isNum(practiceLoc.lng)) {
        const radius = isNum(practiceLoc.radius_m) ? Number(practiceLoc.radius_m) : CONFIG.DEFAULT_RADIUS_M;
        distance = haversine(geo.lat, geo.lng, Number(practiceLoc.lat), Number(practiceLoc.lng));

        // Add GPS error margin to the allowed radius.
        // Without this, legitimate indoor Wi-Fi-based positioning gets rejected constantly.
        const tolerance = radius + Math.min(Number(geo.accuracy) || 0, 300);
        geoStatus = distance <= tolerance ? 'OK' : 'OUT_OF_RANGE';

      } else {
        // PORTABLE, or a practice location without coordinates set: skip verification, just log the address.
        address = CONFIG.REVERSE_GEOCODE ? reverseGeocode(geo.lat, geo.lng) : '';
      }
    }

    // ── Compute worked hours.
    // If there's a Check In with no Leave Early, the employee is treated as having
    // worked the full scheduled hours. Arriving late or leaving early both reduce
    // the total. Multiple Check In / Leave Early pairs in one day (split shifts)
    // are supported.
    const now = new Date();
    const bDate = businessDate(now);
    const todaySoFar = getTodayEvents(emp.employee_code, bDate);
    const schedule = findSchedule(bDate, emp.employee_code, practiceLoc.loc_id);
    const worked = computeDayWorked(
      schedule, bDate, todaySoFar.concat([{ type: type, punched_at: now }]));

    // ── Append the row. The server clock is authoritative for the timestamp.
    // Writing by header name means the values stay correct even if the sheet's
    // physical column order has drifted from the definition order.
    appendRowByHeader(SHEETS.EVENTS, {
      event_id: Utilities.getUuid(),
      employee_code: emp.employee_code,
      employee_name: emp.name,
      loc_id: mode,
      loc_name: MODE_LABEL[mode],
      punched_at: now,
      punch_type: type,
      practice_loc_id: practiceLoc.loc_id,
      practice_loc_name: practiceLoc.name,
      business_date: bDate,
      lat: isNum(geo.lat) ? geo.lat : '',
      lng: isNum(geo.lng) ? geo.lng : '',
      accuracy_m: isNum(geo.accuracy) ? Math.round(Number(geo.accuracy)) : '',
      distance_m: distance,
      geo_status: geoStatus,
      address: address,
      client_uuid: req.client_uuid || '',
      client_time: req.client_time || '',
      user_agent: String(req.ua || '').slice(0, 60),
      is_voided: false,
      worked_hours: worked.hours,
      worked_minutes: worked.minutes
    });

    if (req.client_uuid) cache.put(uuidKey, '1', 21600); // 6 hours

    return {
      ok: true,
      type: type,
      punched_at: now.toISOString(),
      loc_name: MODE_LABEL[mode],
      geo_status: geoStatus,
      distance: distance,
      address: address,
      practice_loc_label: practiceLoc.name,
      day_worked_hours: worked.hours,
      day_worked_minutes: worked.minutes,
      scheduled_hours: schedule ? Number(schedule.scheduled_hours) : ''
    };

  } finally {
    lock.releaseLock();
  }
}


/* ════════════════════════════════════════════
   5. Authentication and data lookups
   ════════════════════════════════════════════ */

function authenticate(token) {
  if (!token) return null;

  const t = readTable(SHEETS.DEVICES);
  for (let i = 0; i < t.rows.length; i++) {
    const d = t.rows[i];
    if (String(d.token) === String(token) && truthy(d.is_active)) {
      // Update last-used timestamp (used to spot dormant devices).
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

/**
 * The l parameter is not a specific location ID — it represents the link's
 * mode (FIXED / PORTABLE). Since only these two shared links are distributed,
 * we validate the mode string directly instead of looking up an individual
 * row in the locations sheet.
 */
function validateMode(l) {
  const m = String(l || '').trim().toUpperCase();
  return (m === 'FIXED' || m === 'PORTABLE') ? m : null;
}

/** For the practice-location buttons. Only locations rows whose type matches the link's mode are considered. */
function findLocationByType(locId, type) {
  if (!locId) return null;
  const t = readTable(SHEETS.LOCATIONS);
  const key = String(locId).trim();
  for (const s of t.rows) {
    if (String(s.loc_id).trim() === key && truthy(s.is_active) &&
        String(s.type).trim().toUpperCase() === type) return s;
  }
  return null;
}

function listLocationsByType(type) {
  const t = readTable(SHEETS.LOCATIONS);
  return t.rows
    .filter(s => truthy(s.is_active) && String(s.type).trim().toUpperCase() === type)
    .map(s => ({ id: String(s.loc_id), label: String(s.name) }));
}

/**
 * Treats a row as "not filled in" if scheduled_start / scheduled_hours is blank or invalid.
 * Using such a row as-is would silently leave that day's worked hours blank.
 */
function isValidSchedule(r) {
  const startStr = normalizeTimeStr(r.scheduled_start);
  const hours = Number(r.scheduled_hours);
  return !!startStr && r.scheduled_hours !== '' && r.scheduled_hours !== null &&
         r.scheduled_hours !== undefined && isNum(hours) && hours > 0;
}

/**
 * Finds the scheduled hours for a given day, employee, and practice location.
 * employee_code and practice_loc_id are both optional in the schedules sheet —
 * a blank employee_code means "everyone", and a blank practice_loc_id means
 * "any location". Different staff at the same location on the same day can
 * therefore have different scheduled hours (per-employee rows), and the same
 * employee can have different hours depending on which location they're
 * checking in at (per-location rows).
 *
 * The most specific match wins, in this order:
 *   1. this employee + this location
 *   2. this employee + any location
 *   3. any employee   + this location
 *   4. any employee   + any location (fully blanket row)
 * Rows with a blank/invalid start time or hours are skipped in favor of the
 * next candidate. If nothing matches, falls back to the CONFIG default
 * (a safety net).
 */
function findSchedule(businessDateStr, employeeCode, practiceLocId) {
  const t = readTable(SHEETS.SCHEDULES);
  const code = String(employeeCode || '').trim().toUpperCase();
  const loc = String(practiceLocId || '').trim().toUpperCase();

  let best = null;
  let bestScore = -1;

  for (const r of t.rows) {
    if (normalizeDateStr(r.business_date) !== businessDateStr) continue;
    if (!isValidSchedule(r)) continue;

    const rowCode = String(r.employee_code || '').trim().toUpperCase();
    if (rowCode && rowCode !== code) continue; // row is for a different specific employee

    const rowLoc = String(r.practice_loc_id || '').trim().toUpperCase();
    if (rowLoc && rowLoc !== loc) continue; // row is for a different specific location

    const score = (rowCode ? 2 : 0) + (rowLoc ? 1 : 0);
    if (score >= bestScore) { // later rows win ties, so a correction row added below an old one takes over
      bestScore = score;
      best = r;
    }
  }

  if (best) return best;
  return {
    business_date: businessDateStr,
    employee_code: '',
    practice_loc_id: '',
    scheduled_start: CONFIG.DEFAULT_SCHEDULED_START,
    scheduled_hours: CONFIG.DEFAULT_SCHEDULED_HOURS
  };
}

/**
 * Computes worked hours for the day.
 *   - Each Check In → Leave Early pair adds up the time in that window.
 *   - A Check In with no matching Leave Early is assumed to run until the
 *     scheduled end time.
 *   - Every window is clipped to [scheduled start, scheduled end]
 *     (arriving late or leaving early both shrink it below the scheduled hours).
 *   - The day's total never exceeds the scheduled hours.
 * Returns both hours (2 decimal places) and minutes (integer). Minutes is the
 * raw value, useful for tracking early leavers precisely without being
 * rounded away by the hours figure.
 */
function computeDayWorked(schedule, businessDateStr, events) {
  if (!schedule) return { hours: '', minutes: '' };

  const startStr = normalizeTimeStr(schedule.scheduled_start);
  const hours = Number(schedule.scheduled_hours);
  if (!startStr || !isNum(hours)) return { hours: '', minutes: '' };

  const start = Utilities.parseDate(businessDateStr + ' ' + startStr, CONFIG.TZ, 'yyyy-MM-dd HH:mm');
  const end = new Date(start.getTime() + hours * 3600000);

  const sorted = events.slice().sort((a, b) => a.punched_at - b.punched_at);
  let totalMs = 0;
  let pendingIn = null;

  sorted.forEach(ev => {
    if (ev.type === 'IN') {
      if (pendingIn === null) pendingIn = ev.punched_at;
    } else if (ev.type === 'OUT' && pendingIn !== null) {
      const s = new Date(Math.max(pendingIn.getTime(), start.getTime()));
      const e = new Date(Math.min(ev.punched_at.getTime(), end.getTime()));
      totalMs += Math.max(0, e.getTime() - s.getTime());
      pendingIn = null;
    }
  });

  if (pendingIn !== null) {
    const s = new Date(Math.max(pendingIn.getTime(), start.getTime()));
    totalMs += Math.max(0, end.getTime() - s.getTime());
  }

  const cappedMs = Math.min(hours * 3600000, totalMs);
  return {
    hours: Math.round(cappedMs / 3600000 * 100) / 100,
    minutes: Math.round(cappedMs / 60000)
  };
}

/**
 * Returns all of this employee's punches for today, oldest first.
 * Scans backward from the end of the sheet, so performance doesn't
 * degrade as the row count grows.
 */
function getTodayEvents(code, businessDateStr) {
  const sh = sheet(SHEETS.EVENTS);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const t = readHeader(SHEETS.EVENTS);
  const from = Math.max(2, lastRow - 500);
  const values = sh.getRange(from, 1, lastRow - from + 1, sh.getLastColumn()).getValues();

  const out = [];
  for (const r of values) {
    if (String(r[t.col.employee_code]) !== String(code)) continue;
    if (truthy(r[t.col.is_voided])) continue;
    if (normalizeDateStr(r[t.col.business_date]) !== businessDateStr) continue;
    out.push({ type: String(r[t.col.punch_type]), punched_at: new Date(r[t.col.punched_at]) });
  }
  out.sort((a, b) => a.punched_at - b.punched_at);
  return out;
}

function getLastPunchToday(code) {
  const events = getTodayEvents(code, businessDate(new Date()));
  if (!events.length) return null;
  const last = events[events.length - 1];
  return { punch_type: last.type, punched_at: last.punched_at };
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
   6. Utilities
   ════════════════════════════════════════════ */

function sheet(name) {
  const sh = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(name);
  if (!sh) throw new Error('Sheet not found: ' + name + ' → run setup()');
  return sh;
}

/** Builds a column-name → index map from the header row (so column reordering doesn't break anything). */
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

/**
 * Appends a row keyed by header name, instead of a positional appendRow([...]) array,
 * so values still land in the right column even if the sheet's physical column
 * order has drifted from the definition order after past column additions.
 */
function appendRowByHeader(name, valuesObj) {
  const t = readHeader(name);
  const row = new Array(t.headers.length).fill('');
  t.headers.forEach((h, i) => {
    const key = String(h).trim();
    if (Object.prototype.hasOwnProperty.call(valuesObj, key)) row[i] = valuesObj[key];
  });
  t.sheet.appendRow(row);
}

/** Accepts both a real checkbox true and the string "TRUE". */
function truthy(v) {
  if (v === true) return true;
  const s = String(v).trim().toUpperCase();
  return s === 'TRUE' || s === 'YES' || s === '1' || s === 'Y';
}

function isNum(v) {
  return v !== null && v !== undefined && v !== '' && !isNaN(Number(v));
}

/** Business day. Punches before the boundary hour count as the previous day (for night shifts crossing midnight). */
function businessDate(d) {
  const shifted = new Date(d.getTime() - CONFIG.DAY_BOUNDARY_HOUR * 3600 * 1000);
  return Utilities.formatDate(shifted, CONFIG.TZ, 'yyyy-MM-dd');
}

/** Adds n days (can be negative) to a 'yyyy-MM-dd' string. Used to compute the weekly summary range. */
function addDays(dateStr, n) {
  const d = Utilities.parseDate(dateStr, CONFIG.TZ, 'yyyy-MM-dd');
  return Utilities.formatDate(new Date(d.getTime() + n * 86400000), CONFIG.TZ, 'yyyy-MM-dd');
}

/** Returns the Monday ('yyyy-MM-dd') of the week containing this date. */
function mondayOfWeek(dateStr) {
  const d = Utilities.parseDate(dateStr, CONFIG.TZ, 'yyyy-MM-dd');
  const day = d.getDay(); // 0=Sun,1=Mon,...6=Sat
  const diff = day === 0 ? -6 : 1 - day;
  return Utilities.formatDate(new Date(d.getTime() + diff * 86400000), CONFIG.TZ, 'yyyy-MM-dd');
}

/** Spreadsheet date cells can come back as a Date object, so normalize to a plain string. */
function normalizeDateStr(v) {
  if (v instanceof Date) return Utilities.formatDate(v, CONFIG.TZ, 'yyyy-MM-dd');
  return String(v || '').trim();
}

/** Time cells can also come back as a Date object (1899-12-30 + time), so normalize to HH:mm. */
function normalizeTimeStr(v) {
  if (v instanceof Date) return Utilities.formatDate(v, CONFIG.TZ, 'HH:mm');
  return String(v || '').trim();
}

/** Distance between two points, in meters. */
function haversine(lat1, lng1, lat2, lng2) {
  const R = 6371000, rad = d => d * Math.PI / 180;
  const dLat = rad(lat2 - lat1), dLng = rad(lng2 - lng1);
  const a = Math.sin(dLat / 2) ** 2 +
            Math.cos(rad(lat1)) * Math.cos(rad(lat2)) * Math.sin(dLng / 2) ** 2;
  return Math.round(R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a)));
}

/** Reverse geocoding. A failure here must never block the punch itself. */
function reverseGeocode(lat, lng) {
  const key = 'geo:' + Number(lat).toFixed(4) + ',' + Number(lng).toFixed(4);
  const cache = CacheService.getScriptCache();
  const hit = cache.get(key);
  if (hit !== null) return hit;

  try {
    const res = Maps.newGeocoder().setLanguage('en').reverseGeocode(lat, lng);
    const addr = (res.results && res.results[0]) ? res.results[0].formatted_address : '';
    cache.put(key, addr, 21600);
    return addr;
  } catch (e) {
    return '';
  }
}

/** PINs are never stored in plain text — salted + peppered SHA-256. */
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
  } catch (e) { /* A logging failure must not stop the main flow. */ }
}


/* ════════════════════════════════════════════
   7. Initial setup (run manually from the editor)
   ════════════════════════════════════════════ */

/**
 * Run this once, the first time.
 * Creates the required sheets, headers, and pepper. Never touches an existing sheet's data.
 * If an existing sheet is missing a column, it's appended at the end (existing data is preserved).
 */
function setup() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  ss.setSpreadsheetTimeZone(CONFIG.TZ);

  const defs = {
    [SHEETS.EMPLOYEES]: ['employee_code', 'name', 'email', 'pin_salt', 'pin_hash', 'is_active'],
    [SHEETS.DEVICES]:   ['token', 'employee_code', 'registered_at', 'last_used_at', 'label', 'is_active'],
    [SHEETS.LOCATIONS]: ['loc_id', 'name', 'type', 'lat', 'lng', 'radius_m', 'is_active'],
    [SHEETS.EVENTS]:    ['event_id', 'employee_code', 'employee_name', 'loc_id', 'loc_name',
                         'punched_at', 'punch_type', 'practice_loc_id', 'practice_loc_name',
                         'business_date', 'lat', 'lng', 'accuracy_m', 'distance_m', 'geo_status',
                         'address', 'client_uuid', 'client_time', 'user_agent', 'is_voided',
                         'worked_hours', 'worked_minutes'],
    [SHEETS.CORRECTIONS]: ['correction_id', 'original_event_id', 'employee_code', 'field',
                           'old_value', 'new_value', 'reason', 'requested_by', 'approved_by', 'corrected_at'],
    [SHEETS.SCHEDULES]: ['business_date', 'employee_code', 'practice_loc_id', 'scheduled_start', 'scheduled_hours', 'note'],
    [SHEETS.WEEKLY_SUMMARY]: ['week_start', 'week_end', 'employee_code', 'employee_name',
                              'worked_hours', 'worked_minutes', 'generated_at'],
    'errors': ['at', 'message', 'stack']
  };

  Object.keys(defs).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(defs[name]);
      sh.getRange(1, 1, 1, defs[name].length).setFontWeight('bold');
      sh.setFrozenRows(1);
    } else {
      // Append only the columns missing from an existing sheet (never reorders or deletes data).
      const existing = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
      const missing = defs[name].filter(h => existing.indexOf(h) === -1);
      if (missing.length) {
        sh.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
        sh.getRange(1, existing.length + 1, 1, missing.length).setFontWeight('bold');
      }
    }
  });

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('PEPPER')) {
    props.setProperty('PEPPER', Utilities.getUuid() + Utilities.getUuid());
  }

  // Seed one sample practice location of each type (edit the coordinates for the real site).
  // A FIXED row shows up on the l=FIXED link; a PORTABLE row shows up on the l=PORTABLE link.
  const loc = ss.getSheetByName(SHEETS.LOCATIONS);
  if (loc.getLastRow() === 1) {
    appendRowByHeader(SHEETS.LOCATIONS,
      { loc_id: 'HQ01', name: 'Main Entrance', type: 'FIXED', lat: 1.3521, lng: 103.8198, radius_m: 100, is_active: true });
    appendRowByHeader(SHEETS.LOCATIONS,
      { loc_id: 'MOBILE_A01', name: 'Off-site A', type: 'PORTABLE', lat: '', lng: '', radius_m: '', is_active: true });
  }

  Logger.log('Setup complete. Run addEmployees() next.' +
             ' Enter each date\'s scheduled hours into the schedules sheet, and practice locations (type=FIXED/PORTABLE) into the locations sheet.' +
             ' To automate the weekly summary, run installWeeklyTrigger() once.');
}

/**
 * Bulk-registers employees and issues initial PINs.
 * Edit the list below before running; distribute the PINs shown in the log to each person.
 * PINs are stored hashed, so they can never be shown again after this — record them now.
 */
function addEmployees() {
  const list = [
    { code: 'E001', name: 'John Smith', email: '' },
    { code: 'E002', name: 'Jane Doe', email: '' }
  ];

  const out = [];

  list.forEach(p => {
    if (findEmployee(p.code)) {
      out.push(p.code + ' : already registered (skipped)');
      return;
    }
    const pin  = String(Math.floor(100000 + Math.random() * 900000)); // 6 digits
    const salt = Utilities.getUuid();
    appendRowByHeader(SHEETS.EMPLOYEES, {
      employee_code: p.code.toUpperCase(), name: p.name, email: p.email || '',
      pin_salt: salt, pin_hash: hashPin(pin, salt), is_active: true
    });
    out.push(p.code + ' / ' + p.name + ' → PIN: ' + pin);
  });

  Logger.log('─── Issued PINs (visible only now — this screen won\'t show them again) ───\n' + out.join('\n'));
}

/** Reissues a PIN. Run this when someone forgets theirs. */
function resetPin() {
  const code = 'E001';   // ← change to the target employee code

  const t = readTable(SHEETS.EMPLOYEES);
  const idx = t.rows.findIndex(e =>
    String(e.employee_code).trim().toUpperCase() === code.trim().toUpperCase());
  if (idx < 0) return Logger.log('Not found: ' + code);

  const pin  = String(Math.floor(100000 + Math.random() * 900000));
  const salt = Utilities.getUuid();
  t.sheet.getRange(idx + 2, t.col.pin_salt + 1).setValue(salt);
  t.sheet.getRange(idx + 2, t.col.pin_hash + 1).setValue(hashPin(pin, salt));

  Logger.log(code + ' new PIN: ' + pin);
}

/** Deactivates a device. Run this when a device is lost or replaced. */
function deactivateDevices() {
  const code = 'E001';   // ← change to the target employee code

  const t = readTable(SHEETS.DEVICES);
  let n = 0;
  t.rows.forEach((d, i) => {
    if (String(d.employee_code).trim().toUpperCase() === code.trim().toUpperCase()
        && truthy(d.is_active)) {
      t.sheet.getRange(i + 2, t.col.is_active + 1).setValue(false);
      n++;
    }
  });
  Logger.log('Deactivated ' + n + ' device(s) for ' + code + '. Ask them to register again.');
}

/** Voids a punch (never physically deleted). Set event_id and run this. */
function voidEvent() {
  const eventId = '';          // ← paste the event_id here
  const reason  = 'Punch mistake';   // ← reason

  if (!eventId) return Logger.log('Please specify an event_id.');

  const t = readTable(SHEETS.EVENTS);
  const idx = t.rows.findIndex(r => String(r.event_id) === eventId);
  if (idx < 0) return Logger.log('Not found: ' + eventId);

  t.sheet.getRange(idx + 2, t.col.is_voided + 1).setValue(true);
  appendRowByHeader(SHEETS.CORRECTIONS, {
    correction_id: Utilities.getUuid(),
    original_event_id: eventId,
    employee_code: t.rows[idx].employee_code,
    field: 'is_voided',
    old_value: 'FALSE',
    new_value: 'TRUE',
    reason: reason,
    requested_by: Session.getActiveUser().getEmail(),
    approved_by: '',
    corrected_at: new Date()
  });
  Logger.log('Voided: ' + eventId);
}


/* ════════════════════════════════════════════
   8. Weekly summary (Mon–Sun)
   ════════════════════════════════════════════ */

/**
 * Sums last week's (Mon–Sun) worked hours per employee and writes one row per
 * employee to weekly_summary. Intended to be called by the time-driven
 * trigger set up in installWeeklyTrigger().
 *
 * How the aggregation works:
 *   Each punch_events row's worked_hours / worked_minutes is "the running
 *   best estimate of that day's total, overwritten in full every time a
 *   punch is made that day." So, for a given day and employee, the value on
 *   the last row (i.e. the one appended last) is that day's final value.
 *   Those final daily values are collected across Monday–Sunday and summed.
 */
function buildWeeklySummary() {
  const today = businessDate(new Date());
  const thisMonday = mondayOfWeek(today);
  const weekStart = addDays(thisMonday, -7); // last week's Monday
  const weekEnd = addDays(thisMonday, -1);   // last week's Sunday

  const t = readHeader(SHEETS.EVENTS);
  const lastRow = t.sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('punch_events has no data.');
    return;
  }

  const values = t.sheet.getRange(2, 1, lastRow - 1, t.headers.length).getValues();

  // Overwrite by (date, employee) key with each matching row's value in sheet order
  // (appendRow always appends at the end = chronological order, so whatever value
  // survives is that day's final one).
  const perDay = {};
  values.forEach(r => {
    const bDate = normalizeDateStr(r[t.col.business_date]);
    if (bDate < weekStart || bDate > weekEnd) return;
    if (truthy(r[t.col.is_voided])) return;

    const code = String(r[t.col.employee_code]).trim();
    if (!code) return;

    perDay[bDate + '|' + code] = {
      name: String(r[t.col.employee_name]),
      hours: isNum(r[t.col.worked_hours]) ? Number(r[t.col.worked_hours]) : 0,
      minutes: isNum(r[t.col.worked_minutes]) ? Number(r[t.col.worked_minutes]) : 0
    };
  });

  const totals = {};
  Object.keys(perDay).forEach(key => {
    const code = key.split('|')[1];
    const d = perDay[key];
    if (!totals[code]) totals[code] = { name: d.name, hours: 0, minutes: 0 };
    totals[code].hours += d.hours;
    totals[code].minutes += d.minutes;
  });

  // Remove any rows already written for this week before appending fresh ones,
  // so a duplicate trigger firing or a manual re-run never doubles up the totals.
  removeWeeklySummaryRows(weekStart);

  const now = new Date();
  const codes = Object.keys(totals);
  codes.forEach(code => {
    appendRowByHeader(SHEETS.WEEKLY_SUMMARY, {
      week_start: weekStart,
      week_end: weekEnd,
      employee_code: code,
      employee_name: totals[code].name,
      worked_hours: Math.round(totals[code].hours * 100) / 100,
      worked_minutes: totals[code].minutes,
      generated_at: now
    });
  });

  Logger.log('Wrote the ' + weekStart + ' to ' + weekEnd + ' weekly summary for ' + codes.length + ' employee(s) to weekly_summary.');
}

/** Deletes any existing weekly_summary rows for the given week_start, so buildWeeklySummary() can be safely re-run. */
function removeWeeklySummaryRows(weekStart) {
  const t = readHeader(SHEETS.WEEKLY_SUMMARY);
  const lastRow = t.sheet.getLastRow();
  if (lastRow < 2) return;

  const values = t.sheet.getRange(2, 1, lastRow - 1, t.headers.length).getValues();
  // Delete from the bottom up so row indices stay valid as rows are removed.
  for (let i = values.length - 1; i >= 0; i--) {
    if (normalizeDateStr(values[i][t.col.week_start]) === weekStart) {
      t.sheet.deleteRow(i + 2);
    }
  }
}

/**
 * Sets up a trigger that automatically runs buildWeeklySummary() every
 * Monday at 3am (CONFIG.TZ). Run this once, initially. Any existing trigger
 * with the same handler is deleted first, so running this again never
 * creates a duplicate.
 */
function installWeeklyTrigger() {
  ScriptApp.getProjectTriggers().forEach(tr => {
    if (tr.getHandlerFunction() === 'buildWeeklySummary') ScriptApp.deleteTrigger(tr);
  });

  ScriptApp.newTrigger('buildWeeklySummary')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(3)
    .inTimezone(CONFIG.TZ)
    .create();

  Logger.log('buildWeeklySummary() will now run automatically every Monday at 3am (' + CONFIG.TZ + ').');
}


/* ════════════════════════════════════════════
   9. Attendance report (admin, password-protected)
   ════════════════════════════════════════════ */

/**
 * Sets (or changes) the password required to generate the attendance report
 * from report.html. Anyone with this password can see every employee's
 * hours, so treat it like any other shared admin credential.
 * Edit the password below, then run this function once.
 */
function setReportPassword() {
  const password = 'CHANGE_ME';   // ← set the report password here, then run this once

  PropertiesService.getScriptProperties().setProperty('REPORT_PASSWORD', password);
  Logger.log('Report password has been set.');
}

function checkReportPassword(pw) {
  const expected = PropertiesService.getScriptProperties().getProperty('REPORT_PASSWORD') || '';
  return !!expected && String(pw || '') === expected;
}

/**
 * Returns the filter options (practice locations, employees) for the report
 * screen's dropdowns. Requires the report password — this is the only data
 * report.html can see before authenticating.
 */
function handleReportMeta(req) {
  if (!checkReportPassword(req.password)) {
    return { ok: false, error: 'BAD_PASSWORD', message: 'Incorrect password.' };
  }

  const locations = readTable(SHEETS.LOCATIONS).rows
    .filter(r => truthy(r.is_active))
    .map(r => ({ id: String(r.loc_id), label: String(r.name) + ' (' + String(r.type) + ')' }));

  const employees = readTable(SHEETS.EMPLOYEES).rows
    .filter(r => truthy(r.is_active))
    .map(r => ({ code: String(r.employee_code), name: String(r.name) }));

  return { ok: true, locations: locations, employees: employees };
}

/**
 * Builds the attendance report for a date range (optionally filtered by
 * practice location and/or employee) and writes it to the "report" sheet,
 * overwriting whatever was there before.
 *
 * Layout: one row per employee, one column per date in the range, plus
 * Total Hours / Scheduled Hours / Attendance Rate columns.
 *   - Total Hours   : sum of that employee's worked_hours (from punch_events)
 *                     across the dates in range that match the filters.
 *   - Scheduled Hours: sum of findSchedule()'s scheduled_hours for every date
 *                     in the range (via the same employee/location priority
 *                     rules used for the live worked-hours calculation).
 *                     Note: every date in the range is treated as an expected
 *                     work day — this system has no explicit "day off" flag,
 *                     so a range that includes non-working days will inflate
 *                     the scheduled total and understate the attendance rate.
 *   - Attendance Rate: Total Hours / Scheduled Hours, as a percentage.
 */
function handleReport(req) {
  if (!checkReportPassword(req.password)) {
    return { ok: false, error: 'BAD_PASSWORD', message: 'Incorrect password.' };
  }

  const startDate = String(req.start_date || '').trim();
  const endDate = String(req.end_date || '').trim();
  if (!startDate || !endDate || startDate > endDate) {
    return { ok: false, error: 'BAD_RANGE', message: 'Please select a valid date range.' };
  }

  const dates = [];
  for (let d = startDate; d <= endDate; d = addDays(d, 1)) {
    dates.push(d);
    if (dates.length > 366) {
      return { ok: false, error: 'RANGE_TOO_LARGE', message: 'Please select a range of 366 days or fewer.' };
    }
  }

  const locFilter = String(req.practice_loc_id || '').trim();       // '' = all locations
  const empFilter = String(req.employee_code || '').trim().toUpperCase(); // '' = all employees

  const t = readHeader(SHEETS.EVENTS);
  const lastRow = t.sheet.getLastRow();

  const perEmployeeDate = {}; // employee_code -> { business_date -> worked_hours }
  const employeeNames = {};   // employee_code -> name

  if (lastRow >= 2) {
    const values = t.sheet.getRange(2, 1, lastRow - 1, t.headers.length).getValues();
    values.forEach(r => {
      if (truthy(r[t.col.is_voided])) return;

      const bDate = normalizeDateStr(r[t.col.business_date]);
      if (bDate < startDate || bDate > endDate) return;

      const code = String(r[t.col.employee_code]).trim();
      if (!code) return;
      if (empFilter && code.toUpperCase() !== empFilter) return;
      if (locFilter && String(r[t.col.practice_loc_id]).trim() !== locFilter) return;

      employeeNames[code] = String(r[t.col.employee_name]);
      if (!perEmployeeDate[code]) perEmployeeDate[code] = {};
      // Rows are in chronological (append) order, so the last matching row
      // for a given day is that day's final worked-hours value — same rule
      // buildWeeklySummary() uses.
      perEmployeeDate[code][bDate] = isNum(r[t.col.worked_hours]) ? Number(r[t.col.worked_hours]) : 0;
    });
  }

  const employeeCodes = Object.keys(perEmployeeDate).sort();
  if (!employeeCodes.length) {
    return { ok: false, error: 'NO_DATA', message: 'No matching punches were found for this filter.' };
  }

  const headerRow = ['Employee'].concat(dates, ['Total Hours', 'Scheduled Hours', 'Attendance Rate']);
  const dataRows = employeeCodes.map(code => {
    let totalActual = 0;
    let totalScheduled = 0;
    const perDateCells = dates.map(d => {
      const hrs = perEmployeeDate[code][d] || 0;
      totalActual += hrs;
      const sched = findSchedule(d, code, locFilter);
      totalScheduled += (sched && isNum(sched.scheduled_hours)) ? Number(sched.scheduled_hours) : 0;
      return hrs;
    });
    const rate = totalScheduled > 0 ? Math.round((totalActual / totalScheduled) * 1000) / 10 + '%' : '';
    return [employeeNames[code] || code].concat(
      perDateCells,
      [Math.round(totalActual * 100) / 100, Math.round(totalScheduled * 100) / 100, rate]
    );
  });

  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let sh = ss.getSheetByName('report');
  if (!sh) sh = ss.insertSheet('report');
  sh.clear();

  const title = 'Attendance Report — ' + (locFilter || 'All locations') + ' — ' +
                (empFilter || 'All employees') + ' — ' + startDate + ' to ' + endDate +
                ' (generated ' + Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy-MM-dd HH:mm') + ')';
  sh.getRange(1, 1).setValue(title);
  sh.getRange(2, 1, 1, headerRow.length).setValues([headerRow]);
  sh.getRange(2, 1, 1, headerRow.length).setFontWeight('bold');
  if (dataRows.length) {
    sh.getRange(3, 1, dataRows.length, headerRow.length).setValues(dataRows);
  }
  sh.setFrozenRows(2);

  return {
    ok: true,
    sheet_url: ss.getUrl() + '#gid=' + sh.getSheetId(),
    employee_count: dataRows.length,
    date_count: dates.length
  };
}
