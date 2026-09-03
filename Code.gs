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
 *      schedules sheet by hand — or use schedule.html (see below) to add
 *      many dates at once. employee_code and practice_loc_id are both
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
 *      → This password gates report.html AND schedule.html, the two
 *        admin-only pages (see below).
 *
 *  【Only 2 links are distributed to staff】
 *   Fixed locations : index.html?l=FIXED     (one shared link for all fixed sites)
 *   Portable        : index.html?l=PORTABLE
 *   l is not a specific location ID — it's one of these two mode names.
 *   Either link then lets staff pick the actual place from the
 *   "Practice Location" buttons, sourced from locations rows matching
 *   that type.
 *
 *  【Admin-only pages (both on GitHub Pages, both password-protected)】
 *   report.html   — builds a pivot report (employees × dates) and writes
 *                    it into the "report" sheet.
 *   schedule.html — bulk-adds rows to one category's schedule sheet
 *                    (schedule_A/S/N/Y — see SCHEDULE_CATEGORY_LABELS): pick
 *                    a date range + which weekdays practice falls on, plus
 *                    the (usually fixed) start time / hours / location, and
 *                    it creates one row per matching date. Keeps that sheet
 *                    sorted by date automatically (sortScheduleSheetByName()).
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
  SCHEDULES_LEGACY: 'schedules', // pre-category-split sheet, kept only for migrateSchedulesToCategories()
  WEEKLY_SUMMARY: 'weekly_summary'
};

const TYPE_LABEL = {
  IN: 'Check In', OUT: 'Leave Early'
};

// Schedules are split one sheet per player category, since different
// categories can now practice at different times/places on the same date.
// The category is derived from the first letter of the employee_code (e.g.
// A003 → 'A'); a code whose first letter isn't one of these (e.g. E-prefix
// staff) has no schedule sheet of its own, so findSchedule() always falls
// back to the CONFIG default for them (matched:false — never a "scheduled
// practice" for attendance-rate purposes).
const SCHEDULE_CATEGORY_LABELS = { A: 'National A', S: 'Sparring', N: 'National Service', Y: 'NYDS' };
const SCHEDULE_CATEGORIES = Object.keys(SCHEDULE_CATEGORY_LABELS); // ['A','S','N','Y']

function scheduleSheetName(category) {
  return 'schedule_' + category;
}

/** Maps an employee code to its schedule category ('A'/'S'/'N'/'Y'), or '' if it doesn't have one. */
function categoryForEmployee(employeeCode) {
  const c = String(employeeCode || '').trim().toUpperCase().charAt(0);
  return SCHEDULE_CATEGORIES.indexOf(c) !== -1 ? c : '';
}

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
      case 'register':      return json(handleRegister(req));
      case 'state':         return json(handleState(req));
      case 'punch':         return json(handlePunch(req));
      case 'report_meta':   return json(handleReportMeta(req));
      case 'report':        return json(handleReport(req));
      case 'schedule_meta': return json(handleScheduleMeta(req));
      case 'add_schedule':  return json(handleAddSchedule(req));
      case 'sort_schedule': return json(handleSortSchedule(req));
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
    employee_code: emp.employee_code,
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
    const schedule = findSchedule(bDate, emp.employee_code, practiceLoc.loc_id, now);
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
 * Schedules live one sheet per category (schedule_A/S/N/Y — see
 * categoryForEmployee()); an employee whose code doesn't map to one of those
 * categories (e.g. E-prefix staff) has no schedule sheet at all, so this
 * always falls straight to the CONFIG default (matched:false) for them.
 *
 * Within that category's sheet, employee_code and practice_loc_id are both
 * optional — a blank employee_code means "everyone in this category", and a
 * blank practice_loc_id means "any location". Passing '' for practiceLocId
 * means the same thing from the caller's side — "don't restrict by
 * location" — so a row with a specific location still matches; this is what
 * handleReport() passes when its own location filter is left on "All
 * Locations". Different players in the same category can therefore have
 * different scheduled hours (per-employee rows), and the same employee can
 * have different hours depending on which location they're checking in at
 * (per-location rows). The same location can also have more than one
 * session on the same day (e.g. a morning and an evening session at AQC) —
 * referenceTime (the actual punch time, when known) disambiguates between
 * them; see below.
 *
 * The most specific match wins, in this order:
 *   1. this employee + this location
 *   2. this employee + any location
 *   3. any employee   + this location
 *   4. any employee   + any location (fully blanket row)
 * Rows with a blank/invalid start time or hours are skipped in favor of the
 * next candidate. If nothing matches, falls back to the CONFIG default
 * (a safety net).
 *
 * When multiple rows tie at the same specificity (e.g. two AQC sessions on
 * the same day), referenceTime picks whichever session's scheduled_start is
 * closest to it — otherwise a 05:39 check-in could get matched against an
 * 18:30 session just because it happens to sort last. Pass the actual punch
 * time whenever one is available. Only when referenceTime is omitted (e.g.
 * summing scheduled hours for a whole date range with no single punch to
 * anchor to) does the tie-break fall back to "the last matching row wins".
 */
function findSchedule(businessDateStr, employeeCode, practiceLocId, referenceTime) {
  const category = categoryForEmployee(employeeCode);
  if (!category) {
    return {
      business_date: businessDateStr,
      employee_code: '',
      practice_loc_id: '',
      scheduled_start: CONFIG.DEFAULT_SCHEDULED_START,
      scheduled_hours: CONFIG.DEFAULT_SCHEDULED_HOURS,
      matched: false
    };
  }

  const t = readTable(scheduleSheetName(category));
  const code = String(employeeCode || '').trim().toUpperCase();
  const loc = String(practiceLocId || '').trim().toUpperCase();

  let best = null;
  let bestScore = -1;
  let bestDiff = Infinity;

  for (const r of t.rows) {
    if (normalizeDateStr(r.business_date) !== businessDateStr) continue;
    if (!isValidSchedule(r)) continue;

    const rowCode = String(r.employee_code || '').trim().toUpperCase();
    if (rowCode && rowCode !== code) continue; // row is for a different specific employee

    const rowLoc = String(r.practice_loc_id || '').trim().toUpperCase();
    if (loc && rowLoc && rowLoc !== loc) continue; // row is for a different specific location

    const score = (rowCode ? 2 : 0) + (rowLoc ? 1 : 0);
    if (score < bestScore) continue; // a less specific row never overrides a better one

    let diff = 0;
    if (referenceTime) {
      const rowStart = Utilities.parseDate(
        businessDateStr + ' ' + normalizeTimeStr(r.scheduled_start), CONFIG.TZ, 'yyyy-MM-dd HH:mm');
      diff = Math.abs(referenceTime.getTime() - rowStart.getTime());
    }

    if (score > bestScore) {
      // Strictly more specific — always take it, resetting the time tiebreak.
      best = r; bestScore = score; bestDiff = diff;
    } else if (!referenceTime || diff <= bestDiff) {
      // Equally specific: with a reference time, the closer session wins;
      // without one, the last matching row wins (old behavior).
      best = r; bestDiff = diff;
    }
  }

  if (best) {
    best.matched = true;
    return best;
  }
  return {
    business_date: businessDateStr,
    employee_code: '',
    practice_loc_id: '',
    scheduled_start: CONFIG.DEFAULT_SCHEDULED_START,
    scheduled_hours: CONFIG.DEFAULT_SCHEDULED_HOURS,
    matched: false
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
 * Returns the distinct scheduled_start times (zero-padded 'HH:mm', sorted
 * ascending) across the given categories' schedule sheets for this date,
 * matching the location filter ('' = every location, same "blank row = any
 * location" rule as findSchedule()). An empty result means no practice was
 * scheduled that day for any of these categories — handleReport() uses that
 * to drop the date from the report entirely, and a non-empty result to
 * detect a day with more than one session so it can split that date into
 * separate per-session columns instead of merging every session's punches
 * into one.
 */
function distinctSessionStarts(businessDateStr, locFilter, categories) {
  const loc = String(locFilter || '').trim().toUpperCase();
  const starts = {};
  categories.forEach(cat => {
    const t = readTable(scheduleSheetName(cat));
    t.rows.forEach(r => {
      if (normalizeDateStr(r.business_date) !== businessDateStr) return;
      if (!isValidSchedule(r)) return;
      const rowLoc = String(r.practice_loc_id || '').trim().toUpperCase();
      if (loc && rowLoc && rowLoc !== loc) return;
      starts[padTimeStr(r.scheduled_start)] = true;
    });
  });
  return Object.keys(starts).sort();
}

/**
 * Splits one day's full event list into one sub-list per session, given that
 * date's sorted session start times. A punch belongs to session i if it's at
 * or after (start_i minus 30 minutes) and before (start_(i+1) minus 30
 * minutes) — so with sessions at 09:00 and 15:00, anything from 14:30 onward
 * counts as the second session, everything earlier as the first. Nothing
 * before the first session's threshold is dropped; it still lands in
 * session 0, since there's no earlier session to claim it.
 */
function splitBySession(dayEvents, starts, businessDateStr) {
  const boundaries = starts.map((s, i) => {
    if (i === 0) return -Infinity;
    const t = Utilities.parseDate(businessDateStr + ' ' + s, CONFIG.TZ, 'yyyy-MM-dd HH:mm');
    return t.getTime() - 30 * 60000;
  });
  const buckets = starts.map(() => []);
  dayEvents.forEach(ev => {
    let idx = 0;
    for (let i = 0; i < boundaries.length; i++) {
      if (ev.punched_at.getTime() >= boundaries[i]) idx = i;
    }
    buckets[idx].push(ev);
  });
  return buckets;
}

/**
 * Recomputes one employee's worked hours for one business date from that
 * day's raw punch_events rows, using whatever is CURRENTLY in the schedules
 * sheet. This is the single source of truth used by both buildWeeklySummary()
 * and handleReport() — neither one trusts the worked_hours/worked_minutes
 * columns already sitting in punch_events, so correcting a schedule after
 * the fact (a last-minute change entered once practice is over) is picked
 * up automatically the next time either is (re-)generated, with no separate
 * recompute step.
 *
 * dayEvents must contain every non-voided event for this employee on this
 * date, regardless of practice location — all of them are needed to pair
 * Check In / Leave Early correctly. The schedule is looked up against the
 * *last* event's practice location and timestamp, matching how handlePunch()
 * decides it live (each punch recomputes the whole day using its own location
 * and time, so whichever punch was most recent is what determines the day's
 * total — the timestamp also disambiguates between multiple same-day sessions
 * at the same location, e.g. a morning and an evening practice at AQC).
 */
function recomputeDayWorked(employeeCode, businessDateStr, dayEvents) {
  if (!dayEvents.length) return { hours: 0, minutes: 0 };
  const sorted = dayEvents.slice().sort((a, b) => a.punched_at - b.punched_at);
  const lastEvent = sorted[sorted.length - 1];
  const schedule = findSchedule(businessDateStr, employeeCode, lastEvent.practice_loc_id, lastEvent.punched_at);
  const worked = computeDayWorked(schedule, businessDateStr, sorted);
  return {
    hours: isNum(worked.hours) ? worked.hours : 0,
    minutes: isNum(worked.minutes) ? worked.minutes : 0
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

/** Zero-pads a time string to 'HH:mm' (e.g. "9:30" -> "09:30") so it sorts correctly as plain text. */
function padTimeStr(v) {
  const s = normalizeTimeStr(v);
  const m = s.match(/^(\d{1,2}):(\d{2})$/);
  return m ? m[1].padStart(2, '0') + ':' + m[2] : s;
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
    [SHEETS.WEEKLY_SUMMARY]: ['week_start', 'week_end', 'employee_code', 'employee_name',
                              'worked_hours', 'worked_minutes', 'generated_at'],
    'errors': ['at', 'message', 'stack']
  };

  const scheduleColumns = ['business_date', 'employee_code', 'practice_loc_id', 'scheduled_start', 'scheduled_hours', 'note'];
  defs[SHEETS.SCHEDULES_LEGACY] = scheduleColumns; // kept only so migrateSchedulesToCategories() can still read it
  SCHEDULE_CATEGORIES.forEach(cat => { defs[scheduleSheetName(cat)] = scheduleColumns; });

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
    { code: 'E001', name: 'Kan', email: '' },
    { code: 'E002', name: 'Kenta', email: '' },
    { code: 'E003', name: 'Rio', email: '' },
    { code: 'E004', name: 'Issei', email: '' },
    { code: 'E005', name: 'Vincent', email: '' },
    { code: 'E006', name: 'Kai Yang', email: '' },
    { code: 'E007', name: 'Shizuki', email: '' },
    { code: 'E008', name: 'Suzuku', email: '' },
    { code: 'E009', name: 'Yusuke', email: '' },
    { code: 'A001', name: 'Ken', email: '' },
    { code: 'A002', name: 'Zhi Zhi', email: '' },
    { code: 'A003', name: 'Al', email: '' },
    { code: 'A004', name: 'CK', email: '' },
    { code: 'A005', name: 'Darren', email: '' },
    { code: 'A006', name: 'Wai Chun', email: '' },
    { code: 'A007', name: 'Dominic', email: '' },
    { code: 'A008', name: 'Nicholas', email: '' },
    { code: 'A009', name: 'Sanjiv', email: '' },
    { code: 'A010', name: 'Wen Zhe', email: '' },
    { code: 'A011', name: 'Derek', email: '' },
    { code: 'A012', name: 'Shaunn', email: '' },
    { code: 'A013', name: 'Isaac', email: '' },
    { code: 'A014', name: 'Ivaac', email: '' },
    { code: 'A015', name: 'Matthias', email: '' },
    { code: 'A016', name: 'Justin', email: '' },
    { code: 'A017', name: 'Gabriel Low', email: '' },
    { code: 'A018', name: 'Yong Jun', email: '' },
    { code: 'A019', name: 'Javier', email: '' },
    { code: 'A020', name: 'Zahar', email: '' },
    { code: 'N001', name: 'Cayden', email: '' },
    { code: 'N002', name: 'Joshua', email: '' },
    { code: 'N003', name: 'Hong Kai', email: '' },
    { code: 'N004', name: 'Jonathan', email: '' },
    { code: 'N005', name: 'Abriel', email: '' },
    { code: 'N006', name: 'Maximus', email: '' },
    { code: 'N007', name: 'Mikel', email: '' },
    { code: 'S001', name: 'Adrian', email: '' },
    { code: 'Y001', name: 'Merrill', email: '' },
    { code: 'Y002', name: 'Jon-Wy', email: '' },
    { code: 'Y003', name: 'Kaien', email: '' },
    { code: 'Y004', name: 'Rhys', email: '' },
    { code: 'Y005', name: 'Qays', email: '' },
    { code: 'Y006', name: 'Jalen', email: '' },
    { code: 'Y007', name: 'Ethan', email: '' },
    { code: 'Y008', name: 'Gerald', email: '' },
    { code: 'Y009', name: 'Kae Mann', email: '' },
    { code: 'Y010', name: 'Ruiye', email: '' },
    { code: 'Y011', name: 'Nathan', email: '' },
    { code: 'Y012', name: 'Dylan', email: '' },
    { code: 'Y013', name: 'Evan', email: '' },
    { code: 'Y014', name: 'Zaiver', email: '' },
    { code: 'Y015', name: 'Gabriel', email: '' },
    { code: 'Y016', name: 'Skyler', email: '' },
    { code: 'Y017', name: 'Asher Poon', email: '' },
    { code: 'Y018', name: 'Tristan', email: '' },
    { code: 'Y019', name: 'Alden', email: '' },
    { code: 'Y020', name: 'Ryan', email: '' },
    { code: 'Y021', name: 'Zachary', email: '' },
    { code: 'Y022', name: 'Russell', email: '' },
    { code: 'Y023', name: 'Elijah', email: '' },
    { code: 'Y024', name: 'Jordan', email: '' },
    { code: 'Y025', name: 'Darius', email: '' },
    { code: 'Y026', name: 'Lucas', email: '' },
    { code: 'Y027', name: 'Clarence', email: '' },
    { code: 'Y028', name: 'Ayden', email: '' },
    { code: 'Y029', name: 'Joshua', email: '' },
    { code: 'Y030', name: 'Jarrod', email: '' },
    { code: 'Y031', name: 'Z-Hean', email: '' },
    { code: 'Y032', name: 'Savva', email: '' },
    { code: 'Y033', name: 'Bransten', email: '' },
    { code: 'Y034', name: 'Jayven', email: '' },
    { code: 'Y035', name: 'Asher Sim', email: '' },
    { code: 'Y036', name: 'Raphael', email: '' },
    { code: 'Y037', name: 'Lennon', email: '' },
    { code: 'Y038', name: 'Jeroy', email: '' },
    { code: 'Y039', name: 'Kaysan', email: '' }
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
 *   Every non-voided punch_events row in range is grouped by (date, employee)
 *   into its raw events (type + time + location), then recomputeDayWorked()
 *   recalculates that day's hours from scratch against whatever is currently
 *   in the schedules sheet. Nothing here trusts the worked_hours /
 *   worked_minutes columns already sitting in punch_events, so a schedule
 *   correction made after the fact is reflected automatically the next time
 *   this runs — just re-run it (safe: existing rows for the week are
 *   replaced, never duplicated).
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

  // Group raw events by (date, employee) so each day's hours can be
  // recomputed live from the current schedules sheet.
  const eventsByDayEmployee = {}; // 'date|code' -> [{type, punched_at, practice_loc_id}]
  const employeeNames = {};       // employee_code -> name

  values.forEach(r => {
    if (truthy(r[t.col.is_voided])) return;
    const bDate = normalizeDateStr(r[t.col.business_date]);
    if (bDate < weekStart || bDate > weekEnd) return;

    const code = String(r[t.col.employee_code]).trim();
    if (!code) return;

    employeeNames[code] = String(r[t.col.employee_name]);
    const key = bDate + '|' + code;
    (eventsByDayEmployee[key] = eventsByDayEmployee[key] || []).push({
      type: String(r[t.col.punch_type]),
      punched_at: new Date(r[t.col.punched_at]),
      practice_loc_id: String(r[t.col.practice_loc_id] || '')
    });
  });

  const totals = {};
  Object.keys(eventsByDayEmployee).forEach(key => {
    const sep = key.indexOf('|');
    const bDate = key.slice(0, sep);
    const code = key.slice(sep + 1);
    const worked = recomputeDayWorked(code, bDate, eventsByDayEmployee[key]);
    if (!totals[code]) totals[code] = { name: employeeNames[code], hours: 0, minutes: 0 };
    totals[code].hours += worked.hours;
    totals[code].minutes += worked.minutes;
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
 * Sets (or changes) the shared admin password. This gates BOTH report.html
 * (viewing every employee's hours) and schedule.html (bulk-adding rows to
 * the schedules sheet). Treat it like any other shared admin credential.
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
 * Builds the attendance report for a date range and writes it to the
 * "report" sheet, overwriting whatever was there before. Filters:
 *   - practice_loc_id   : single location id, or '' for all locations.
 *   - employee_codes    : array of specific employee codes to include.
 *   - employee_prefixes : array of single letters (e.g. 'E'); every employee
 *                         whose code starts with one of them is included.
 * A row matches if it's in employee_codes OR its prefix is in
 * employee_prefixes; if both arrays are empty, every employee is included.
 *
 * Every active, registered employee matching the filter gets a row, even
 * one with zero punches in the whole range — a player who never checked in
 * shows up as 0 hours / 0% rather than being silently omitted.
 *
 * Layout: one row per employee, one column per *session* in the range, plus
 * Total Hours / Attendance Rate columns. Every session column and Total
 * Hours are formatted to always show 2 decimal places (e.g. "3.00"), since
 * that's precise enough to distinguish individual minutes.
 *
 * A date with no practice scheduled at all (across the categories implied
 * by the employee filter — see distinctSessionStarts()) gets no column at
 * all; it's simply left out of the report rather than showing as a
 * misleading all-zero day.
 *
 * Otherwise a date normally gets one column, labeled with just the date.
 * But if distinctSessionStarts() finds 2+ distinct scheduled_start times on
 * that date (matching the location filter) — e.g. a morning and an evening
 * practice — it becomes multiple columns instead, labeled "yyyy-MM-dd(1)",
 * "yyyy-MM-dd(2)", etc. in start-time order. splitBySession() assigns each
 * punch to a column using a 30-minutes-before-start threshold: a punch
 * counts toward session N once it's within 30 minutes of session N's start
 * (and not yet within 30 minutes of session N+1's start). This keeps a
 * double-practice day from having one session's hours silently swallowed
 * into the other's schedule window, which is what happened when the whole
 * day shared a single schedule for clipping.
 *   - Total Hours   : for each session column, that session's own punches
 *                     (see splitBySession() above) are clipped against that
 *                     session's own schedule row — via findSchedule() /
 *                     computeDayWorked(), never the worked_hours snapshot
 *                     stored on the row itself — so a schedule correction
 *                     made after the fact is reflected the next time the
 *                     report is generated. A column only shows hours if the
 *                     employee has at least one non-voided punch in that
 *                     session matching the location filter (if any).
 *   - Attendance Rate: count-based, not hours-based. For every session
 *                     column, findSchedule() (same employee/location
 *                     priority rules as the live worked-hours calculation,
 *                     anchored to that session's own start time) checks
 *                     whether an actual practice was scheduled for that
 *                     employee in that session (schedule.matched — the
 *                     CONFIG fallback used when nothing matches does NOT
 *                     count as a scheduled practice). That count is the
 *                     denominator. The numerator is how many of those
 *                     scheduled sessions the employee actually has a punch
 *                     in. Rate = attended / scheduled sessions, as a
 *                     percentage — arriving late or leaving early doesn't
 *                     reduce it, only being fully absent from a scheduled
 *                     session does.
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

  // Employees can be picked individually (employee_codes) and/or by the
  // first letter of their code (employee_prefixes, e.g. 'E' for E001/E002/...).
  // A row matches if either list includes it; if both lists are empty, every
  // employee matches (the "all employees" default).
  const empCodesFilter = (Array.isArray(req.employee_codes) ? req.employee_codes : [])
    .map(c => String(c).trim().toUpperCase()).filter(Boolean);
  const empPrefixFilter = (Array.isArray(req.employee_prefixes) ? req.employee_prefixes : [])
    .map(c => String(c).trim().toUpperCase()).filter(Boolean);
  const hasEmpFilter = empCodesFilter.length > 0 || empPrefixFilter.length > 0;

  function employeeMatches(code) {
    if (!hasEmpFilter) return true;
    const upper = code.toUpperCase();
    return empCodesFilter.indexOf(upper) !== -1 || empPrefixFilter.indexOf(upper.charAt(0)) !== -1;
  }

  // Which categories' schedule sheets can possibly contain a relevant date/
  // session — with no employee filter, every category; with a filter, only
  // the categories it could actually match. Used to decide which dates have
  // a practice scheduled at all (see Pass 2 below).
  const categories = !hasEmpFilter
    ? SCHEDULE_CATEGORIES.slice()
    : Array.from(new Set(
        SCHEDULE_CATEGORIES.filter(cat =>
          empPrefixFilter.indexOf(cat) !== -1 ||
          empCodesFilter.some(c => c.charAt(0) === cat))
      ));

  const t = readHeader(SHEETS.EVENTS);
  const lastRow = t.sheet.getLastRow();

  // Pass 1: collect every non-voided event in range, grouped by (date, employee),
  // regardless of the location filter — pairing Check In / Leave Early correctly
  // requires seeing the whole day.
  const eventsByDayEmployee = {}; // 'date|code' -> [{type, punched_at, practice_loc_id}]
  const employeeNames = {};       // employee_code -> name

  if (lastRow >= 2) {
    const values = t.sheet.getRange(2, 1, lastRow - 1, t.headers.length).getValues();
    values.forEach(r => {
      if (truthy(r[t.col.is_voided])) return;

      const bDate = normalizeDateStr(r[t.col.business_date]);
      if (bDate < startDate || bDate > endDate) return;

      const code = String(r[t.col.employee_code]).trim();
      if (!code) return;

      employeeNames[code] = String(r[t.col.employee_name]);
      const key = bDate + '|' + code;
      (eventsByDayEmployee[key] = eventsByDayEmployee[key] || []).push({
        type: String(r[t.col.punch_type]),
        punched_at: new Date(r[t.col.punched_at]),
        practice_loc_id: String(r[t.col.practice_loc_id] || '')
      });
    });
  }

  // Pass 2: build the column list — one per date normally, or one per
  // session (see distinctSessionStarts()) on a date with 2+ practices. A
  // date with no scheduled practice at all for these categories is dropped
  // entirely — an unscheduled day has no business appearing in the report.
  const columns = []; // [{label, date, sessionStart}] — sessionStart is null for a single-session date
  dates.forEach(d => {
    const starts = distinctSessionStarts(d, locFilter, categories);
    if (starts.length === 0) return;
    if (starts.length === 1) {
      columns.push({ label: d, date: d, sessionStart: null });
    } else {
      starts.forEach((s, i) => columns.push({ label: d + '(' + (i + 1) + ')', date: d, sessionStart: s }));
    }
  });

  // Pass 3: recompute hours + attendance per (column, employee), from that
  // column's own slice of the day's events. A column only gets an entry —
  // and only then counts toward attendance — if at least one of its punches
  // matches the location filter (if any); this mirrors how a location
  // filter has always worked, so a player who trained only at some other
  // location that day still shows 0 for this location's report instead of
  // an unrelated location's hours leaking in.
  const perEmployeeCol = {}; // employee_code -> { column_label -> worked_hours }
  const colIncluded = {};    // 'column_label|code' -> true

  function matchesLocFilter(evts) {
    return locFilter ? evts.some(ev => ev.practice_loc_id === locFilter) : evts.length > 0;
  }

  Object.keys(eventsByDayEmployee).forEach(key => {
    const sep = key.indexOf('|');
    const d = key.slice(0, sep);
    const code = key.slice(sep + 1);
    if (!employeeMatches(code)) return;
    const dayEvents = eventsByDayEmployee[key];
    if (!perEmployeeCol[code]) perEmployeeCol[code] = {};

    const starts = distinctSessionStarts(d, locFilter, categories);
    if (starts.length === 0) return; // no practice scheduled that day — dropped from columns, so skip
    if (starts.length === 1) {
      if (!matchesLocFilter(dayEvents)) return;
      const worked = recomputeDayWorked(code, d, dayEvents);
      perEmployeeCol[code][d] = isNum(worked.hours) ? Number(worked.hours) : 0;
      colIncluded[d + '|' + code] = true;
    } else {
      const buckets = splitBySession(dayEvents, starts, d);
      buckets.forEach((bucketEvents, i) => {
        if (!matchesLocFilter(bucketEvents)) return;
        const label = d + '(' + (i + 1) + ')';
        const sessionStart = Utilities.parseDate(d + ' ' + starts[i], CONFIG.TZ, 'yyyy-MM-dd HH:mm');
        const schedule = findSchedule(d, code, locFilter, sessionStart);
        const worked = computeDayWorked(schedule, d, bucketEvents);
        perEmployeeCol[code][label] = isNum(worked.hours) ? Number(worked.hours) : 0;
        colIncluded[label + '|' + code] = true;
      });
    }
  });

  // Every active, registered employee matching the filter appears in the
  // report even with zero punches, so a player who never showed up still
  // shows as absent (0 hours, 0% attendance) instead of being left out.
  readTable(SHEETS.EMPLOYEES).rows
    .filter(r => truthy(r.is_active))
    .forEach(r => {
      const code = String(r.employee_code || '').trim();
      if (!code || !employeeMatches(code)) return;
      employeeNames[code] = String(r.name);
      if (!perEmployeeCol[code]) perEmployeeCol[code] = {};
    });

  const employeeCodes = Object.keys(perEmployeeCol).sort();
  if (!employeeCodes.length) {
    return { ok: false, error: 'NO_DATA', message: 'No matching players were found for this filter.' };
  }

  const headerRow = ['Employee'].concat(columns.map(c => c.label), ['Total Hours', 'Attendance Rate']);
  const dataRows = employeeCodes.map(code => {
    let totalActual = 0;
    let scheduledCount = 0;
    let attendedCount = 0;
    const cells = columns.map(c => {
      const hrs = (perEmployeeCol[code] && perEmployeeCol[code][c.label]) || 0;
      totalActual += hrs;

      const sched = c.sessionStart
        ? findSchedule(c.date, code, locFilter, Utilities.parseDate(c.date + ' ' + c.sessionStart, CONFIG.TZ, 'yyyy-MM-dd HH:mm'))
        : findSchedule(c.date, code, locFilter);
      if (sched && sched.matched) {
        scheduledCount++;
        if (colIncluded[c.label + '|' + code]) attendedCount++;
      }
      return hrs;
    });
    const rate = scheduledCount > 0 ? Math.round((attendedCount / scheduledCount) * 1000) / 10 + '%' : '';
    return [employeeNames[code] || code].concat(
      cells,
      [Math.round(totalActual * 100) / 100, rate]
    );
  });

  const empDescParts = [];
  if (empCodesFilter.length) empDescParts.push(empCodesFilter.join(', '));
  if (empPrefixFilter.length) empDescParts.push(empPrefixFilter.map(p => p + '*').join(', '));
  const empDesc = empDescParts.length ? empDescParts.join(' + ') : 'All employees';

  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let sh = ss.getSheetByName('report');
  if (!sh) sh = ss.insertSheet('report');
  sh.clear();

  const title = 'Attendance Report — ' + (locFilter || 'All locations') + ' — ' +
                empDesc + ' — ' + startDate + ' to ' + endDate +
                ' (generated ' + Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy-MM-dd HH:mm') + ')';
  sh.getRange(1, 1).setValue(title);
  sh.getRange(2, 1, 1, headerRow.length).setValues([headerRow]);
  sh.getRange(2, 1, 1, headerRow.length).setFontWeight('bold');
  if (dataRows.length) {
    sh.getRange(3, 1, dataRows.length, headerRow.length).setValues(dataRows);
    // Force every session column plus Total Hours to always show 2 decimal
    // places (e.g. "3.00" instead of "3"), so partial-hour/minute detail
    // is visible even for whole-hour totals.
    const hourColumns = columns.length + 1; // each session column + Total Hours
    sh.getRange(3, 2, dataRows.length, hourColumns).setNumberFormat('0.00');
  }
  sh.setFrozenRows(2);

  return {
    ok: true,
    sheet_url: ss.getUrl() + '#gid=' + sh.getSheetId(),
    employee_count: dataRows.length,
    date_count: dates.length
  };
}


/* ════════════════════════════════════════════
   10. Bulk schedule entry (admin, password-protected)
   ════════════════════════════════════════════ */

/**
 * Returns the practice-location dropdown options, and the category list
 * (National A / Sparring / National Service / NYDS), for schedule.html.
 * Requires the shared admin password (same one as report.html).
 */
function handleScheduleMeta(req) {
  if (!checkReportPassword(req.password)) {
    return { ok: false, error: 'BAD_PASSWORD', message: 'Incorrect password.' };
  }

  const locations = readTable(SHEETS.LOCATIONS).rows
    .filter(r => truthy(r.is_active))
    .map(r => ({ id: String(r.loc_id), label: String(r.name) + ' (' + String(r.type) + ')' }));

  const categories = SCHEDULE_CATEGORIES.map(id => ({ id: id, label: SCHEDULE_CATEGORY_LABELS[id] }));

  return { ok: true, locations: locations, categories: categories };
}

/**
 * Adds one row to the chosen category's schedule sheet (schedule_A/S/N/Y —
 * see SCHEDULE_CATEGORY_LABELS) for every matching date, all sharing the
 * same start time / hours / location / employee / note. Meant for the
 * common case where practice happens at the same time and place for weeks
 * at a stretch, so the admin doesn't have to type one row per date. Dates
 * come from two optional, combinable sources — at least one must produce a
 * date:
 *   - A recurring pattern: start_date + end_date + weekdays (0=Sun..6=Sat).
 *     All three are required together if any of them is given.
 *   - Individually picked dates: the dates array ('yyyy-MM-dd' strings),
 *     for one-off days that don't fit a weekly pattern.
 * The two sources are merged and de-duplicated, so a recurring pattern and a
 * few extra individual dates can be submitted in the same request.
 *
 * employee_code and practice_loc_id are optional, same as manual entry:
 * blank employee_code = everyone in this category, blank practice_loc_id =
 * any location.
 *
 * After adding the rows, that category's sheet is re-sorted by date and
 * time (sortScheduleSheetByName()) so it stays in chronological order.
 */
function handleAddSchedule(req) {
  if (!checkReportPassword(req.password)) {
    return { ok: false, error: 'BAD_PASSWORD', message: 'Incorrect password.' };
  }

  const category = String(req.category || '').trim().toUpperCase();
  if (SCHEDULE_CATEGORIES.indexOf(category) === -1) {
    return { ok: false, error: 'BAD_CATEGORY', message: 'Please choose a category.' };
  }

  const startDate = String(req.start_date || '').trim();
  const endDate = String(req.end_date || '').trim();
  const weekdays = (Array.isArray(req.weekdays) ? req.weekdays : []).map(Number);
  const individualDates = (Array.isArray(req.dates) ? req.dates : [])
    .map(d => String(d).trim()).filter(Boolean);

  const rangeDates = [];
  const wantsRange = startDate || endDate || weekdays.length;
  if (wantsRange) {
    if (!startDate || !endDate || startDate > endDate) {
      return { ok: false, error: 'BAD_RANGE', message: 'Please select a valid date range.' };
    }
    if (!weekdays.length) {
      return { ok: false, error: 'NO_WEEKDAYS', message: 'Please select at least one day of the week.' };
    }
    for (let d = startDate; d <= endDate; d = addDays(d, 1)) {
      const dow = Utilities.parseDate(d, CONFIG.TZ, 'yyyy-MM-dd').getDay(); // 0=Sun..6=Sat, same convention as mondayOfWeek()
      if (weekdays.indexOf(dow) !== -1) rangeDates.push(d);
      if (rangeDates.length > 366) {
        return { ok: false, error: 'RANGE_TOO_LARGE', message: 'That range is too large — please narrow it down.' };
      }
    }
  }

  const dates = Array.from(new Set(rangeDates.concat(individualDates))).sort();
  if (!dates.length) {
    return { ok: false, error: 'NO_DATES', message: 'Please add at least one date — individually, or via a date range and weekdays.' };
  }
  if (dates.length > 366) {
    return { ok: false, error: 'RANGE_TOO_LARGE', message: 'That’s too many dates at once — please narrow it down.' };
  }

  const scheduledStart = String(req.scheduled_start || '').trim();
  const scheduledHours = Number(req.scheduled_hours);
  if (!scheduledStart || !isNum(scheduledHours) || scheduledHours <= 0) {
    return { ok: false, error: 'BAD_SCHEDULE', message: 'Please enter a valid start time and number of hours.' };
  }

  const employeeCode = String(req.employee_code || '').trim().toUpperCase();
  const practiceLocId = String(req.practice_loc_id || '').trim();
  const note = String(req.note || '').trim();
  const sheetName = scheduleSheetName(category);

  dates.forEach(d => {
    appendRowByHeader(sheetName, {
      business_date: d,
      employee_code: employeeCode,
      practice_loc_id: practiceLocId,
      scheduled_start: scheduledStart,
      scheduled_hours: scheduledHours,
      note: note
    });
  });

  sortScheduleSheetByName(sheetName);

  return { ok: true, added: dates.length };
}

/** Re-sorts every category's schedule sheet by date without adding anything, for cleaning up rows added by hand. */
function handleSortSchedule(req) {
  if (!checkReportPassword(req.password)) {
    return { ok: false, error: 'BAD_PASSWORD', message: 'Incorrect password.' };
  }
  SCHEDULE_CATEGORIES.forEach(cat => sortScheduleSheetByName(scheduleSheetName(cat)));
  return { ok: true };
}

/**
 * Sorts every row in the given schedule sheet by business_date, then by
 * scheduled_start within the same date, earliest first. Also normalizes
 * both to a consistent text format while sorting — some rows were typed
 * inconsistently (e.g. "2026-8-23" instead of "2026-08-23", or "9:30"
 * instead of "09:30"), which sorts wrong as plain text even though the
 * values are the same. Safe to run any time; it only reorders rows and
 * rewrites the business_date/scheduled_start columns, nothing else changes.
 */
function sortScheduleSheetByName(sheetName) {
  const t = readTable(sheetName);
  if (!t.rows.length) return;

  const rows = t.rows.map(r => {
    const copy = Object.assign({}, r);
    copy.business_date = normalizeDateStr(r.business_date);
    copy.scheduled_start = padTimeStr(r.scheduled_start);
    return copy;
  });
  rows.sort((a, b) => {
    if (a.business_date !== b.business_date) return a.business_date < b.business_date ? -1 : 1;
    if (a.scheduled_start !== b.scheduled_start) return a.scheduled_start < b.scheduled_start ? -1 : 1;
    return 0;
  });

  const values = rows.map(r => t.headers.map(h => {
    const key = String(h).trim();
    return r[key] !== undefined ? r[key] : '';
  }));
  t.sheet.getRange(2, 1, values.length, t.headers.length).setValues(values);
}

/**
 * One-time migration: copies every row from the old unified 'schedules'
 * sheet into ALL FOUR category sheets (schedule_A/S/N/Y), since every row
 * in that sheet was a blanket entry (blank employee_code) that applied to
 * the whole team regardless of category. Run this once, from the Apps
 * Script editor, after running setup() to create the new sheets — it makes
 * every practice already entered keep applying to everyone until it's
 * edited to diverge per category. Safe to run more than once: each
 * category sheet's existing rows are cleared first, so re-running just
 * re-copies from 'schedules' rather than duplicating.
 */
function migrateSchedulesToCategories() {
  const old = readTable(SHEETS.SCHEDULES_LEGACY);
  SCHEDULE_CATEGORIES.forEach(cat => {
    const name = scheduleSheetName(cat);
    const sh = sheet(name);
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
    old.rows.forEach(r => {
      appendRowByHeader(name, {
        business_date: normalizeDateStr(r.business_date),
        employee_code: r.employee_code || '',
        practice_loc_id: r.practice_loc_id || '',
        scheduled_start: padTimeStr(r.scheduled_start),
        scheduled_hours: r.scheduled_hours,
        note: r.note || ''
      });
    });
    sortScheduleSheetByName(name);
  });
  Logger.log('Copied ' + old.rows.length + ' rows from "schedules" into each of: ' +
             SCHEDULE_CATEGORIES.map(scheduleSheetName).join(', '));
}
