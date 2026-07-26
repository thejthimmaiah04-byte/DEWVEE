/**
 * Weather Pod -> Google Sheets receiver
 * -------------------------------------
 * Deploy this as a Web App (see the deploy steps your assistant gave you).
 * The ESP32 POSTs a JSON batch; this appends one row per reading.
 *
 * Expected POST body:
 *   { "device": "weatherpod-1",
 *     "rows": [ { "t": <utc_epoch_seconds>, "temp": 24.31,
 *                 "hum": 58.2, "pct": 87, "v": 3.94 }, ... ] }
 *
 * Columns written: Timestamp | Device | Temp (C) | Humidity (%) | Battery (%) | Battery (V)
 * Tip: set the Sheet's timezone (File > Settings > Time zone) to IST so
 *      the Timestamp column displays correctly.
 *
 * Duplicate protection (fully automatic, no manual steps required):
 *  - doPost() skips any incoming row whose (Timestamp, Device) pair
 *    already exists in the sheet, or is repeated within the same batch.
 *  - It then runs a full-sheet sweep (cleanupDuplicateRows) on every
 *    single upload, deleting any duplicate row that could still have
 *    slipped through, keeping the first occurrence of each pair.
 *  - removeDuplicateTimestamps() remains available to run by hand from
 *    the Apps Script editor if you ever want to trigger a sweep outside
 *    of an upload, but it is no longer required for normal operation.
 */

var SHEET_NAME = 'Data';
var TIMEZONE   = 'Asia/Kolkata';   // change if deploying elsewhere

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var device = body.device || 'unknown';
    var rows   = body.rows || [];

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    // Write a header row once.
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp', 'Device', 'Temp (C)',
                       'Humidity (%)', 'Battery (%)', 'Battery (V)']);
      sheet.setFrozenRows(1);
    }

    // Build a set of "timestamp|device" keys already present, so we can
    // skip anything that would duplicate an existing row.
    var existingKeys = {};
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var existing = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // A:B
      for (var i = 0; i < existing.length; i++) {
        existingKeys[existing[i][0] + '|' + existing[i][1]] = true;
      }
    }

    var startRow = sheet.getLastRow() + 1;
    var out = [];
    var skipped = 0;

    rows.forEach(function (r) {
      var ts = r.t
        ? Utilities.formatDate(new Date(r.t * 1000), TIMEZONE, 'yyyy-MM-dd HH:mm:ss')
        : '';
      var key = ts + '|' + device;

      if (existingKeys[key]) {          // already in the sheet, or already
        skipped++;                      // queued earlier in this same batch
        return;
      }
      existingKeys[key] = true;         // claim it so within-batch dupes are caught too

      out.push([ ts, device, r.temp, r.hum, r.pct, r.v ]);
    });

    if (out.length > 0) {
      // Set text format on column A BEFORE writing so Sheets stores the
      // timestamp as a plain string, not as a date serial that loses zero-padding.
      sheet.getRange(startRow, 1, out.length, 1).setNumberFormat('@STRING@');
      var range = sheet.getRange(startRow, 1, out.length, out[0].length);
      range.setValues(out);
    }

    // Automatic safety sweep: even though rows are checked against
    // existing data before insert (above), this catches anything that
    // could still slip through - e.g. two near-simultaneous requests
    // racing each other - so duplicates never require a manual cleanup.
    var removed = cleanupDuplicateRows(sheet);

    return ContentService
      .createTextOutput(JSON.stringify({
        ok: true, added: out.length, skipped: skipped, duplicatesRemoved: removed
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * doGet routes off the same /exec URL:
 *   (no action)         -> serves the dashboard web app (Dashboard.html)
 *   ?action=devices     -> JSON: one entry per pod (latest battery, last-seen)
 *   ?action=data&...    -> JSON: down-sampled time-series for the chart
 *   ?action=export&...  -> JSON: RAW rows for a [from,to] date range (CSV export)
 * The dashboard itself calls the getDevices()/getReadings()/getReadingsRange()
 * functions directly via google.script.run (same-origin, so no CORS), but
 * the JSON endpoints are kept so a future native app can read the same data.
 */
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || '';
  if (action === 'devices') return jsonOut(getDevices());
  if (action === 'data')    return jsonOut(getReadings(e.parameter));
  if (action === 'export')  return jsonOut(getReadingsRange(e.parameter));
  if (action === 'compare') return jsonOut(getReadingsCompare(e.parameter));
  if (action === 'ping')    return jsonOut({ ok: true, msg: 'weather pod endpoint alive' });

  return HtmlService.createHtmlOutputFromFile('Dashboard')
    .setTitle('Dewvee')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------
//  Data access used by the dashboard (and the JSON endpoints)
// ---------------------------------------------------------------

var LOW_BATTERY_PCT = 20;                 // alert threshold
// Must be LONGER than a pod's upload interval, or a healthy pod looks offline
// between uploads. Pods upload every ~6 h, so 8 h keeps on-time pods "online"
// while still flagging one that has missed its window. (Once per-pod upload
// intervals come from BLE config, we can compute this exactly per device.)
var OFFLINE_AFTER_MS = 12 * 60 * 60 * 1000;

// Sheet timestamps are plain strings in TIMEZONE (IST). Parse one back to a
// millisecond epoch deterministically, without relying on Date string parsing.
function tsToEpochMs(ts) {
  // expected: 'yyyy-MM-dd HH:mm:ss'
  var m = /^(\d{4})-(\d{2})-(\d{2})[ T](\d{1,2}):(\d{2}):(\d{2})/.exec(String(ts));
  if (!m) return NaN;
  var utc = Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6]);
  return utc - 5.5 * 3600 * 1000;   // Asia/Kolkata is UTC+5:30
}

// A cell in column A is normally our forced-text timestamp, but a manually
// entered row could come back as a real Date object - handle both.
function rowEpochMs(v0) {
  if (v0 instanceof Date) return v0.getTime();
  return tsToEpochMs(v0);
}

// Battery % from resting voltage. The ESP now applies its own ADC/divider
// calibration before sending, so the voltage arriving here is the TRUE pack
// voltage (matches a multimeter) - 4.20 V is the standard single-cell
// Li-ion/LiPo full-charge voltage, so that's the anchor for 100%. Linearly
// interpolated between the table points so the reading moves smoothly
// (effectively continuous, not in coarse 10% steps).
function battPctFromVolts(v) {
  var T = [[4.20,100],[4.10,90],[4.00,80],[3.90,70],[3.80,60],[3.75,50],
           [3.70,40],[3.65,30],[3.60,20],[3.50,12],[3.40,6],[3.30,3],[3.00,0]];
  if (v >= T[0][0]) return 100;
  if (v <= T[T.length - 1][0]) return 0;
  for (var i = 0; i < T.length - 1; i++) {
    var hi = T[i], lo = T[i + 1];
    if (v <= hi[0] && v >= lo[0]) {
      var f = (v - lo[0]) / (hi[0] - lo[0]);
      return Math.round(lo[1] + f * (hi[1] - lo[1]));
    }
  }
  return 0;
}

// Pull the whole data sheet once as an array of row objects.
function readAllRows_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var v = values[i];
    var t = rowEpochMs(v[0]);
    if (isNaN(t)) continue;
    var volt = Number(v[5]);
    var pct  = (!isNaN(volt) && volt > 0) ? battPctFromVolts(volt) : Number(v[4]);
    out.push({
      t: t, device: String(v[1]),
      temp: Number(v[2]), hum: Number(v[3]),
      pct: pct, volt: volt
    });
  }
  out.sort(function (a, b) { return a.t - b.t; });
  return out;
}

// One summary entry per pod: latest reading, battery, online/low flags.
function getDevices() {
  var rows = readAllRows_();
  var now = Date.now();
  var byDev = {};
  rows.forEach(function (r) {
    var d = byDev[r.device];
    if (!d || r.t > d.lastSeen) {
      byDev[r.device] = {
        device: r.device, lastSeen: r.t,
        pct: r.pct, volt: r.volt, temp: r.temp, hum: r.hum,
        count: (d ? d.count : 0) + 1
      };
    } else {
      d.count++;
    }
  });
  var list = Object.keys(byDev).map(function (k) {
    var d = byDev[k];
    d.online   = (now - d.lastSeen) <= OFFLINE_AFTER_MS;
    d.lowBatt  = d.pct <= LOW_BATTERY_PCT;
    d.ageMs    = now - d.lastSeen;
    return d;
  });
  list.sort(function (a, b) { return a.device < b.device ? -1 : 1; });
  return { now: now, lowThreshold: LOW_BATTERY_PCT, devices: list };
}

// Down-sampled time-series for one or more pods over a named range.
// params: { devices?: 'a,b', range?: 'day|week|month|year|all' }
function getReadings(params) {
  params = params || {};
  var range = (params.range || 'day').toLowerCase();
  var wanted = params.devices ? String(params.devices).split(',') : null;

  var spans = {
    day:   { ms: 24 * 3600e3,        bucket: 5 * 60e3      },  // raw-ish
    week:  { ms: 7 * 24 * 3600e3,    bucket: 30 * 60e3     },
    month: { ms: 30 * 24 * 3600e3,   bucket: 2 * 3600e3    },
    year:  { ms: 365 * 24 * 3600e3,  bucket: 24 * 3600e3   },
    all:   { ms: Infinity,           bucket: 24 * 3600e3   }
  };
  var span = spans[range] || spans.day;
  var now = Date.now();
  var cutoff = span.ms === Infinity ? 0 : now - span.ms;

  var rows = readAllRows_().filter(function (r) {
    if (r.t < cutoff) return false;
    if (wanted && wanted.indexOf(r.device) < 0) return false;
    return true;
  });

  // Bucket -> average temp/hum, last battery, per device.
  var acc = {};   // device -> bucketStart -> {tS,hS,n,pct,volt,t}
  rows.forEach(function (r) {
    var b = Math.floor(r.t / span.bucket) * span.bucket;
    var dev = acc[r.device] || (acc[r.device] = {});
    var slot = dev[b] || (dev[b] = { t: b, tS: 0, hS: 0, n: 0, pct: r.pct, volt: r.volt });
    if (!isNaN(r.temp)) { slot.tS += r.temp; }
    if (!isNaN(r.hum))  { slot.hS += r.hum; }
    slot.n++;
    slot.pct = r.pct; slot.volt = r.volt;   // last-in-bucket wins for battery
  });

  var series = {};
  Object.keys(acc).forEach(function (dev) {
    series[dev] = Object.keys(acc[dev]).map(function (b) {
      var s = acc[dev][b];
      return {
        t: s.t,
        temp: s.n ? +(s.tS / s.n).toFixed(2) : null,
        hum:  s.n ? +(s.hS / s.n).toFixed(1) : null,
        pct:  s.pct
      };
    }).sort(function (a, b) { return a.t - b.t; });
  });

  return { range: range, now: now, bucketMs: span.bucket, series: series };
}

// RAW (not down-sampled) rows for an explicit [from, to] date range - used
// by the CSV export so it contains every actual reading in the window, not
// averaged buckets like the chart does. params: { from, to (ms epoch),
// devices?: 'a,b' }. "to" is treated as inclusive through the end of that
// calendar day (see the +1 day on the client) so picking the same start and
// end date exports that whole day.
function getReadingsRange(params) {
  params = params || {};
  var from = Number(params.from), to = Number(params.to);
  var wanted = params.devices ? String(params.devices).split(',') : null;
  if (!isFinite(from)) from = 0;
  if (!isFinite(to))   to   = Date.now();

  var rows = readAllRows_().filter(function (r) {
    if (r.t < from || r.t > to) return false;
    if (wanted && wanted.indexOf(r.device) < 0) return false;
    return true;
  });

  return {
    from: from, to: to,
    rows: rows.map(function (r) {
      return { t: r.t, device: r.device, temp: r.temp, hum: r.hum, pct: r.pct, volt: r.volt };
    })
  };
}

// Down-sampled time-series for an explicit [from, to] window used by the
// compare overlay on the dashboard. Identical bucketing logic to getReadings()
// but accepts epoch-ms bounds instead of a named range, and auto-sizes the
// bucket so it always returns ~200 data points per device regardless of window
// size.  params: { from, to (ms epoch), devices? }
function getReadingsCompare(params) {
  params = params || {};
  var from = Number(params.from), to = Number(params.to);
  var wanted = params.devices ? String(params.devices).split(',') : null;
  if (!isFinite(from)) from = 0;
  if (!isFinite(to))   to   = Date.now();

  var windowMs = to - from;
  // Round bucket up to the nearest 5-minute boundary, targeting ~200 points.
  var MIN_BUCKET = 5 * 60e3;
  var bucket = Math.max(MIN_BUCKET, Math.ceil(windowMs / 200 / MIN_BUCKET) * MIN_BUCKET);

  var rows = readAllRows_().filter(function (r) {
    if (r.t < from || r.t > to) return false;
    if (wanted && wanted.indexOf(r.device) < 0) return false;
    return true;
  });

  var acc = {};
  rows.forEach(function (r) {
    var b = Math.floor(r.t / bucket) * bucket;
    var dev = acc[r.device] || (acc[r.device] = {});
    var slot = dev[b] || (dev[b] = { t: b, tS: 0, hS: 0, n: 0 });
    if (!isNaN(r.temp)) slot.tS += r.temp;
    if (!isNaN(r.hum))  slot.hS += r.hum;
    slot.n++;
  });

  var series = {};
  Object.keys(acc).forEach(function (dev) {
    series[dev] = Object.keys(acc[dev]).map(function (b) {
      var s = acc[dev][b];
      return {
        t: s.t,
        temp: s.n ? +(s.tS / s.n).toFixed(2) : null,
        hum:  s.n ? +(s.hS / s.n).toFixed(1) : null
      };
    }).sort(function (a, b) { return a.t - b.t; });
  });

  return { from: from, to: to, bucketMs: bucket, series: series };
}

// ---------------------------------------------------------------
//  Battery / offline alerts by email (works even when the app is closed)
//  Set a time-driven trigger on this function (e.g. every 3 hours):
//  Apps Script editor -> Triggers -> Add Trigger -> checkBatteries,
//  Time-driven, Hour timer, every 3 hours.
// ---------------------------------------------------------------
function checkBatteries() {
  var info = getDevices();
  var alerts = [];
  info.devices.forEach(function (d) {
    if (d.lowBatt) alerts.push('LOW BATTERY: ' + d.device + ' at ' + d.pct + '%');
    if (!d.online) alerts.push('OFFLINE: ' + d.device +
                    ' (no data for ' + Math.round(d.ageMs / 3600e3) + ' h)');
  });
  if (!alerts.length) return;

  // De-dupe so you are not emailed the same alert every run.
  var props = PropertiesService.getScriptProperties();
  var last = props.getProperty('lastAlert') || '';
  var now = alerts.join('\n');
  if (now === last) return;
  props.setProperty('lastAlert', now);

  MailApp.sendEmail({
    to: Session.getEffectiveUser().getEmail(),
    subject: 'Dewvee alert (' + alerts.length + ')',
    body: alerts.join('\n')
  });
}

/**
 * Scans the whole sheet and deletes any row whose (Timestamp, Device)
 * pair already appeared earlier in the sheet, keeping the first
 * occurrence. Returns the number of rows removed.
 *
 * Called automatically at the end of every doPost() - no manual step
 * needed. You can also still run it by hand from the Apps Script editor
 * (function dropdown -> removeDuplicateTimestamps -> Run) if you ever
 * want to trigger a cleanup outside of an upload.
 */
function cleanupDuplicateRows(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return 0;   // header + fewer than 2 data rows: nothing to dedup

  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();   // Timestamp, Device
  var seen = {};
  var rowsToDelete = [];

  for (var i = 0; i < data.length; i++) {
    var key = data[i][0] + '|' + data[i][1];
    var sheetRow = i + 2;   // +2: 1-indexed, plus header row
    if (seen[key]) {
      rowsToDelete.push(sheetRow);
    } else {
      seen[key] = true;
    }
  }

  for (var d = rowsToDelete.length - 1; d >= 0; d--) {
    sheet.deleteRow(rowsToDelete[d]);
  }

  if (rowsToDelete.length > 0) {
    Logger.log('Auto-cleanup removed %s duplicate row(s).', rowsToDelete.length);
  }
  return rowsToDelete.length;
}

/** Manual entry point - same logic, callable by hand from the editor. */
function removeDuplicateTimestamps() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { Logger.log('Sheet "%s" not found.', SHEET_NAME); return; }
  var removed = cleanupDuplicateRows(sheet);
  Logger.log('Removed %s duplicate row(s).', removed);
}
