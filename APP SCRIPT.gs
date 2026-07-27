/**
 * Weather Pod -> Google Sheets receiver
 * -------------------------------------
 * Deploy this as a Web App (Execute as: Me, Who has access: Anyone).
 *
 * Each device gets its own sheet named after it (e.g. "DEWVEE:01").
 * Sheets are created automatically the first time a new device sends data.
 * If Google's default "Sheet1" tab exists when a device first registers,
 * it is renamed to that device's name instead of inserting a new tab.
 *
 * Expected POST body:
 *   { "device": "DEWVEE:01",
 *     "rows": [ { "t": <utc_epoch_seconds>, "temp": 24.31,
 *                 "hum": 58.2, "pct": 87, "v": 3.94 }, ... ] }
 *
 * Columns written:
 *   Timestamp | Device | Temp (C) | Humidity (%) | Battery (%) | Battery (V)
 *
 * Tip: set the spreadsheet timezone (File > Settings > Time zone) to IST so
 * the Timestamp column displays correctly.
 */

var TIMEZONE = 'Asia/Kolkata';   // change if deploying elsewhere

// ---------------------------------------------------------------
//  Sheet routing — one sheet per device, auto-created on arrival
// ---------------------------------------------------------------

/**
 * Returns the sheet for a device, creating it if necessary.
 * If Google's default "Sheet1" still exists, it is renamed to the device
 * name instead of inserting a brand-new tab — keeps the tab order clean.
 */
function getSheetForDevice_(ss, device) {
  var sheet = ss.getSheetByName(device);
  if (sheet) return sheet;

  // Repurpose the blank "Sheet1" if it exists, otherwise insert a new tab.
  var sheet1 = ss.getSheetByName('Sheet1');
  if (sheet1) {
    sheet1.setName(device);
    sheet = sheet1;
  } else {
    sheet = ss.insertSheet(device);
  }

  // Write header and freeze it on a fresh sheet.
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Timestamp', 'Device', 'Temp (C)',
                     'Humidity (%)', 'Battery (%)', 'Battery (V)']);
    sheet.setFrozenRows(1);
    sheet.getRange('A:A').setNumberFormat('@STRING@');
  }

  Logger.log('Sheet created/renamed for device: ' + device);
  return sheet;
}

// ---------------------------------------------------------------
//  POST handler — receives batches from the ESP32 pods
// ---------------------------------------------------------------

function doPost(e) {
  try {
    var body   = JSON.parse(e.postData.contents);
    var device = body.device || 'unknown';
    var rows   = body.rows   || [];

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getSheetForDevice_(ss, device);

    // Build a set of "timestamp|device" keys already in this sheet so we can
    // skip anything that would duplicate an existing row.
    var existingKeys = {};
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var existing = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
      for (var i = 0; i < existing.length; i++) {
        existingKeys[existing[i][0] + '|' + existing[i][1]] = true;
      }
    }

    var startRow = sheet.getLastRow() + 1;
    var out      = [];
    var skipped  = 0;

    rows.forEach(function (r) {
      var ts = r.t
        ? Utilities.formatDate(new Date(r.t * 1000), TIMEZONE, 'yyyy-MM-dd HH:mm:ss')
        : '';
      var key = ts + '|' + device;
      if (existingKeys[key]) { skipped++; return; }
      existingKeys[key] = true;
      out.push([ts, device, r.temp, r.hum, r.pct, r.v]);
    });

    if (out.length > 0) {
      // Set text format BEFORE writing so Sheets stores the timestamp as a
      // plain string and never auto-converts it to a date serial.
      sheet.getRange(startRow, 1, out.length, 1).setNumberFormat('@STRING@');
      sheet.getRange(startRow, 1, out.length, out[0].length).setValues(out);
    }

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

// ---------------------------------------------------------------
//  GET handler — dashboard + JSON endpoints
// ---------------------------------------------------------------

/**
 * doGet routes off the same /exec URL:
 *   (no action)         -> serves the dashboard web app (Dashboard.html)
 *   ?action=devices     -> JSON: one entry per pod (latest battery, last-seen)
 *   ?action=data&...    -> JSON: down-sampled time-series for the chart
 *   ?action=export&...  -> JSON: raw rows for a [from,to] date range
 *   ?action=compare&... -> JSON: down-sampled time-series for compare overlay
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
//  Data access — used by the dashboard and JSON endpoints
// ---------------------------------------------------------------

var LOW_BATTERY_PCT  = 20;
var OFFLINE_AFTER_MS = 12 * 60 * 60 * 1000;   // 12 h

// Sheet timestamps are plain strings in TIMEZONE (IST). Parse one back to
// a millisecond epoch deterministically, without relying on Date.parse().
function tsToEpochMs(ts) {
  var m = /^(\d{4})-(\d{2})-(\d{2})[ T](\d{1,2}):(\d{2}):(\d{2})/.exec(String(ts));
  if (!m) return NaN;
  var utc = Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6]);
  return utc - 5.5 * 3600 * 1000;   // Asia/Kolkata = UTC+5:30
}

// A cell in column A is normally our forced-text timestamp, but a manually
// entered row could come back as a real Date object — handle both.
function rowEpochMs(v0) {
  if (v0 instanceof Date) return v0.getTime();
  return tsToEpochMs(v0);
}

// Battery % from resting voltage (linearly interpolated between table points).
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

/**
 * Reads all device sheets and returns a merged, time-sorted array of row
 * objects. Any sheet in the spreadsheet whose data rows have a parseable
 * timestamp in column A is treated as a device sheet automatically —
 * this includes the old "Data" sheet if it still exists, so no old
 * readings are lost after the migration to per-device sheets.
 */
function readAllRows_() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var out = [];

  ss.getSheets().forEach(function (sheet) {
    if (sheet.getLastRow() < 2) return;   // empty or header-only — skip
    var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
    for (var i = 0; i < values.length; i++) {
      var v = values[i];
      var t = rowEpochMs(v[0]);
      if (isNaN(t)) continue;             // not a data row (e.g. non-device sheet)
      var volt = Number(v[5]);
      var pct  = (!isNaN(volt) && volt > 0) ? battPctFromVolts(volt) : Number(v[4]);
      out.push({
        t: t, device: String(v[1]),
        temp: Number(v[2]), hum: Number(v[3]),
        pct: pct, volt: volt
      });
    }
  });

  out.sort(function (a, b) { return a.t - b.t; });
  return out;
}

// One summary entry per pod: latest reading, battery, online/low flags.
function getDevices() {
  var rows = readAllRows_();
  var now  = Date.now();
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
    d.online  = (now - d.lastSeen) <= OFFLINE_AFTER_MS;
    d.lowBatt = d.pct <= LOW_BATTERY_PCT;
    d.ageMs   = now - d.lastSeen;
    return d;
  });
  list.sort(function (a, b) { return a.device < b.device ? -1 : 1; });
  return { now: now, lowThreshold: LOW_BATTERY_PCT, devices: list };
}

// Down-sampled time-series for one or more pods over a named range.
// params: { devices?: 'a,b', range?: 'day|week|month|year|all' }
function getReadings(params) {
  params = params || {};
  var range  = (params.range || 'day').toLowerCase();
  var wanted = params.devices ? String(params.devices).split(',') : null;

  var spans = {
    day:   { ms: 24 * 3600e3,       bucket: 5 * 60e3    },
    week:  { ms: 7 * 24 * 3600e3,   bucket: 30 * 60e3   },
    month: { ms: 30 * 24 * 3600e3,  bucket: 2 * 3600e3  },
    year:  { ms: 365 * 24 * 3600e3, bucket: 24 * 3600e3 },
    all:   { ms: Infinity,          bucket: 24 * 3600e3 }
  };
  var span   = spans[range] || spans.day;
  var now    = Date.now();
  var cutoff = span.ms === Infinity ? 0 : now - span.ms;

  var rows = readAllRows_().filter(function (r) {
    if (r.t < cutoff) return false;
    if (wanted && wanted.indexOf(r.device) < 0) return false;
    return true;
  });

  var acc = {};
  rows.forEach(function (r) {
    var b    = Math.floor(r.t / span.bucket) * span.bucket;
    var dev  = acc[r.device] || (acc[r.device] = {});
    var slot = dev[b] || (dev[b] = { t: b, tS: 0, hS: 0, n: 0, pct: r.pct, volt: r.volt });
    if (!isNaN(r.temp)) slot.tS += r.temp;
    if (!isNaN(r.hum))  slot.hS += r.hum;
    slot.n++;
    slot.pct  = r.pct;
    slot.volt = r.volt;
  });

  var series = {};
  Object.keys(acc).forEach(function (dev) {
    series[dev] = Object.keys(acc[dev]).map(function (b) {
      var s = acc[dev][b];
      return {
        t:    s.t,
        temp: s.n ? +(s.tS / s.n).toFixed(2) : null,
        hum:  s.n ? +(s.hS / s.n).toFixed(1) : null,
        pct:  s.pct
      };
    }).sort(function (a, b) { return a.t - b.t; });
  });

  return { range: range, now: now, bucketMs: span.bucket, series: series };
}

// Raw (not down-sampled) rows for an explicit [from, to] window — used by
// CSV export. params: { from, to (ms epoch), devices?: 'a,b' }
function getReadingsRange(params) {
  params = params || {};
  var from   = Number(params.from), to = Number(params.to);
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

// Down-sampled time-series for an explicit [from, to] window — used by the
// compare overlay. Auto-sizes the bucket to ~200 points per device.
// params: { from, to (ms epoch), devices? }
function getReadingsCompare(params) {
  params = params || {};
  var from   = Number(params.from), to = Number(params.to);
  var wanted = params.devices ? String(params.devices).split(',') : null;
  if (!isFinite(from)) from = 0;
  if (!isFinite(to))   to   = Date.now();

  var windowMs  = to - from;
  var MIN_BUCKET = 5 * 60e3;
  var bucket    = Math.max(MIN_BUCKET, Math.ceil(windowMs / 200 / MIN_BUCKET) * MIN_BUCKET);

  var rows = readAllRows_().filter(function (r) {
    if (r.t < from || r.t > to) return false;
    if (wanted && wanted.indexOf(r.device) < 0) return false;
    return true;
  });

  var acc = {};
  rows.forEach(function (r) {
    var b    = Math.floor(r.t / bucket) * bucket;
    var dev  = acc[r.device] || (acc[r.device] = {});
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
        t:    s.t,
        temp: s.n ? +(s.tS / s.n).toFixed(2) : null,
        hum:  s.n ? +(s.hS / s.n).toFixed(1) : null
      };
    }).sort(function (a, b) { return a.t - b.t; });
  });

  return { from: from, to: to, bucketMs: bucket, series: series };
}

// ---------------------------------------------------------------
//  Battery / offline alerts by email
//  Set a time-driven trigger: Triggers -> Add Trigger ->
//  checkBatteries, Time-driven, Hour timer, every 3 hours.
// ---------------------------------------------------------------
function checkBatteries() {
  var info   = getDevices();
  var alerts = [];
  info.devices.forEach(function (d) {
    if (d.lowBatt) alerts.push('LOW BATTERY: ' + d.device + ' at ' + d.pct + '%');
    if (!d.online) alerts.push('OFFLINE: '     + d.device +
                    ' (no data for ' + Math.round(d.ageMs / 3600e3) + ' h)');
  });
  if (!alerts.length) return;

  var props = PropertiesService.getScriptProperties();
  var last  = props.getProperty('lastAlert') || '';
  var now   = alerts.join('\n');
  if (now === last) return;
  props.setProperty('lastAlert', now);

  MailApp.sendEmail({
    to:      Session.getEffectiveUser().getEmail(),
    subject: 'Dewvee alert (' + alerts.length + ')',
    body:    alerts.join('\n')
  });
}

// ---------------------------------------------------------------
//  Duplicate cleanup
// ---------------------------------------------------------------

/**
 * Scans a sheet and deletes any row whose (Timestamp, Device) pair already
 * appeared earlier, keeping the first occurrence. Returns rows removed.
 * Called automatically at the end of every doPost().
 */
function cleanupDuplicateRows(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return 0;

  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var seen = {}, rowsToDelete = [];
  for (var i = 0; i < data.length; i++) {
    var key = data[i][0] + '|' + data[i][1];
    var sheetRow = i + 2;
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
    Logger.log('Auto-cleanup removed %s duplicate row(s) from "%s".',
               rowsToDelete.length, sheet.getName());
  }
  return rowsToDelete.length;
}

/**
 * Manual entry point — runs cleanupDuplicateRows on every sheet.
 * Run from the Apps Script editor (function dropdown -> Run) whenever
 * you want to trigger a sweep outside of a normal upload.
 */
function removeDuplicateTimestamps() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var total = 0;
  ss.getSheets().forEach(function (sheet) {
    if (sheet.getLastRow() >= 3) {
      total += cleanupDuplicateRows(sheet);
    }
  });
  Logger.log('Removed %s duplicate row(s) across all sheets.', total);
}
