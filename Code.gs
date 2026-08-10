// ================================================================
//  DEWVEE  —  Google Apps Script backend
// ================================================================
//  Sheet layout (auto-created):
//    "<DeviceName>" — one sheet per device, all readings for that device
//    "Devices"      — one row per unique device: metadata + last seen
//
//  Properties (Project Properties → Script Properties):
//    UPLOAD_KEY     optional shared secret; if set, every POST must
//                   include a matching "key" field or be rejected.
//    EXPORT_EMAILS  JSON array of recipient email addresses
//    EXPORT_HOUR    hour (0-23, UTC) to send the daily export
//    EXPORT_STATUS  last run result written by sendDailyExport()
//    ALERT_SETTINGS JSON map  { "DEWVEE:01": { tempMin, tempMax,
//                               humMin, humMax } }
// ================================================================

var DATA_SHEET    = "Data";
var DEVICES_SHEET = "Devices";

// ── Column indices (1-based) in the Data sheet ─────────────────
var DC = { TS:1, DEVICE:2, TEMP:3, HUM:4, PCT:5, VOLT:6 };

// ── Column indices in the Devices sheet ────────────────────────
var DEV = { DEVICE:1, LOCATION:2, SAMPLE_MIN:3,
            FIRST_SEEN:4, LAST_SEEN:5, LAST_TEMP:6,
            LAST_HUM:7, LAST_PCT:8, UPLOAD_SEC:9 };

// ================================================================
//  Sheet helpers
// ================================================================
function getOrCreateSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (name === DEVICES_SHEET) {
      sh.appendRow(['Device', 'Location', 'SampleMin', 'FirstSeen',
                    'LastSeen', 'LastTemp', 'LastHum', 'LastPct', 'UploadSec']);
      sh.getRange(1, 1, 1, 9).setFontWeight('bold');
      sh.getRange('D2:E').setNumberFormat('yyyy-MM-dd HH:mm:ss');
    }
  }
  return sh;
}

// Returns the per-device data sheet, creating it with headers if needed.
function getOrCreateDeviceSheet_(deviceName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(deviceName);
  if (!sh) {
    sh = ss.insertSheet(deviceName);
    sh.appendRow(['Timestamp', 'Device', 'Temperature', 'Humidity',
                  'Battery%', 'Voltage']);
    sh.getRange(1, 1, 1, 6).setFontWeight('bold');
  }
  // Ensure the timestamp column always shows date + time, not just the date.
  // Applied on every call so existing sheets are fixed automatically.
  sh.getRange('A2:A').setNumberFormat('yyyy-MM-dd HH:mm:ss');
  return sh;
}

function getDeviceRow_(devSheet, deviceName) {
  var data = devSheet.getDataRange().getValues();
  for (var r = 1; r < data.length; r++) {
    if (data[r][DEV.DEVICE - 1] === deviceName) return r + 1; // 1-based sheet row
  }
  return -1;
}

// ================================================================
//  POST handler — receives readings from pods
// ================================================================
function doPost(e) {
  var out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);

  try {
    var body = JSON.parse(e.postData.contents);

    // Browser-generated report — route before sensor auth
    if (body.type === 'sendReport') {
      out.setContent(JSON.stringify(handleBrowserReport_(body)));
      return out;
    }

    var device    = String(body.device    || '').trim();
    var location  = String(body.location  || '').trim();
    var sampleSec = parseInt(body.sampleSec) || (parseInt(body.sampleMin) * 60) || 300;
    var uploadSec = parseInt(body.uploadSec) || 43200;
    var sampleMin = Math.round(sampleSec / 60) || 5;
    var rows      = body.rows;
    var incomingKey = String(body.key || '');

    if (!device)   { out.setContent(JSON.stringify({ok:false,error:'missing device'})); return out; }
    if (!rows || !rows.length) { out.setContent(JSON.stringify({ok:false,error:'no rows'})); return out; }

    // ── Authenticate ─────────────────────────────────────────────
    var storedKey = PropertiesService.getScriptProperties().getProperty('UPLOAD_KEY') || '';
    if (storedKey && storedKey !== incomingKey) {
      out.setContent(JSON.stringify({ok:false,error:'unauthorized'}));
      return out;
    }

    // ── Validate + sanitise individual readings ───────────────────
    var valid = [];
    rows.forEach(function(r) {
      var temp = parseFloat(r.temp);
      var hum  = parseFloat(r.hum);
      var pct  = parseInt(r.pct);
      var volt = parseFloat(r.v);
      var ts   = parseInt(r.t);
      if (isNaN(temp) || temp < -40 || temp > 85)  return; // out of sensor range
      if (isNaN(hum)  || hum  <   0 || hum  > 100) return;
      if (isNaN(ts)   || ts   < 1700000000)         return; // pre-2023 epoch
      if (isNaN(pct))  pct  = 0;
      if (isNaN(volt)) volt = 0;
      valid.push([new Date(ts * 1000), device,
                  Math.round(temp * 100) / 100,
                  Math.round(hum  * 10)  / 10,
                  pct, Math.round(volt * 1000) / 1000]);
    });

    if (valid.length === 0) {
      out.setContent(JSON.stringify({ok:false,error:'all rows failed validation'}));
      return out;
    }

    // ── Write to per-device sheet ─────────────────────────────────
    var devDataSheet = getOrCreateDeviceSheet_(device);
    devDataSheet.getRange(devDataSheet.getLastRow() + 1, 1, valid.length, 6)
                .setValues(valid);

    // ── Update / insert Devices sheet row ────────────────────────
    var devSheet = getOrCreateSheet_(DEVICES_SHEET);
    var lastRow  = valid[valid.length - 1];
    var devRow   = getDeviceRow_(devSheet, device);
    var now      = new Date();
    if (devRow < 0) {
      devSheet.appendRow([device, location, sampleMin, now, now,
                          lastRow[2], lastRow[3], lastRow[4], uploadSec]);
    } else {
      var existing = devSheet.getRange(devRow, 1, 1, 9).getValues()[0];
      devSheet.getRange(devRow, 1, 1, 9).setValues([[
        device,
        location || existing[DEV.LOCATION - 1],
        sampleMin,
        existing[DEV.FIRST_SEEN - 1] || now,
        now,
        lastRow[2], lastRow[3], lastRow[4], uploadSec
      ]]);
    }

    out.setContent(JSON.stringify({ok:true, written:valid.length}));
  } catch(err) {
    out.setContent(JSON.stringify({ok:false, error:err.message}));
  }
  return out;
}

// ================================================================
//  GET handler — serves dashboard data
// ================================================================
function doGet(e) {
  var p   = e.parameter;
  var out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);

  try {
    var action = p.action || '';
    var result;

    if      (action === 'devices')          result = getDevicesData_();
    else if (action === 'data')             result = getReadingsData_(p);
    else if (action === 'export')           result = getReadingsRange_(p);
    else if (action === 'compare')          result = getCompareData_(p);
    else if (action === 'getExportSettings')result = getExportSettings_();
    else if (action === 'setExportSettings')result = setExportSettings_(p);
    else if (action === 'getAlertSettings') result = getAlertSettings_();
    else if (action === 'setAlertSettings') result = setAlertSettings_(p);
    else                                    result = {error:'unknown action'};

    out.setContent(JSON.stringify(result));
  } catch(err) {
    out.setContent(JSON.stringify({error:err.message}));
  }
  return out;
}

// ================================================================
//  Data query helpers
// ================================================================
function parseRangeMs_(range) {
  var ms = { '1h':3600e3, '6h':6*3600e3, '12h':12*3600e3, 'day':86400e3,
             'week':7*86400e3, 'month':30*86400e3, 'year':365*86400e3 };
  return ms[range] || 86400e3;
}

// Battery voltage constants (LiPo)
var BATT_CUTOFF_V = 3.20;  // pod shuts off below this (matches firmware BATT_DEAD_V)
var BATT_FULL_V   = 4.32;  // resting full-charge voltage for this battery

// Derive battery % from voltage with 0.1% precision (3-decimal voltage → fine-grained %)
function voltToPct_(volt) {
  if (!volt || volt <= 0) return null;
  return Math.round(Math.max(0, Math.min(100,
    (volt - BATT_CUTOFF_V) / (BATT_FULL_V - BATT_CUTOFF_V) * 100
  )) * 10) / 10;
}

// Scans ALL sheets (skipping the Devices metadata sheet) and returns merged
// voltage map, battery readings history, and latest-reading snapshot per device.
// Handles both the old per-device sheet format ("DEWVEE:01" etc. with multi-line
// headers) and the new unified "Data" sheet format.
function scanAllSheets_(opts) {
  // opts: { fromMs, toMs, devices }  — optional filter for series mode
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var sheets  = ss.getSheets();
  var devSet  = (opts && opts.devices && opts.devices.length) ? opts.devices : null;
  var fromMs  = (opts && opts.fromMs) || 0;
  var toMs    = (opts && opts.toMs)   || Infinity;
  var seriesMode = !!(opts && opts.fromMs);

  var voltMap      = {};
  var readingsMap  = {};
  var readingsSeen = {};  // dedup: {dev: {tsKey: true}}
  var dataSnap     = {};
  var series       = {};  // used in seriesMode
  var seriesSeen   = {};  // dedup: {dev: {tsKey: true}}

  sheets.forEach(function(sh) {
    if (sh.getName() === DEVICES_SHEET) return;
    var rows = sh.getLastRow();
    if (rows < 2) return;

    var allD = sh.getDataRange().getValues();
    // Normalise headers: lower-case, collapse whitespace/newlines
    var headers = allD[0].map(function(h) {
      return String(h).toLowerCase().replace(/[\n\r]+/g, ' ').replace(/\s+/g, ' ').trim();
    });

    // Detect columns by header content (works for old AND new header naming)
    var colTs = -1, colDev = -1, colTemp = -1, colHum = -1, colPct = -1, colVolt = -1;
    headers.forEach(function(h, i) {
      if (colTs   === -1 && (h === 'timestamp' || h.indexOf('time') !== -1)) colTs   = i;
      if (colDev  === -1 && h.indexOf('device') !== -1) colDev  = i;
      if (colTemp === -1 && h.indexOf('temp') !== -1)   colTemp = i;
      if (colHum  === -1 && (h.indexOf('humid') !== -1 || h.indexOf('relative') !== -1)) colHum = i;
      // Battery % — matches "battery (%)" / "battery%" but NOT "battery voltage"
      if (colPct  === -1 && h.indexOf('batt') !== -1 && h.indexOf('volt') === -1) colPct  = i;
      if (colVolt === -1 && h.indexOf('volt') !== -1) colVolt = i;
    });

    if (colTs === -1 || colTemp === -1) return; // not a sensor data sheet

    var sheetName = sh.getName();

    for (var ri = 1; ri < allD.length; ri++) {
      var row  = allD[ri];
      var rawTs = row[colTs];
      var ts   = rawTs instanceof Date ? rawTs.getTime() : new Date(rawTs).getTime();
      if (!ts || isNaN(ts) || ts <= 0) continue;

      var dev = colDev !== -1 ? String(row[colDev]).trim() : sheetName;
      if (!dev || dev === 'undefined' || dev === '') dev = sheetName;

      // Apply device filter when set
      if (devSet && devSet.indexOf(dev) === -1) continue;

      var temp = parseFloat(row[colTemp]) || 0;
      var hum  = colHum  !== -1 ? parseFloat(row[colHum])  || 0 : 0;
      var pct  = colPct  !== -1 ? parseInt(row[colPct])         : 0;
      var volt = colVolt !== -1 ? parseFloat(row[colVolt]) || 0 : 0;

      // Track readings for battery projection — include voltage for high-precision regression
      voltMap[dev] = volt;

      // Track any row with a valid voltage — voltage regression doesn't need valid pct
      if (volt > 0) {
        if (!readingsMap[dev]) { readingsMap[dev] = []; readingsSeen[dev] = {}; }
        var rKey = String(ts);
        if (!readingsSeen[dev][rKey]) {
          readingsSeen[dev][rKey] = true;
          readingsMap[dev].push({ ts: ts, pct: (!isNaN(pct) && pct >= 0 && pct <= 100) ? pct : 0, volt: volt });
        }
      }

      // Derive pct from voltage (0.1% precision) rather than the ESP32's integer pct
      var derivedPct = voltToPct_(volt);
      if (derivedPct === null) derivedPct = isNaN(pct) ? 0 : pct;
      if (!dataSnap[dev] || ts > dataSnap[dev].lastTs) {
        dataSnap[dev] = { lastTs: ts, lastTemp: temp, lastHum: hum,
                          lastPct: derivedPct, lastVolt: volt };
      }

      // Series data (for charts) — deduplicate by exact timestamp
      if (seriesMode && ts >= fromMs && ts <= toMs) {
        if (!series[dev]) { series[dev] = []; seriesSeen[dev] = {}; }
        var tsKey = String(ts);
        if (!seriesSeen[dev][tsKey]) {
          seriesSeen[dev][tsKey] = true;
          series[dev].push({ t: Math.round(ts - fromMs), temp: temp, hum: hum,
                             pct: isNaN(pct) ? 0 : pct, v: volt });
        }
      }
    }
  });

  // Sort series and battery readings ascending (Chart.js normalized:true requires sorted data)
  for (var d in series)      series[d].sort(function(a, b) { return a.t - b.t; });
  for (var d in readingsMap) readingsMap[d].sort(function(a, b) { return a.ts - b.ts; });

  return { voltMap: voltMap, readingsMap: readingsMap, dataSnap: dataSnap, series: series };
}

function computeBatteryProjection_(readings) {
  // readings: [{ts:ms, pct:number, volt:number}] sorted ascending
  var noDataBase = { note: 'insufficient_data', dataPoints: 0, spanMs: 0 };
  if (!readings || readings.length < 4) return Object.assign({}, noDataBase, { dataPoints: readings ? readings.length : 0 });

  // Find last recharge: voltage rose > 0.05V (or pct rose > 5pp if no voltage)
  var startIdx = 0;
  for (var i = readings.length - 1; i > 0; i--) {
    var prev = readings[i - 1], cur = readings[i];
    var rose = (cur.volt > 0 && prev.volt > 0)
      ? (cur.volt - prev.volt) > 0.05
      : (cur.pct  - prev.pct)  > 5;
    if (rose) { startIdx = i; break; }
  }
  var pts = readings.slice(startIdx);
  if (pts.length < 4) return Object.assign({}, noDataBase, { dataPoints: pts.length });

  var spanMs = pts[pts.length - 1].ts - pts[0].ts;

  // Adaptive minimum span: more readings = shorter wall-clock time needed.
  // A device sampling every 10 s accumulates 360 pts/hr — reliable regression in 1 h.
  // A device sampling every 12 hrs needs weeks of data before the slope is trustworthy.
  var minSpanMs = pts.length >= 120 ? 3600000       // ≥120 pts → 1 h minimum
               : pts.length >= 30  ? 3 * 3600000    // ≥30 pts  → 3 h minimum
               :                     6 * 3600000;   // otherwise → 6 h minimum

  if (spanMs < minSpanMs) return { note: 'insufficient_data', dataPoints: pts.length, spanMs: spanMs };

  // Prefer voltage regression (3-decimal precision) over integer pct
  var hasVolt = pts.filter(function(p) { return p.volt > 0; }).length > pts.length / 2;

  // Linear regression: x = days elapsed, y = voltage (or pct)
  var t0 = pts[0].ts;
  var n = pts.length, sx=0, sy=0, sxy=0, sxx=0;
  for (var j = 0; j < n; j++) {
    var x = (pts[j].ts - t0) / 86400000;  // days
    var y = hasVolt ? pts[j].volt : pts[j].pct;
    sx += x; sy += y; sxy += x*y; sxx += x*x;
  }
  var mx = sx/n, my = sy/n;
  var ssxy = sxy - n*mx*my, ssxx = sxx - n*mx*mx;
  if (ssxx < 1e-9) return { note: 'insufficient_data', dataPoints: n, spanMs: spanMs };

  var slope = ssxy / ssxx;    // V/day or %/day — negative = draining
  if (slope >= 0) return { note: 'insufficient_data', dataPoints: n, spanMs: spanMs };

  var drainPerDay = -slope;   // positive value

  var lastPt      = pts[pts.length - 1];
  var currentVolt = lastPt.volt;
  var currentPct  = hasVolt ? voltToPct_(currentVolt) : lastPt.pct;

  // Days = usable energy remaining / daily drain
  var daysLeft;
  if (hasVolt && currentVolt > 0) {
    daysLeft = (currentVolt - BATT_CUTOFF_V) / drainPerDay;
  } else {
    daysLeft = currentPct / drainPerDay;
  }

  if (daysLeft <= 0) return { note: 'insufficient_data', dataPoints: n, spanMs: spanMs };

  // R²
  var intercept = my - slope * mx;
  var ssRes = 0, ssTot = 0;
  for (var k = 0; k < n; k++) {
    var xk = (pts[k].ts - t0) / 86400000;
    var yk = hasVolt ? pts[k].volt : pts[k].pct;
    var pred = slope * xk + intercept;
    ssRes += Math.pow(yk - pred, 2);
    ssTot += Math.pow(yk - my,   2);
  }
  var r2 = ssTot > 1e-9 ? 1 - ssRes / ssTot : 0;

  // Convert drain to %/day regardless of which unit was used for regression
  var drainPctPerDay = hasVolt
    ? drainPerDay / (BATT_FULL_V - BATT_CUTOFF_V) * 100
    : drainPerDay;

  // mAh/day — 2400 mAh usable
  var drainMahPerDay = drainPctPerDay / 100 * 2400;

  var confidence = r2 >= 0.90 ? 'high' : r2 >= 0.65 ? 'medium' : 'low';

  return {
    daysLeft:       Math.round(daysLeft       * 10)  / 10,
    drainPctPerDay: Math.round(drainPctPerDay * 100) / 100,
    drainMahPerDay: Math.round(drainMahPerDay * 10)  / 10,
    r2:             Math.round(r2             * 1000) / 1000,
    dataPoints:     n,
    spanDays:       Math.round(spanMs / 86400000 * 10) / 10,
    deadByMs:       Math.round(lastPt.ts + daysLeft * 86400000),
    confidence:     confidence,
    currentVolt:    Math.round(currentVolt * 1000) / 1000
  };
}

function getDevicesData_() {
  var devSheet = getOrCreateSheet_(DEVICES_SHEET);
  var devData  = devSheet.getDataRange().getValues();
  var now      = Date.now();
  var result   = [];

  // Scan ALL sheets for sensor data (handles old per-device sheets + new "Data" sheet)
  var scan = scanAllSheets_({});
  var voltMap    = scan.voltMap;
  var readingsMap = scan.readingsMap;
  var dataSnap   = scan.dataSnap;

  // Primary: Devices metadata sheet (populated by doPost on new uploads)
  var seenDevs = {};
  for (var r = 1; r < devData.length; r++) {
    var row       = devData[r];
    var device    = String(row[DEV.DEVICE    - 1] || '');
    var location  = row[DEV.LOCATION  - 1];
    var sampleMin = parseInt(row[DEV.SAMPLE_MIN - 1]) || 10;
    var uploadSec = parseInt(row[DEV.UPLOAD_SEC - 1]) || 43200;
    var lastSeen  = row[DEV.LAST_SEEN - 1];
    var lastTemp  = row[DEV.LAST_TEMP - 1];
    var lastHum   = row[DEV.LAST_HUM  - 1];
    var lastPct   = row[DEV.LAST_PCT  - 1];
    if (!device) continue;

    seenDevs[device] = true;
    var tsMs = lastSeen ? new Date(lastSeen).getTime() : 0;
    var snap  = dataSnap[device];
    var derivedPct = (snap && snap.lastPct !== null) ? snap.lastPct : (parseInt(lastPct) || 0);
    result.push({
      device:    device,
      location:  location || '',
      sampleMin: sampleMin,
      sampleSec: sampleMin * 60,
      uploadSec: uploadSec,
      temp:      parseFloat(lastTemp)  || 0,
      hum:       parseFloat(lastHum)   || 0,
      pct:       derivedPct,
      volt:      voltMap[device]       || 0,
      ts:        Math.floor(tsMs / 1000),
      lowBatt:   derivedPct <= 20,
      ageMs:     tsMs > 0 ? now - tsMs : Infinity,
      online:    tsMs > 0 && (now - tsMs) < 86400000,
      battProj:  computeBatteryProjection_(readingsMap[device] || null)
    });
  }

  // Fallback: any device found in sensor sheets but not yet in Devices sheet
  // (covers old per-device sheets and fresh deployments)
  for (var dev in dataSnap) {
    if (seenDevs[dev]) continue;
    var snap      = dataSnap[dev];
    var tsMs      = snap.lastTs || 0;
    result.push({
      device:    dev,
      location:  '',
      sampleMin: 10,
      sampleSec: 600,
      uploadSec: 43200,
      temp:      snap.lastTemp,
      hum:       snap.lastHum,
      pct:       snap.lastPct,
      volt:      snap.lastVolt,
      ts:        Math.floor(tsMs / 1000),
      lowBatt:   snap.lastPct <= 20,
      ageMs:     tsMs > 0 ? now - tsMs : Infinity,
      online:    tsMs > 0 && (now - tsMs) < 86400000,
      battProj:  computeBatteryProjection_(readingsMap[dev] || null)
    });
  }

  result.sort(function(a, b) {
    if (a.online !== b.online) return a.online ? -1 : 1;
    return String(a.device).localeCompare(String(b.device));
  });

  return { devices: result };
}


function getReadingsData_(p) {
  var range      = p.range      || 'day';
  var dayOffset  = parseInt(p.dayOffset) || 0;
  var devices    = p.devices ? p.devices.split(',') : [];
  var nowMs      = Date.now();
  var shiftMs    = dayOffset * 86400000;
  var toMs       = nowMs - shiftMs;
  var fromMs     = (range === 'all') ? new Date('2024-01-01').getTime() : toMs - parseRangeMs_(range);
  return buildSeriesFromSheet_(devices, fromMs, toMs, range);
}

function getReadingsRange_(p) {
  var devices = p.devices ? p.devices.split(',') : [];
  var fromMs  = parseInt(p.from) || (Date.now() - 86400e3);
  var toMs    = parseInt(p.to)   || Date.now();
  var scan    = scanAllSheets_({ fromMs: fromMs, toMs: toMs, devices: devices.length ? devices : null });
  var rows    = [];
  for (var dev in scan.series) {
    scan.series[dev].forEach(function(pt) {
      rows.push({ t: fromMs + pt.t, device: dev,
                  temp: pt.temp, hum: pt.hum, pct: pt.pct, volt: pt.v });
    });
  }
  rows.sort(function(a, b) { return a.t - b.t; });
  return { rows: rows };
}

function getCompareData_(p) {
  var devices = p.devices ? p.devices.split(',') : [];
  var fromMs  = parseInt(p.from) || (Date.now() - 86400e3);
  var toMs    = parseInt(p.to)   || Date.now();
  return buildSeriesFromSheet_(devices, fromMs, toMs, 'compare');
}

function buildSeriesFromSheet_(devices, fromMs, toMs, range) {
  var scan = scanAllSheets_({ fromMs: fromMs, toMs: toMs, devices: devices.length ? devices : null });
  return { series: scan.series, from: fromMs, to: toMs, range: range };
}

// ================================================================
//  Export settings
// ================================================================
function getExportSettings_() {
  var p = PropertiesService.getScriptProperties();
  var emailProp = p.getProperty('EXPORT_EMAILS') || '[]';
  var emails;
  try { emails = JSON.parse(emailProp); } catch(e) { emails = []; }
  // backward compat: legacy single-email property
  if (!emails.length) {
    var legacy = p.getProperty('EXPORT_EMAIL');
    if (legacy) emails = [legacy];
  }
  return {
    emails: emails,
    hour:   parseInt(p.getProperty('EXPORT_HOUR') || '8'),
    status: p.getProperty('EXPORT_STATUS') || ''
  };
}

function setExportSettings_(p) {
  var data = p.data ? JSON.parse(decodeURIComponent(p.data)) : {};
  var sp   = PropertiesService.getScriptProperties();
  if (data.emails !== undefined) sp.setProperty('EXPORT_EMAILS', JSON.stringify(data.emails));
  if (data.hour   !== undefined) sp.setProperty('EXPORT_HOUR',   String(parseInt(data.hour)||8));
  setupDailyTrigger();
  return { ok: true };
}

// ================================================================
//  Alert settings
// ================================================================
function getAlertSettings_() {
  var raw = PropertiesService.getScriptProperties().getProperty('ALERT_SETTINGS') || '{}';
  try { return JSON.parse(raw); } catch(e) { return {}; }
}

function setAlertSettings_(p) {
  var data = p.data ? JSON.parse(decodeURIComponent(p.data)) : {};
  PropertiesService.getScriptProperties().setProperty('ALERT_SETTINGS', JSON.stringify(data));
  return { ok: true };
}

// ================================================================
//  Daily export trigger
// ================================================================
function setupDailyTrigger() {
  // Remove existing export triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sendDailyExport') ScriptApp.deleteTrigger(t);
  });
  var sp   = PropertiesService.getScriptProperties();
  var hour = parseInt(sp.getProperty('EXPORT_HOUR') || '8');
  ScriptApp.newTrigger('sendDailyExport').timeBased().atHour(hour).everyDays(1).create();
}

function sendDailyExport() {
  var sp = PropertiesService.getScriptProperties();
  try {
    var emailProp = sp.getProperty('EXPORT_EMAILS') || '[]';
    var emails;
    try { emails = JSON.parse(emailProp); } catch(e) { emails = []; }
    if (!emails.length) {
      var legacy = sp.getProperty('EXPORT_EMAIL');
      if (legacy) emails = [legacy];
    }
    if (!emails.length) { sp.setProperty('EXPORT_STATUS','No recipients configured'); return; }
    sendWeeklyReport_(emails.join(','));
    sp.setProperty('EXPORT_STATUS', 'OK:' + new Date().toISOString());
  } catch(err) {
    sp.setProperty('EXPORT_STATUS', 'ERR:' + err.message + ':' + new Date().toISOString());
    Logger.log('sendDailyExport failed: ' + err.message);
  }
}

// ================================================================
//  Rich Weekly Report — charts via QuickChart.io + HTML email
// ================================================================
var REPORT_COLORS_ = ['#a78bfa','#60a5fa','#34d399','#fbbf24','#f87171'];

function dewPoint_(t, h) {
  var a = 17.27, b = 237.7;
  var g = Math.log(Math.max(h, 1) / 100) + a * t / (b + t);
  return Math.round(b * g / (a - g) * 10) / 10;
}

function hexToRgb_(hex) {
  return parseInt(hex.slice(1,3),16)+','+parseInt(hex.slice(3,5),16)+','+parseInt(hex.slice(5,7),16);
}

function weekStats_(pts, snap, readingsMap) {
  if (!pts || !pts.length) return null;
  var temps = [], hums = [];
  pts.forEach(function(p){ if(!isNaN(p.temp)&&p.temp>0) temps.push(p.temp); if(!isNaN(p.hum)&&p.hum>0) hums.push(p.hum); });
  if (!temps.length) return null;
  var tAvg = Math.round(temps.reduce(function(a,b){return a+b;},0)/temps.length*10)/10;
  var hAvg = Math.round(hums.reduce(function(a,b){return a+b;},0)/hums.length*10)/10;
  var proj = computeBatteryProjection_(readingsMap || null);
  return {
    tMin: Math.round(Math.min.apply(null,temps)*10)/10,
    tMax: Math.round(Math.max.apply(null,temps)*10)/10,
    tAvg: tAvg,
    hMin: Math.round(Math.min.apply(null,hums)*10)/10,
    hMax: Math.round(Math.max.apply(null,hums)*10)/10,
    hAvg: hAvg,
    dAvg: dewPoint_(tAvg, hAvg),
    battDays: proj && proj.daysLeft ? Math.round(proj.daysLeft) : null,
    battPct:  snap ? Math.round(snap.lastPct || 0) : null
  };
}

// Daily averages — used for the all-device overview charts
function dailyAvg_(pts, fromMs) {
  var days = {};
  var DAY_NAMES = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  pts.forEach(function(pt) {
    var idx = Math.floor(pt.t / 86400000);
    if (!days[idx]) {
      var d = new Date(fromMs + pt.t);
      days[idx] = { idx:idx, label: DAY_NAMES[d.getDay()], t:[], h:[] };
    }
    if (!isNaN(pt.temp)&&pt.temp>0) days[idx].t.push(pt.temp);
    if (!isNaN(pt.hum) &&pt.hum>0)  days[idx].h.push(pt.hum);
  });
  return Object.keys(days).sort(function(a,b){return a-b;}).map(function(k){
    var d = days[k];
    var avgT = d.t.length ? d.t.reduce(function(a,b){return a+b;},0)/d.t.length : 0;
    var avgH = d.h.length ? d.h.reduce(function(a,b){return a+b;},0)/d.h.length : 0;
    return {
      label: d.label,
      avgT: Math.round(avgT*10)/10, avgH: Math.round(avgH*10)/10,
      minT: d.t.length?Math.round(Math.min.apply(null,d.t)*10)/10:null,
      maxT: d.t.length?Math.round(Math.max.apply(null,d.t)*10)/10:null
    };
  });
}

// Hourly averages — used for per-device detail charts
function hourlyAvg_(pts, fromMs) {
  var hours = {};
  var DAY_NAMES = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  pts.forEach(function(pt) {
    var idx = Math.floor(pt.t / 3600000);
    if (!hours[idx]) {
      var d = new Date(fromMs + pt.t);
      hours[idx] = { idx:idx, label: DAY_NAMES[d.getDay()]+' '+d.getHours()+'h', t:[], h:[] };
    }
    if (!isNaN(pt.temp)&&pt.temp>0) hours[idx].t.push(pt.temp);
    if (!isNaN(pt.hum) &&pt.hum>0)  hours[idx].h.push(pt.hum);
  });
  return Object.keys(hours).sort(function(a,b){return a-b;}).map(function(k){
    var h = hours[k];
    var avgT = h.t.length ? h.t.reduce(function(a,b){return a+b;},0)/h.t.length : 0;
    var avgH = h.h.length ? h.h.reduce(function(a,b){return a+b;},0)/h.h.length : 0;
    return { idx:parseInt(k), label:h.label,
             avgT:Math.round(avgT*10)/10, avgH:Math.round(avgH*10)/10 };
  });
}

function fetchQuickChart_(chartCfg, width, height) {
  try {
    var resp = UrlFetchApp.fetch('https://quickchart.io/chart', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ width:width||600, height:height||200,
                                backgroundColor:'#0f1117', format:'png', chart:chartCfg }),
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) {
      Logger.log('QuickChart ' + resp.getResponseCode() + ': ' + resp.getContentText().substring(0,300));
      return null;
    }
    return resp.getBlob().setContentType('image/png');
  } catch(e) { Logger.log('QuickChart error: ' + e.message); return null; }
}

// Chart.js v2 config for the all-device overview (daily averages)
function overviewChartCfg_(dailyByDev, devOrder, devColors, field) {
  var labels = [];
  if (devOrder.length) labels = (dailyByDev[devOrder[0]]||[]).map(function(d){return d.label;});
  var datasets = devOrder.map(function(dev,i){
    var col = devColors[i], rgb = hexToRgb_(col);
    return { label:dev, data:(dailyByDev[dev]||[]).map(function(d){return d[field];}),
             borderColor:col, backgroundColor:'rgba('+rgb+',0.08)', fill:true,
             borderWidth:2.5, tension:0.4, pointRadius:5,
             pointBackgroundColor:col, pointBorderColor:'#0f1117', pointBorderWidth:1.5 };
  });
  return { type:'line', data:{ labels:labels, datasets:datasets },
    options:{ scales:{
      yAxes:[{ ticks:{ fontColor:'#9ca3af', fontSize:10 },
               gridLines:{ color:'rgba(255,255,255,0.07)', zeroLineColor:'rgba(255,255,255,0.1)' } }],
      xAxes:[{ ticks:{ fontColor:'#9ca3af', fontSize:10 },
               gridLines:{ color:'rgba(255,255,255,0.07)' } }] },
    legend:{ display:true, labels:{ fontColor:'#d1d5db', fontSize:11, padding:16 } },
    layout:{ padding:{ left:8, right:8, top:10, bottom:8 } } } };
}

// Chart.js v2 config for an individual device (hourly, dual y-axis)
// dewAvg: dew point °C to draw as dashed reference line on yT
// t/hHighIdx, t/hLowIdx: hourlyData indices where the extreme values occur
function deviceChartCfg_(hourlyData, dewAvg, tHighIdx, tLowIdx, hHighIdx, hLowIdx) {
  var labels = hourlyData.map(function(h){ return h.label; });
  var tData  = hourlyData.map(function(h){ return h.avgT; });
  var hData  = hourlyData.map(function(h){ return h.avgH; });
  var n      = hourlyData.length;

  // Single-point null arrays for markers (Chart.js renders only non-null points)
  function mkr(arr, idx) { return arr.map(function(v,i){ return i===idx ? v : null; }); }

  var tHigh = tData[tHighIdx], tLow = tData[tLowIdx];
  var hHigh = hData[hHighIdx], hLow = hData[hLowIdx];

  return { type:'line', data:{ labels:labels, datasets:[
    // main temp line
    { label:'Temp (°C)', data:tData,
      borderColor:'#f0a54e', backgroundColor:'rgba(240,165,78,0.08)',
      fill:true, borderWidth:2, tension:0.4, pointRadius:0, yAxisID:'yT' },
    // main humidity line
    { label:'Humidity (%)', data:hData,
      borderColor:'#5b8cff', backgroundColor:'transparent',
      fill:false, borderWidth:1.8, tension:0.4, pointRadius:0, yAxisID:'yH' },
    // dew point reference line — dashed green on yT
    { label:'Dew '+dewAvg+'°C', data:Array(n).fill(dewAvg),
      borderColor:'#3ecf8e', borderDash:[6,4], borderWidth:1.5,
      fill:false, pointRadius:0, tension:0, yAxisID:'yT' },
    // temp HIGH marker — orange dot
    { label:'T▲ '+tHigh+'°', data:mkr(tData, tHighIdx),
      borderColor:'#f0a54e', backgroundColor:'#f0a54e',
      showLine:false, fill:false, pointRadius:7, yAxisID:'yT' },
    // temp LOW marker — blue dot
    { label:'T▼ '+tLow+'°', data:mkr(tData, tLowIdx),
      borderColor:'#60a5fa', backgroundColor:'#60a5fa',
      showLine:false, fill:false, pointRadius:7, yAxisID:'yT' },
    // hum HIGH marker — sky dot
    { label:'H▲ '+hHigh+'%', data:mkr(hData, hHighIdx),
      borderColor:'#38bdf8', backgroundColor:'#38bdf8',
      showLine:false, fill:false, pointRadius:7, yAxisID:'yH' },
    // hum LOW marker — indigo dot
    { label:'H▼ '+hLow+'%', data:mkr(hData, hLowIdx),
      borderColor:'#818cf8', backgroundColor:'#818cf8',
      showLine:false, fill:false, pointRadius:7, yAxisID:'yH' }
  ]}, options:{ scales:{
    yAxes:[
      { id:'yT', position:'left',
        ticks:{ fontColor:'rgba(240,165,78,0.85)', fontSize:10 },
        gridLines:{ color:'rgba(255,255,255,0.06)', zeroLineColor:'rgba(255,255,255,0.08)' } },
      { id:'yH', position:'right',
        ticks:{ fontColor:'rgba(91,140,255,0.75)', fontSize:10 },
        gridLines:{ drawOnChartArea:false } }
    ],
    xAxes:[{ ticks:{ fontColor:'#4b5563', fontSize:9, maxTicksLimit:7, autoSkip:true },
             gridLines:{ color:'rgba(255,255,255,0.05)' } }] },
  legend:{ display:true, labels:{ fontColor:'#d1d5db', fontSize:10, boxWidth:10 } },
  layout:{ padding:{ left:6, right:6, top:10, bottom:6 } } } };
}

function buildDeviceSummary_(dev, s) {
  var humWord = s.hAvg > 82 ? 'persistently high' : s.hAvg > 74 ? 'elevated' : 'moderate';
  var hSpike  = s.hMax >= 88 ? ' Spiked to ' + s.hMax.toFixed(0) + '% — ventilation advised.' : '';
  return 'Avg temp <b>' + s.tAvg.toFixed(1) + '°C</b> (range ' + s.tMin.toFixed(1) + '–' + s.tMax.toFixed(1) + '°C). ' +
         'Humidity ' + humWord + ' at ' + s.hAvg.toFixed(0) + '% avg.' + hSpike;
}

function buildSummaryText_(statsMap, devOrder) {
  var hiDev = '', hiTemp = -99, loDev = '', loTemp = 99, humWarnings = [];
  devOrder.forEach(function(dev){
    var s = statsMap[dev]; if(!s) return;
    if (s.tMax > hiTemp){ hiTemp = s.tMax; hiDev = dev; }
    if (s.tMin < loTemp){ loTemp = s.tMin; loDev = dev; }
    if (s.hMax >= 85) humWarnings.push(dev + ' (' + s.hMax.toFixed(0) + '%)');
  });
  var avgTemps = devOrder.map(function(d){ return statsMap[d]?statsMap[d].tAvg:null; }).filter(Boolean);
  var grandAvg = avgTemps.length ? Math.round(avgTemps.reduce(function(a,b){return a+b;},0)/avgTemps.length*10)/10 : 0;
  var txt = 'Indoor temperatures averaged <b>' + grandAvg + '°C across all rooms</b> this week. ';
  txt += 'The week\'s peak was <b>' + hiTemp.toFixed(1) + '°C on ' + hiDev + '</b>, lowest was <b>' + loTemp.toFixed(1) + '°C</b>. ';
  txt += humWarnings.length
    ? 'Humidity exceeded 85% on <b>' + humWarnings.join(', ') + '</b> — ventilation is recommended. '
    : 'Humidity stayed within manageable levels across all sensors. ';
  txt += 'All sensors reported successfully. See the attached CSV for full raw data.';
  return txt;
}

function buildReportHTML_(series, statsMap, fromMs, imageIds, devOrder, devColors) {
  var dateStr = Utilities.formatDate(new Date(fromMs), 'UTC', 'MMM d') +
                ' &ndash; ' + Utilities.formatDate(new Date(), 'UTC', 'MMM d, yyyy');
  var W = 600;

  function td(style, content){ return '<td style="' + style + '">' + content + '</td>'; }
  function row(tds){ return '<tr>' + tds + '</tr>'; }
  function lbl(txt, col){ return '<p style="margin:0 0 8px;font-size:10px;font-weight:700;color:' + col + ';text-transform:uppercase;letter-spacing:0.09em;">' + txt + '</p>'; }
  function img(cid, w){ return '<img src="cid:' + cid + '" width="' + w + '" style="display:block;width:100%;max-width:' + w + 'px;border-radius:6px;" />'; }

  var h = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>';
  h += '<body style="margin:0;padding:20px 0;background:#06070a;font-family:Arial,Helvetica,sans-serif;">';
  h += '<table width="' + W + '" cellpadding="0" cellspacing="0" align="center" style="background:#0f1117;">';

  // Header
  h += row(td('padding:22px 28px;background:#161924;border-bottom:1px solid rgba(255,255,255,0.07);',
    '<table width="100%" cellpadding="0" cellspacing="0"><tr>' +
    td('', '<span style="font-size:24px;font-weight:900;color:#f2f4f7;letter-spacing:0.15em;">DEWVEE</span><br>' +
           '<span style="font-size:10px;color:#4b5563;letter-spacing:0.1em;font-weight:700;">WEEKLY CLIMATE REPORT</span>') +
    td('text-align:right;', '<span style="font-size:14px;font-weight:600;color:#d1d5db;">' + dateStr + '</span><br>' +
           '<span style="font-size:10px;color:#374151;">' + devOrder.length + ' sensors active</span>') +
    '</tr></table>'));

  // Temp overview chart
  h += row(td('padding:16px 20px 4px;',
    lbl('Temperature &mdash; All Devices (Daily Average)', '#f0a54e') +
    (imageIds.chartTemp ? img(imageIds.chartTemp, W-40) : '<p style="color:#4b5563;font-size:11px;">Chart unavailable</p>')));

  // Hum overview chart
  h += row(td('padding:4px 20px 16px;',
    lbl('Relative Humidity &mdash; All Devices (Daily Average)', '#5b8cff') +
    (imageIds.chartHum ? img(imageIds.chartHum, W-40) : '<p style="color:#4b5563;font-size:11px;">Chart unavailable</p>')));

  // Summary block
  h += row(td('padding:0 20px 16px;',
    '<div style="background:rgba(91,140,255,0.06);border:1px solid rgba(91,140,255,0.15);border-radius:8px;padding:14px 16px;">' +
    '<p style="margin:0 0 6px;font-size:10px;font-weight:700;color:#5b8cff;text-transform:uppercase;letter-spacing:0.1em;">&#128202; Week Summary</p>' +
    '<p style="margin:0;font-size:12px;color:#9ca3af;line-height:1.8;">' + buildSummaryText_(statsMap, devOrder) + '</p>' +
    '</div>'));

  // Section divider
  h += row(td('padding:0 20px;', '<hr style="border:none;border-top:1px solid rgba(255,255,255,0.07);margin:0;" />'));
  h += row(td('padding:10px 20px 4px;font-size:10px;font-weight:700;color:#4b5563;text-transform:uppercase;letter-spacing:0.09em;', 'Individual Device Reports'));

  // Per-device sections
  devOrder.forEach(function(dev, i) {
    var s = statsMap[dev]; if (!s) return;
    var col = devColors[i] || '#ffffff';
    var cid = imageIds['dev_' + i];
    var battStr = s.battDays ? '~' + s.battDays + 'd remaining' : (s.battPct != null ? s.battPct + '%' : 'N/A');
    var battCol = (s.battDays||99) > 45 ? '#4ec46a' : (s.battDays||99) > 20 ? '#e8a020' : '#e05050';

    h += row(td('padding:4px 20px;',
      '<table width="100%" cellpadding="0" cellspacing="0" style="background:#151820;border:1px solid rgba(255,255,255,0.08);border-radius:8px;">' +
      // device header
      row(td('padding:10px 14px;border-bottom:1px solid rgba(255,255,255,0.06);',
        '<span style="display:inline-block;width:9px;height:9px;border-radius:50%;background:' + col + ';margin-right:8px;vertical-align:middle;"></span>' +
        '<span style="font-size:13px;font-weight:700;color:#e5e7eb;">' + dev + '</span>')) +
      // chart
      row(td('padding:10px 14px 4px;',
        cid ? img(cid, W-68) : '<p style="font-size:11px;color:#4b5563;margin:0;">Chart unavailable</p>')) +
      // stats row
      row(td('padding:4px 14px 6px;',
        '<table width="100%" cellpadding="4" cellspacing="0">' +
        // temp stats
        '<tr><td style="font-size:9px;color:#4b5563;font-weight:700;text-transform:uppercase;width:80px;">Temperature</td>' +
        '<td style="text-align:center;background:rgba(255,255,255,0.04);border-radius:4px;"><span style="display:block;font-size:12px;font-weight:700;color:#e5e7eb;">' + s.tMin.toFixed(1) + '°</span><span style="font-size:8px;color:#6b7280;">LOW</span></td>' +
        '<td style="text-align:center;background:rgba(255,255,255,0.06);border-radius:4px;"><span style="display:block;font-size:13px;font-weight:700;color:#f0a54e;">' + s.tAvg.toFixed(1) + '°</span><span style="font-size:8px;color:#6b7280;">AVG</span></td>' +
        '<td style="text-align:center;background:rgba(255,255,255,0.04);border-radius:4px;"><span style="display:block;font-size:12px;font-weight:700;color:#e5e7eb;">' + s.tMax.toFixed(1) + '°</span><span style="font-size:8px;color:#6b7280;">HIGH</span></td></tr>' +
        // hum stats
        '<tr><td style="font-size:9px;color:#4b5563;font-weight:700;text-transform:uppercase;">Humidity</td>' +
        '<td style="text-align:center;background:rgba(255,255,255,0.04);border-radius:4px;"><span style="display:block;font-size:12px;font-weight:700;color:#e5e7eb;">' + s.hMin.toFixed(0) + '%</span><span style="font-size:8px;color:#6b7280;">LOW</span></td>' +
        '<td style="text-align:center;background:rgba(255,255,255,0.06);border-radius:4px;"><span style="display:block;font-size:13px;font-weight:700;color:#f0a54e;">' + s.hAvg.toFixed(0) + '%</span><span style="font-size:8px;color:#6b7280;">AVG</span></td>' +
        '<td style="text-align:center;background:rgba(255,255,255,0.04);border-radius:4px;"><span style="display:block;font-size:12px;font-weight:700;color:#e5e7eb;">' + s.hMax.toFixed(0) + '%</span><span style="font-size:8px;color:#6b7280;">HIGH</span></td></tr>' +
        // dew + battery
        '<tr><td style="font-size:9px;color:#4b5563;font-weight:700;text-transform:uppercase;">Dew Point</td>' +
        '<td colspan="2" style="background:rgba(62,207,142,0.07);border-radius:4px;text-align:center;">' +
        '<span style="font-size:13px;font-weight:700;color:#3ecf8e;">' + s.dAvg.toFixed(1) + '°C</span></td>' +
        '<td style="background:rgba(255,255,255,0.04);border-radius:4px;text-align:center;">' +
        '<span style="font-size:11px;font-weight:700;color:' + battCol + ';">' + battStr + '</span>' +
        '<span style="display:block;font-size:8px;color:#4b5563;">BATTERY</span></td></tr>' +
        '</table>')) +
      // summary text
      row(td('padding:6px 14px 12px;font-size:11px;color:#9ca3af;line-height:1.7;border-top:1px solid rgba(255,255,255,0.05);', buildDeviceSummary_(dev, s))) +
      '</table>'));
    h += row(td('height:8px;', ''));  // spacer between devices
  });

  // Footer
  h += row(td('padding:10px 20px 14px;border-top:1px solid rgba(255,255,255,0.06);',
    '<table width="100%" cellpadding="0" cellspacing="0"><tr>' +
    td('font-size:10px;color:#374151;', 'Generated by DEWVEE &middot; ' +
       Utilities.formatDate(new Date(), 'UTC', 'dd MMM yyyy HH:mm') + ' UTC') +
    td('font-size:10px;color:#374151;text-align:right;', 'Attached: dewvee_export.csv') +
    '</tr></table>'));

  h += '</table></body></html>';
  return h;
}

// Main report function — fetch data, render charts, send email
function sendWeeklyReport_(recipients) {
  var fromMs = Date.now() - 7 * 86400000;
  var toMs   = Date.now();
  var scan   = scanAllSheets_({ fromMs: fromMs, toMs: toMs });
  var series = scan.series;
  var devOrder = Object.keys(series).sort();
  if (!devOrder.length) { Logger.log('sendWeeklyReport_: no data for last 7 days'); return; }

  var devColors = devOrder.map(function(d,i){ return REPORT_COLORS_[i % REPORT_COLORS_.length]; });

  // Stats per device
  var statsMap = {};
  devOrder.forEach(function(dev){
    statsMap[dev] = weekStats_(series[dev], scan.dataSnap[dev], scan.readingsMap[dev]);
  });

  // Aggregate data for charts
  var dailyByDev = {}, hourlyByDev = {};
  devOrder.forEach(function(dev){
    dailyByDev[dev]  = dailyAvg_(series[dev],  fromMs);
    hourlyByDev[dev] = hourlyAvg_(series[dev], fromMs);
  });

  // Fetch chart images
  var blobs = {}, imageIds = {};

  var tb = fetchQuickChart_(overviewChartCfg_(dailyByDev, devOrder, devColors, 'avgT'), 600, 210);
  if (tb) { tb.setName('ct.png'); blobs['chartTemp'] = tb; imageIds.chartTemp = 'chartTemp'; }

  var hb = fetchQuickChart_(overviewChartCfg_(dailyByDev, devOrder, devColors, 'avgH'), 600, 210);
  if (hb) { hb.setName('ch.png'); blobs['chartHum']  = hb; imageIds.chartHum  = 'chartHum'; }

  devOrder.forEach(function(dev, i){
    var hourly = hourlyByDev[dev];
    if (!hourly || !hourly.length) return;

    // Find indices of temperature and humidity extremes in the hourly series
    var tHiIdx=0, tLoIdx=0, hHiIdx=0, hLoIdx=0;
    var tHi=-Infinity, tLo=Infinity, hHi=-Infinity, hLo=Infinity;
    hourly.forEach(function(h, idx){
      if (h.avgT > tHi){ tHi=h.avgT; tHiIdx=idx; }
      if (h.avgT < tLo){ tLo=h.avgT; tLoIdx=idx; }
      if (h.avgH > hHi){ hHi=h.avgH; hHiIdx=idx; }
      if (h.avgH < hLo){ hLo=h.avgH; hLoIdx=idx; }
    });
    var dewAvg = (statsMap[dev] && statsMap[dev].dAvg) ? statsMap[dev].dAvg : 20;

    var blob = fetchQuickChart_(deviceChartCfg_(hourly, dewAvg, tHiIdx, tLoIdx, hHiIdx, hLoIdx), 560, 185);
    if (blob) {
      var cid = 'dev' + i;
      blob.setName(cid + '.png');
      blobs[cid] = blob;
      imageIds['dev_' + i] = cid;
    }
  });

  // CSV
  var csvRows = [['Timestamp','Device','Temperature(C)','Humidity(%)','Battery(%)','Voltage(V)']];
  devOrder.forEach(function(dev){
    series[dev].forEach(function(pt){
      csvRows.push([ Utilities.formatDate(new Date(fromMs+pt.t),'UTC','yyyy-MM-dd HH:mm:ss'),
                     dev, pt.temp, pt.hum, pt.pct, pt.v ]);
    });
  });
  csvRows.sort(function(a,b){ return a[0]<b[0]?-1:a[0]>b[0]?1:0; });
  var csv = csvRows.map(function(r){
    return r.map(function(c){ return '"'+String(c).replace(/"/g,'""')+'"'; }).join(',');
  }).join('\n');

  var subject = 'DEWVEE Weekly Climate Report — ' +
    Utilities.formatDate(new Date(fromMs),'UTC','MMM d') + ' to ' +
    Utilities.formatDate(new Date(),'UTC','MMM d, yyyy');

  MailApp.sendEmail({
    to:           recipients,
    subject:      subject,
    body:         'DEWVEE Weekly Climate Report — open with an HTML-capable email client to view charts.',
    htmlBody:     buildReportHTML_(series, statsMap, fromMs, imageIds, devOrder, devColors),
    attachments:  [Utilities.newBlob(csv, 'text/csv', 'dewvee_export.csv')],
    inlineImages: blobs
  });
  Logger.log('Report sent to ' + recipients + ' — ' + Object.keys(blobs).length + ' charts.');
}

// Run this from the Apps Script editor to send an immediate demo report
function sendDemoReport() {
  sendWeeklyReport_('Thejthimmaiah04@gmail.com');
}

// Handle a browser-rendered report POST: decode base64 JPEG images, send email
function handleBrowserReport_(body) {
  try {
    var to       = String(body.to      || '').trim();
    var subject  = String(body.subject || 'DEWVEE Weekly Climate Report').trim();
    var html     = String(body.htmlBody || '');
    var images   = body.images   || {};
    var csvFiles = body.csvFiles || {};
    if (!to)   return {ok: false, error: 'missing recipient'};
    if (!html) return {ok: false, error: 'missing htmlBody'};

    var inlineImages = {};
    Object.keys(images).forEach(function(key) {
      var decoded = Utilities.base64Decode(images[key]);
      inlineImages[key] = Utilities.newBlob(decoded, 'image/jpeg', key + '.jpg');
    });

    var attachments = [];
    Object.keys(csvFiles).forEach(function(fname) {
      var decoded = Utilities.base64Decode(csvFiles[fname]);
      attachments.push(Utilities.newBlob(decoded, 'text/csv', fname));
    });

    MailApp.sendEmail({to: to, subject: subject, htmlBody: html,
                       inlineImages: inlineImages, attachments: attachments});
    return {ok: true};
  } catch(e) {
    return {ok: false, error: e.message};
  }
}

// ================================================================
//  Alert checking — runs on the daily trigger alongside export,
//  or can be wired to its own time-based trigger.
// ================================================================
function checkAlerts() {
  var sp       = PropertiesService.getScriptProperties();
  var alertRaw = sp.getProperty('ALERT_SETTINGS') || '{}';
  var alerts;
  try { alerts = JSON.parse(alertRaw); } catch(e) { return; }
  if (!Object.keys(alerts).length) return;

  var emailProp = sp.getProperty('EXPORT_EMAILS') || '[]';
  var emails;
  try { emails = JSON.parse(emailProp); } catch(e) { emails = []; }
  if (!emails.length) return;

  var devData = getOrCreateSheet_(DEVICES_SHEET).getDataRange().getValues();
  var breaches = [];

  for (var r = 1; r < devData.length; r++) {
    var device  = String(devData[r][DEV.DEVICE   - 1]);
    var location= String(devData[r][DEV.LOCATION  - 1] || '');
    var lastTemp= parseFloat(devData[r][DEV.LAST_TEMP - 1]);
    var lastHum = parseFloat(devData[r][DEV.LAST_HUM  - 1]);
    var lastSeen= new Date(devData[r][DEV.LAST_SEEN - 1]).getTime();
    var cfg     = alerts[device];
    if (!cfg) continue;
    if (isNaN(lastTemp) || isNaN(lastHum)) continue;

    var msgs = [];
    if (cfg.tempMin !== undefined && lastTemp < cfg.tempMin)
      msgs.push('Temperature ' + lastTemp.toFixed(1) + '°C below minimum ' + cfg.tempMin + '°C');
    if (cfg.tempMax !== undefined && lastTemp > cfg.tempMax)
      msgs.push('Temperature ' + lastTemp.toFixed(1) + '°C above maximum ' + cfg.tempMax + '°C');
    if (cfg.humMin  !== undefined && lastHum  < cfg.humMin)
      msgs.push('Humidity '    + lastHum.toFixed(0)  + '% below minimum ' + cfg.humMin  + '%');
    if (cfg.humMax  !== undefined && lastHum  > cfg.humMax)
      msgs.push('Humidity '    + lastHum.toFixed(0)  + '% above maximum ' + cfg.humMax  + '%');

    if (msgs.length) {
      var label = location ? device + ' (' + location + ')' : device;
      breaches.push('⚠ ' + label + ':\n  ' + msgs.join('\n  '));
    }
  }

  if (!breaches.length) return;

  var body = 'DEWVEE Alert — ' + Utilities.formatDate(new Date(), 'UTC', 'dd MMM yyyy HH:mm') + ' UTC\n\n' +
             breaches.join('\n\n') +
             '\n\nLog in to the DEWVEE dashboard to review the latest readings.';
  MailApp.sendEmail({
    to:      emails.join(','),
    subject: '⚠ DEWVEE Alert: ' + breaches.length + ' threshold breach(es)',
    body:    body
  });
  Logger.log('Alert email sent to ' + emails.join(',') + ' — ' + breaches.length + ' breach(es).');
}

// ================================================================
//  Data retention — removes rows older than RETAIN_DAYS from all
//  per-device sheets. Run weekly via setupAllTriggers or manually.
// ================================================================
var RETAIN_DAYS = 90;

function pruneOldData_() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var cutoff  = Date.now() - RETAIN_DAYS * 86400000;
  var pruned  = 0;
  ss.getSheets().forEach(function(sh) {
    if (sh.getName() === DEVICES_SHEET) return;
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return;
    var all = sh.getDataRange().getValues();
    var kept = [all[0]];          // always keep the header row
    for (var i = 1; i < all.length; i++) {
      var raw = all[i][0];
      var ts  = raw instanceof Date ? raw.getTime() : new Date(raw).getTime();
      if (!ts || isNaN(ts) || ts >= cutoff) {
        kept.push(all[i]);
      } else {
        pruned++;
      }
    }
    if (kept.length < all.length) {
      sh.clearContents();
      sh.getRange(1, 1, kept.length, kept[0].length).setValues(kept);
    }
  });
  Logger.log('pruneOldData_: removed ' + pruned + ' rows older than ' + RETAIN_DAYS + ' days.');
}

// ================================================================
//  Weekly backup — copies the Data sheet to a timestamped file
//  in the same Google Drive folder as this spreadsheet.
// ================================================================
function weeklyBackup() {
  try {
    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var file    = DriveApp.getFileById(ss.getId());
    var folder  = file.getParents().next();
    var name    = 'DEWVEE Backup ' + Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd');
    ss.copy(name).getParents().next(); // copy creates in root; move below
    var copies  = DriveApp.getFilesByName(name);
    if (copies.hasNext()) folder.addFile(copies.next());
    Logger.log('Weekly backup created: ' + name);
  } catch(err) {
    Logger.log('weeklyBackup failed: ' + err.message);
  }
}

// ================================================================
//  Setup helpers — run once manually from the Apps Script editor
// ================================================================
function setupAllTriggers() {
  // Clear everything first
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });
  // Daily export + alert at configured hour (default 8 UTC)
  var hour = parseInt(PropertiesService.getScriptProperties().getProperty('EXPORT_HOUR') || '8');
  ScriptApp.newTrigger('sendDailyExport').timeBased().atHour(hour).everyDays(1).create();
  ScriptApp.newTrigger('checkAlerts').timeBased().atHour(hour).everyDays(1).create();
  // Weekly backup every Sunday at 02:00 UTC
  ScriptApp.newTrigger('weeklyBackup').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(2).create();
  // Weekly data prune every Sunday at 03:00 UTC (removes rows > 90 days old)
  ScriptApp.newTrigger('pruneOldData_').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(3).create();
  Logger.log('All triggers created.');
}
