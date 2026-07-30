// ================================================================
//  DEWVEE  —  Google Apps Script backend
// ================================================================
//  Sheet layout (auto-created):
//    "Data"    — all sensor readings, one row per reading
//    "Devices" — one row per unique device: metadata + last seen
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
            LAST_HUM:7, LAST_PCT:8 };

// ================================================================
//  Sheet helpers
// ================================================================
function getOrCreateSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (name === DATA_SHEET) {
      sh.appendRow(['Timestamp', 'Device', 'Temperature', 'Humidity',
                    'Battery%', 'Voltage']);
      sh.getRange(1, 1, 1, 6).setFontWeight('bold');
    } else if (name === DEVICES_SHEET) {
      sh.appendRow(['Device', 'Location', 'SampleMin', 'FirstSeen',
                    'LastSeen', 'LastTemp', 'LastHum', 'LastPct']);
      sh.getRange(1, 1, 1, 8).setFontWeight('bold');
    }
  }
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
    var device    = String(body.device    || '').trim();
    var location  = String(body.location  || '').trim();
    var sampleMin = parseInt(body.sampleMin) || 5;
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

    // ── Write to Data sheet ───────────────────────────────────────
    var dataSheet = getOrCreateSheet_(DATA_SHEET);
    dataSheet.getRange(dataSheet.getLastRow() + 1, 1, valid.length, 6)
             .setValues(valid);

    // ── Update / insert Devices sheet row ────────────────────────
    var devSheet = getOrCreateSheet_(DEVICES_SHEET);
    var lastRow  = valid[valid.length - 1];
    var devRow   = getDeviceRow_(devSheet, device);
    var now      = new Date();
    if (devRow < 0) {
      devSheet.appendRow([device, location, sampleMin, now, now,
                          lastRow[2], lastRow[3], lastRow[4]]);
    } else {
      var existing = devSheet.getRange(devRow, 1, 1, 8).getValues()[0];
      devSheet.getRange(devRow, 1, 1, 8).setValues([[
        device,
        location || existing[DEV.LOCATION - 1],
        sampleMin,
        existing[DEV.FIRST_SEEN - 1] || now,
        now,
        lastRow[2], lastRow[3], lastRow[4]
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
  var ms = { '6h':6*3600e3, '12h':12*3600e3, 'day':86400e3,
             'week':7*86400e3, 'month':30*86400e3, 'year':365*86400e3 };
  return ms[range] || 86400e3;
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

  var voltMap    = {};
  var readingsMap = {};
  var dataSnap   = {};
  var series     = {};   // used in seriesMode

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

      // Always track voltage and readings for battery projection
      voltMap[dev] = volt;

      if (!isNaN(pct) && pct >= 0 && pct <= 100) {
        if (!readingsMap[dev]) readingsMap[dev] = [];
        readingsMap[dev].push({ ts: ts, pct: pct });
      }

      if (!dataSnap[dev] || ts > dataSnap[dev].lastTs) {
        dataSnap[dev] = { lastTs: ts, lastTemp: temp, lastHum: hum,
                          lastPct: isNaN(pct) ? 0 : pct, lastVolt: volt };
      }

      // Series data (for charts)
      if (seriesMode && ts >= fromMs && ts <= toMs) {
        if (!series[dev]) series[dev] = [];
        series[dev].push({ t: Math.round(ts - fromMs), temp: temp, hum: hum,
                           pct: isNaN(pct) ? 0 : pct, v: volt });
      }
    }
  });

  // Sort battery readings ascending
  for (var d in readingsMap) {
    readingsMap[d].sort(function(a, b) { return a.ts - b.ts; });
  }

  return { voltMap: voltMap, readingsMap: readingsMap, dataSnap: dataSnap, series: series };
}

function computeBatteryProjection_(readings) {
  // readings: [{ts:ms, pct:number}] sorted ascending — all readings for this device
  if (!readings || readings.length < 4) return null;

  // Walk backward to find last recharge event (pct rise > 5 pp)
  var startIdx = 0;
  for (var i = readings.length - 1; i > 0; i--) {
    if (readings[i].pct - readings[i - 1].pct > 5) { startIdx = i; break; }
  }
  var pts = readings.slice(startIdx);
  if (pts.length < 4) return null;

  var spanMs = pts[pts.length - 1].ts - pts[0].ts;
  if (spanMs < 6 * 3600000) return { note: 'insufficient_data' };

  // Linear regression: x = hours elapsed, y = pct
  var t0 = pts[0].ts;
  var n = pts.length, sx=0, sy=0, sxy=0, sxx=0;
  for (var j = 0; j < n; j++) {
    var x = (pts[j].ts - t0) / 3600000;
    var y = pts[j].pct;
    sx += x; sy += y; sxy += x*y; sxx += x*x;
  }
  var mx = sx/n, my = sy/n;
  var ssxy = sxy - n*mx*my, ssxx = sxx - n*mx*mx;
  if (ssxx < 0.001 || ssxy >= 0) return { note: 'stable' }; // not draining

  var slope = ssxy / ssxx; // pct/hour (negative)
  var intercept = my - slope * mx;

  // R² — how clean the discharge trend is
  var ssRes = 0, ssTot = 0;
  for (var k = 0; k < n; k++) {
    var pred = slope * ((pts[k].ts - t0) / 3600000) + intercept;
    ssRes += Math.pow(pts[k].pct - pred, 2);
    ssTot += Math.pow(pts[k].pct - my,   2);
  }
  var r2 = ssTot > 0.001 ? 1 - ssRes / ssTot : 0;

  var drainPerHour = -slope;              // positive %/h
  if (drainPerHour < 0.001) return { note: 'stable' };

  var currentPct  = pts[pts.length - 1].pct;
  var hoursLeft   = currentPct / drainPerHour;
  var daysLeft    = hoursLeft / 24;
  var drainPerDay = drainPerHour * 24;

  return {
    daysLeft:      Math.round(daysLeft    * 10) / 10,
    drainPctPerDay:Math.round(drainPerDay * 10) / 10,
    r2:            Math.round(r2          * 100) / 100,
    dataPoints:    n,
    spanDays:      Math.round(spanMs / 86400000 * 10) / 10
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
    var lastSeen  = row[DEV.LAST_SEEN - 1];
    var lastTemp  = row[DEV.LAST_TEMP - 1];
    var lastHum   = row[DEV.LAST_HUM  - 1];
    var lastPct   = row[DEV.LAST_PCT  - 1];
    if (!device) continue;

    seenDevs[device] = true;
    var tsMs = lastSeen ? new Date(lastSeen).getTime() : 0;
    result.push({
      device:    device,
      location:  location || '',
      sampleMin: sampleMin,
      temp:      parseFloat(lastTemp)  || 0,
      hum:       parseFloat(lastHum)   || 0,
      pct:       parseInt(lastPct)     || 0,
      volt:      voltMap[device]       || 0,
      ts:        Math.floor(tsMs / 1000),
      lowBatt:   parseInt(lastPct) <= 20,
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
    var sampleMin = 10;
    result.push({
      device:    dev,
      location:  '',
      sampleMin: sampleMin,
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
  var rangeMs    = parseRangeMs_(range);
  var shiftMs    = dayOffset * 86400000;
  var toMs       = nowMs - shiftMs;
  var fromMs     = toMs  - rangeMs;
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
  var sp     = PropertiesService.getScriptProperties();
  var status = '';
  try {
    var emailProp = sp.getProperty('EXPORT_EMAILS') || '[]';
    var emails;
    try { emails = JSON.parse(emailProp); } catch(e) { emails = []; }
    if (!emails.length) {
      var legacy = sp.getProperty('EXPORT_EMAIL');
      if (legacy) emails = [legacy];
    }
    if (!emails.length) { sp.setProperty('EXPORT_STATUS','No recipients configured'); return; }

    // Build CSV for the last 7 days
    var sheet = getOrCreateSheet_(DATA_SHEET);
    var data  = sheet.getDataRange().getValues();
    var from  = Date.now() - 7 * 86400e3;
    var rows  = [['Timestamp','Device','Location','Temperature(C)','Humidity(%)','Battery(%)','Voltage(V)']];

    // Build device → location map
    var locMap = {};
    try {
      var devData = getOrCreateSheet_(DEVICES_SHEET).getDataRange().getValues();
      for (var r = 1; r < devData.length; r++)
        locMap[devData[r][DEV.DEVICE-1]] = devData[r][DEV.LOCATION-1] || '';
    } catch(e) {}

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var ts  = new Date(row[DC.TS - 1]);
      if (ts.getTime() < from) continue;
      rows.push([
        Utilities.formatDate(ts, 'UTC', 'yyyy-MM-dd HH:mm:ss'),
        row[DC.DEVICE - 1],
        locMap[row[DC.DEVICE - 1]] || '',
        row[DC.TEMP - 1], row[DC.HUM - 1], row[DC.PCT - 1], row[DC.VOLT - 1]
      ]);
    }

    var csv = rows.map(function(r) {
      return r.map(function(c) { return '"' + String(c).replace(/"/g,'""') + '"'; }).join(',');
    }).join('\n');

    var subject = 'DEWVEE Weekly Export – ' +
      Utilities.formatDate(new Date(), 'UTC', 'dd MMM yyyy');
    MailApp.sendEmail({
      to:          emails.join(','),
      subject:     subject,
      body:        'Attached is the DEWVEE sensor data for the last 7 days.',
      attachments: [Utilities.newBlob(csv, 'text/csv', 'dewvee_export.csv')]
    });
    status = 'OK:' + new Date().toISOString();
  } catch(err) {
    status = 'ERR:' + err.message + ':' + new Date().toISOString();
    Logger.log('sendDailyExport failed: ' + err.message);
  }
  sp.setProperty('EXPORT_STATUS', status);
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
  Logger.log('All triggers created.');
}
