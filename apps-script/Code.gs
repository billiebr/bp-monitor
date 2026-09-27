/**
 * מד לחץ דם — צד השרת (Google Apps Script).
 *
 * מה הקוד עושה:
 *   • analyze — מקבל צילום של מסך מכשיר לחץ הדם, מזהה את הספרות בעזרת זיהוי הטקסט (OCR)
 *               של Google Drive בחשבון שלך — בלי מפתחות ובלי עלות — ומחזיר סיסטולי / דיאסטולי / דופק.
 *   • save    — מוסיף שורה ללשונית "מדידות" בגיליון (כולל מניעת כפילויות).
 *   • list    — מחזיר את המדידות האחרונות.
 *
 * התקנה: מדביקים את הקובץ, מריצים setup פעם אחת (מאשרים הרשאות), ומפרסמים גרסה חדשה.
 */

// מזהה הגיליון (החלק שבכתובת שלו בין /d/ ל-/edit). נדרש כשהסקריפט לא נפתח מתוך הגיליון
// (הרחבות ← Apps Script); אם הסקריפט נפתח מתוך הגיליון אפשר להשאיר ריק.
var SPREADSHEET_ID = '1bsXluNMZu4tYzahqACGpnlSbUcfpP6wZspXc8t5T_gA';

var SHEET_NAME = 'מדידות';
var HEADERS = ['תאריך ושעה', 'סיסטולי', 'דיאסטולי', 'דופק', 'סיווג', 'הערות', 'מקור', 'מזהה'];

// ---------------------------------------------------------------- setup / menu

function setup() {
  var sheet = getSheet_();
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('APP_TOKEN')) {
    props.setProperty('APP_TOKEN', Utilities.getUuid().replace(/-/g, '').slice(0, 16));
  }
  buildChart_(sheet);
  Logger.log('הגיליון מוכן. קוד הגישה לאפליקציה: ' + props.getProperty('APP_TOKEN'));
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('לחץ דם')
    .addItem('הצגת קוד גישה לאפליקציה', 'showToken')
    .addItem('בניית גרף מחדש', 'rebuildChart')
    .addToUi();
}

function showToken() {
  var token = PropertiesService.getScriptProperties().getProperty('APP_TOKEN');
  SpreadsheetApp.getUi().alert(token ? 'קוד הגישה: ' + token : 'הריצו קודם את setup מתוך עורך הסקריפט.');
}

function rebuildChart() {
  buildChart_(getSheet_());
}

function getSpreadsheet_() {
  return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME, 0);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setRightToLeft(true);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#fde8ea');
    sheet.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm');
    sheet.setColumnWidth(1, 140);
    sheet.setColumnWidth(5, 170);
    sheet.setColumnWidth(6, 220);
    sheet.hideColumns(8);
    addCategoryColors_(sheet);
  }
  return sheet;
}

function addCategoryColors_(sheet) {
  var range = sheet.getRange('E2:E');
  var colors = [
    ['נמוך', '#dbeafe'],
    ['תקין', '#dcfce7'],
    ['תקין-גבוה', '#fef9c3'],
    ['יתר לחץ דם דרגה 1', '#fed7aa'],
    ['יתר לחץ דם דרגה 2', '#fecaca'],
    ['יתר לחץ דם דרגה 3', '#f87171'],
  ];
  var rules = colors.map(function (c) {
    return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(c[0]).setBackground(c[1]).setRanges([range]).build();
  });
  sheet.setConditionalFormatRules(rules);
}

function buildChart_(sheet) {
  sheet.getCharts().forEach(function (c) { sheet.removeChart(c); });
  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sheet.getRange('A1:D1000'))
    .setNumHeaders(1)
    .setPosition(2, 10, 0, 0)
    .setOption('title', 'לחץ דם ודופק לאורך זמן')
    .setOption('colors', ['#dc2626', '#2563eb', '#16a34a'])
    .setOption('pointSize', 5)
    .setOption('width', 700)
    .setOption('height', 380)
    .build();
  sheet.insertChart(chart);
}

// ---------------------------------------------------------------- web app

function doGet() {
  return json_({ ok: true, app: 'bp-monitor' });
}

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var token = PropertiesService.getScriptProperties().getProperty('APP_TOKEN');
    if (!token || body.token !== token) {
      return json_({ ok: false, error: 'קוד גישה שגוי' });
    }
    switch (body.action) {
      case 'ping':
        return json_({ ok: true, sheetUrl: getSpreadsheet_().getUrl() });
      case 'analyze':
        return json_({ ok: true, reading: analyzeImage_(body.image, body.mediaType || 'image/jpeg') });
      case 'save':
        return json_(saveReading_(body.reading || {}));
      case 'list':
        return json_(listReadings_(Number(body.limit) || 200));
      default:
        return json_({ ok: false, error: 'פעולה לא מוכרת' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------- readings

function classify_(sys, dia) {
  if (sys >= 180 || dia >= 110) return 'יתר לחץ דם דרגה 3';
  if (sys >= 160 || dia >= 100) return 'יתר לחץ דם דרגה 2';
  if (sys >= 140 || dia >= 90) return 'יתר לחץ דם דרגה 1';
  if (sys >= 130 || dia >= 85) return 'תקין-גבוה';
  if (sys < 90 || dia < 60) return 'נמוך';
  return 'תקין';
}

function saveReading_(r) {
  var sys = Math.round(Number(r.systolic));
  var dia = Math.round(Number(r.diastolic));
  var pulse = r.pulse === '' || r.pulse == null ? '' : Math.round(Number(r.pulse));
  if (!(sys >= 50 && sys <= 300) || !(dia >= 30 && dia <= 200) || dia >= sys) {
    return { ok: false, error: 'הערכים לא נראים תקינים' };
  }
  var id = String(r.id || Utilities.getUuid());
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = getSheet_();
    // Offline retries resend the same id — never write it twice.
    if (sheet.getLastRow() > 1) {
      var found = sheet.getRange(2, 8, sheet.getLastRow() - 1, 1).createTextFinder(id).matchEntireCell(true).findNext();
      if (found) return { ok: true, duplicate: true };
    }
    var when = r.takenAt ? new Date(r.takenAt) : new Date();
    if (isNaN(when.getTime())) when = new Date();
    sheet.appendRow([when, sys, dia, pulse, classify_(sys, dia), String(r.note || ''), String(r.source || 'ידני'), id]);
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

function listReadings_(limit) {
  var sheet = getSheet_();
  var last = sheet.getLastRow();
  var readings = [];
  if (last > 1) {
    var start = Math.max(2, last - limit + 1);
    var rows = sheet.getRange(start, 1, last - start + 1, 8).getValues();
    rows.forEach(function (row) {
      if (!row[1]) return;
      readings.push({
        takenAt: row[0] instanceof Date ? row[0].toISOString() : String(row[0]),
        systolic: row[1],
        diastolic: row[2],
        pulse: row[3],
        category: row[4],
        note: row[5],
        source: row[6],
        id: row[7],
      });
    });
  }
  readings.sort(function (a, b) { return new Date(b.takenAt) - new Date(a.takenAt); });
  return { ok: true, readings: readings, sheetUrl: getSpreadsheet_().getUrl() };
}

// ---------------------------------------------------------------- photo reading (Google Drive OCR)

function analyzeImage_(base64, mediaType) {
  if (!base64) throw new Error('לא התקבלה תמונה');
  var text = ocr_(Utilities.base64Decode(base64), mediaType);
  return parseReading_(text);
}

// Uploads the photo to Drive as a Google Doc — Drive runs OCR on the way in —
// reads the recognised text, then moves the temporary file to the trash.
function ocr_(bytes, mediaType) {
  var boundary = 'bpBoundary' + Date.now();
  var meta = JSON.stringify({ name: 'bp-ocr-temp', mimeType: 'application/vnd.google-apps.document' });
  var head = '--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' + meta + '\r\n' +
    '--' + boundary + '\r\nContent-Type: ' + mediaType + '\r\n\r\n';
  var tail = '\r\n--' + boundary + '--';
  var body = Utilities.newBlob(head).getBytes().concat(bytes, Utilities.newBlob(tail).getBytes());

  var res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&ocrLanguage=en', {
    method: 'post',
    contentType: 'multipart/related; boundary=' + boundary,
    payload: body,
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('זיהוי התמונה נכשל (' + res.getResponseCode() + ')');
  }
  var id = JSON.parse(res.getContentText()).id;
  try {
    return DocumentApp.openById(id).getBody().getText();
  } finally {
    DriveApp.getFileById(id).setTrashed(true);
  }
}

// Picks SYS / DIA / PULSE out of the OCR text. Displays list them top to bottom,
// so we take the first plausible trio in reading order; dates and times are skipped.
function parseReading_(text) {
  var cleaned = String(text || '')
    .replace(/(\d{1,2})\s*[:.]\s*(\d{2})/g, ' ')          // clock 10:25
    .replace(/\d{1,4}\s*[\/-]\s*\d{1,2}(\s*[\/-]\s*\d{1,4})?/g, ' ') // dates 27/9
    .replace(/[Oo]/g, '0').replace(/[lI|]/g, '1').replace(/[Ss]/g, '5').replace(/B/g, '8');
  var nums = (cleaned.match(/\d{2,3}/g) || []).map(Number);

  for (var i = 0; i < nums.length - 1; i++) {
    var sys = nums[i], dia = nums[i + 1];
    if (sys >= 70 && sys <= 260 && dia >= 35 && dia <= 160 && dia < sys) {
      var pulse = nums[i + 2];
      return {
        readable: true,
        systolic: sys,
        diastolic: dia,
        pulse: pulse >= 30 && pulse <= 220 ? pulse : 0,
        device_datetime: '',
        note: '',
        category: classify_(sys, dia),
      };
    }
  }
  return { readable: false, systolic: 0, diastolic: 0, pulse: 0, device_datetime: '', note: '', rawText: String(text || '').slice(0, 200) };
}
