/**
 * מד לחץ דם — צד השרת (Google Apps Script, מחובר לגיליון Google Sheets).
 *
 * מה הקוד עושה:
 *   • analyze — מקבל צילום של מסך מכשיר לחץ הדם, שולח אותו ל-Claude
 *               ומחזיר סיסטולי / דיאסטולי / דופק.
 *   • save    — מוסיף שורה לגיליון "מדידות" (כולל מניעת כפילויות).
 *   • list    — מחזיר את המדידות האחרונות לאפליקציה (היסטוריה וגרף).
 *
 * התקנה — ראו bp-monitor/README.md. בקצרה:
 *   1. בגיליון: הרחבות ← Apps Script, מדביקים את הקובץ הזה ושומרים.
 *   2. מריצים את הפונקציה setup פעם אחת (ומאשרים הרשאות).
 *   3. בגיליון: תפריט "לחץ דם" ← "הגדרת מפתח Claude".
 *   4. פריסה ← פריסה חדשה ← אפליקציית אינטרנט, "מי יכול לגשת: כולם".
 */

var SHEET_NAME = 'מדידות';
var HEADERS = ['תאריך ושעה', 'סיסטולי', 'דיאסטולי', 'דופק', 'סיווג', 'הערות', 'מקור', 'מזהה'];
var CLAUDE_MODEL = 'claude-opus-5';

// ---------------------------------------------------------------- setup / menu

function setup() {
  var sheet = getSheet_();
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('APP_TOKEN')) {
    props.setProperty('APP_TOKEN', Utilities.getUuid().replace(/-/g, '').slice(0, 16));
  }
  buildChart_(sheet);
  Logger.log('הגיליון מוכן. קוד הגישה לאפליקציה: ' + props.getProperty('APP_TOKEN'));
  if (!props.getProperty('ANTHROPIC_API_KEY')) {
    Logger.log('חסר מפתח Claude: בגיליון ← תפריט "לחץ דם" ← "הגדרת מפתח Claude".');
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('לחץ דם')
    .addItem('הצגת קוד גישה לאפליקציה', 'showToken')
    .addItem('הגדרת מפתח Claude', 'promptApiKey')
    .addItem('בניית גרף מחדש', 'rebuildChart')
    .addToUi();
}

function showToken() {
  var token = PropertiesService.getScriptProperties().getProperty('APP_TOKEN');
  SpreadsheetApp.getUi().alert(token ? 'קוד הגישה: ' + token : 'הריצו קודם את setup מתוך עורך הסקריפט.');
}

function promptApiKey() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('מפתח Claude', 'הדביקו את מפתח ה-API (מתחיל ב-sk-ant-):', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var key = res.getResponseText().trim();
  if (key.indexOf('sk-ant-') !== 0) {
    ui.alert('המפתח לא נראה תקין — הוא אמור להתחיל ב-sk-ant-');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('ANTHROPIC_API_KEY', key);
  ui.alert('המפתח נשמר ✔');
}

function rebuildChart() {
  buildChart_(getSheet_());
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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
        return json_({ ok: true, sheetUrl: SpreadsheetApp.getActiveSpreadsheet().getUrl() });
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
  return { ok: true, readings: readings, sheetUrl: SpreadsheetApp.getActiveSpreadsheet().getUrl() };
}

// ---------------------------------------------------------------- Claude vision

var READING_SCHEMA = {
  type: 'object',
  properties: {
    readable: { type: 'boolean', description: 'true only if systolic and diastolic are clearly legible' },
    systolic: { type: 'integer', description: 'SYS value in mmHg, 0 if not readable' },
    diastolic: { type: 'integer', description: 'DIA value in mmHg, 0 if not readable' },
    pulse: { type: 'integer', description: 'PULSE per minute, 0 if not shown' },
    device_datetime: { type: 'string', description: 'date/time shown on the display as ISO 8601, or empty string' },
    note: { type: 'string', description: 'short Hebrew note: e.g. irregular heartbeat icon, error code, or why unreadable; empty if nothing' },
  },
  required: ['readable', 'systolic', 'diastolic', 'pulse', 'device_datetime', 'note'],
  additionalProperties: false,
};

function analyzeImage_(base64, mediaType) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) throw new Error('לא הוגדר מפתח Claude בגיליון (תפריט "לחץ דם")');
  if (!base64) throw new Error('לא התקבלה תמונה');

  var payload = {
    model: CLAUDE_MODEL,
    max_tokens: 4000,
    fallbacks: 'default',
    output_config: { effort: 'low', format: { type: 'json_schema', schema: READING_SCHEMA } },
    system:
      'You read home blood-pressure monitor displays from phone photos. ' +
      'These screens usually use 7-segment digits: SYS (systolic, top, largest), DIA (diastolic, middle) and PULSE (bottom, often next to a heart icon). ' +
      'Watch for 7-segment look-alikes (1/7, 5/6, 8/0/9) and glare. Never guess: if SYS or DIA is not clearly legible, set readable=false. ' +
      'Include a short Hebrew note if an irregular-heartbeat or movement icon, an error code, or low battery is shown.',
    messages: [
      {
        role: 'user',
        content: [
          { type: 'image', source: { type: 'base64', media_type: mediaType, data: base64 } },
          { type: 'text', text: 'Read the blood pressure measurement on this device.' },
        ],
      },
    ],
  };

  var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'anthropic-beta': 'server-side-fallback-2026-07-01',
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  var data = JSON.parse(res.getContentText());
  if (res.getResponseCode() !== 200) {
    throw new Error('שגיאה מ-Claude: ' + (data.error && data.error.message ? data.error.message : res.getResponseCode()));
  }
  if (data.stop_reason === 'refusal') throw new Error('Claude לא הצליח לעבד את התמונה — נסו שוב או הזינו ידנית');
  var text = (data.content || []).filter(function (b) { return b.type === 'text'; }).map(function (b) { return b.text; }).join('');
  var reading = JSON.parse(text);
  if (reading.readable) reading.category = classify_(reading.systolic, reading.diastolic);
  return reading;
}
