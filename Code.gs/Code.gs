/**
 * קוד סנכרון חכם - Biscotti Manager
 * מותאם לכותרות בעברית מהגיליון
 */

const CONFIG = {
  'שעות תקן': {
    'תאריך': 'date',
    'מכירות': 'sales',
    'שעות': 'hours'
  },
  'ניהול עובדים': {
    'מספר עובד': 'id',
    'שם': 'name',
    'תפקיד': 'position',
    'חברה': 'company',
    'טלפון': 'phone',
    'תאריך התחלה': 'startDate',
    'הערות': 'notes'
  },
  'מלאי': {
    'מזהה': 'id',
    'כמות': 'qty',
    'מינימום': 'min'
  },
  'משימות לדוד': {
    'ID': 'id',
    'משימה': 'name',
    'בוצע': 'checked'
  }
};

function doGet(e) {
  const action = e.parameter.action;
  try {
    if (action === 'getBootstrap') {
      return jsonResponse({
        status: 'success',
        standardHours: getMappedData('שעות תקן'),
        employees: getMappedData('ניהול עובדים'),
        inventory: getMappedData('מלאי'),
        tasks: getMappedData('משימות לדוד')
      });
    }
    return jsonResponse({ status: 'error', message: 'Unknown action' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

function getMappedData(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const mapping = CONFIG[sheetName] || {};

  return data.map(row => {
    let obj = {};
    headers.forEach((h, i) => {
      const key = mapping[h] || h;
      let val = row[i];
      // טיפול מיוחד בהערות (JSON)
      if (key === 'notes' && val) {
        try { val = JSON.parse(val); } catch(e) { val = []; }
      }
      obj[key] = val;
    });
    return obj;
  });
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
