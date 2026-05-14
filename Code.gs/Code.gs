/**
 * Biscotti Manager - Server Side API (code.gs)
 * סנכרון מלא עמודות A-H | ניהול מלאי והזמנות
 */

function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    const action = e.parameter.action;
    let payload = null;
    try {
      if (e.parameter.payload) {
        payload = JSON.parse(decodeURIComponent(e.parameter.payload));
      }

      let result;
      switch(action) {
        case 'getBootstrap':       result = getBootstrap();              break;
        case 'saveEmployeeData':   result = saveEmployeeData(payload);   break;
        case 'updateInventory':    result = updateInventory(payload);    break;
        case 'updateTask':         result = updateTask(payload);                   break;
        case 'resetAllTasks':      result = resetAllTasks();                       break;
        case 'updateTaskPersonal': result = updateTaskPersonal(payload);           break;
        case 'resetAllTasksPersonal': result = resetAllTasksPersonal();            break;
        case 'saveNoteToEmployee': result = saveNoteToEmployee(payload); break;
        case 'deleteNoteFromEmployee': result = deleteNoteFromEmployee(payload); break;
        case 'syncFromYerakot':    result = syncFromYerakot();             break;
        case 'debugSync':          result = debugSync();                   break;
        default: result = { status: 'error', message: 'Action not found: ' + action };
      }
      return ContentService.createTextOutput(JSON.stringify({ status: 'success', data: result }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('ביסקוטי - ניהול חכם')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) { return doGet(e); }

// ─── Bootstrap ────────────────────────────────────────────────────────────────

function getBootstrap() {
  const ss = SpreadsheetApp.openById(BISCOTTI_SHEET_ID);
  return {
    employees:     getSheetData(ss, 'ניהול עובדים'),
    standardHours: getSheetData(ss, 'שעות תקן'),
    inventory:     getSheetData(ss, 'מלאי'),
    tasks:         getSheetData(ss, 'משימות לדוד'),
    tasksPersonal: getSheetData(ss, 'משימות אישיות')
  };
}

// קריאה גנרית — מחזירה מערך אובייקטים מותאם לכל גיליון
function getSheetData(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) return [];
  const data        = sheet.getDataRange().getValues();
  const displayData = sheet.getDataRange().getDisplayValues(); // שומר פורמט שעות/טלפונים
  if (data.length <= 1) return [];

  const tz = ss.getSpreadsheetTimeZone();
  let result = [];

  for (let i = 1; i < data.length; i++) {
    const r    = data[i];
    const dRow = displayData[i];
    if (r[0] === "" && r[1] === "") continue;

    if (name === 'שעות תקן') {
      result.push({ date: formatDate(r[0], ss), sales: r[1] || 0, hours: r[2] || 0 });

    } else if (name === 'מלאי') {
      result.push({ id: r[0].toString().trim(), qty: r[1] || 0, timestamp: r[2], min: r[3] || 0 });

    } else if (name === 'משימות לדוד' || name === 'משימות אישיות') {
      result.push({ id: i, name: r[0], done: r[1] || 0, checked: Number(r[1]) > 0 });

    } else if (name === 'ניהול עובדים') {
      let notes = [];
      try { notes = r[5] ? JSON.parse(r[5]) : []; } catch(e) {}

      // שעת התחלה (עמודה G) — נרמול HH:MM
      let st = dRow[6] ? String(dRow[6]).trim().replace(/^'/, "") : "";
      if (st && st.includes(':')) {
        const parts = st.split(':');
        st = parts[0].padStart(2, '0') + ':' + parts[1].substring(0, 2).padStart(2, '0');
      }

      // תאריך התחלה (עמודה H)
      let sDate = "";
      if (r[7] instanceof Date) {
        sDate = Utilities.formatDate(r[7], tz, "yyyy-MM-dd");
      } else if (r[7]) {
        sDate = String(r[7]).trim().replace(/^'/, "");
      }

      result.push({
        position:  dRow[0],
        name:      dRow[1],
        id:        dRow[2],
        company:   dRow[3],
        phone:     dRow[4].replace(/^'/, ""),
        notes:     notes,
        startTime: st,
        startDate: sDate
      });
    }
  }
  // deduplicate inventory by id — keep last row per item
  if (name === 'מלאי') {
    const seen = new Map();
    result.forEach(r => seen.set(r.id, r));
    return [...seen.values()];
  }

  return result;
}

// ─── Employees ────────────────────────────────────────────────────────────────

function saveEmployeeData(emp) {
  const ss    = SpreadsheetApp.openById(BISCOTTI_SHEET_ID);
  const sheet = getOrCreateSheet(ss, 'ניהול עובדים');
  const data  = sheet.getDataRange().getValues();

  let foundRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] && data[i][2].toString() === emp.id.toString()) { foundRow = i + 1; break; }
  }

  const notesStr = emp.notes
    ? (typeof emp.notes === 'string' ? emp.notes : JSON.stringify(emp.notes))
    : "[]";

  // מגן פורמטים עם ' כדי למנוע המרת טלפון/שעה לפורמט מספרי
  const newRow = [
    emp.position  || "",
    emp.name      || "",
    emp.id        || "",
    emp.company   || "",
    emp.phone     ? "'" + emp.phone     : "",
    notesStr,
    emp.startTime ? "'" + emp.startTime : "",
    emp.startDate ? "'" + emp.startDate : ""
  ];

  if (foundRow > -1) { sheet.getRange(foundRow, 1, 1, 8).setValues([newRow]); }
  else               { sheet.appendRow(newRow); }
  return { status: 'success' };
}

function saveNoteToEmployee(data) {
  const sheet = SpreadsheetApp.openById(BISCOTTI_SHEET_ID).getSheetByName('ניהול עובדים');
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][2] && rows[i][2].toString() === data.employeeId.toString()) {
      let notes = rows[i][5] ? JSON.parse(rows[i][5]) : [];
      notes.unshift({ text: data.note, date: new Date().toLocaleString('he-IL') });
      sheet.getRange(i + 1, 6).setValue(JSON.stringify(notes));
      return { status: 'success' };
    }
  }
  return { status: 'not_found' };
}

function deleteNoteFromEmployee(payload) {
  const { empId, noteIndex } = payload;
  const sheet = SpreadsheetApp.openById(BISCOTTI_SHEET_ID).getSheetByName('ניהול עובדים');
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][2] && rows[i][2].toString() === empId.toString()) {
      let notes = rows[i][5] ? JSON.parse(rows[i][5]) : [];
      notes.splice(noteIndex, 1);
      sheet.getRange(i + 1, 6).setValue(JSON.stringify(notes));
      return { status: 'success' };
    }
  }
  return { status: 'not_found' };
}

// ─── Inventory ────────────────────────────────────────────────────────────────

function updateInventory(payload) {
  const { id, qty, min } = payload;
  const sheet    = getOrCreateSheet(SpreadsheetApp.openById(BISCOTTI_SHEET_ID), 'מלאי');
  const data     = sheet.getDataRange().getValues();
  const ts       = new Date().toLocaleString('he-IL');
  let foundRow   = -1;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString() === id.toString()) { foundRow = i + 1; break; }
  }

  if (foundRow > -1) {
    if (qty !== undefined && qty !== null) sheet.getRange(foundRow, 2).setValue(qty);
    if (min !== undefined && min !== null) sheet.getRange(foundRow, 4).setValue(min);
    sheet.getRange(foundRow, 3).setValue(ts);
  } else {
    sheet.appendRow([id, qty || 0, ts, min || 0]);
  }
  return { status: 'success' };
}

// ─── Tasks ────────────────────────────────────────────────────────────────────

function updateTask(task) {
  const sheet  = getOrCreateSheet(SpreadsheetApp.openById(BISCOTTI_SHEET_ID), 'משימות לדוד');
  const data   = sheet.getDataRange().getValues();
  let foundRow = -1;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === task.name) { foundRow = i + 1; break; }
  }

  if (foundRow > -1) { sheet.getRange(foundRow, 2).setValue(task.done); }
  else               { sheet.appendRow([task.name, task.done]); }
  return { status: 'success' };
}

function resetAllTasks() {
  const sheet = SpreadsheetApp.openById(BISCOTTI_SHEET_ID).getSheetByName('משימות לדוד');
  if (!sheet) return { status: 'error' };
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 2, lastRow - 1, 1).setValues(Array(lastRow - 1).fill([0]));
  }
  return { status: 'success' };
}

function updateTaskPersonal(task) {
  const sheet  = getOrCreateSheet(SpreadsheetApp.openById(BISCOTTI_SHEET_ID), 'משימות אישיות');
  const data   = sheet.getDataRange().getValues();
  let foundRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === task.name) { foundRow = i + 1; break; }
  }
  if (foundRow > -1) { sheet.getRange(foundRow, 2).setValue(task.done); }
  else               { sheet.appendRow([task.name, task.done]); }
  return { status: 'success' };
}

function resetAllTasksPersonal() {
  const sheet = SpreadsheetApp.openById(BISCOTTI_SHEET_ID).getSheetByName('משימות אישיות');
  if (!sheet) return { status: 'error' };
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 2, lastRow - 1, 1).setValues(Array(lastRow - 1).fill([0]));
  }
  return { status: 'success' };
}

// ─── Yerakot Sync ─────────────────────────────────────────────────────────────

// מחזיר את ה-IDs משני הגיליונות לצורך בדיקה
function debugSync() {
  const yerakotRows  = SpreadsheetApp.openById(YERAKOT_SHEET_ID).getSheetByName('inventory').getDataRange().getValues();
  const biscottiRows = SpreadsheetApp.openById(BISCOTTI_SHEET_ID).getSheetByName('מלאי').getDataRange().getValues();
  return {
    yerakot:  yerakotRows.map(r => ({ id: r[0], qty: r[1] })),
    biscotti: biscottiRows.map(r => ({ id: r[0], qty: r[1] }))
  };
}

const YERAKOT_SHEET_ID  = '1GLkHEZXCupy8__Gn173Lqa1CSB25Xwx2evHFsrfcm0k';
const BISCOTTI_SHEET_ID = '17PNmqOGJrKwqgjUkwFGzGnEBHIDn9UfowLtNflJvYiU';

// מסנכרן כמויות מספירת עובד (yerakot) לגיליון המלאי של ביסקוטי
function syncFromYerakot() {
  const yerakotSheet = SpreadsheetApp.openById(YERAKOT_SHEET_ID).getSheetByName('inventory');
  if (!yerakotSheet) return { status: 'error', message: 'yerakot inventory sheet not found' };

  const yerakotRows = yerakotSheet.getDataRange().getValues(); // [id, qty, timestamp]
  const biscottiSheet = getOrCreateSheet(SpreadsheetApp.openById(BISCOTTI_SHEET_ID), 'מלאי');
  const biscottiRows  = biscottiSheet.getDataRange().getValues();

  let updated = 0;

  for (let y = 1; y < yerakotRows.length; y++) {
    const id  = yerakotRows[y][0];
    const qty = Number(yerakotRows[y][1]);
    const ts  = yerakotRows[y][2] || new Date().toLocaleString('he-IL');

    let found = false;
    for (let b = 1; b < biscottiRows.length; b++) {
      if (biscottiRows[b][0] && biscottiRows[b][0].toString() === id.toString()) {
        biscottiSheet.getRange(b + 1, 2).setValue(qty);
        biscottiSheet.getRange(b + 1, 3).setValue(ts);
        found = true;
        updated++;
        break;
      }
    }
    // פריט חדש שלא קיים במלאי ביסקוטי — מוסיף שורה
    if (!found) {
      biscottiSheet.appendRow([id, qty, ts, 0]);
      updated++;
    }
  }

  return { status: 'success', updated };
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function formatDate(date, ss) {
  if (!date) return "";
  if (Object.prototype.toString.call(date) === '[object Date]') {
    return Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
  }
  return String(date);
}
