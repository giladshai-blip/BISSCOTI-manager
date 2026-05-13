const SHEET_NAME = 'מסד נתונים';

function doGet(e) {
  if (e.parameter.action === 'login') return handleLogin(e.parameter.username, e.parameter.password);
  if (e.parameter.action === 'getUsers') return getUsers();

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const tickets = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      if ((h === 'tags' || h === 'managerNotes') && row[i]) {
        try { obj[h] = JSON.parse(row[i]); } catch { obj[h] = []; }
      } else {
        obj[h] = row[i];
      }
    });
    return obj;
  }).filter(t => t.id);

  return ContentService
    .createTextOutput(JSON.stringify(tickets))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const payload = JSON.parse(e.postData.contents);
  const { action, ticket } = payload;

  // ניהול משתמשים
  if (action === 'addUser') {
    const usersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('users');
    usersSheet.appendRow([ticket.username, ticket.password, ticket.name, ticket.role]);
    return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'deleteUser') {
    const usersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('users');
    const usersData = usersSheet.getDataRange().getValues();
    const rowIndex = usersData.findIndex(row => row[0] === ticket.username);
    if (rowIndex > 0) usersSheet.deleteRow(rowIndex + 1);
    return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
  }

  // ניהול קריאות
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  if (action === 'ADD') {
    const newRow = headers.map(h => {
      if (h === 'tags' || h === 'managerNotes') return JSON.stringify(ticket[h] || []);
      return ticket[h] !== undefined ? ticket[h] : '';
    });
    sheet.appendRow(newRow);
  }

  if (action === 'UPDATE') {
    const rowIndex = data.findIndex(row => row[0] === ticket.id);
    if (rowIndex > 0) {
      headers.forEach((h, i) => {
        let val = ticket[h] !== undefined ? ticket[h] : '';
        if (h === 'tags' || h === 'managerNotes') val = JSON.stringify(val || []);
        sheet.getRange(rowIndex + 1, i + 1).setValue(val);
      });
    }
  }

  if (action === 'DELETE') {
    const rowIndex = data.findIndex(row => row[0] === ticket.id);
    if (rowIndex > 0) sheet.deleteRow(rowIndex + 1);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleLogin(username, password) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('users');
  const data = sheet.getDataRange().getValues();
  const user = data.slice(1).find(row =>
    String(row[0]).trim() === username.trim() &&
    String(row[1]).trim() === password.trim()
  );
  if (user) {
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      user: { username: user[0], name: user[2], role: user[3] }
    })).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput(JSON.stringify({ success: false })).setMimeType(ContentService.MimeType.JSON);
}

function getUsers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('users');
  const data = sheet.getDataRange().getValues();
  const users = data.slice(1)
    .filter(row => row[0])
    .map(row => ({ username: row[0], name: row[2], role: row[3] }));
  return ContentService.createTextOutput(JSON.stringify({ users })).setMimeType(ContentService.MimeType.JSON);
}
