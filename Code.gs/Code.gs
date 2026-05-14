// ===== BISSCOTI Manager – Server Router =====
// עדכן את ה-SHEET_ID למזהה ה-Google Sheet שלך
const SHEET_ID      = 'YOUR_SHEET_ID_HERE';
const SHEET_TICKETS = 'tickets';
const SHEET_USERS   = 'users';

// GET router – login / getUsers / loadTickets
function doGet(e) {
  const action = e.parameter.action;
  if (action === 'login')    return login(e.parameter.username, e.parameter.password);
  if (action === 'getUsers') return getUsers();
  return loadTickets();
}

// POST router – ticket CRUD + user management
function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const { action, ticket, user, username } = body;

  if (action === 'ADD')        return addTicket(ticket);
  if (action === 'UPDATE')     return updateTicket(ticket);
  if (action === 'DELETE')     return deleteTicket(ticket.id);
  if (action === 'addUser')    return addUser(user);
  if (action === 'deleteUser') return deleteUser(username);

  return json({ error: 'unknown action: ' + action });
}

// ─── Tickets ──────────────────────────────────────────────────────────────────

function loadTickets() {
  const sheet = getSheet(SHEET_TICKETS);
  const rows  = sheet.getDataRange().getValues();
  const headers = rows[0];
  const tickets = rows.slice(1).map(row =>
    Object.fromEntries(headers.map((h, i) => [h, row[i]]))
  );
  return json({ tickets });
}

function addTicket(t) {
  const sheet   = getSheet(SHEET_TICKETS);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  sheet.appendRow(headers.map(h => t[h] ?? ''));
  return json({ ok: true });
}

function updateTicket(t) {
  const sheet = getSheet(SHEET_TICKETS);
  const data  = sheet.getDataRange().getValues();
  const idCol = data[0].indexOf('id');
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === t.id) {
      data[0].forEach((h, j) => { if (t[h] !== undefined) sheet.getRange(i + 1, j + 1).setValue(t[h]); });
      return json({ ok: true });
    }
  }
  return json({ error: 'ticket not found: ' + t.id });
}

function deleteTicket(id) {
  const sheet = getSheet(SHEET_TICKETS);
  const data  = sheet.getDataRange().getValues();
  const idCol = data[0].indexOf('id');
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === id) { sheet.deleteRow(i + 1); return json({ ok: true }); }
  }
  return json({ error: 'ticket not found: ' + id });
}

// ─── Users ────────────────────────────────────────────────────────────────────

function login(username, password) {
  const users = getUserRows();
  const match = users.find(u => u.username === username && u.password === password);
  if (!match) return json({ error: 'invalid credentials' });
  return json({ user: { username: match.username, name: match.name, role: match.role } });
}

function getUsers() {
  return json({ users: getUserRows().map(u => ({ username: u.username, name: u.name, role: u.role })) });
}

function addUser(user) {
  const sheet   = getSheet(SHEET_USERS);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  sheet.appendRow(headers.map(h => user[h] ?? ''));
  return json({ ok: true });
}

function deleteUser(username) {
  const sheet = getSheet(SHEET_USERS);
  const data  = sheet.getDataRange().getValues();
  const uCol  = data[0].indexOf('username');
  for (let i = 1; i < data.length; i++) {
    if (data[i][uCol] === username) { sheet.deleteRow(i + 1); return json({ ok: true }); }
  }
  return json({ error: 'user not found: ' + username });
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function getSheet(name) {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
}

function getUserRows() {
  const sheet   = getSheet(SHEET_USERS);
  const rows    = sheet.getDataRange().getValues();
  const headers = rows[0];
  return rows.slice(1).map(row => Object.fromEntries(headers.map((h, i) => [h, row[i]])));
}

// מחזיר JSON response תקני
function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
