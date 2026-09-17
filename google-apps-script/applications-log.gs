/**
 * BSN Personal Loan — application log (Fasa 3).
 *
 * SETUP (one-time):
 * 1. Create a new blank Google Sheet (e.g. "BSN Applications Log").
 * 2. Extensions -> Apps Script. Paste this whole file in, replacing Code.gs.
 * 3. Change STATUS_ADMIN_PASSWORD and SUBMIT_TOKEN below to values of your
 *    own choosing (not the placeholders).
 * 4. Run `setupSheet` once from the editor (select it in the toolbar
 *    dropdown, click Run) — this creates the header row and prompts Google
 *    to ask for permission (Sheets access). Approve it once.
 * 5. Deploy -> New deployment -> select type "Web app".
 *      - Execute as: Me
 *      - Who has access: Anyone
 *    Deploy, then copy the Web App URL.
 * 6. Give the Web App URL and the SUBMIT_TOKEN you chose (NOT the admin
 *    password) back to whoever is wiring up loan.lateffitazri.my.
 * 7. Use STATUS_ADMIN_PASSWORD yourself to log into status.lateffitazri.my.
 *
 * WHAT IT DOES:
 * - doPost with action=submit + the correct SUBMIT_TOKEN appends a new
 *   application row. This is the only thing the public loan site can do —
 *   the token may end up visible in that site's source, so it intentionally
 *   grants nothing beyond "add one row".
 * - doGet / doPost with action=updateStatus require STATUS_ADMIN_PASSWORD,
 *   checked here on the server — status.lateffitazri.my cannot read or
 *   change anything without it, even if its own source is inspected.
 */

// ── Configuration — change these ────────────────────────────────────────

const STATUS_ADMIN_PASSWORD = 'TUKAR_INI'; // used by status.lateffitazri.my to log in
const SUBMIT_TOKEN = 'TUKAR_INI_JUGA';      // embedded in loan.lateffitazri.my's public source

const SHEET_NAME = 'Applications';
const COLUMNS = ['id', 'timestamp', 'nama', 'produk', 'sektor', 'tarikhMohon', 'status', 'dsrStatus', 'penyatuanHutang'];

// ── One-time setup ──────────────────────────────────────────────────────

function setupSheet() {
  const sheet = getSheet_();
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(COLUMNS);
    sheet.setFrozenRows(1);
  }
}

// ── HTTP entry points ────────────────────────────────────────────────────

// GET is used for everything status.lateffitazri.my needs to READ a response
// from (listing + status updates) — Apps Script Web App GET responses are
// readable cross-origin via fetch(), unlike doPost's (see doPost comment
// below), so "updateStatus" is deliberately routed through doGet too even
// though it's a mutation.
function doGet(e) {
  const password = (e.parameter.password || '').trim();
  if (password !== STATUS_ADMIN_PASSWORD) {
    return jsonResponse_({ ok: false, error: 'Kod rahsia tidak sah.' });
  }
  try {
    if (e.parameter.action === 'updateStatus') {
      return jsonResponse_(handleUpdateStatus_(e.parameter));
    }
    const rows = readAllRows_();
    rows.sort((a, b) => (b.tarikhMohon || '').localeCompare(a.tarikhMohon || ''));
    return jsonResponse_({ ok: true, rows });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

// POST is only used by loan.lateffitazri.my's fire-and-forget "submit" call,
// sent with mode:'no-cors' (so its response is opaque/unreadable — fine,
// since the applicant's flow never needs to read it back).
function doPost(e) {
  const action = e.parameter.action;
  try {
    if (action === 'submit') {
      return jsonResponse_(handleSubmit_(e.parameter));
    }
    return jsonResponse_({ ok: false, error: 'Unknown action.' });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

// ── Handlers ─────────────────────────────────────────────────────────────

function handleSubmit_(p) {
  if ((p.token || '') !== SUBMIT_TOKEN) {
    return { ok: false, error: 'Token tidak sah.' };
  }
  const sheet = getSheet_();
  const id = Utilities.getUuid();
  const now = new Date();
  sheet.appendRow([
    id,
    Utilities.formatDate(now, 'GMT+8', 'yyyy-MM-dd HH:mm:ss'),
    p.nama || '',
    p.produk || '',
    p.sektor || '',
    Utilities.formatDate(now, 'GMT+8', 'yyyy-MM-dd'),
    p.status || 'Mohon',
    p.dsrStatus || '',
    p.penyatuanHutang || '',
  ]);
  return { ok: true, id };
}

function handleUpdateStatus_(p) {
  if ((p.password || '') !== STATUS_ADMIN_PASSWORD) {
    return { ok: false, error: 'Kod rahsia tidak sah.' };
  }
  const id = p.id || '';
  const newStatus = p.status || '';
  if (!id || !newStatus) {
    return { ok: false, error: 'Missing id or status.' };
  }
  const sheet = getSheet_();
  const idCol = COLUMNS.indexOf('id') + 1;
  const statusCol = COLUMNS.indexOf('status') + 1;
  const values = sheet.getRange(2, idCol, Math.max(sheet.getLastRow() - 1, 0), 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === id) {
      sheet.getRange(i + 2, statusCol).setValue(newStatus);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Row not found for id: ' + id };
}

// ── Helpers ──────────────────────────────────────────────────────────────

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(COLUMNS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function readAllRows_() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, COLUMNS.length).getValues();
  return values.map((row) => {
    const obj = {};
    COLUMNS.forEach((col, i) => { obj[col] = row[i]; });
    return obj;
  });
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
