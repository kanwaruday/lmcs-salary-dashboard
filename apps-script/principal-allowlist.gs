// ═══════════════════════════════════════════════════════════════════
// Principal Dashboard — Allowlist API
//
// Standalone Apps Script project, deliberately separate from Code.gs
// (which already has three tangled deployments — see the project's
// vault note). Backs principal-admin.html's add/edit/delete UI and
// principal.html's live "who's allowed in" check.
//
// SETUP:
//   1. Go to script.google.com → New project → paste this file
//   2. Deploy → New deployment → Web App
//      Execute as: Me  |  Who has access: Anyone
//      (same simple mode as Code.gs's other deployments — reads are
//      public by design, writes are gated by verifyAdminToken below,
//      not by Apps Script's own access settings)
//   3. Copy the Web App URL into principal.html and principal-admin.html
//      (ALLOWLIST_API_URL constant in each)
// ═══════════════════════════════════════════════════════════════════

const SHEET_ID = '1NZu0ElismFytG395Nxjz29vAz7OfkmJtZhs70bOwT58'; // "LMCS Principal Allowlist"
const ADMIN_EMAILS = ['uday.kanwar@lms.org.in'];
const GOOGLE_CLIENT_ID = '697999989724-mvi85iobr20g4mm8a8nrjd1rms2o8tf6.apps.googleusercontent.com';

function doGet(e) {
  const action = (e.parameter.action || 'list').toLowerCase();
  try {
    if (action === 'list') {
      return jsonOut({ success: true, entries: readEntries() });
    }

    // Every mutating action requires a Google ID token that verifies,
    // server-side, to one of ADMIN_EMAILS — not just "logged in."
    const email = verifyAdminToken(e.parameter.idToken);
    if (!email) {
      return jsonOut({ success: false, error: 'Not authorized' });
    }

    if (action === 'add') {
      addEntry(e.parameter.email, e.parameter.name, e.parameter.campusId);
    } else if (action === 'edit') {
      editEntry(e.parameter.email, e.parameter.name, e.parameter.campusId);
    } else if (action === 'delete') {
      deleteEntry(e.parameter.email);
    } else {
      return jsonOut({ success: false, error: 'Unknown action: ' + action });
    }

    return jsonOut({ success: true, entries: readEntries() });
  } catch (err) {
    return jsonOut({ success: false, error: err.message });
  }
}

function getSheet() {
  return SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
}

function readEntries() {
  const rows = getSheet().getDataRange().getValues();
  const entries = [];
  for (let i = 1; i < rows.length; i++) {
    const [email, name, campusId] = rows[i];
    if (email) {
      entries.push({
        email: String(email).trim().toLowerCase(),
        name: String(name || '').trim(),
        campusId: String(campusId || '').trim().toUpperCase(),
      });
    }
  }
  return entries;
}

function findRowIndex(sheet, email) {
  const values = sheet.getDataRange().getValues();
  const target = String(email).trim().toLowerCase();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim().toLowerCase() === target) return i + 1; // 1-based row number
  }
  return -1;
}

function addEntry(email, name, campusId) {
  if (!email || !campusId) throw new Error('email and campusId are required');
  const sheet = getSheet();
  if (findRowIndex(sheet, email) !== -1) throw new Error('That email is already on the list — use edit instead');
  sheet.appendRow([email.trim().toLowerCase(), name || '', campusId.trim().toUpperCase()]);
}

function editEntry(email, name, campusId) {
  const sheet = getSheet();
  const row = findRowIndex(sheet, email);
  if (row === -1) throw new Error('Email not found: ' + email);
  sheet.getRange(row, 2).setValue(name || '');
  sheet.getRange(row, 3).setValue((campusId || '').trim().toUpperCase());
}

function deleteEntry(email) {
  const sheet = getSheet();
  const row = findRowIndex(sheet, email);
  if (row === -1) throw new Error('Email not found: ' + email);
  sheet.deleteRow(row);
}

// Verifies a Google Identity Services ID token via Google's own tokeninfo
// endpoint (signature + audience + expiry all checked by Google, not us).
// Returns the verified email if it's an admin, otherwise null.
function verifyAdminToken(idToken) {
  if (!idToken) return null;
  try {
    const res = UrlFetchApp.fetch(
      'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
      { muteHttpExceptions: true }
    );
    if (res.getResponseCode() !== 200) return null;
    const payload = JSON.parse(res.getContentText());
    if (payload.aud !== GOOGLE_CLIENT_ID) return null; // token wasn't issued for our client
    if (payload.email_verified !== 'true' && payload.email_verified !== true) return null;
    const email = (payload.email || '').toLowerCase();
    return ADMIN_EMAILS.indexOf(email) !== -1 ? email : null;
  } catch (err) {
    return null;
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
