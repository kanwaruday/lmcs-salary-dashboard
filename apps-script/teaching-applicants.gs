// ═══════════════════════════════════════════════════════════════════
// Hiring Dashboard — Teaching Applicants API
//
// Standalone Apps Script project (same pattern as principal-allowlist.gs)
// bound to the "Teaching Applicants" Google Form response sheet. Backs
// teaching-applicants.html: lists applicants and lets a signed-in
// principal update an applicant's Status/Notes.
//
// Unlike principal-allowlist.gs, `list` here is NOT open to anyone —
// this sheet holds applicants' phone numbers and CV links, so both
// reads and writes require a Google ID token that verifies to someone
// on the live Principal Allowlist sheet.
//
// SETUP:
//   1. Go to script.google.com → New project → paste this file
//   2. Deploy → New deployment → Web App
//      Execute as: Me  |  Who has access: Anyone
//      (access is public at the Apps Script layer, same as every other
//      deployment in this repo — the real gate is verifyPrincipal below)
//   3. Copy the Web App URL into teaching-applicants.html's API_URL constant
// ═══════════════════════════════════════════════════════════════════

const SHEET_ID = '1aSCQ3IGO-ZP_5yRjtnMpZdlauTdWj3814ATD9qwskPQ'; // "Teaching Applicants"
const SHEET_GID = 800656734; // targets the exact tab regardless of its name
const ALLOWLIST_SHEET_ID = '1NZu0ElismFytG395Nxjz29vAz7OfkmJtZhs70bOwT58'; // "LMCS Principal Allowlist"
const GOOGLE_CLIENT_ID = '697999989724-mvi85iobr20g4mm8a8nrjd1rms2o8tf6.apps.googleusercontent.com';

// Columns in the response sheet (1-indexed)
const COL = {
  TIMESTAMP: 1, NAME: 2, PHONE: 3, AGE: 4, SUBJECTS: 5, BRANCHES: 6,
  QUALIFICATION: 7, BED: 8, GENDER: 9, CV: 10, STATUS: 11, NOTES: 12,
};

function doGet(e) {
  const action = (e.parameter.action || 'list').toLowerCase();
  try {
    const principal = verifyPrincipal(e.parameter.idToken);
    if (!principal) {
      return jsonOut({ success: false, error: 'Not authorized — sign in via Principal Dashboard.' });
    }

    if (action === 'list') {
      return jsonOut({ success: true, applicants: readApplicants(), principal: principal });
    }

    if (action === 'updatestatus') {
      const row = parseInt(e.parameter.row, 10);
      if (!row || row < 2) throw new Error('Invalid row');
      updateStatus(row, e.parameter.status || '', e.parameter.notes || '', principal);
      return jsonOut({ success: true, applicants: readApplicants(), principal: principal });
    }

    return jsonOut({ success: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonOut({ success: false, error: err.message });
  }
}

function getSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheets().find(s => s.getSheetId() === SHEET_GID);
  if (!sheet) throw new Error('Teaching Applicants sheet tab not found');
  return sheet;
}

function readApplicants() {
  const values = getSheet().getDataRange().getValues();
  const applicants = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (!r[COL.NAME - 1]) continue;
    applicants.push({
      row: i + 1, // 1-based sheet row, used to write status/notes back
      timestamp: r[COL.TIMESTAMP - 1] ? new Date(r[COL.TIMESTAMP - 1]).toISOString() : '',
      name: String(r[COL.NAME - 1] || '').trim(),
      phone: String(r[COL.PHONE - 1] || '').trim(),
      age: String(r[COL.AGE - 1] || '').trim(),
      subjects: String(r[COL.SUBJECTS - 1] || '').trim(),
      branches: String(r[COL.BRANCHES - 1] || '').trim(),
      qualification: String(r[COL.QUALIFICATION - 1] || '').trim(),
      bed: String(r[COL.BED - 1] || '').trim(),
      gender: String(r[COL.GENDER - 1] || '').trim(),
      cv: String(r[COL.CV - 1] || '').trim(),
      status: String(r[COL.STATUS - 1] || '').trim(),
      notes: String(r[COL.NOTES - 1] || '').trim(),
    });
  }
  return applicants;
}

function updateStatus(row, status, notes, principal) {
  const sheet = getSheet();
  // Defense in depth: a campus-locked principal may only update applicants
  // who actually applied to their own campus (the client already filters
  // the view this way, but the server re-checks before writing).
  if (principal.campusId !== 'ALL') {
    const branches = String(sheet.getRange(row, COL.BRANCHES).getValue() || '');
    const tag = 'LMS-' + principal.campusId.replace(/[^0-9]/g, '');
    if (branches.indexOf(tag) === -1) {
      throw new Error('This applicant did not apply to your campus');
    }
  }
  sheet.getRange(row, COL.STATUS).setValue(status);
  sheet.getRange(row, COL.NOTES).setValue(notes);
}

// Verifies a Google ID token (signature/audience/expiry via Google's own
// tokeninfo endpoint) and checks the email against the live Principal
// Allowlist sheet. Returns { email, name, campusId } or null.
function verifyPrincipal(idToken) {
  if (!idToken) return null;
  try {
    const res = UrlFetchApp.fetch(
      'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
      { muteHttpExceptions: true }
    );
    if (res.getResponseCode() !== 200) return null;
    const payload = JSON.parse(res.getContentText());
    if (payload.aud !== GOOGLE_CLIENT_ID) return null;
    if (payload.email_verified !== 'true' && payload.email_verified !== true) return null;
    const email = (payload.email || '').toLowerCase();

    const rows = SpreadsheetApp.openById(ALLOWLIST_SHEET_ID).getSheets()[0].getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim().toLowerCase() === email) {
        return { email: email, name: String(rows[i][1] || '').trim(), campusId: String(rows[i][2] || '').trim().toUpperCase() };
      }
    }
    return null;
  } catch (err) {
    return null;
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
