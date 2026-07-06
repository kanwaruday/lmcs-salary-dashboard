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
  INTERVIEW_AT: 13, // added for the "Interview Scheduled" date/time — trailing column, Forms won't touch it
};

// "Interview Report Form (Responses)" — filled in by whoever conducts the
// interview, independent of this app. We only read it, to warn (not block)
// if a report can't be found once an applicant is marked Interviewed.
const INTERVIEW_REPORT_SHEET_ID = '18NKMBGxT2DuGcrL_m-Y0JGcrmi6oqm6qtLB4Jb_d0Ow';
const IR_COL = { NAME: 1, PHONE: 2, BRANCH: 3, REMARKS: 5, PDF: 6 }; // 0-indexed

// Per-campus "Staff Document Submission form (Responses)" sheets — used to
// warn (not block) if required documents are missing once Hired.
const DOC_SHEET_IDS = {
  LMS1: '1PJ2acPOHopbHzzwGes8X4YPWSDBdXnd0zl46zvFCvIg',
  LMS2: '10xxFM1li1-6xzWEBEEPkZIF4MJcilddoVJkEfVbkScM',
  LMS3: '1qdQnXrWWTrgnCt82MSk_fN6Fqaox-EoenuF_XToypAc',
  LMS4: '1x8eFlnFcEKaCwF9HMRan_EQwTgSPq9xHubZ7n5vyR8k',
  LMS5: '1-bhS4eu4OHFhRzllsALtWLRWiGFt-KutbNke3nNlCk8',
  LMS6: '1-3_3BT7zHVJXy7UqxwZuZmpzUJp4uaibHNaA7onI0jU',
};
// Keyword fragments used to find the relevant columns in each doc sheet by
// header text, since the real sheets have ~24 document columns spread
// across two merged form revisions rather than one clean set. A document
// counts as present if ANY column matching the keyword is non-empty.
const REQUIRED_DOC_KEYWORDS = {
  'Bio-Data': ['BIODATA', 'BIO-DATA', 'BIO DATA'],
  'Hiring Slip': ['HIRING SLIP'],
  'PAN Card': ['PAN CARD'],
  'Aadhar Card': ['AADHAR'],
  'Bank/Cancelled Cheque': ['CANCELLED CHEQUE', 'BANK COPY'],
  'Qualification Certificate': ['BACHELOR', 'MASTER CERTIFICATE', 'PROFESSIONAL DEGREE', 'HIGHEST QUALIFICATION'],
  'Police Verification': ['POLICE VERIFICATION', 'CHARACTER CERTIFICATE'],
  'Medical Certificate': ['MEDICAL CERTIFICATE'],
};

// "LMCS Staff Master" — the network-wide roster. Only ever appended to via
// the explicit "Add to Staff Master" confirm step, never automatically.
const STAFF_MASTER_SHEET_ID = '1OjVMUvpLM8JkdAwjmljCtZUI1VUqGLbic36cW9dm0C0';
const SM_COL = { SYSTEM_ID: 1, DOJ: 2, SCHOOL: 3, EMP_CODE: 4, OLD_ID: 5, NEW_CODE: 6, NAME: 7, STATUS: 8 };
const CAMPUS_CODE_PREFIX = { LMS1: 'KUL', LMS2: 'KEL', LMS3: 'DUN', LMS4: 'NCM', LMS5: 'SAY', LMS6: 'JOG' };

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
      updateStatus(row, e.parameter.status || '', e.parameter.notes || '', e.parameter.interviewAt || '', principal);
      return jsonOut({ success: true, applicants: readApplicants(), principal: principal });
    }

    if (action === 'checkinterviewreport') {
      return jsonOut({ success: true, result: checkInterviewReport(e.parameter.phone || '') });
    }

    if (action === 'checkdocuments') {
      return jsonOut({ success: true, result: checkDocuments(e.parameter.campusId || '', e.parameter.name || '') });
    }

    if (action === 'suggeststaffcode') {
      return jsonOut({ success: true, result: suggestStaffCode(e.parameter.campusId || '', e.parameter.doj || '') });
    }

    if (action === 'addstaffmaster') {
      addStaffMaster(e.parameter.name || '', e.parameter.campusId || '', e.parameter.doj || '', e.parameter.systemId || '', e.parameter.employeeCode || '');
      return jsonOut({ success: true });
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
      interviewAt: r[COL.INTERVIEW_AT - 1] ? new Date(r[COL.INTERVIEW_AT - 1]).toISOString() : '',
    });
  }
  return applicants;
}

function updateStatus(row, status, notes, interviewAt, principal) {
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
  if (interviewAt) sheet.getRange(row, COL.INTERVIEW_AT).setValue(new Date(interviewAt));
}

function normalizePhone(p) {
  const digits = String(p || '').replace(/\D/g, '');
  return digits.length >= 10 ? digits.slice(-10) : digits;
}

function normalizeName(n) {
  return String(n || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

// Looks up the "Interview Report Form (Responses)" sheet by phone number
// (the one reliable shared key — that sheet has no applicant-name
// normalization we can trust). Read-only, used only to warn, never block.
function checkInterviewReport(phone) {
  const target = normalizePhone(phone);
  if (!target) return { found: false };
  const values = SpreadsheetApp.openById(INTERVIEW_REPORT_SHEET_ID).getSheets()[0].getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (normalizePhone(r[IR_COL.PHONE]) === target) {
      return {
        found: true,
        remarks: String(r[IR_COL.REMARKS] || '').trim(),
        pdf: String(r[IR_COL.PDF] || '').trim(),
      };
    }
  }
  return { found: false };
}

// Looks up a campus's "Staff Document Submission" sheet by name (these
// sheets have no phone number column, so name is the only key available —
// matching is best-effort, hence this only ever warns, never blocks).
// Finds required documents by header keyword rather than fixed column
// index, since the real sheets have two merged, inconsistent form
// revisions' worth of document columns.
function checkDocuments(campusId, name) {
  const sheetId = DOC_SHEET_IDS[campusId];
  if (!sheetId) return { found: false, reason: 'No document sheet configured for ' + campusId };
  const values = SpreadsheetApp.openById(sheetId).getDataRange().getValues();
  if (!values.length) return { found: false };
  const header = values[0].map(h => String(h || '').toUpperCase());

  const target = normalizeName(name);
  let row = null;
  for (let i = 1; i < values.length; i++) {
    if (normalizeName(values[i][2]) === target) { row = values[i]; break; } // col C = Employee Name
  }
  if (!row) return { found: false };

  const missing = [];
  Object.keys(REQUIRED_DOC_KEYWORDS).forEach(label => {
    const keywords = REQUIRED_DOC_KEYWORDS[label];
    const matchingCols = header
      .map((h, idx) => (keywords.some(k => h.indexOf(k) !== -1) ? idx : -1))
      .filter(idx => idx !== -1);
    const present = matchingCols.some(idx => String(row[idx] || '').trim() !== '');
    if (!present) missing.push(label);
  });
  return { found: true, complete: missing.length === 0, missing: missing };
}

// Next SystemID/EmployeeCode for a campus — a suggestion only. The
// principal confirms (and can edit) before addStaffMaster actually writes,
// since sequence/padding conventions drift slightly between campuses in
// the real sheet.
function suggestStaffCode(campusId, doj) {
  const prefix = CAMPUS_CODE_PREFIX[campusId];
  if (!prefix) return { error: 'Unknown campus ' + campusId };
  const values = SpreadsheetApp.openById(STAFF_MASTER_SHEET_ID).getSheets()[0].getDataRange().getValues();
  let maxId = 0;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][SM_COL.SCHOOL - 1]).trim() === campusId) {
      const n = parseInt(values[i][SM_COL.SYSTEM_ID - 1], 10);
      if (!isNaN(n) && n > maxId) maxId = n;
    }
  }
  const nextId = maxId + 1;
  const d = doj ? new Date(doj) : new Date();
  const yy = String(d.getFullYear()).slice(-2);
  const mm = ('0' + (d.getMonth() + 1)).slice(-2);
  // LMS1–4 zero-pad the sequence to 3 digits in existing data; LMS5–6 don't.
  const seq = (campusId === 'LMS5' || campusId === 'LMS6') ? String(nextId) : ('00' + nextId).slice(-3);
  return { systemId: nextId, employeeCode: prefix + '/' + yy + '/' + mm + '/' + seq };
}

function addStaffMaster(name, campusId, doj, systemId, employeeCode) {
  if (!name || !campusId || !doj || !systemId || !employeeCode) {
    throw new Error('Name, campus, date of joining, system ID and employee code are all required');
  }
  const sheet = SpreadsheetApp.openById(STAFF_MASTER_SHEET_ID).getSheets()[0];
  sheet.appendRow([Number(systemId), new Date(doj), campusId, employeeCode, '', '', name, '', '']);
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
