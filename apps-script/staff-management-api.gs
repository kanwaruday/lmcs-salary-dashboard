// ═══════════════════════════════════════════════════════════════════
// LMCS Staff Management API — write-capable Web App for the
// Employee Master workbook (1OjVMUvpLM8JkdAwjmljCtZUI1VUqGLbic36cW9dm0C0).
//
// Deliberately a SEPARATE project from "LMCS Employee Roster Proxy" --
// that one is public-read-only by design (see its own file header); this
// one writes real HR data, so every mutating call is verified against the
// SAME "LMCS Principal Allowlist" sheet principal-allowlist.gs backs,
// re-derived server-side from a real Google ID token -- never trusted
// from the client. Only 'Coordinator' (their own locked campus only) or
// 'Owner' (any campus) can call the write actions; everyone else gets
// "Not authorized," full stop.
//
// Reads for dropdowns (existing employee names/schools) are NOT
// duplicated here -- the frontend calls the existing Employee Roster
// Proxy's ?action=employees for that (already excludes departed staff).
//
// SETUP:
//   1. script.google.com -> New project -> paste this file
//   2. Deploy -> New deployment -> Web App
//      Execute as: Me | Who has access: Anyone
//      (same "reads open, writes gated by verified token" model as
//      principal-allowlist.gs -- Apps Script's own access setting isn't
//      the real gate here)
//   3. Copy the Web App URL into staff/add-employee.html's
//      STAFF_API_URL constant
// ═══════════════════════════════════════════════════════════════════

const STAFF_EMP_SHEET_ID = '1OjVMUvpLM8JkdAwjmljCtZUI1VUqGLbic36cW9dm0C0';
const STAFF_ALLOWLIST_SHEET_ID = '1NZu0ElismFytG395Nxjz29vAz7OfkmJtZhs70bOwT58'; // "LMCS Principal Allowlist"
const STAFF_GOOGLE_CLIENT_ID = '697999989724-mvi85iobr20g4mm8a8nrjd1rms2o8tf6.apps.googleusercontent.com';

// EmployeeCode prefix per campus -- matches EMP_PREFIX_TO_SCHOOL in
// employee-roster.gs (that one maps prefix -> school; this is the reverse,
// keyed by campusId the portal already uses, e.g. 'LMS1' not 'LMS 1').
const STAFF_SCHOOL_PREFIX = {
  LMS1: 'KUL', LMS2: 'KEL', LMS3: 'DUN', LMS4: 'NCM', LMS5: 'SAY', LMS6: 'JOG',
};

function doGet(e) {
  try {
    const action = (e.parameter.action || '').toLowerCase();
    const caller = verifyStaffManagerToken_(e.parameter.idToken);
    if (!caller) return staffJsonOut_({ success: false, error: 'Not authorized' });

    const data = e.parameter.data ? JSON.parse(e.parameter.data) : {};
    let result;
    if (action === 'nextcode') result = { employeeCode: previewNextCode_(data) };
    else if (action === 'addnewhire') result = addNewHire_(data, caller);
    else if (action === 'transfer') result = transferEmployee_(data, caller);
    else if (action === 'markinactive') result = markInactive_(data, caller);
    else return staffJsonOut_({ success: false, error: 'Unknown action: ' + action });

    return staffJsonOut_(Object.assign({ success: true }, result));
  } catch (err) {
    return staffJsonOut_({ success: false, error: err.message });
  }
}

// ── Auth: verified Google ID token -> {email, campusId, role}, re-derived
//    from the allowlist sheet every call (never trusts client-passed
//    role/campusId). Returns null if the token doesn't verify, or the
//    email isn't on the list, or their role isn't Coordinator/Owner. ──
function verifyStaffManagerToken_(idToken) {
  if (!idToken) return null;
  try {
    const res = UrlFetchApp.fetch(
      'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
      { muteHttpExceptions: true }
    );
    if (res.getResponseCode() !== 200) return null;
    const payload = JSON.parse(res.getContentText());
    if (payload.aud !== STAFF_GOOGLE_CLIENT_ID) return null;
    if (payload.email_verified !== 'true' && payload.email_verified !== true) return null;
    const email = (payload.email || '').toLowerCase();

    const rows = SpreadsheetApp.openById(STAFF_ALLOWLIST_SHEET_ID).getSheets()[0].getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0] || '').trim().toLowerCase() !== email) continue;
      const campusId = String(rows[i][2] || '').trim().toUpperCase();
      const role = String(rows[i][3] || '').trim();
      if (role !== 'Coordinator' && role !== 'Owner' && campusId !== 'ALL') return null;
      return { email: email, campusId: campusId, role: role };
    }
    return null; // not on the allowlist at all
  } catch (err) {
    return null;
  }
}

// A Coordinator may only act on their own locked campus; Owner (campusId
// ALL) may act on any. For a single-campus action (new hire, mark
// inactive) `schools` has one entry, which must equal the caller's campus.
function assertScope_(caller, schools) {
  if (caller.campusId === 'ALL') return;
  schools.forEach(function (s) {
    if (s !== caller.campusId) throw new Error('Coordinators can only manage their own campus');
  });
}

// Transfers touch TWO campuses (old + new) -- a Coordinator locked to one
// campus can never match BOTH, so assertScope_'s all-must-match rule would
// make every cross-campus transfer impossible for them, which defeats the
// point (Uday, 2026-08-29: transfers must work for Coordinators from day
// one). Instead: a Coordinator may act if EITHER side of the transfer is
// their own campus -- releasing one of their own staff, or receiving one
// -- but not for a transfer between two campuses neither of which is
// theirs. Owner (campusId ALL) is unrestricted either way.
function assertTransferScope_(caller, oldSchool, newSchool) {
  if (caller.campusId === 'ALL') return;
  if (caller.campusId !== oldSchool && caller.campusId !== newSchool) {
    throw new Error('Coordinators can only manage transfers involving their own campus');
  }
}

function openEmpWorkbook_() {
  return SpreadsheetApp.openById(STAFF_EMP_SHEET_ID);
}

// ── EmployeeCode generation ──────────────────────────────────────────
// Next-unused sequence number for a school prefix, computed as
// MAX(existing seq for that prefix, including departed/transferred
// employees) + 1 -- a number is never reused, so an old payroll/salary
// record referencing a departed employee's code can never collide with a
// new hire's. Per Uday (2026-08-29): the YY/MM segment is normally the
// employee's real join date, EXCEPT on a transfer between LMS campuses,
// where it carries over from their original code unchanged (that's how
// movement between campuses is tracked) -- see transferEmployee_ below.
function nextSeqForPrefix_(prefix) {
  const rows = openEmpWorkbook_().getSheetByName('EmpMaster').getDataRange().getValues();
  const header = rows[0];
  const codeCol = header.indexOf('EmployeeCode');
  let maxSeq = 0;
  for (let i = 1; i < rows.length; i++) {
    const code = String(rows[i][codeCol] || '').trim();
    if (!code) continue;
    const parts = code.split('/');
    if (parts[0] !== prefix || parts.length < 4) continue;
    const seq = parseInt(parts[3], 10);
    if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
  }
  return maxSeq + 1;
}

function buildEmployeeCode_(prefix, yy, mm) {
  return prefix + '/' + yy + '/' + mm + '/' + String(nextSeqForPrefix_(prefix)).padStart(3, '0');
}

// Preview-only (no write) so the frontend can show the code that WOULD be
// assigned before the user confirms.
function previewNextCode_(data) {
  const prefix = STAFF_SCHOOL_PREFIX[data.school];
  if (!prefix) throw new Error('Unknown school: ' + data.school);
  const doj = new Date(data.dateOfJoining);
  if (isNaN(doj.getTime())) throw new Error('Invalid dateOfJoining');
  const yy = String(doj.getFullYear()).slice(-2);
  const mm = String(doj.getMonth() + 1).padStart(2, '0');
  return buildEmployeeCode_(prefix, yy, mm);
}

// ── Class/Subject label <-> EmpAcademic code (reverse of empClassLabel /
//    empSubjectLabel in employee-roster.gs) ──────────────────────────
const STAFF_SUBJECT_TO_CODE = {
  'Social Science': 'SocialScience',
  'Cognitive Activities': 'CognitiveActivities',
  'E.V.S.': 'EVS',
  'GK': 'GK',
  'Political Science': 'PolSc',
  'Business Studies': 'BussStud',
  'P.Ed': 'PEd',
};
function classToCode_(label) {
  const l = String(label || '').trim();
  if (l === 'M-I') return 'M1';
  if (l === 'M-II') return 'M2';
  if (l === 'M-III') return 'M3';
  const m = l.match(/^Class\s*(\d+)$/i);
  return m ? 'C' + m[1] : l;
}
function subjectToCode_(label) {
  return STAFF_SUBJECT_TO_CODE[label] || String(label || '').replace(/[^A-Za-z0-9]/g, '');
}
function classSubjectCode_(cls, subj) {
  return classToCode_(cls) + '_' + subjectToCode_(subj);
}

// ── Row builders -- always by HEADER NAME, never a hardcoded column
//    index, so a future column reorder in EmpMaster/EmpAcademic can't
//    silently misalign a write (see fixMisalignedRows in Code.gs for
//    exactly the kind of past incident this guards against). Fields not
//    passed in `values` are left blank, e.g. AuthEmail/OldSystemID/
//    NewEmployeeCode -- the latter two are reserved for a future
//    database migration and deliberately untouched here. ──
function buildRow_(header, values) {
  const row = new Array(header.length).fill('');
  header.forEach(function (h, i) {
    if (Object.prototype.hasOwnProperty.call(values, h)) row[i] = values[h];
  });
  return row;
}

function appendAcademicRow_(ss, employeeCode, name, classSubjects) {
  if (!classSubjects || !classSubjects.length) return;
  const sheet = ss.getSheetByName('EmpAcademic');
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const codeCol = header.indexOf('EmployeeCode');
  const nameCol = header.indexOf('Name');
  const subjectCols = [];
  for (let i = 1; i <= 10; i++) {
    const idx = header.indexOf('Class_Subject' + i);
    if (idx >= 0) subjectCols.push(idx);
  }
  const row = new Array(header.length).fill('');
  if (codeCol >= 0) row[codeCol] = employeeCode;
  if (nameCol >= 0) row[nameCol] = name;
  classSubjects.slice(0, subjectCols.length).forEach(function (cs, i) {
    row[subjectCols[i]] = classSubjectCode_(cs.class, cs.subject);
  });
  sheet.appendRow(row);
}

// ── Actions ───────────────────────────────────────────────────────────

/** data: {name, dateOfJoining ('YYYY-MM-DD'), school (campusId), role,
 *  reportsTo (EmployeeCode), classSubjects: [{class,subject}, ...]} */
function addNewHire_(data, caller) {
  if (!data.name || !data.dateOfJoining || !data.school) {
    throw new Error('name, dateOfJoining, and school are required');
  }
  assertScope_(caller, [data.school]);
  const prefix = STAFF_SCHOOL_PREFIX[data.school];
  if (!prefix) throw new Error('Unknown school: ' + data.school);

  const doj = new Date(data.dateOfJoining);
  if (isNaN(doj.getTime())) throw new Error('Invalid dateOfJoining');
  const yy = String(doj.getFullYear()).slice(-2);
  const mm = String(doj.getMonth() + 1).padStart(2, '0');
  const code = buildEmployeeCode_(prefix, yy, mm);

  const ss = openEmpWorkbook_();
  const master = ss.getSheetByName('EmpMaster');
  const header = master.getRange(1, 1, 1, master.getLastColumn()).getValues()[0];
  master.appendRow(buildRow_(header, {
    DateOfJoining: doj, SchoolCode: data.school, EmployeeCode: code,
    Name: data.name, Role: data.role || '', ReportsTo: data.reportsTo || '', Status: 'Active',
  }));
  appendAcademicRow_(ss, code, data.name, data.classSubjects);

  return { employeeCode: code };
}

/** data: {oldEmployeeCode, newSchool (campusId), role, reportsTo,
 *  classSubjects: [{class,subject}, ...]}
 *  Original DateOfJoining and the YY/MM segment of the code carry over
 *  unchanged -- that's how movement between campuses is tracked, per
 *  Uday (2026-08-29). The old row is marked Transferred (excluded from
 *  the roster the same way Inactive is, via employee-roster.gs's Status
 *  filter -- no separate change needed there) and stamped with today's
 *  date in DoRel; a new row is created at the new campus with a fresh
 *  sequence number. */
function transferEmployee_(data, caller) {
  if (!data.oldEmployeeCode || !data.newSchool) {
    throw new Error('oldEmployeeCode and newSchool are required');
  }
  const ss = openEmpWorkbook_();
  const master = ss.getSheetByName('EmpMaster');
  const rows = master.getDataRange().getValues();
  const header = rows[0];
  const codeCol = header.indexOf('EmployeeCode');
  const nameCol = header.indexOf('Name');
  const dojCol = header.indexOf('DateOfJoining');
  const schoolCol = header.indexOf('SchoolCode');
  const statusCol = header.indexOf('Status');
  const dorelCol = header.indexOf('DoRel');

  let oldRowNum = -1, oldName = '', oldDoj = null, oldSchool = '';
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][codeCol] || '').trim() === data.oldEmployeeCode) {
      oldRowNum = i + 1;
      oldName = rows[i][nameCol];
      oldDoj = rows[i][dojCol];
      oldSchool = String(rows[i][schoolCol] || '').trim().toUpperCase();
      break;
    }
  }
  if (oldRowNum === -1) throw new Error('Employee not found: ' + data.oldEmployeeCode);
  assertTransferScope_(caller, oldSchool, data.newSchool);

  const newPrefix = STAFF_SCHOOL_PREFIX[data.newSchool];
  if (!newPrefix) throw new Error('Unknown school: ' + data.newSchool);
  const oldParts = data.oldEmployeeCode.split('/');
  const newCode = buildEmployeeCode_(newPrefix, oldParts[1], oldParts[2]);

  if (statusCol >= 0) master.getRange(oldRowNum, statusCol + 1).setValue('Transferred');
  if (dorelCol >= 0) master.getRange(oldRowNum, dorelCol + 1).setValue(new Date());

  master.appendRow(buildRow_(header, {
    DateOfJoining: oldDoj, SchoolCode: data.newSchool, EmployeeCode: newCode,
    Name: oldName, Role: data.role || '', ReportsTo: data.reportsTo || '', Status: 'Active',
  }));
  appendAcademicRow_(ss, newCode, oldName, data.classSubjects);

  return { employeeCode: newCode };
}

/** data: {employeeCode, dateOfResignation ('YYYY-MM-DD', optional),
 *  dateOfRelieving ('YYYY-MM-DD', optional)}. Never deletes the row --
 *  preserves salary/incentive history that may still reference this
 *  EmployeeCode. */
function markInactive_(data, caller) {
  if (!data.employeeCode) throw new Error('employeeCode is required');
  const master = openEmpWorkbook_().getSheetByName('EmpMaster');
  const rows = master.getDataRange().getValues();
  const header = rows[0];
  const codeCol = header.indexOf('EmployeeCode');
  const schoolCol = header.indexOf('SchoolCode');
  const statusCol = header.indexOf('Status');
  const doresCol = header.indexOf('DoRes');
  const dorelCol = header.indexOf('DoRel');

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][codeCol] || '').trim() !== data.employeeCode) continue;
    const school = String(rows[i][schoolCol] || '').trim().toUpperCase();
    assertScope_(caller, [school]);
    const rowNum = i + 1;
    if (statusCol >= 0) master.getRange(rowNum, statusCol + 1).setValue('Inactive');
    if (doresCol >= 0 && data.dateOfResignation) {
      const d = new Date(data.dateOfResignation);
      if (!isNaN(d.getTime())) master.getRange(rowNum, doresCol + 1).setValue(d);
    }
    if (dorelCol >= 0 && data.dateOfRelieving) {
      const d = new Date(data.dateOfRelieving);
      if (!isNaN(d.getTime())) master.getRange(rowNum, dorelCol + 1).setValue(d);
    }
    return { success: true };
  }
  throw new Error('Employee not found: ' + data.employeeCode);
}

function staffJsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
