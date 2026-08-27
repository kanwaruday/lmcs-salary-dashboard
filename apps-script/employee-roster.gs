// ═══════════════════════════════════════════════════════════════════
// Employee Roster Proxy — narrow, public-read Web App in front of the
// Employee Master workbook (1OjVMUvpLM8JkdAwjmljCtZUI1VUqGLbic36cW9dm0C0).
//
// WHY THIS EXISTS (2026-08-27): that workbook is currently shared
// "anyone with the link" so the portal's client-side JS can read it
// directly — but Google Sheets sharing is per-FILE, not per-tab, and
// the same workbook also has EmpPersonal/EmpProfessional/EmpKeyNumbers
// tabs with home addresses, DOB, phone numbers, Aadhar, PAN, and bank
// account details for all 158 employees. Public sharing on the file
// exposes all of that, not just the two tabs the portal actually needs.
//
// The fix: this script runs AS THE OWNER (Execute as: Me), so it can
// read the workbook even after its sharing is locked down to fully
// private. It reads ONLY EmpMaster and EmpAcademic, and returns ONLY
// {employeeCode, name, school, class, subject} rows -- never anything
// from the PII tabs. Once this is deployed and the workbook's sharing
// is restricted, the raw workbook becomes unreachable to the public;
// only this filtered subset is.
//
// Reads are public (no ID token) -- same model as every other read in
// this app (Daily sheet, Course Mapping, the old Allocation sheet were
// all directly public before this). Nothing this endpoint returns is
// more sensitive than "so-and-so teaches Class 3 Science at LMS 6",
// which was already visible in the (also public) old Allocation Sheet.
//
// SETUP:
//   1. script.google.com -> New project -> paste this file
//   2. Deploy -> New deployment -> Web App
//      Execute as: Me | Who has access: Anyone
//   3. Copy the Web App URL into course-mapping-admin.html's
//      "Daily Sheet & Proxy" config (Employee Roster Proxy URL field)
//   4. ONLY THEN: go to the Employee Master sheet -> Share -> Restricted
//      (remove "Anyone with the link"). Do this AFTER step 2/3 are
//      confirmed working, not before -- verify with curl that the
//      deployed URL still returns real data once sharing is locked
//      down, since Apps Script owner-execution should be unaffected,
//      but confirm before assuming.
// ═══════════════════════════════════════════════════════════════════

const EMP_ROSTER_SHEET_ID = '1OjVMUvpLM8JkdAwjmljCtZUI1VUqGLbic36cW9dm0C0';

// EmployeeCode prefix -> School Code, matching the LMS Campuses convention
// used throughout the portal. Confirmed against the live sheet 2026-08-27.
const EMP_PREFIX_TO_SCHOOL = {
  KUL: 'LMS 1', KEL: 'LMS 2', DUN: 'LMS 3', NCM: 'LMS 4', SAY: 'LMS 5', JOG: 'LMS 6',
};

// C1..C12 -> "Class 1".."Class 12", M1..M3 -> "M-I".."M-III" -- matches
// mapDailyClass() in daily-progress.html exactly, so the client needs no
// translation step for what this endpoint returns.
function empClassLabel(code) {
  const m = String(code).match(/^C(\d+)$/);
  if (m) return 'Class ' + m[1];
  if (code === 'M1') return 'M-I';
  if (code === 'M2') return 'M-II';
  if (code === 'M3') return 'M-III';
  return code;
}

// EmpAcademic's no-space subject codes -> the spelling already used
// elsewhere in the portal (Allocation Sheet, Course Mapping). Every code
// actually observed in the sheet as of 2026-08-27 is listed; anything
// else passes through unchanged (visible as-is rather than silently
// dropped, so a newly added subject code doesn't just disappear).
const EMP_SUBJECT_LABEL = {
  SocialScience: 'Social Science',
  CognitiveActivities: 'Cognitive Activities',
  EVS: 'E.V.S.',
  GK: 'GK',
  PolSc: 'Political Science',
  BussStud: 'Business Studies',
  PEd: 'P.Ed',
};
function empSubjectLabel(code) {
  return EMP_SUBJECT_LABEL[code] || code;
}

function doGet(e) {
  const action = (e.parameter.action || 'roster').toLowerCase();
  try {
    if (action === 'roster') return jsonOut({ success: true, rows: readRosterRows() });
    if (action === 'employees') return jsonOut({ success: true, employees: readEmployees() });
    return jsonOut({ success: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonOut({ success: false, error: err.message });
  }
}

function openWorkbook_() {
  return SpreadsheetApp.openById(EMP_ROSTER_SHEET_ID);
}

/** EmpMaster: EmployeeCode -> {name, school}. Never reads AuthEmail/ReportsTo
 *  or any other tab -- deliberately narrow. */
function readEmployees() {
  const rows = openWorkbook_().getSheetByName('EmpMaster').getDataRange().getValues();
  const header = rows[0];
  const codeCol = header.indexOf('EmployeeCode');
  const nameCol = header.indexOf('Name');
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const code = String(rows[i][codeCol] || '').trim();
    if (!code) continue;
    const prefix = code.split('/')[0];
    out.push({
      employeeCode: code,
      name: String(rows[i][nameCol] || '').trim(),
      school: EMP_PREFIX_TO_SCHOOL[prefix] || null,
    });
  }
  return out;
}

/** EmpAcademic, flattened to one row per (school, class, subject, teacher) --
 *  same shape as the old Allocation Sheet's rows, so it's a drop-in
 *  replacement in the client's matching logic. */
function readRosterRows() {
  const rows = openWorkbook_().getSheetByName('EmpAcademic').getDataRange().getValues();
  const header = rows[0];
  const codeCol = header.indexOf('EmployeeCode');
  const nameCol = header.indexOf('Name');
  const subjectCols = [];
  for (let i = 1; i <= 10; i++) {
    const idx = header.indexOf('Class_Subject' + i);
    if (idx >= 0) subjectCols.push(idx);
  }

  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const code = String(rows[i][codeCol] || '').trim();
    if (!code) continue;
    const prefix = code.split('/')[0];
    const school = EMP_PREFIX_TO_SCHOOL[prefix];
    if (!school) continue; // unknown prefix -- skip rather than mis-attribute
    const name = String(rows[i][nameCol] || '').trim();

    subjectCols.forEach(function (col) {
      const val = String(rows[i][col] || '').trim();
      if (!val) return;
      const m = val.match(/^([A-Za-z]+\d+)_(.+)$/);
      if (!m) return; // shouldn't happen -- confirmed 0 malformed cells 2026-08-27
      out.push({
        employeeCode: code,
        school: school,
        class: empClassLabel(m[1]),
        subject: empSubjectLabel(m[2]),
        teacher: name,
      });
    });
  }
  return out;
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
