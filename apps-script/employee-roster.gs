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
// STATUS FILTERING (2026-08-29): EmpMaster has a Status column (blank or
// "Active" = still employed, anything else e.g. "Left" = departed).
// readDepartedCodes_() reads it once per request and both readEmployees()
// and readRosterRows() exclude anyone marked departed -- so flipping one
// employee's Status in EmpMaster cascades everywhere this proxy feeds
// (Missing Uploads, Course Mapping teacher assignment, the Chapter
// Tracker's teacher lookup) without hunting down every place their name
// might still appear.
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
    if (action === 'designations') return jsonOut({ success: true, designations: readDesignationSummary() });
    return jsonOut({ success: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonOut({ success: false, error: err.message });
  }
}

/** Distinct Designation values from EmpSalary, with counts -- NEVER any
 *  name, email, or other per-employee field. Built 2026-09-01 so Claude
 *  could see the real job-title taxonomy (for designing the
 *  PortalAccessTier mapping) without needing raw access to the workbook,
 *  which is intentionally Restricted -- this is a read of the same narrow
 *  shape as everything else in this file, just aggregated instead of
 *  per-row. */
function readDesignationSummary() {
  const rows = openWorkbook_().getSheetByName('EmpSalary').getDataRange().getValues();
  const header = rows[0];
  const desigCol = header.indexOf('Designation');
  const deptCol = header.indexOf('Department');
  const statusCol = header.indexOf('Status');
  const counts = {}; // "Designation|Department" -> {designation, department, count, active, inactive}
  for (let i = 1; i < rows.length; i++) {
    const designation = String(rows[i][desigCol] || '').trim();
    if (!designation) continue;
    const department = deptCol >= 0 ? String(rows[i][deptCol] || '').trim() : '';
    const status = statusCol >= 0 ? String(rows[i][statusCol] || '').trim() : '';
    const key = designation + '|' + department;
    if (!counts[key]) counts[key] = { designation, department, count: 0, active: 0, inactive: 0 };
    counts[key].count++;
    if (status.toLowerCase() === 'active' || !status) counts[key].active++;
    else counts[key].inactive++;
  }
  return Object.values(counts).sort((a, b) => b.count - a.count);
}

function openWorkbook_() {
  return SpreadsheetApp.openById(EMP_ROSTER_SHEET_ID);
}

/** EmpMaster's Status column -> {EmployeeCode: true} for everyone marked
 *  departed (anything other than blank/"Active", case-insensitively).
 *  Read once per request; shared by readEmployees() and readRosterRows()
 *  below -- see STATUS FILTERING note at the top of this file. */
function readDepartedCodes_() {
  const rows = openWorkbook_().getSheetByName('EmpMaster').getDataRange().getValues();
  const header = rows[0];
  const codeCol = header.indexOf('EmployeeCode');
  const statusCol = header.indexOf('Status');
  const departed = {};
  if (statusCol < 0) return departed; // no Status column -- nothing to exclude
  for (let i = 1; i < rows.length; i++) {
    const code = String(rows[i][codeCol] || '').trim();
    if (!code) continue;
    const status = String(rows[i][statusCol] || '').trim().toLowerCase();
    if (status && status !== 'active') departed[code] = true;
  }
  return departed;
}

/** EmpMaster: EmployeeCode -> {name, school}. Never reads AuthEmail/ReportsTo
 *  or any other tab -- deliberately narrow. Excludes departed staff. */
function readEmployees() {
  const departed = readDepartedCodes_();
  const rows = openWorkbook_().getSheetByName('EmpMaster').getDataRange().getValues();
  const header = rows[0];
  const codeCol = header.indexOf('EmployeeCode');
  const nameCol = header.indexOf('Name');
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const code = String(rows[i][codeCol] || '').trim();
    if (!code || departed[code]) continue;
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
 *  replacement in the client's matching logic. Excludes departed staff'
 *  class/subject rows too, not just their name in the roster list -- a
 *  Status flip in EmpMaster is enough, no need to also clear EmpAcademic. */
function readRosterRows() {
  const departed = readDepartedCodes_();
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
    if (!code || departed[code]) continue;
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
