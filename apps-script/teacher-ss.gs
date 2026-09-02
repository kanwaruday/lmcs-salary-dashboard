// ═══════════════════════════════════════════════════════════════════
// Teacher SS — dashboard stats for the "Teacher SS" (Support Session
// teacher-evaluation rubric) part of the Principal's Daily Reporting
// hub. One concern-specific file inside the single "LMCS Principal's
// Daily Reporting Backend" project -- see main.gs's header for the
// full project layout and naming convention. doGet lives in main.gs,
// which calls teacherSsStats_() below for action=teacherstats.
//
// Reads "LMCS Teacher SS 2026 (Responses)", the response sheet for the
// 6 campus "LMS N Teacher SS (Qualitative) 2026 for Heads" Google
// Forms -- those Forms stay the real submission mechanism (confirmed
// by Uday 2026-09-02, correcting an earlier "build fresh, no Forms"
// direction); this file only aggregates their responses for the
// portal's Dashboard tab. Never writes to the sheet.
//
// TERMINOLOGY (2026-09-02): "DR" (Daily Report) for teachers was
// renamed to "SS" (Support Session) -- this file, the 6 Forms, and the
// response sheet's Drive file title all now say SS. The sheet's
// internal TAB names below (TSS_CAMPUS_TO_TAB) still say "Teacher DR"
// deliberately -- Google Forms doesn't rename an already-linked
// destination tab when you rename the form, and renaming the tabs
// themselves isn't reachable with the tools used for everything else
// here. Don't "fix" TSS_CAMPUS_TO_TAB to say SS without first renaming
// the actual tabs in the sheet, or every read here breaks.
// ═══════════════════════════════════════════════════════════════════

const TSS_RESPONSES_SHEET_ID = '1cP-f8ShSJJPMBXkQCfkSvmajPA1db6fMBOPCbLH3oIc'; // "LMCS Teacher SS 2026 (Responses)"

// campusId (as used by the portal, e.g. 'LMS1') -> the exact tab name
// in the responses sheet. Still "Teacher DR" -- see TERMINOLOGY above.
const TSS_CAMPUS_TO_TAB = {
  LMS1: 'LMS 1 Teacher DR', LMS2: 'LMS 2 Teacher DR', LMS3: 'LMS 3 Teacher DR',
  LMS4: 'LMS 4 Teacher DR', LMS5: 'LMS 5 Teacher DR', LMS6: 'LMS 6 Teacher DR',
};

// Same 10 rubric columns every tab has, in sheet order — must match
// ss-forms-sync.gs's rubricTitles / the sheet's real headers exactly.
const TSS_RUBRIC_COLUMNS = [
  'Pedagogy', 'Teacher Content', 'Student Response', 'Activity Based Approach',
  'Notebook Checking of Students', 'BUTCHER Scheme',
  'Discipline- English, Aerobics, Morning Assembly & Dispersal',
  'Extra Curricular (Non Academic Competitions)', 'Attitude Towards School',
  'Dedication Towards Class',
];

/** Called from main.gs's doGet for action=teacherstats. Caller's own
 *  campus only, unless Owner (sees all 6). */
function teacherSsStats_(caller) {
  const campuses = caller.campusId === 'ALL' ? Object.keys(TSS_CAMPUS_TO_TAB) : [caller.campusId];
  const byCampus = {};
  campuses.forEach(function (campusId) {
    byCampus[campusId] = readCampusStats_(campusId);
  });
  return { success: true, rubricParams: TSS_RUBRIC_COLUMNS, campuses: byCampus };
}

/** One campus's tab -> per-teacher aggregates: submission count,
 *  average total score (sum of the 10 rubric columns, 0-100), average
 *  per rubric parameter, and most recent submission date. */
function readCampusStats_(campusId) {
  const tabName = TSS_CAMPUS_TO_TAB[campusId];
  const sheet = SpreadsheetApp.openById(TSS_RESPONSES_SHEET_ID).getSheetByName(tabName);
  if (!sheet) return { teachers: [], responseCount: 0 };

  const values = sheet.getDataRange().getValues();
  const header = values[0];
  const teacherNameCol = header.indexOf(campusId.replace('LMS', 'LMS ') + ' Teacher Name');
  const timestampCol = header.indexOf('Timestamp');
  const rubricCols = TSS_RUBRIC_COLUMNS.map(function (title) { return header.indexOf(title); });

  const byTeacher = {}; // teacherName -> {count, sumTotal, sumPerParam[10], lastDate}
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const teacher = String(row[teacherNameCol] || '').trim();
    if (!teacher) continue;

    if (!byTeacher[teacher]) {
      byTeacher[teacher] = { count: 0, sumTotal: 0, sumPerParam: TSS_RUBRIC_COLUMNS.map(function () { return 0; }), lastDate: null };
    }
    const entry = byTeacher[teacher];
    entry.count++;

    let rowTotal = 0;
    rubricCols.forEach(function (col, idx) {
      const v = col >= 0 ? Number(row[col]) || 0 : 0;
      rowTotal += v;
      entry.sumPerParam[idx] += v;
    });
    entry.sumTotal += rowTotal;

    const ts = row[timestampCol];
    if (ts && (!entry.lastDate || new Date(ts) > new Date(entry.lastDate))) {
      entry.lastDate = ts instanceof Date ? ts.toISOString() : String(ts);
    }
  }

  const teachers = Object.keys(byTeacher).map(function (name) {
    const e = byTeacher[name];
    return {
      teacher: name,
      submissionCount: e.count,
      avgTotal: e.count ? Math.round((e.sumTotal / e.count) * 10) / 10 : 0, // out of 100
      avgPerParam: e.sumPerParam.map(function (sum) { return e.count ? Math.round((sum / e.count) * 10) / 10 : 0; }),
      lastSubmission: e.lastDate,
    };
  });
  teachers.sort(function (a, b) { return a.teacher.localeCompare(b.teacher); });

  return { teachers: teachers, responseCount: values.length - 1 };
}
