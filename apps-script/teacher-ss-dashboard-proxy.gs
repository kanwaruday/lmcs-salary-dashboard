// ═══════════════════════════════════════════════════════════════════
// Teacher SS Dashboard Proxy — read-only Web App in front of
// "LMCS Teacher SS 2026 (Responses)" (spreadsheet id
// 1cP-f8ShSJJPMBXkQCfkSvmajPA1db6fMBOPCbLH3oIc), the response sheet
// for the 6 campus "LMS N Teacher SS (Qualitative) 2026 for Heads"
// Google Forms.
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
//
// Per Uday (2026-09-02, correcting an earlier "build fresh, no Forms"
// direction from a different session): the Forms stay as the real
// submission mechanism — this proxy just aggregates their responses
// for the portal's Teaching SS dashboard, same "narrow read proxy in
// front of someone else's Sheet" shape as employee-roster.gs. Never
// writes to the sheet.
//
// GATING: unlike Employee Roster (public), teacher performance scores
// are more sensitive, so this requires a verified Google ID token and
// only serves Principal/Coordinator/Owner — same allowlist + token
// pattern as staff-management-api.gs, just for reads. A caller's own
// campus only, unless Owner (sees all 6).
//
// SETUP:
//   1. script.google.com -> "LMCS Teacher SS Backend" project ->
//      paste this file in as teacher-ss-dashboard-proxy.gs
//   2. Deploy -> New deployment -> Web App
//      Execute as: Me | Who has access: Anyone
//      (Apps Script's own access setting isn't the real gate --
//      the verified-token check inside is)
//   3. Copy the Web App URL into teacher-ss/index.html's
//      DASHBOARD_API_URL constant
// ═══════════════════════════════════════════════════════════════════

const TSS_RESPONSES_SHEET_ID = '1cP-f8ShSJJPMBXkQCfkSvmajPA1db6fMBOPCbLH3oIc'; // "LMCS Teacher SS 2026 (Responses)"
const TSS_ALLOWLIST_SHEET_ID = '1NZu0ElismFytG395Nxjz29vAz7OfkmJtZhs70bOwT58'; // "LMCS Principal Allowlist"
const TSS_GOOGLE_CLIENT_ID = '697999989724-mvi85iobr20g4mm8a8nrjd1rms2o8tf6.apps.googleusercontent.com';

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

// School's OFFICIAL Calendar per campus (2026-09-02, confirmed live/working
// -- same IDs already proven in ChapterTracker.gs's CHAPTER_CALENDAR_IDS and
// in WhatsApp Hub's hub-config.json, just re-keyed to this portal's 'LMS1'..
// 'LMS6' convention instead of ChapterTracker's 1..6). Used by the
// 'monthactivities' action below. LMS2 and LMS3 deliberately share one
// calendar (separate one "coming 2027" per hub-config.json's note).
const TSS_SCHOOL_CALENDAR_IDS = {
  LMS1: 'c_88402c40e6c9a435cbed4810bfae632d781ac9e7ca167b5435e4a9586bda7595@group.calendar.google.com',
  LMS2: 'c_25n5qbggh7u9s5o6vq31iho4vc@group.calendar.google.com',
  LMS3: 'c_25n5qbggh7u9s5o6vq31iho4vc@group.calendar.google.com', // shared with LMS2
  LMS4: 'lms.org.in_snflpd3f357jv8l380s8sc4ju0@group.calendar.google.com',
  LMS5: 'c_d367ca2cb9bbab4f6a5b6864d6e03c28fbfccf02b5212621e89daf2f9f3f8523@group.calendar.google.com',
  LMS6: 'c_3db6cb2edb78d889cf166771c426f0fe3ba78b04cf9ca8e566329c888d023635@group.calendar.google.com',
};

function doGet(e) {
  try {
    const caller = verifySSDashboardToken_(e.parameter.idToken);
    if (!caller) return tssJsonOut_({ success: false, error: 'Not authorized' });

    const action = (e.parameter.action || 'stats').toLowerCase();

    if (action === 'monthactivities') {
      const campusId = String(e.parameter.campusId || '').trim().toUpperCase();
      if (caller.campusId !== 'ALL' && campusId !== caller.campusId) {
        return tssJsonOut_({ success: false, error: 'Not authorized for that campus' });
      }
      return tssJsonOut_({ success: true, activities: readNextMonthActivities_(campusId) });
    }

    const campuses = caller.campusId === 'ALL' ? Object.keys(TSS_CAMPUS_TO_TAB) : [caller.campusId];
    const byCampus = {};
    campuses.forEach(function (campusId) {
      byCampus[campusId] = readCampusStats_(campusId);
    });
    return tssJsonOut_({ success: true, rubricParams: TSS_RUBRIC_COLUMNS, campuses: byCampus });
  } catch (err) {
    return tssJsonOut_({ success: false, error: err.message });
  }
}

/** Next calendar month's events (not the current month) on a campus's
 *  Official school Calendar, e.g. "3rd — Literary Fair". Confirmed
 *  scope per Uday 2026-09-02: NEXT month, not current. */
function readNextMonthActivities_(campusId) {
  const calId = TSS_SCHOOL_CALENDAR_IDS[campusId];
  if (!calId) return [];
  const cal = CalendarApp.getCalendarById(calId);
  if (!cal) return [];

  const now = new Date();
  const rangeStart = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  const rangeEnd = new Date(now.getFullYear(), now.getMonth() + 2, 1); // exclusive

  const events = cal.getEvents(rangeStart, rangeEnd);
  events.sort(function (a, b) { return a.getStartTime() - b.getStartTime(); });
  return events.map(function (ev) {
    return ordinal_(ev.getStartTime().getDate()) + ' — ' + ev.getTitle();
  });
}

function ordinal_(day) {
  if (day >= 11 && day <= 13) return day + 'th';
  var lastDigit = day % 10;
  if (lastDigit === 1) return day + 'st';
  if (lastDigit === 2) return day + 'nd';
  if (lastDigit === 3) return day + 'rd';
  return day + 'th';
}

// Verified Google ID token -> {email, campusId, role}, re-derived from
// the allowlist every call. Only Principal/Coordinator/Owner may read;
// everyone else (e.g. Teacher role, or not on the list) gets null.
function verifySSDashboardToken_(idToken) {
  if (!idToken) return null;
  try {
    const res = UrlFetchApp.fetch(
      'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
      { muteHttpExceptions: true }
    );
    if (res.getResponseCode() !== 200) return null;
    const payload = JSON.parse(res.getContentText());
    if (payload.aud !== TSS_GOOGLE_CLIENT_ID) return null;
    if (payload.email_verified !== 'true' && payload.email_verified !== true) return null;
    const email = (payload.email || '').toLowerCase();

    const rows = SpreadsheetApp.openById(TSS_ALLOWLIST_SHEET_ID).getSheets()[0].getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0] || '').trim().toLowerCase() !== email) continue;
      const campusId = String(rows[i][2] || '').trim().toUpperCase();
      const role = String(rows[i][3] || '').trim();
      if (role !== 'Principal' && role !== 'Coordinator' && role !== 'Owner') return null;
      return { email: email, campusId: campusId, role: role };
    }
    return null; // not on the allowlist at all
  } catch (err) {
    return null;
  }
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

function tssJsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
