// ═══════════════════════════════════════════════════════════════════
// Principal DR — the principal's own daily operational report (the
// "Principal DR" tab). One concern-specific file inside the single
// "LMCS Principal's Daily Reporting Backend" project -- see main.gs's
// header for the full project layout and naming convention. doGet AND
// doPost (and the shared auth/JSON helpers) live in main.gs; this file
// only implements what's specific to Principal DR, called from there.
// Reads go through doGet (?action=...); writes (Submit, Planned
// Activities CRUD) go through doPost as a JSON body, since Apps Script
// Web Apps don't implement doOptions and a JSON-Content-Type POST would
// trigger a CORS preflight that fails -- see the frontend's pdrPost_()
// comment for the full reasoning.
//
// Implemented so far: Month Activities (School's OFFICIAL Calendar,
// merged with still-open Planned Activities due the same month, tagged
// "School Specific"), Support Session count (Principal's PERSONAL
// Calendar, simple "SS"-tag count, not yet cross-referenced against
// Teacher SS submissions), Daily Reports persistence (Morning Assembly
// marking, Tasks Completed, Tasks for Tomorrow, Registers Crosschecked,
// Important Message -- one upserted row per campus per day in the
// "Daily Reports" tab), the Tasks Completed suggestion lookup (nearest
// prior WORKING day's Tasks for Tomorrow, skipping Sundays/2nd
// Saturdays/GH Calendar holidays), and Planned Activities CRUD (its own
// "Planned Activities" tab, soft-deleted via Status, never row-removed).
// This stays ONE file per Uday, no further splitting even as more gets
// added.
// ═══════════════════════════════════════════════════════════════════

// School's OFFICIAL Calendar per campus (2026-09-02, confirmed
// live/working -- same IDs already proven in ChapterTracker.gs's
// CHAPTER_CALENDAR_IDS and in WhatsApp Hub's hub-config.json, just
// re-keyed to this portal's 'LMS1'..'LMS6' convention instead of
// ChapterTracker's 1..6). LMS2 and LMS3 deliberately share one
// calendar (separate one "coming 2027" per hub-config.json's note).
const PDR_SCHOOL_CALENDAR_IDS = {
  LMS1: 'c_88402c40e6c9a435cbed4810bfae632d781ac9e7ca167b5435e4a9586bda7595@group.calendar.google.com',
  LMS2: 'c_25n5qbggh7u9s5o6vq31iho4vc@group.calendar.google.com',
  LMS3: 'c_25n5qbggh7u9s5o6vq31iho4vc@group.calendar.google.com', // shared with LMS2
  LMS4: 'lms.org.in_snflpd3f357jv8l380s8sc4ju0@group.calendar.google.com',
  LMS5: 'c_d367ca2cb9bbab4f6a5b6864d6e03c28fbfccf02b5212621e89daf2f9f3f8523@group.calendar.google.com',
  LMS6: 'c_3db6cb2edb78d889cf166771c426f0fe3ba78b04cf9ca8e566329c888d023635@group.calendar.google.com',
};

// Each Principal's OWN PERSONAL Calendar (their email doubles as the
// Calendar id) -- confirmed 2026-09-02. Support Sessions are logged
// here, tagged "SS" per the tag convention (MA/DISP/SS/NBK/MEET).
// LMS5 and LMS6 principals haven't shared their calendars yet -- those
// two return 0 until they do, not an error.
const PDR_PRINCIPAL_PERSONAL_CALENDARS = {
  LMS1: 'nidhi.kant@lms.org.in',
  LMS2: 'arti.sharma@lms.org.in',
  LMS3: 'suresh.prasher@lms.org.in',
  LMS4: 'nisha.lms@lms.org.in',
  // LMS5: not yet shared
  // LMS6: not yet shared
};

// "LMCS Principal DR — Daily Reports" Sheet (2026-09-02, created by
// Uday to match the schema agreed in chat -- see
// project_principals_daily_reporting.md for the full design writeup).
// Two tabs, both used below.
const PDR_DAILY_REPORTS_SHEET_ID = '1GzGdakmjjkou0GGgljY6ugXvOM21MLkyeyc85Ke3i1Y';
const PDR_DAILY_REPORTS_TAB = 'Daily Reports';
const PDR_PLANNED_ACTIVITIES_TAB = 'Planned Activities';

// Daily Reports column layout (0-indexed, matches the tab's header row
// exactly): Timestamp, Date, CampusId, PrincipalEmail, MA_ClassOrHouse,
// MA_Score, TasksCompleted, TasksForTomorrow, RegistersCrosschecked,
// ImportantMessage. List-valued cells (TasksCompleted, TasksForTomorrow,
// RegistersCrosschecked, ImportantMessage) are newline-joined -- plain
// text, deliberately human-readable directly in Sheets (e.g. for
// management reading Important Message), not JSON.

// Planned Activities column layout (0-indexed): Id, CampusId, Title,
// DueDate, Assignee, CreatedAt, CreatedBy, Status, CompletedAt.
// Status is 'Active' / 'Completed' / 'Deleted' -- three values, not
// two: the UI has a separate checkbox (complete) and trash icon
// (delete), and 'Deleted' is a SOFT delete, the row is never removed.

function pdrDailyReportsSheet_() {
  return SpreadsheetApp.openById(PDR_DAILY_REPORTS_SHEET_ID).getSheetByName(PDR_DAILY_REPORTS_TAB);
}

function pdrPlannedActivitiesSheet_() {
  return SpreadsheetApp.openById(PDR_DAILY_REPORTS_SHEET_ID).getSheetByName(PDR_PLANNED_ACTIVITIES_TAB);
}

// "dd/mm/yy" -> Date at local midnight, or null if unparseable. Mirrors
// the frontend's parseDDMMYY() exactly -- keep the two in sync.
function pdrParseDDMMYY_(s) {
  var m = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/.exec(String(s || '').trim());
  if (!m) return null;
  var dd = parseInt(m[1], 10), mm = parseInt(m[2], 10) - 1, yy = parseInt(m[3], 10);
  if (yy < 100) yy += 2000;
  var d = new Date(yy, mm, dd);
  d.setHours(0, 0, 0, 0);
  return isNaN(d.getTime()) ? null : d;
}

function pdrFormatDDMMYY_(d) {
  var dd = String(d.getDate()).padStart(2, '0');
  var mm = String(d.getMonth() + 1).padStart(2, '0');
  var yy = String(d.getFullYear()).slice(-2);
  return dd + '/' + mm + '/' + yy;
}

function pdrJoinList_(arr) {
  return (arr || []).map(function (s) { return String(s || '').trim(); }).filter(Boolean).join('\n');
}

function pdrSplitList_(cellValue) {
  return String(cellValue || '').split('\n').map(function (s) { return s.trim(); }).filter(Boolean);
}

// ── Off-day check (Sunday / 2nd Saturday / GH Calendar events) ──────
// Reused from the same "LMS Holiday Rules" convention already used
// elsewhere (submission trackers): Sunday, the 2nd Saturday of the
// month, and any all-day Calendar event whose title starts with "GH"
// (General Holiday) on the campus's OFFICIAL Calendar -- confirmed
// real examples from testMonthActivities()'s own output: "GH-Dussehra",
// "GH-Karva Chauth", etc.

function pdrIsOffDay_(date, campusId) {
  if (date.getDay() === 0) return true; // Sunday
  if (date.getDay() === 6 && Math.ceil(date.getDate() / 7) === 2) return true; // 2nd Saturday
  return pdrHasGHEventOn_(date, campusId);
}

function pdrHasGHEventOn_(date, campusId) {
  var calId = PDR_SCHOOL_CALENDAR_IDS[campusId];
  if (!calId) return false;
  var cal = CalendarApp.getCalendarById(calId);
  if (!cal) return false;
  var dayStart = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  var dayEnd = new Date(date.getFullYear(), date.getMonth(), date.getDate() + 1);
  return cal.getEvents(dayStart, dayEnd).some(function (ev) { return ev.getTitle().indexOf('GH') === 0; });
}

// ── Month Activities (School's OFFICIAL Calendar + Planned Activities) ──

/** Called from main.gs's doGet for action=monthactivities. `caller` is
 *  the already-verified {email, campusId, role} from main.gs. */
function principalDrMonthActivities_(caller, campusIdParam) {
  const campusId = String(campusIdParam || '').trim().toUpperCase();
  if (caller.campusId !== 'ALL' && campusId !== caller.campusId) {
    return { success: false, error: 'Not authorized for that campus' };
  }
  return { success: true, activities: pdrReadMergedMonthActivities_(campusId) };
}

function pdrNextMonthRange_() {
  const now = new Date();
  return {
    start: new Date(now.getFullYear(), now.getMonth() + 1, 1),
    end: new Date(now.getFullYear(), now.getMonth() + 2, 1), // exclusive
  };
}

/** Next calendar month's official-Calendar events MERGED with still-
 *  open Planned Activities due that same month, sorted together
 *  chronologically. Planned Activities entries carry tag:"School
 *  Specific" so the frontend can badge them differently -- confirmed
 *  scope per Uday 2026-09-02: same window as the official Calendar,
 *  not a separate "due soon" window. */
function pdrReadMergedMonthActivities_(campusId) {
  const range = pdrNextMonthRange_();
  const items = [];

  const calId = PDR_SCHOOL_CALENDAR_IDS[campusId];
  const cal = calId ? CalendarApp.getCalendarById(calId) : null;
  if (cal) {
    cal.getEvents(range.start, range.end).forEach(function (ev) {
      items.push({ sortKey: ev.getStartTime(), text: pdrFormatEventDate_(ev) + ' — ' + ev.getTitle(), tag: null });
    });
  }

  pdrReadPlannedActivities_(campusId).forEach(function (a) {
    if (a.completed) return; // only still-open planned activities show here
    const due = pdrParseDDMMYY_(a.date);
    if (!due || due < range.start || due >= range.end) return;
    items.push({
      sortKey: due,
      text: pdrOrdinal_(due.getDate()) + ' ' + pdrMonthAbbrev_(due.getMonth()) + ' — ' + a.activity,
      tag: 'School Specific',
    });
  });

  items.sort(function (a, b) { return a.sortKey - b.sortKey; });
  return items.map(function (i) { return { text: i.text, tag: i.tag }; });
}

/** e.g. "18th Oct" for a single-day event, "18th–20th Oct" for a
 *  multi-day event within one month, or "29th Sep – 2nd Oct" when it
 *  spans a month boundary. All-day events store an EXCLUSIVE end (a
 *  1-day all-day event's getEndTime() is already the next day), so
 *  that's pulled back a day before comparing against the start. */
function pdrFormatEventDate_(ev) {
  const start = ev.getStartTime();
  const end = ev.isAllDayEvent() ? new Date(ev.getEndTime().getTime() - 1) : ev.getEndTime();

  const sameDay = start.getFullYear() === end.getFullYear()
    && start.getMonth() === end.getMonth()
    && start.getDate() === end.getDate();
  if (sameDay) {
    return pdrOrdinal_(start.getDate()) + ' ' + pdrMonthAbbrev_(start.getMonth());
  }

  const sameMonth = start.getFullYear() === end.getFullYear() && start.getMonth() === end.getMonth();
  if (sameMonth) {
    return pdrOrdinal_(start.getDate()) + '–' + pdrOrdinal_(end.getDate()) + ' ' + pdrMonthAbbrev_(start.getMonth());
  }

  return pdrOrdinal_(start.getDate()) + ' ' + pdrMonthAbbrev_(start.getMonth())
    + ' – ' + pdrOrdinal_(end.getDate()) + ' ' + pdrMonthAbbrev_(end.getMonth());
}

function pdrMonthAbbrev_(monthIndex0) {
  return ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul',
    'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][monthIndex0];
}

function pdrOrdinal_(day) {
  if (day >= 11 && day <= 13) return day + 'th';
  var lastDigit = day % 10;
  if (lastDigit === 1) return day + 'st';
  if (lastDigit === 2) return day + 'nd';
  if (lastDigit === 3) return day + 'rd';
  return day + 'th';
}

// ── Support Sessions today (Principal's PERSONAL Calendar) ──────────

/** Called from main.gs's doGet for action=supportsessionstoday.
 *  Simple count (and titles) of today's Calendar events tagged "SS" on
 *  the Principal's own PERSONAL calendar -- NOT yet cross-referenced
 *  against "LMCS Teacher SS 2026 (Responses)" (that's still blocked on
 *  the Class/Subject/Teacher naming convention Uday is defining --
 *  see project_principals_daily_reporting.md). This is intentionally
 *  the simpler first cut: count real "SS"-tagged Calendar entries,
 *  nothing parsed out of them yet. */
function principalDrSupportSessionsToday_(caller, campusIdParam) {
  const campusId = String(campusIdParam || '').trim().toUpperCase();
  if (caller.campusId !== 'ALL' && campusId !== caller.campusId) {
    return { success: false, error: 'Not authorized for that campus' };
  }
  return { success: true, sessions: pdrReadTodaysSupportSessions_(campusId) };
}

/** Today's events on campusId's Principal's PERSONAL Calendar whose
 *  title contains the "SS" tag as its own token (e.g. "SS - C8
 *  English - Sushma ma'am") -- word-boundary match so it doesn't false-
 *  positive on something like "ASSEMBLY". Returns their titles. */
function pdrReadTodaysSupportSessions_(campusId) {
  const email = PDR_PRINCIPAL_PERSONAL_CALENDARS[campusId];
  if (!email) return []; // this campus's principal hasn't shared their calendar yet
  const cal = CalendarApp.getCalendarById(email);
  if (!cal) return [];

  const now = new Date();
  const dayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const dayEnd = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1); // exclusive

  const ssTag = /\bSS\b/;
  return cal.getEvents(dayStart, dayEnd)
    .filter(function (ev) { return ssTag.test(ev.getTitle()); })
    .map(function (ev) { return ev.getTitle(); });
}

// ── Daily Reports: save (upsert) + the Tasks Completed suggestion lookup ──

/** Called from main.gs's doPost for action=savedailyreport. Upserts
 *  (not always-appends) the row for (campusId, body.date) -- Timestamp
 *  is always overwritten to "now", so it reflects last-updated time,
 *  not first-submitted time. Support Session count is deliberately NOT
 *  persisted here -- it's Calendar-derived and recomputed live on every
 *  load, not something the Principal edits/submits. */
function principalDrSaveDailyReport_(caller, body) {
  const campusId = String(body.campusId || '').trim().toUpperCase();
  if (caller.campusId !== 'ALL' && campusId !== caller.campusId) {
    return { success: false, error: 'Not authorized for that campus' };
  }
  const dateStr = String(body.date || '').trim();
  if (!pdrParseDDMMYY_(dateStr)) {
    return { success: false, error: 'Invalid date: "' + dateStr + '" (expected dd/mm/yy)' };
  }

  const ma = body.morningAssembly || {};
  const row = [
    new Date(),
    dateStr,
    campusId,
    caller.email,
    String(ma.classHouse || ''),
    String(ma.score || ''),
    pdrJoinList_(body.tasksCompleted),
    pdrJoinList_(body.tasksTomorrow),
    pdrJoinList_(body.registersCrosschecked),
    pdrJoinList_(body.importantMessage),
  ];

  const sheet = pdrDailyReportsSheet_();
  const rowIndex = pdrFindDailyReportRowIndex_(sheet, campusId, dateStr);
  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
  return { success: true };
}

/** 1-indexed sheet row number for (campusId, dateStr)'s Daily Reports
 *  row, or -1 if none exists yet. */
function pdrFindDailyReportRowIndex_(sheet, campusId, dateStr) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][2]).trim() === campusId && String(values[i][1]).trim() === dateStr) {
      return i + 1;
    }
  }
  return -1;
}

/** Called from main.gs's doGet for action=yesterdaystasks. `dateParam`
 *  is the report's OWN date field (a Principal can edit it, so this is
 *  NOT necessarily "today") -- the lookback starts the working day
 *  before THAT date, not before the server's real today. */
function principalDrYesterdaysTasks_(caller, campusIdParam, dateParam) {
  const campusId = String(campusIdParam || '').trim().toUpperCase();
  if (caller.campusId !== 'ALL' && campusId !== caller.campusId) {
    return { success: false, error: 'Not authorized for that campus' };
  }
  const refDate = pdrParseDDMMYY_(dateParam) || new Date();
  return { success: true, tasks: pdrFindPriorWorkingDayTasksForTomorrow_(campusId, refDate) };
}

// Bounds the backward walk so a campus with no submission history
// doesn't loop forever looking for a match that will never exist.
const PDR_SUGGESTION_LOOKBACK_DAYS = 10;

/** Walks backward from refDate, skipping off-days (pdrIsOffDay_), and
 *  returns the first 2 entries of the nearest prior WORKING day's
 *  Daily Reports row that has a non-empty TasksForTomorrow. A working
 *  day with an empty (or missing) TasksForTomorrow does NOT stop the
 *  search -- it keeps walking further back, per Uday's spec. Returns
 *  [] if nothing turns up within the lookback bound -- not an error. */
function pdrFindPriorWorkingDayTasksForTomorrow_(campusId, refDate) {
  const values = pdrDailyReportsSheet_().getDataRange().getValues();
  const cursor = new Date(refDate.getFullYear(), refDate.getMonth(), refDate.getDate());

  for (let i = 0; i < PDR_SUGGESTION_LOOKBACK_DAYS; i++) {
    cursor.setDate(cursor.getDate() - 1);
    if (pdrIsOffDay_(cursor, campusId)) continue;
    const dateStr = pdrFormatDDMMYY_(cursor);

    for (let r = 1; r < values.length; r++) {
      if (String(values[r][2]).trim() !== campusId || String(values[r][1]).trim() !== dateStr) continue;
      const tasksForTomorrow = pdrSplitList_(values[r][7]);
      if (tasksForTomorrow.length) return tasksForTomorrow.slice(0, 2);
      break; // found the day's row but it's empty -- stop scanning rows, keep walking back
    }
  }
  return [];
}

// ── Planned Activities: read + CRUD ──────────────────────────────────

/** Called from main.gs's doGet for action=plannedactivities. */
function principalDrPlannedActivities_(caller, campusIdParam) {
  const campusId = String(campusIdParam || '').trim().toUpperCase();
  if (caller.campusId !== 'ALL' && campusId !== caller.campusId) {
    return { success: false, error: 'Not authorized for that campus' };
  }
  return { success: true, activities: pdrReadPlannedActivities_(campusId) };
}

/** All non-Deleted rows (Active + Completed) for a campus, shaped for
 *  the frontend's renderPlannedActivities(). Reused by the Month
 *  Activities merge (which filters out completed ones itself) and by
 *  the Tasks Completed suggestion lookup below. */
function pdrReadPlannedActivities_(campusId) {
  const values = pdrPlannedActivitiesSheet_().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (String(row[1]).trim() !== campusId) continue;
    const status = String(row[7] || '').trim();
    if (status === 'Deleted') continue;
    out.push({
      id: String(row[0]),
      date: String(row[3]),
      activity: String(row[2]),
      assignee: String(row[4] || ''),
      completed: status === 'Completed',
    });
  }
  return out;
}

/** Called from main.gs's doPost for action=addplannedactivity. */
function principalDrAddPlannedActivity_(caller, body) {
  const campusId = String(body.campusId || '').trim().toUpperCase();
  if (caller.campusId !== 'ALL' && campusId !== caller.campusId) {
    return { success: false, error: 'Not authorized for that campus' };
  }
  const title = String(body.title || '').trim();
  if (!title) return { success: false, error: 'Activity name required' };
  const dueDate = String(body.dueDate || '').trim();
  const assignee = String(body.assignee || '').trim();
  const id = 'pa-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000);

  pdrPlannedActivitiesSheet_().appendRow([id, campusId, title, dueDate, assignee, new Date(), caller.email, 'Active', '']);
  return { success: true, activity: { id: id, date: dueDate, activity: title, assignee: assignee, completed: false } };
}

/** Called from main.gs's doPost for action=setplannedactivitycompleted. */
function principalDrSetPlannedActivityCompleted_(caller, body) {
  const campusId = String(body.campusId || '').trim().toUpperCase();
  if (caller.campusId !== 'ALL' && campusId !== caller.campusId) {
    return { success: false, error: 'Not authorized for that campus' };
  }
  const sheet = pdrPlannedActivitiesSheet_();
  const rowIndex = pdrFindPlannedActivityRow_(sheet, campusId, body.id);
  if (rowIndex < 0) return { success: false, error: 'Activity not found' };

  const completed = !!body.completed;
  sheet.getRange(rowIndex, 8).setValue(completed ? 'Completed' : 'Active'); // Status
  sheet.getRange(rowIndex, 9).setValue(completed ? new Date() : ''); // CompletedAt
  return { success: true };
}

/** Called from main.gs's doPost for action=deleteplannedactivity.
 *  Soft-delete only -- the row is never removed, just marked Deleted
 *  and excluded from every read above. */
function principalDrDeletePlannedActivity_(caller, body) {
  const campusId = String(body.campusId || '').trim().toUpperCase();
  if (caller.campusId !== 'ALL' && campusId !== caller.campusId) {
    return { success: false, error: 'Not authorized for that campus' };
  }
  const sheet = pdrPlannedActivitiesSheet_();
  const rowIndex = pdrFindPlannedActivityRow_(sheet, campusId, body.id);
  if (rowIndex < 0) return { success: false, error: 'Activity not found' };

  sheet.getRange(rowIndex, 8).setValue('Deleted');
  return { success: true };
}

/** 1-indexed sheet row number for a Planned Activity Id within one
 *  campus, or -1 if not found. Campus-scoped so one campus can't
 *  reference/mutate another's activity by guessing an id. */
function pdrFindPlannedActivityRow_(sheet, campusId, id) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id) && String(values[i][1]).trim() === campusId) return i + 1;
  }
  return -1;
}

// ── Diagnostics -- safe to keep here indefinitely, or delete once
//    you're confident everything above is working. Read-only; none of
//    these write to the Daily Reports Sheet, so they're safe to run
//    anytime without polluting real data. ─────────────────────────────

function testCalendarAccess() {
  ['nidhi.kant@lms.org.in', 'arti.sharma@lms.org.in', 'suresh.prasher@lms.org.in', 'nisha.lms@lms.org.in'].forEach(function (email) {
    var cal = CalendarApp.getCalendarById(email);
    Logger.log(email + ' -> ' + (cal ? 'ACCESSIBLE (' + cal.getName() + ')' : 'NOT ACCESSIBLE'));
  });
}

function testMonthActivities() {
  Logger.log(JSON.stringify(pdrReadMergedMonthActivities_('LMS3'), null, 2));
}

function testSupportSessionsToday() {
  Object.keys(PDR_PRINCIPAL_PERSONAL_CALENDARS).forEach(function (campusId) {
    Logger.log(campusId + ' -> ' + JSON.stringify(pdrReadTodaysSupportSessions_(campusId)));
  });
}

function testPlannedActivities() {
  Logger.log(JSON.stringify(pdrReadPlannedActivities_('LMS2'), null, 2));
}

function testYesterdaysTasks() {
  Logger.log(JSON.stringify(pdrFindPriorWorkingDayTasksForTomorrow_('LMS2', new Date()), null, 2));
}
