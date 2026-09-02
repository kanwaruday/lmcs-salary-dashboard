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
// ── INCIDENT (2026-09-02): a real submission got silently corrupted ──
// Uday submitted a real Daily Report for LMS2, 1 Sept 2026 (typed as
// "01/09/26", dd/mm/yy, matching every other date field on this page).
// The next working day, the "Tasks for Tomorrow" suggestions from that
// report never appeared. Root cause: the Sheet's Timestamp column
// proved its locale/default date format is US-style (month-first) --
// a plain string like "01/09/26" that LOOKS like a date is exactly the
// kind of value Google Sheets auto-converts into a real Date cell on
// write, and under a month-first locale it reads "01/09/26" as
// 9 Jan 2026, not 1 Sept 2026 -- silently wrong data, not a crash.
// Even where that mis-parse doesn't occur, the resulting cell becomes
// an actual Date object instead of the plain string this code wrote,
// so a naive `String(cell) === dateStr` comparison breaks anyway (a
// Date's String() is a verbose timestamp, never "dd/mm/yy").
//
// Fix, both layers: (1) every date this file stores switched from
// "dd/mm/yy" to ISO "yyyy-mm-dd" (pdrParseISO_/pdrFormatISO_) --
// unambiguous under ANY locale even if a cell still gets auto-
// converted; (2) every date READ goes through pdrCellDateToISO_(),
// which normalizes whichever type ended up in the cell (string or
// Date) to a canonical ISO string before comparing -- never a raw
// String(cell) comparison again; (3) every date WRITE forces the
// target cell's number format to Plain Text ('@') right before
// setValue, so Sheets stops auto-converting it at all going forward.
// Run pdrEnsureDateColumnsPlainText() once after deploying this to
// also fix the columns' default format for future manual edits.
// The one real corrupted row from this incident needs deleting/
// resubmitting by hand -- it can't be un-corrupted from here.
//
// Implemented so far: Month Activities (School's OFFICIAL Calendar,
// merged with still-open Planned Activities due the same month, tagged
// "School Specific"), Support Session count (Principal's PERSONAL
// Calendar, simple "SS"-tag count, not yet cross-referenced against
// Teacher SS submissions), Daily Reports persistence (Morning Assembly
// marking, Tasks Completed, Tasks for Tomorrow, Registers Crosschecked,
// Important Message -- one upserted row per campus per day in the
// "Daily Reports" tab; list-valued cells are comma-joined, not
// newline-joined, per Uday 2026-09-02, for readability directly in
// Sheets -- NOTE: an item containing a literal comma will incorrectly
// split into two on read-back, a known tradeoff of that choice), the
// Tasks Completed suggestion lookup (nearest prior WORKING day's Tasks
// for Tomorrow, skipping Sundays/2nd Saturdays/GH Calendar holidays),
// reading back an existing day's report so a past submission can be
// reviewed (action=dailyreport, powers the frontend's date-field
// refresh/reload), and Planned Activities CRUD (its own "Planned
// Activities" tab, soft-deleted via Status, never row-removed). This
// stays ONE file per Uday, no further splitting even as more gets
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
// ImportantMessage. Date is ISO "yyyy-mm-dd" (see the INCIDENT note
// above -- was dd/mm/yy, changed 2026-09-02). List-valued cells
// (TasksCompleted, TasksForTomorrow, RegistersCrosschecked,
// ImportantMessage) are comma-joined, human-readable directly in
// Sheets for management, not JSON.

// Planned Activities column layout (0-indexed): Id, CampusId, Title,
// DueDate, Assignee, CreatedAt, CreatedBy, Status, CompletedAt.
// DueDate is ISO "yyyy-mm-dd", same reasoning as above. Status is
// 'Active' / 'Completed' / 'Deleted' -- three values, not two: the UI
// has a separate checkbox (complete) and trash icon (delete), and
// 'Deleted' is a SOFT delete, the row is never removed.

function pdrDailyReportsSheet_() {
  return SpreadsheetApp.openById(PDR_DAILY_REPORTS_SHEET_ID).getSheetByName(PDR_DAILY_REPORTS_TAB);
}

function pdrPlannedActivitiesSheet_() {
  return SpreadsheetApp.openById(PDR_DAILY_REPORTS_SHEET_ID).getSheetByName(PDR_PLANNED_ACTIVITIES_TAB);
}

// "yyyy-mm-dd" -> Date at local midnight, or null if unparseable.
// Mirrors the frontend's parseISODate() exactly -- keep the two in
// sync. This is what native <input type="date"> fields produce, so no
// custom typing/format enforcement is needed client-side any more.
function pdrParseISO_(s) {
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || '').trim());
  if (!m) return null;
  var d = new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
  d.setHours(0, 0, 0, 0);
  return isNaN(d.getTime()) ? null : d;
}

function pdrFormatISO_(d) {
  var yyyy = d.getFullYear();
  var mm = String(d.getMonth() + 1).padStart(2, '0');
  var dd = String(d.getDate()).padStart(2, '0');
  return yyyy + '-' + mm + '-' + dd;
}

/** Normalizes a raw Daily-Reports/Planned-Activities date CELL to a
 *  canonical "yyyy-mm-dd" string, regardless of whether Sheets stored
 *  it as the plain string this code wrote OR silently auto-converted
 *  it into a real Date object -- see the INCIDENT note at the top of
 *  this file. Every date comparison in this file goes through this
 *  instead of a raw String(cell) call. */
function pdrCellDateToISO_(cellValue) {
  if (cellValue instanceof Date) return pdrFormatISO_(cellValue);
  var parsed = pdrParseISO_(cellValue);
  return parsed ? pdrFormatISO_(parsed) : String(cellValue || '').trim();
}

function pdrJoinList_(arr) {
  return (arr || []).map(function (s) { return String(s || '').trim(); }).filter(Boolean).join(', ');
}

/** Splits a comma-joined cell back into an array. NOTE: an item whose
 *  own text contains a literal comma will incorrectly split into two
 *  entries here -- accepted tradeoff for readability in the Sheet UI,
 *  per Uday 2026-09-02. */
function pdrSplitList_(cellValue) {
  return String(cellValue || '').split(/,\s*/).map(function (s) { return s.trim(); }).filter(Boolean);
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
    const due = pdrParseISO_(a.date);
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

// ── Support Sessions (Principal's PERSONAL Calendar) ────────────────

/** Called from main.gs's doGet for action=supportsessionstoday. Simple
 *  count (and titles) of Calendar events tagged "SS" on the Principal's
 *  own PERSONAL calendar, for `dateParam` if given (defaults to real
 *  "today" if absent/unparseable) -- date-parameterized as of
 *  2026-09-02 so viewing a past report via loadReportForDate() shows
 *  THAT day's SS count, not always today's (per Uday: "Can this fetch
 *  the data based on the date selected on the top right"). NOT yet
 *  cross-referenced against "LMCS Teacher SS 2026 (Responses)" (that's
 *  still blocked on the Class/Subject/Teacher naming convention Uday is
 *  defining -- see project_principals_daily_reporting.md). This is
 *  intentionally the simpler first cut: count real "SS"-tagged
 *  Calendar entries, nothing parsed out of them yet. */
function principalDrSupportSessionsToday_(caller, campusIdParam, dateParam) {
  const campusId = String(campusIdParam || '').trim().toUpperCase();
  if (caller.campusId !== 'ALL' && campusId !== caller.campusId) {
    return { success: false, error: 'Not authorized for that campus' };
  }
  const refDate = pdrParseISO_(dateParam) || new Date();
  return { success: true, sessions: pdrReadSupportSessionsForDate_(campusId, refDate) };
}

/** refDate's events on campusId's Principal's PERSONAL Calendar whose
 *  title contains the "SS" tag as its own token (e.g. "SS - C8
 *  English - Sushma ma'am") -- word-boundary match so it doesn't false-
 *  positive on something like "ASSEMBLY". Returns their titles. */
function pdrReadSupportSessionsForDate_(campusId, refDate) {
  const email = PDR_PRINCIPAL_PERSONAL_CALENDARS[campusId];
  if (!email) return []; // this campus's principal hasn't shared their calendar yet
  const cal = CalendarApp.getCalendarById(email);
  if (!cal) return [];

  const dayStart = new Date(refDate.getFullYear(), refDate.getMonth(), refDate.getDate());
  const dayEnd = new Date(refDate.getFullYear(), refDate.getMonth(), refDate.getDate() + 1); // exclusive

  const ssTag = /\bSS\b/;
  return cal.getEvents(dayStart, dayEnd)
    .filter(function (ev) { return ssTag.test(ev.getTitle()); })
    .map(function (ev) { return ev.getTitle(); });
}

// ── Daily Reports: save (upsert), read-back, + the Tasks Completed suggestion lookup ──

/** Called from main.gs's doPost for action=savedailyreport. Upserts
 *  (not always-appends) the row for (campusId, body.date) -- Timestamp
 *  is always overwritten to "now", so it reflects last-updated time,
 *  not first-submitted time. Support Session count is deliberately NOT
 *  persisted here -- it's Calendar-derived and recomputed live on every
 *  load, not something the Principal edits/submits. Forces the Date
 *  cell to Plain Text before writing -- see the INCIDENT note. */
function principalDrSaveDailyReport_(caller, body) {
  const campusId = String(body.campusId || '').trim().toUpperCase();
  if (caller.campusId !== 'ALL' && campusId !== caller.campusId) {
    return { success: false, error: 'Not authorized for that campus' };
  }
  const dateISO = String(body.date || '').trim();
  if (!pdrParseISO_(dateISO)) {
    return { success: false, error: 'Invalid date: "' + dateISO + '" (expected yyyy-mm-dd)' };
  }

  const ma = body.morningAssembly || {};
  const row = [
    new Date(),
    dateISO,
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
  const rowIndex = pdrFindDailyReportRowIndex_(sheet, campusId, dateISO);
  const targetRow = rowIndex > 0 ? rowIndex : sheet.getLastRow() + 1;
  sheet.getRange(targetRow, 2).setNumberFormat('@'); // Date column -- Plain Text, see INCIDENT note
  sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
  return { success: true };
}

/** 1-indexed sheet row number for (campusId, dateISO)'s Daily Reports
 *  row, or -1 if none exists yet. */
function pdrFindDailyReportRowIndex_(sheet, campusId, dateISO) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][2]).trim() === campusId && pdrCellDateToISO_(values[i][1]) === dateISO) {
      return i + 1;
    }
  }
  return -1;
}

/** Called from main.gs's doGet for action=dailyreport -- reads back an
 *  existing day's report so the frontend's Date-field refresh/reload
 *  can show a past submission (Uday 2026-09-02: "Is there a way I can
 *  see a past report that I have submitted?"). Returns
 *  {exists:false} rather than an error when nothing's been submitted
 *  for that day yet -- that's the normal case for most dates. */
function principalDrGetDailyReport_(caller, campusIdParam, dateParam) {
  const campusId = String(campusIdParam || '').trim().toUpperCase();
  if (caller.campusId !== 'ALL' && campusId !== caller.campusId) {
    return { success: false, error: 'Not authorized for that campus' };
  }
  const dateISO = String(dateParam || '').trim();
  if (!pdrParseISO_(dateISO)) {
    return { success: false, error: 'Invalid date: "' + dateISO + '" (expected yyyy-mm-dd)' };
  }

  const values = pdrDailyReportsSheet_().getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][2]).trim() !== campusId) continue;
    if (pdrCellDateToISO_(values[i][1]) !== dateISO) continue;
    return {
      success: true,
      report: {
        exists: true,
        maClassOrHouse: String(values[i][4] || ''),
        maScore: String(values[i][5] || ''),
        tasksCompleted: pdrSplitList_(values[i][6]),
        tasksForTomorrow: pdrSplitList_(values[i][7]),
        registersCrosschecked: pdrSplitList_(values[i][8]),
        importantMessage: pdrSplitList_(values[i][9]),
      },
    };
  }
  return { success: true, report: { exists: false } };
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
  const refDate = pdrParseISO_(dateParam) || new Date();
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
    const dateISO = pdrFormatISO_(cursor);

    for (let r = 1; r < values.length; r++) {
      if (String(values[r][2]).trim() !== campusId || pdrCellDateToISO_(values[r][1]) !== dateISO) continue;
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
 *  the frontend's renderPlannedActivities(). `date` is normalized to
 *  ISO via pdrCellDateToISO_ regardless of the raw cell's type -- see
 *  the INCIDENT note. Reused by the Month Activities merge (which
 *  filters out completed ones itself) and by the Tasks Completed
 *  suggestion logic on the frontend. */
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
      date: pdrCellDateToISO_(row[3]),
      activity: String(row[2]),
      assignee: String(row[4] || ''),
      completed: status === 'Completed',
    });
  }
  return out;
}

/** Called from main.gs's doPost for action=addplannedactivity. Forces
 *  the DueDate cell to Plain Text after writing -- see INCIDENT note. */
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

  const sheet = pdrPlannedActivitiesSheet_();
  sheet.appendRow([id, campusId, title, dueDate, assignee, new Date(), caller.email, 'Active', '']);
  sheet.getRange(sheet.getLastRow(), 4).setNumberFormat('@').setValue(dueDate); // re-enforce Plain Text + re-write, in case appendRow's own write got auto-converted
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

// ── One-time setup ────────────────────────────────────────────────

/** Run ONCE after deploying this file: sets the Date column (Daily
 *  Reports, col B) and DueDate column (Planned Activities, col D) to
 *  Plain Text format for their full length, so Sheets stops auto-
 *  detecting/converting date-looking strings typed into them -- by
 *  code OR by a human editing the Sheet directly. The code also
 *  re-applies '@' per-cell on every write as a second line of defense
 *  (e.g. in case someone resets the column format later), but this
 *  fixes the column-wide default too. See the INCIDENT note up top. */
function pdrEnsureDateColumnsPlainText() {
  pdrDailyReportsSheet_().getRange('B2:B').setNumberFormat('@');
  pdrPlannedActivitiesSheet_().getRange('D2:D').setNumberFormat('@');
  Logger.log('Date columns set to Plain Text. The 2026-09-02 corrupted test row still needs deleting/resubmitting by hand -- this only prevents it happening again.');
}

// ── Diagnostics -- safe to keep here indefinitely, or delete once
//    you're confident everything above is working. Read-only (except
//    pdrEnsureDateColumnsPlainText above, which only changes
//    formatting, never values); none of these write real data, so
//    they're safe to run anytime without polluting anything. ────────

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
    Logger.log(campusId + ' -> ' + JSON.stringify(pdrReadSupportSessionsForDate_(campusId, new Date())));
  });
}

function testPlannedActivities() {
  Logger.log(JSON.stringify(pdrReadPlannedActivities_('LMS2'), null, 2));
}

function testYesterdaysTasks() {
  Logger.log(JSON.stringify(pdrFindPriorWorkingDayTasksForTomorrow_('LMS2', new Date()), null, 2));
}

function testGetDailyReport() {
  Logger.log(JSON.stringify(principalDrGetDailyReport_({ campusId: 'ALL' }, 'LMS2', pdrFormatISO_(new Date())), null, 2));
}

/** Run this FIRST when diagnosing any date-matching issue -- shows
 *  exactly what type each Daily Reports row's Date cell really holds
 *  (string vs Date object) and what it normalizes to, confirming (or
 *  ruling out) the class of bug described in the INCIDENT note. */
function testDailyReportsRawCellTypes() {
  const values = pdrDailyReportsSheet_().getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    Logger.log('Row ' + (i + 1) + ': raw=' + values[i][1] + ' | typeof=' + typeof values[i][1]
      + ' | isDateObject=' + (values[i][1] instanceof Date) + ' | normalizedISO=' + pdrCellDateToISO_(values[i][1]));
  }
}
