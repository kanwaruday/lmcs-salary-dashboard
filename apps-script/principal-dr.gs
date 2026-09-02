// ═══════════════════════════════════════════════════════════════════
// Principal DR — the principal's own daily operational report (the
// "Principal DR" tab). One concern-specific file inside the single
// "LMCS Principal's Daily Reporting Backend" project -- see main.gs's
// header for the full project layout and naming convention. doGet
// (and shared auth/JSON helpers) live in main.gs; this file only
// implements what's specific to Principal DR, called from there.
//
// Implemented so far: Month Activities (the "Activities of the Month"
// card) -- reads the School's OFFICIAL Calendar, NEXT calendar month.
// NOT implemented yet (still mocked on the frontend, each blocked on a
// real requirement -- see project_principals_daily_reporting.md):
// Support Sessions (blocked on Uday's naming-convention meeting),
// Tasks-for-Tomorrow persistence, Planned Activities persistence, and
// Submit (all three need a new daily-reports Sheet that doesn't exist
// yet). Add all of that HERE when it's built -- this stays ONE file
// per Uday, no further splitting even as more gets added.
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

/** Called from main.gs's doGet for action=monthactivities. `caller` is
 *  the already-verified {email, campusId, role} from main.gs. */
function principalDrMonthActivities_(caller, campusIdParam) {
  const campusId = String(campusIdParam || '').trim().toUpperCase();
  if (caller.campusId !== 'ALL' && campusId !== caller.campusId) {
    return { success: false, error: 'Not authorized for that campus' };
  }
  return { success: true, activities: pdrReadNextMonthActivities_(campusId) };
}

/** Next calendar month's events (not the current month) on a campus's
 *  Official school Calendar, e.g. "3rd — Literary Fair". Confirmed
 *  scope per Uday 2026-09-02: NEXT month, not current. */
function pdrReadNextMonthActivities_(campusId) {
  const calId = PDR_SCHOOL_CALENDAR_IDS[campusId];
  if (!calId) return [];
  const cal = CalendarApp.getCalendarById(calId);
  if (!cal) return [];

  const now = new Date();
  const rangeStart = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  const rangeEnd = new Date(now.getFullYear(), now.getMonth() + 2, 1); // exclusive

  const events = cal.getEvents(rangeStart, rangeEnd);
  events.sort(function (a, b) { return a.getStartTime() - b.getStartTime(); });
  return events.map(function (ev) {
    return pdrFormatEventDate_(ev) + ' — ' + ev.getTitle();
  });
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

// ── Diagnostics -- safe to keep here indefinitely, or delete once
//    you're confident everything above is working. ──────────────────

function testCalendarAccess() {
  ['nidhi.kant@lms.org.in', 'arti.sharma@lms.org.in', 'suresh.prasher@lms.org.in', 'nisha.lms@lms.org.in'].forEach(function (email) {
    var cal = CalendarApp.getCalendarById(email);
    Logger.log(email + ' -> ' + (cal ? 'ACCESSIBLE (' + cal.getName() + ')' : 'NOT ACCESSIBLE'));
  });
}

function testMonthActivities() {
  Logger.log(JSON.stringify(pdrReadNextMonthActivities_('LMS3'), null, 2));
}

function testSupportSessionsToday() {
  Object.keys(PDR_PRINCIPAL_PERSONAL_CALENDARS).forEach(function (campusId) {
    Logger.log(campusId + ' -> ' + JSON.stringify(pdrReadTodaysSupportSessions_(campusId)));
  });
}
