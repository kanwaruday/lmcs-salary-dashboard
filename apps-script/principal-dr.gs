// ═══════════════════════════════════════════════════════════════════
// Principal DR — everything specific to the Principal's own Daily
// Report tab (teacher-ss/index.html's "Principal DR" tab). Per Uday
// (2026-09-02): all Principal DR code should live strictly in this ONE
// file, separate from teacher-ss-dashboard-proxy.gs, which is Teacher
// SS -- a different concept (see that file's own header; "DR" here is
// the principal's own daily operational report, "SS" there is Support
// Session, the teacher-evaluation rubric).
//
// Apps Script only allows ONE doGet per project, so
// teacher-ss-dashboard-proxy.gs's doGet stays the single Web App entry
// point and just calls OUT to the function(s) below for any 'action'
// that's Principal-DR-specific (currently: action=monthactivities ->
// pdrReadNextMonthActivities_). Don't move doGet itself here; DO keep
// adding new Principal DR logic/data here, not in the Teacher SS file,
// as more of the tab gets wired to real data.
//
// Implemented so far: Month Activities (the "Activities of the Month"
// card) -- reads the School's OFFICIAL Calendar, NEXT calendar month.
// NOT implemented yet (still mocked on the frontend, each blocked on a
// real requirement -- see project_principals_daily_reporting.md):
// Support Sessions (blocked on Uday's naming-convention meeting),
// Tasks-for-Tomorrow persistence, Planned Activities persistence, and
// Submit (all three need a new daily-reports Sheet that doesn't exist
// yet).
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

/** Next calendar month's events (not the current month) on a campus's
 *  Official school Calendar, e.g. "3rd — Literary Fair". Confirmed
 *  scope per Uday 2026-09-02: NEXT month, not current. Called from
 *  teacher-ss-dashboard-proxy.gs's doGet for action=monthactivities --
 *  that file only dispatches to this, the logic itself lives here. */
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
    return pdrOrdinal_(ev.getStartTime().getDate()) + ' — ' + ev.getTitle();
  });
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
