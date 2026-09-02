// ═══════════════════════════════════════════════════════════════════
// Principal DR Backend — its OWN standalone Apps Script project
// (2026-09-02, per Uday: Teacher SS and Principal DR should each have
// "everything associated to this" in its own project, not share one --
// this file previously lived inside "LMCS Teacher SS Backend" as a
// second file; it's now fully self-contained with its own doGet).
//
// Powers teacher-ss/index.html's "Principal DR" tab -- the principal's
// own daily operational report. NOT the same thing as Teacher SS
// (Support Session, the teacher-evaluation rubric, a different Apps
// Script project entirely) -- if you're looking for Teacher SS code,
// it isn't here, and Principal DR code shouldn't go in that project.
//
// Implemented so far: Month Activities (the "Activities of the Month"
// card) -- reads the School's OFFICIAL Calendar, NEXT calendar month.
// NOT implemented yet (still mocked on the frontend, each blocked on a
// real requirement -- see project_principals_daily_reporting.md):
// Support Sessions (blocked on Uday's naming-convention meeting),
// Tasks-for-Tomorrow persistence, Planned Activities persistence, and
// Submit (all three need a new daily-reports Sheet that doesn't exist
// yet). Add all of that HERE when it's built.
//
// GATING: same allowlist + verified-Google-ID-token pattern as every
// other backend in this portal -- Principal/Coordinator/Owner only.
//
// SETUP:
//   1. script.google.com -> New project -> name it "LMCS Principal DR
//      Backend" -> paste this file in as principal-dr.gs
//   2. Deploy -> New deployment -> Web App
//      Execute as: Me | Who has access: Anyone
//      (Apps Script's own access setting isn't the real gate --
//      the verified-token check inside is)
//   3. Copy the Web App URL into teacher-ss/index.html's
//      PRINCIPAL_DR_API_URL constant
// ═══════════════════════════════════════════════════════════════════

const PDR_ALLOWLIST_SHEET_ID = '1NZu0ElismFytG395Nxjz29vAz7OfkmJtZhs70bOwT58'; // "LMCS Principal Allowlist"
const PDR_GOOGLE_CLIENT_ID = '697999989724-mvi85iobr20g4mm8a8nrjd1rms2o8tf6.apps.googleusercontent.com'; // same client ID assets/auth.js signs in with

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

function doGet(e) {
  try {
    const caller = pdrVerifyToken_(e.parameter.idToken);
    if (!caller) return pdrJsonOut_({ success: false, error: 'Not authorized' });

    const action = (e.parameter.action || '').toLowerCase();

    if (action === 'monthactivities') {
      const campusId = String(e.parameter.campusId || '').trim().toUpperCase();
      if (caller.campusId !== 'ALL' && campusId !== caller.campusId) {
        return pdrJsonOut_({ success: false, error: 'Not authorized for that campus' });
      }
      return pdrJsonOut_({ success: true, activities: pdrReadNextMonthActivities_(campusId) });
    }

    return pdrJsonOut_({ success: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return pdrJsonOut_({ success: false, error: err.message });
  }
}

// Verified Google ID token -> {email, campusId, role}, re-derived from
// the allowlist every call. Only Principal/Coordinator/Owner may read;
// everyone else (e.g. Teacher role, or not on the list) gets null.
function pdrVerifyToken_(idToken) {
  if (!idToken) return null;
  try {
    const res = UrlFetchApp.fetch(
      'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
      { muteHttpExceptions: true }
    );
    if (res.getResponseCode() !== 200) return null;
    const payload = JSON.parse(res.getContentText());
    if (payload.aud !== PDR_GOOGLE_CLIENT_ID) return null;
    if (payload.email_verified !== 'true' && payload.email_verified !== true) return null;
    const email = (payload.email || '').toLowerCase();

    const rows = SpreadsheetApp.openById(PDR_ALLOWLIST_SHEET_ID).getSheets()[0].getDataRange().getValues();
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

function pdrJsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
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
