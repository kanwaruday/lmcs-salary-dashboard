// ═══════════════════════════════════════════════════════════════════
// LMCS Principal's Daily Reporting — main dispatcher for the single
// "LMCS Principal's Daily Reporting Backend" Apps Script project.
// Powers principals-daily-reporting/index.html (renamed 2026-09-02
// from teacher-ss/, which undersold what this page actually covers).
//
// FINAL architecture as of 2026-09-02 (after two earlier same-day
// attempts: first "shared project, separate files", then "fully
// separate projects per concern" -- both superseded by this, per
// Uday): ONE project, ONE doGet (Apps Script's hard limit), MANY
// per-concern files. Add a new concern (PTI SS, IT SS, Clerk SS,
// Helpers SS, the SS completion/pending tracker, etc.) as ITS OWN NEW
// FILE in this same project, following the naming convention below --
// never a new project, never folded into an unrelated file.
//
// Files in this project (2026-09-02):
//   main.gs            -- this file: doGet dispatch + shared
//                         auth/JSON helpers every concern reuses.
//   teacher-ss.gs       -- Teacher SS (Support Session teacher-
//                         evaluation rubric) dashboard stats.
//   principal-dr.gs     -- Principal DR (the principal's own daily
//                         operational report): Month Activities,
//                         Support Session count, Daily Reports
//                         persistence + the Tasks Completed suggestion
//                         lookup, Planned Activities CRUD. Stays ONE
//                         file even as more gets added -- no further
//                         splitting, per Uday.
//   ss-forms-sync.gs    -- Run-menu only, no doGet -- fixes/keeps live
//                         the Google Forms behind Teacher SS (and,
//                         once built, PTI/IT/Clerk/Helpers SS -- it's
//                         already a generic multi-role engine).
//   [future] pti-ss.gs, it-ss.gs, clerk-ss.gs, helpers-ss.gs,
//     ss-tracker.gs -- not built yet, add when each is.
//
// Naming convention: shared helpers/constants (this file) have no
// prefix. Concern-specific files use a short prefix matching their
// name (TSS_/tss for teacher-ss.gs, PDR_/pdr for principal-dr.gs, and
// so on for future files) so nothing collides across files sharing
// this one project's global scope.
//
// GATING: every action requires a verified Google ID token, checked
// against the "LMCS Principal Allowlist" sheet -- Principal/
// Coordinator/Owner only, same as every other backend in this portal.
//
// SETUP:
//   1. script.google.com -> New project -> name it "LMCS Principal's
//      Daily Reporting Backend"
//   2. Paste this file in as main.gs, plus every other file listed
//      above as its own file with that exact name
//   3. Deploy -> New deployment -> Web App
//      Execute as: Me | Who has access: Anyone
//      (Apps Script's own access setting isn't the real gate -- the
//      verified-token check inside is)
//   4. Copy the ONE resulting Web App URL into
//      principals-daily-reporting/index.html's PDR_BACKEND_URL
//      constant
// ═══════════════════════════════════════════════════════════════════

const ALLOWLIST_SHEET_ID = '1NZu0ElismFytG395Nxjz29vAz7OfkmJtZhs70bOwT58'; // "LMCS Principal Allowlist"
const GOOGLE_CLIENT_ID = '697999989724-mvi85iobr20g4mm8a8nrjd1rms2o8tf6.apps.googleusercontent.com'; // same client ID assets/auth.js signs in with

function doGet(e) {
  try {
    const caller = verifyCallerToken_(e.parameter.idToken);
    if (!caller) return jsonOut_({ success: false, error: 'Not authorized' });

    const action = (e.parameter.action || 'teacherstats').toLowerCase();

    if (action === 'teacherstats') return jsonOut_(teacherSsStats_(caller));
    if (action === 'monthactivities') return jsonOut_(principalDrMonthActivities_(caller, e.parameter.campusId));
    if (action === 'supportsessionstoday') return jsonOut_(principalDrSupportSessionsToday_(caller, e.parameter.campusId));
    if (action === 'plannedactivities') return jsonOut_(principalDrPlannedActivities_(caller, e.parameter.campusId));
    if (action === 'yesterdaystasks') return jsonOut_(principalDrYesterdaysTasks_(caller, e.parameter.campusId, e.parameter.date));
    if (action === 'dailyreport') return jsonOut_(principalDrGetDailyReport_(caller, e.parameter.campusId, e.parameter.date));

    return jsonOut_({ success: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonOut_({ success: false, error: err.message });
  }
}

// Writes (Submit, Planned Activities CRUD) go through doPost as a JSON
// body instead of doGet query params -- two reasons: (1) Daily Reports
// payloads carry variable-length arrays (Tasks Completed, Tasks for
// Tomorrow, etc.) that are awkward/unsafe to URL-encode; (2) the
// idToken travels in the body instead of the URL, which is marginally
// better practice for a bearer credential either way. The frontend
// deliberately sends a plain-text body (no Content-Type header) so
// browsers treat it as CORS-safelisted and skip a preflight OPTIONS
// request -- this project has no doOptions, so a preflight would fail.
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const caller = verifyCallerToken_(body.idToken);
    if (!caller) return jsonOut_({ success: false, error: 'Not authorized' });

    const action = String(body.action || '').toLowerCase();

    if (action === 'savedailyreport') return jsonOut_(principalDrSaveDailyReport_(caller, body));
    if (action === 'addplannedactivity') return jsonOut_(principalDrAddPlannedActivity_(caller, body));
    if (action === 'setplannedactivitycompleted') return jsonOut_(principalDrSetPlannedActivityCompleted_(caller, body));
    if (action === 'deleteplannedactivity') return jsonOut_(principalDrDeletePlannedActivity_(caller, body));

    return jsonOut_({ success: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonOut_({ success: false, error: err.message });
  }
}

// Verified Google ID token -> {email, campusId, role}, re-derived from
// the allowlist every call. Only Principal/Coordinator/Owner may call
// any action in this project; everyone else (e.g. Teacher role, or not
// on the list) gets null.
function verifyCallerToken_(idToken) {
  if (!idToken) return null;
  try {
    const res = UrlFetchApp.fetch(
      'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
      { muteHttpExceptions: true }
    );
    if (res.getResponseCode() !== 200) return null;
    const payload = JSON.parse(res.getContentText());
    if (payload.aud !== GOOGLE_CLIENT_ID) return null;
    if (payload.email_verified !== 'true' && payload.email_verified !== true) return null;
    const email = (payload.email || '').toLowerCase();

    const rows = SpreadsheetApp.openById(ALLOWLIST_SHEET_ID).getSheets()[0].getDataRange().getValues();
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

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
