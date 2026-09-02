// ═══════════════════════════════════════════════════════════════════
// SS Forms Sync — a generic engine for keeping every LMCS "Support
// Session" (formerly "Daily Report"/DR) Google Form correct and its
// name dropdown live, driven by EmpMaster. One role (Teacher) is
// fully wired below; the others (PTI, IT/Computer, Fee Clerk cum PRO,
// Principal Self-SS) are stubbed out — see "ADDING A NEW ROLE" at the
// bottom.
//
// TERMINOLOGY (2026-09-02): "DR" (Daily Report) for teachers was
// renamed to "SS" (Support Session) by Uday. The 6 Teacher Forms and
// their response sheet's file title have been renamed to match
// ("LMS N Teacher SS (Qualitative) 2026 for Heads",
// "LMCS Teacher SS 2026 (Responses)") — only the response SHEET'S
// internal TAB names ("LMS 1 Teacher DR" etc.) still say DR, since
// Google Forms doesn't auto-rename an already-linked destination tab
// and tab renames aren't reachable via the tools used to rename
// everything else here. teacher-ss-dashboard-proxy.gs's
// TSS_CAMPUS_TO_TAB still points at the DR-named tabs deliberately —
// don't "fix" that without renaming the actual tabs first, or the
// dashboard will break. The OTHER 4 roles' names (PTI DR, IT DR, Fee
// Clerk DR, Principal DR) were NOT part of this rename — Uday's
// correction was specifically "for teachers".
//
// WHY GENERIC (2026-09-01): per Uday, Support Session forms are
// planned for every employee category, not just teachers — see the
// "Principal's Daily Reporting" project note. The other 4 roles' forms
// already exist (PTI DR, IT DR, Fee Clerk DR, Principal DR — all
// single forms covering all 6 campuses via an in-form School question,
// unlike Teacher's one-form-per-campus layout) but are still on the
// OLD rubric/structure and are explicitly not-yet-viewable/finalized —
// their rubric, Principal-performance factors, and comms system are
// deferred to Uday (see project_principals_daily_reporting.md). DO NOT
// guess their question structure — read the live form with FormApp
// first (see ADDING A NEW ROLE) once Uday finalizes each one, rather
// than assuming it matches Teacher's shape.
//
// IMPORTANT (2026-09-02, confirmed by Uday, correcting an earlier
// "fresh backend, no Forms" direction from a different session): the
// 6 Teacher SS Google Forms ARE the real, ongoing submission mechanism
// for the Teaching-SS portal module — this script and the Forms it
// fixes are NOT being replaced. Keep this file; do not delete it in
// favor of a Forms-free rebuild without checking with Uday first.
//
// SETUP:
//   1. script.google.com -> "LMCS Principal's Daily Reporting Backend"
//      project (2026-09-02: renamed/consolidated, see main.gs's header) ->
//      paste this file in as ss-forms-sync.gs
//   2. Run syncAllSSForms() once manually (Run menu) — authorize when
//      prompted (edit access to the forms in SS_ROLE_CONFIGS, read
//      access to the Employee Master workbook)
//   3. Run installDailySSSyncTrigger() once — re-runs syncAllSSForms()
//      daily so every configured role's teacher list stays live and
//      any range fixes stay applied.
// ═══════════════════════════════════════════════════════════════════

const SS_EMP_SHEET_ID = '1OjVMUvpLM8JkdAwjmljCtZUI1VUqGLbic36cW9dm0C0'; // LMCSStaffMasters
const SS_PREFIX_TO_SCHOOL = {
  KUL: 'LMS 1', KEL: 'LMS 2', DUN: 'LMS 3', NCM: 'LMS 4', SAY: 'LMS 5', JOG: 'LMS 6',
};

// KNOWN DATA-QUALITY CAVEAT (2026-09-02, from the "Designation & Department
// Tags" audit): EmpSalary's Department column is the canonical 3-way split
// (Teaching / Non-Teaching / Administrative), joined here by EmployeeCode --
// same join key EmpAcademic already uses against EmpMaster. BUT as of the
// audit, a handful of clearly-teaching records (some TGT, a PTI "Games
// Teacher") had a BLANK Department cell instead of "Teaching", not yet
// corrected. If a real teacher unexpectedly disappears from a campus's
// dropdown after this filter went live, check their EmpSalary row's
// Department cell first before assuming a bug here.
function getDepartmentByCode_() {
  const rows = SpreadsheetApp.openById(SS_EMP_SHEET_ID)
    .getSheetByName('EmpSalary')
    .getDataRange()
    .getValues();
  const header = rows[0];
  const codeCol = header.indexOf('EmployeeCode');
  const deptCol = header.indexOf('Department');
  const map = {};
  if (codeCol < 0 || deptCol < 0) return map; // sheet shape unexpected -- fail open to empty, not a crash
  for (let i = 1; i < rows.length; i++) {
    const code = String(rows[i][codeCol] || '').trim();
    if (!code) continue;
    map[code] = String(rows[i][deptCol] || '').trim();
  }
  return map;
}

// ── Role configs ──────────────────────────────────────────────────
// Each entry describes ONE SS form family. `perCampus: true` means one
// form per school (keyed 'LMS 1'..'LMS 6', like Teacher); `perCampus:
// false` means a single form covering all schools (name field doesn't
// vary by school, and the employee-choice list should span every
// school this role applies to, typically distinguished by including a
// School question elsewhere in the form that this engine leaves
// alone).
const SS_ROLE_CONFIGS = {
  teacher: {
    label: 'Teacher SS',
    perCampus: true,
    forms: {
      'LMS 1': '1BvUT49YVbGPAH-BLBdjkZ6FoBmMLnzNDQlRu3j2bIEw',
      'LMS 2': '16T084-3WDTnpINA76qfSKODXXhOM0wO5B4qMDcE8w4M',
      'LMS 3': '1rrM4Y6jE_k23WLMLqjeB9VPNC_EzsTu3YT_Z-GRqBQk',
      'LMS 4': '1lL1gF2D_2WgsRKk_4pAMesuFCgwvpIGlYpyuPXwNIro',
      'LMS 5': '11otlIiFKWQ5tZ2JfxyyL4pn_5ManI8Oc1LoIbj3o2-0',
      'LMS 6': '1VZhZ2b1nZlRT2X_YiqJhvArlpnd5WjOaXGv1WcROXpI',
    },
    rubricTitles: [
      'Pedagogy',
      'Teacher Content',
      'Student Response',
      'Activity Based Approach',
      'Notebook Checking of Students',
      'BUTCHER Scheme',
      'Discipline- English, Aerobics, Morning Assembly & Dispersal',
      'Extra Curricular (Non Academic Competitions)',
      'Attitude Towards School',
      'Dedication Towards Class',
    ],
    rubricMax: 10,
    // Per-campus form -> its name-question title is literally that
    // campus's label + ' Teacher Name' (matches the live forms exactly
    // -- this field's title was NOT part of the DR->SS rename).
    nameFieldTitle: function (school) { return school + ' Teacher Name'; },
    // Which EmpMaster rows count as "a Teacher" for this campus: same
    // campus AND EmpSalary.Department === 'Teaching' (case-insensitive,
    // trimmed -- see the data-quality caveat above getDepartmentByCode_()).
    matchesEmployee: function (empRow, school) {
      return empRow.school === school && empRow.department.toLowerCase() === 'teaching';
    },
  },

  // ── TODO stubs — fill in once Uday finalizes each role's form ──
  // (rubric + Principal-performance factors + comms system are all
  // still open per project_principals_daily_reporting.md). For each:
  //   1. Open the live/rebuilt form and confirm its real question
  //      titles + current structure with FormApp (don't assume it
  //      mirrors Teacher's) — e.g. run:
  //        Logger.log(FormApp.openById('<id>').getItems().map(function(i){
  //          return i.getTitle() + ' [' + i.getType() + ']';
  //        }));
  //   2. Fill in the config below using that real structure.
  //   3. Add the role's key to ACTIVE_SS_ROLES.
  pti: null,              // PTI DR — existing old form: 1oAtyo-Q3bm3bbAcrngoPOcUgYO3jgPd9vcFOCgscJYY
  itComputer: null,       // IT DR — existing old form: 1E1hsC4EDjeeCszK3HviOFK_3NrwZ24E1QMBm04WBcgc
  feeClerkPRO: null,      // Fee Clerk DR — existing old form: 1mpjTQisZ6BwsPRI2tjc2MgNO05-Y63X5blLdRzSUooM
  principalSelfDR: null,  // Principal DR — existing old form: 1IypepIVAQ7n4vR-EvMLAoz4QzjeHoQVeg4JZnL5Vd3o
  // NOTE: these 4 keep the "DR" label deliberately -- Uday's 2026-09-02
  // rename was specifically "for teachers", not these roles.
};

// Which of the keys above actually run. Add a key here once its config
// above is filled in — keeps syncAllSSForms() from erroring on the
// still-null stubs.
const ACTIVE_SS_ROLES = ['teacher'];

// ── Generic engine — role-agnostic, do not edit per-role ───────────

function syncAllSSForms() {
  const allResults = {};
  ACTIVE_SS_ROLES.forEach(function (roleKey) {
    const config = SS_ROLE_CONFIGS[roleKey];
    if (!config) throw new Error('ACTIVE_SS_ROLES references unconfigured role: ' + roleKey);
    allResults[roleKey] = syncOneSSRole_(config);
  });
  Logger.log(JSON.stringify(allResults, null, 2));
  return allResults;
}

function syncOneSSRole_(config) {
  const results = [];
  Object.keys(config.forms).forEach(function (formKey) {
    // formKey is a school ('LMS 1'..) when perCampus, otherwise a
    // single fixed key (e.g. 'all') — see per-role config comments.
    const form = FormApp.openById(config.forms[formKey]);
    const employees = getActiveEmployeeChoices_(config, formKey);
    let rangeFixed = 0;
    let nameListUpdated = false;

    form.getItems().forEach(function (item) {
      const title = item.getTitle();

      if (config.rubricTitles.indexOf(title) >= 0 && item.getType() === FormApp.ItemType.TEXT) {
        item.asTextItem().setValidation(
          FormApp.createTextValidation()
            .requireNumberBetween(0, config.rubricMax)
            .setHelpText('Enter a score between 0 and ' + config.rubricMax + '.')
            .build()
        );
        rangeFixed++;
      }

      if (title === config.nameFieldTitle(formKey) && item.getType() === FormApp.ItemType.LIST) {
        item.asListItem().setChoiceValues(employees);
        nameListUpdated = true;
      }
    });

    results.push({ formKey: formKey, rangeFixed: rangeFixed, nameListUpdated: nameListUpdated, employeeCount: employees.length });
  });
  return results;
}

/** EmpMaster -> "EmployeeCode Name" strings, Active only, filtered
 *  through the role's matchesEmployee(), sorted by name. */
function getActiveEmployeeChoices_(config, formKey) {
  const departmentByCode = getDepartmentByCode_();
  const rows = SpreadsheetApp.openById(SS_EMP_SHEET_ID)
    .getSheetByName('EmpMaster')
    .getDataRange()
    .getValues();
  const header = rows[0];
  const codeCol = header.indexOf('EmployeeCode');
  const nameCol = header.indexOf('Name');
  const statusCol = header.indexOf('Status');

  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const code = String(rows[i][codeCol] || '').trim();
    if (!code) continue;
    const prefix = code.split('/')[0];
    const school = SS_PREFIX_TO_SCHOOL[prefix];
    if (!school) continue;
    const status = statusCol >= 0 ? String(rows[i][statusCol] || '').trim().toLowerCase() : '';
    if (status && status !== 'active') continue; // departed/transferred — exclude

    const empRow = {
      code: code,
      name: String(rows[i][nameCol] || '').trim(),
      school: school,
      department: departmentByCode[code] || '',
    };
    if (!config.matchesEmployee(empRow, formKey)) continue;
    out.push(empRow.code + ' ' + empRow.name);
  }
  out.sort(function (a, b) { return a.localeCompare(b); });
  return out;
}

/** Run once to keep every active role's name lists synced going
 *  forward — same daily-refresh pattern as Code.gs's updateCache(). */
function installDailySSSyncTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncAllSSForms') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncAllSSForms')
    .timeBased()
    .everyDays(1)
    .atHour(5)
    .create();
  Logger.log('Daily trigger installed: syncAllSSForms() will run once a day around 5 AM.');
}

// ── ADDING A NEW ROLE ───────────────────────────────────────────────
// 1. Get the live form's real question titles (don't assume — see the
//    Logger.log one-liner in the TODO comment above).
// 2. Fill in SS_ROLE_CONFIGS.<role> with: forms (id map — one key if
//    perCampus:false, one per campus if true), rubricTitles,
//    rubricMax, nameFieldTitle(formKey), matchesEmployee(empRow, formKey).
// 3. Add '<role>' to ACTIVE_SS_ROLES.
// 4. Re-run syncAllSSForms() manually once to confirm it applies
//    cleanly before relying on the daily trigger for it.
