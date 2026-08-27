// ═══════════════════════════════════════════════════════════════════
// Chapter Completion Tracker — writes a small, precomputed status table
// instead of shipping the full academic-year daily log (21,910+ rows) to
// every browser just so it can derive "which chapters are overdue."
//
// Plan reviewed and approved 2026-08-27 (see the chapter-tracker-plan
// Artifact from that session for the full writeup).
//
// ALGORITHM v2 (2026-08-27) — replaces the original chronological-pairing
// port with NAME-based matching, per Uday: "the CW HW Sheet is the actual
// truth... chronology is not important, a teacher has the right to
// complete say chapter 2 before chapter 1." The original v1 (planned[i]
// paired with actual[i] by date order) is gone entirely, not kept as a
// fallback -- guessing by order was shown to misattribute real examples
// (e.g. LMS4 Class 3 English: an actual CWA for "The Special One" got
// wrongly credited to "Lara, the Yellow Ladybird" purely because it was
// the first CWA logged that year). Matching a chapter's own recorded name
// (cwChapter on the daily record) against Course Mapping's chapterName is
// far more reliable -- 95% exact-match rate wherever a plan exists,
// confirmed against live data before this was built. Where no name match
// is found, the event is 'Unmapped' rather than guessed at.
//
// DEADLINE MODEL v3 (2026-08-27) — CALENDAR-INTEGRATED. Course Mapping
// dates are something teachers already plan around the school calendar --
// so rather than a number hand-typed into this file once and left to
// drift, Term/board deadlines are read LIVE from each campus's Google
// Calendar every cache cycle (chapterBuildDeadlineSets, below), the same
// calendars teachers use. Rule, per Uday: "teaching needs to end one day
// before Revision starts." Each campus can have a different date (LMS1's
// Term 2 runs a few days later than LMS2-6), and board classes (10, 12)
// run on an entirely different, earlier calendar since Dec-Feb is taken
// up by two rounds of Pre-Board revision, not new teaching:
//   Regular:  Term 1 ends ~17 Sep      Term 2 ends ~14/17 Feb (LMS1 later)
//   Board:    Term 1 ends ~10 Jul      Term 2 ends ~30 Nov
//   (Pre-Board 1/2, Jan-Feb, are cumulative full-syllabus review -- not a
//   new-content deadline, so they don't get their own boundary here.)
// If a campus's calendar is missing the expected event (renamed, not
// published yet, access hiccup), that ONE campus/phase falls back to the
// hardcoded date above and logs a warning -- see chapterResolveBoundary.
// This is deliberately NOT a hard failure: one stale calendar shouldn't
// break every other campus's deadlines. Run logChapterDeadlines() any
// time to see, at a glance, which campuses are live vs. fell back.
//
// Course Mapping's own Term column is ~40% blank/inconsistent right now
// (see the course-mapping-data-quality export from 2026-08-27) -- rather
// than block on a mid-year data cleanup, a chapter with no usable Term
// value falls back to "due by the end of its own planned month" instead.
//
// TEACHER NAMES v4 (2026-08-27) — same casing bug fixed for Missing
// Uploads earlier this session ("Arpna Sood" vs "ARPNA SOOD" in the daily
// sheet) had resurfaced here: the Teacher column would show BOTH castings
// for the same person, comma-joined, since they were being collected as
// distinct object keys. chapterFetchRosterCanon() now fetches the same
// Employee Roster Proxy Missing Uploads already trusts (EmpAcademic --
// the authoritative teacher-per-class-subject source) once per run, and
// teacherIndex is keyed by LOWERCASED name so same-person casing variants
// collapse into one entry, displayed using the roster's canonical casing
// when that combo is covered there, or the first daily-sheet casing seen
// otherwise. Degrades gracefully (logs a warning, keeps the dedup-only
// behavior) if the proxy is unreachable -- never blocks the main rebuild.
//
// SETUP:
//   1. Paste this file into the SAME Apps Script project as Code.gs --
//      needs CM_SCHOOLS/CM_TABS/ROMAN, already defined there.
//   2. This version calls CalendarApp for the first time -- the NEXT time
//      you run anything in this project from the editor, Google will ask
//      you to re-authorize with an added "See your calendars" permission.
//      That's expected; approve it (it's read-only, and only reads the 6
//      LMS school calendars, not your personal one).
//   3. Run updateChapterTrackerManual() to rebuild immediately. Check the
//      execution log (or run logChapterDeadlines() separately) -- it logs
//      a warning for any campus/phase that couldn't find a live calendar
//      event and used the hardcoded fallback instead.
//   4. Code.gs's updateCache() already calls this (see the bottom of that
//      file) -- from here on it refreshes on the existing 30-minute
//      trigger, no new schedule to set up.
//
// This does NOT change anything visible in the portal yet. Nothing reads
// from this sheet until it's wired in deliberately.
// ═══════════════════════════════════════════════════════════════════

const CHAPTER_TRACKER_SHEET_ID = '11u0Aeonc_CdwDa7ysXrdsF6vs2xYKxckQjOnc7ZEuCM';

// Same "LMCS Employee Roster Proxy" Missing Uploads already trusts for
// canonical teacher-per-class-subject data (EmpAcademic, via a narrow
// read-only proxy in front of the Employee Master workbook). See TEACHER
// NAMES v4 note at the top of this file.
const ROSTER_PROXY_URL = 'https://script.google.com/macros/s/AKfycbyHiaZY_iWK2VTKKFJcCsBNnIbUndJYUSjnPkxvJ-dYavaihiul2xBJuJohPRsP9Spf/exec';

const CHAPTER_CWA_EXEMPT_SUBJECTS = ['general knowledge', 'computer'];

const CHAPTER_TRACKER_HEADERS = [
  'School', 'Class', 'Subject', 'Chapter No.', 'Chapter Name', 'Teacher',
  'Term', 'Planned CWA Date', 'Effective Deadline', 'Actual CWA Date',
  'Status', 'Days Late/Overdue', 'Last Updated',
];

function chapterNormalizeSubject(s) {
  return String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '').slice(0, 12);
}

// Unlike subject normalization, chapter-name matching is NOT truncated --
// chapter names are long and distinctive enough that truncation risks
// false collisions ("The Lost Chi..." matching two different chapters),
// where subjects (short, few distinct values) don't have that risk.
function chapterNormalizeName(s) {
  return String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function chapterIsSubjectExempt(subject) {
  const sl = String(subject || '').toLowerCase();
  return CHAPTER_CWA_EXEMPT_SUBJECTS.some(e => sl === e || sl.includes(e) || e.includes(sl));
}

function chapterHasCWA(tagStr, subject) {
  if (chapterIsSubjectExempt(subject)) return false;
  return String(tagStr || '').split(',').map(t => t.trim()).some(t => t.toUpperCase() === 'CWA');
}

// teacherCanon[lmsIdx|cls|normSubject] -> canonical teacher name from the
// Employee Roster Proxy (EmpAcademic). One fetch per run (called once from
// buildChapterStatusRows, never per-chapter). Falls back to an empty map
// (teacher display still dedupes by lowercase, just without the roster's
// authoritative casing) if the proxy is unreachable -- see TEACHER NAMES
// v4 note at the top of this file.
function chapterFetchRosterCanon() {
  const map = {};
  try {
    const res = UrlFetchApp.fetch(ROSTER_PROXY_URL + '?action=roster', { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) throw new Error('HTTP ' + res.getResponseCode());
    const json = JSON.parse(res.getContentText());
    if (!json.success || !Array.isArray(json.rows)) throw new Error(json.error || 'bad response shape');
    json.rows.forEach(r => {
      const m = String(r.school || '').match(/(\d+)/);
      if (!m || !r.teacher) return;
      const lmsIdx = +m[1] - 1;
      const key = lmsIdx + '|' + r.class + '|' + chapterNormalizeSubject(r.subject);
      map[key] = r.teacher;
    });
  } catch (e) {
    Logger.log('[ChapterTracker] Employee Roster Proxy unreachable, teacher names will fall back to daily-sheet casing (still deduped): ' + e);
  }
  return map;
}

// Board classes (10, 12) run a different calendar entirely -- see the
// DEADLINE MODEL v3 note at the top of this file.
function chapterIsBoardClass(cls) {
  const m = String(cls || '').match(/class\s*(\d+)/i);
  if (!m) return false;
  const n = +m[1];
  return n === 10 || n === 12;
}

function chapterAcademicYear(now) {
  const startYr = now.getMonth() >= 3 ? now.getFullYear() : now.getFullYear() - 1;
  return {
    startYr,
    label: startYr + '-' + String(startYr + 1).slice(2),
    start: new Date(startYr, 3, 1),                 // Apr 1 -- whole-year data window
    end: new Date(startYr + 1, 2, 31, 23, 59, 59),   // Mar 31 next year -- window only, NOT a
                                                      // deadline; real Term/board deadlines come
                                                      // from chapterBuildDeadlineSets below.
  };
}

// ==================== CALENDAR-DERIVED TERM/BOARD DEADLINES ====================

const CHAPTER_CALENDAR_IDS = {
  1: 'c_88402c40e6c9a435cbed4810bfae632d781ac9e7ca167b5435e4a9586bda7595@group.calendar.google.com',
  2: 'c_25n5qbggh7u9s5o6vq31iho4vc@group.calendar.google.com',
  3: 'c_25n5qbggh7u9s5o6vq31iho4vc@group.calendar.google.com', // LMS2 & LMS3 share one calendar
  4: 'lms.org.in_snflpd3f357jv8l380s8sc4ju0@group.calendar.google.com',
  5: 'c_d367ca2cb9bbab4f6a5b6864d6e03c28fbfccf02b5212621e89daf2f9f3f8523@group.calendar.google.com',
  6: 'c_3db6cb2edb78d889cf166771c426f0fe3ba78b04cf9ca8e566329c888d023635@group.calendar.google.com',
};

// Confirmed against the LMS calendars on 2026-08-27 -- used ONLY when the
// live lookup below can't find a matching event for that campus/phase.
// Safe to leave as-is indefinitely; the live lookup supersedes these
// automatically as each new year's calendar is published. Update only if
// you want a fresher manual baseline for some reason.
const CHAPTER_FALLBACK_DEADLINES = {
  regular: {
    term1End: (yr) => new Date(yr, 8, 17, 23, 59, 59),                 // 17 Sep
    term2End: (yr, lms) => lms === 1
      ? new Date(yr + 1, 1, 17, 23, 59, 59)                            // LMS1: 17 Feb
      : new Date(yr + 1, 1, 14, 23, 59, 59),                           // LMS2-6: 14 Feb
  },
  board: {
    term1End: (yr) => new Date(yr, 6, 10, 23, 59, 59),                 // 10 Jul
    term2End: (yr) => new Date(yr, 10, 30, 23, 59, 59),                // 30 Nov
  },
};

// Earliest Revision/Examination event starting inside [windowStart,
// windowEnd) that is/isn't board-specific per `board`, excluding Pre-Board
// checkpoints entirely (those are cumulative review, not a new-content
// deadline -- see file header). Returns the event's start Date, or null.
function chapterFindEventStart(calendarId, windowStart, windowEnd, board) {
  const cal = CalendarApp.getCalendarById(calendarId);
  if (!cal) return null;
  const events = cal.getEvents(windowStart, windowEnd);
  let earliest = null;
  events.forEach(ev => {
    const title = (ev.getTitle() || '').toLowerCase();
    if (!/revision|exam/.test(title)) return;
    if (/pre[\s-]?board/.test(title)) return;
    if (/board/.test(title) !== board) return;
    const start = ev.getStartTime();
    if (!earliest || start < earliest) earliest = start;
  });
  return earliest;
}

function chapterDayBefore(d) {
  const b = new Date(d);
  b.setDate(b.getDate() - 1);
  b.setHours(23, 59, 59, 0);
  return b;
}

// One campus/phase boundary: try the live calendar first, fall back (with
// a logged warning) to the hardcoded date on a miss -- calendar not found,
// no matching event in the window, or an outright lookup error. A miss on
// one campus/phase never affects any other.
function chapterResolveBoundary(lms, calendarId, windowStart, windowEnd, board, fallbackDate, label) {
  try {
    const start = chapterFindEventStart(calendarId, windowStart, windowEnd, board);
    if (start) return chapterDayBefore(start);
    Logger.log('[ChapterTracker] Calendar fallback: no ' + label + ' event found for LMS' + lms +
      ' in ' + windowStart.toDateString() + '..' + windowEnd.toDateString() +
      ' -- using hardcoded ' + fallbackDate.toDateString());
  } catch (e) {
    Logger.log('[ChapterTracker] Calendar lookup FAILED for LMS' + lms + ' ' + label + ': ' + e +
      ' -- using hardcoded ' + fallbackDate.toDateString());
  }
  return fallbackDate;
}

// Builds { 1: {regular:{term1End,term2End}, board:{term1End,term2End}}, ...
// for LMS1-6 }. Call ONCE per run (cache this result, never call inside
// the per-chapter loop) -- 6 campuses x 2 phases x 2 tracks = 24
// CalendarApp lookups total per cache cycle, not per row.
function chapterBuildDeadlineSets(ay) {
  const yr = ay.startYr;
  const sets = {};
  for (let lms = 1; lms <= 6; lms++) {
    const calId = CHAPTER_CALENDAR_IDS[lms];
    sets[lms] = {
      regular: {
        term1End: chapterResolveBoundary(lms, calId, new Date(yr, 7, 15), new Date(yr, 9, 15), false,
          CHAPTER_FALLBACK_DEADLINES.regular.term1End(yr), 'Term 1 Revision/Exam'),
        term2End: chapterResolveBoundary(lms, calId, new Date(yr + 1, 0, 1), new Date(yr + 1, 2, 15), false,
          CHAPTER_FALLBACK_DEADLINES.regular.term2End(yr, lms), 'Term 2 Revision/Exam'),
      },
      board: {
        term1End: chapterResolveBoundary(lms, calId, new Date(yr, 5, 1), new Date(yr, 7, 15), true,
          CHAPTER_FALLBACK_DEADLINES.board.term1End(yr), 'Board Term 1 Revision/Exam'),
        term2End: chapterResolveBoundary(lms, calId, new Date(yr, 10, 1), new Date(yr, 11, 20), true,
          CHAPTER_FALLBACK_DEADLINES.board.term2End(yr), 'Board Term 2 Revision/Exam'),
      },
    };
  }
  return sets;
}

// Run this by itself from the Apps Script editor any time to see, at a
// glance, which campus/phase deadlines came from a live calendar lookup
// vs. a hardcoded fallback -- much faster than reading a full rebuild log.
function logChapterDeadlines() {
  const ay = chapterAcademicYear(new Date());
  const sets = chapterBuildDeadlineSets(ay);
  Logger.log('Academic year: ' + ay.label);
  for (let lms = 1; lms <= 6; lms++) {
    const s = sets[lms];
    Logger.log('LMS' + lms +
      '  Regular T1=' + s.regular.term1End.toDateString() + ' T2=' + s.regular.term2End.toDateString() +
      '  |  Board T1=' + s.board.term1End.toDateString() + ' T2=' + s.board.term2End.toDateString());
  }
}

// Normalizes the free-text Term column (which has real inconsistencies --
// blank ~40% of the time, plus "term 1" lowercase, "Both terms"/"Both
// Terms" casing, stray "Term 3"/"Term 2/1" values) into exactly one of
// 'term1' / 'term2' / 'both' / null (null = unusable, fall back to month).
function chapterNormalizeTerm(term) {
  const t = String(term || '').trim().toLowerCase();
  if (t === 'term 1') return 'term1';
  if (t === 'term 2') return 'term2';
  if (t === 'both terms') return 'both';
  return null; // blank, 'term 3', 'term 2/1', or anything else non-standard
}

// The deadline actually used to judge a chapter On Time vs Late vs
// Overdue. Term 1/Term 2/Both Terms use the real (campus- and
// board-aware) term boundary from `deadlines`; anything else (the ~40%
// currently blank/non-standard) falls back to the end of the month the
// chapter was itself planned for -- always available, since every chapter
// has its own planned CWA date regardless of Term.
function chapterEffectiveDeadline(term, cwaDateObj, deadlines) {
  const norm = chapterNormalizeTerm(term);
  if (norm === 'term1') return deadlines.term1End;
  if (norm === 'term2' || norm === 'both') return deadlines.term2End;
  if (!cwaDateObj) return deadlines.term2End; // no date at all -- last resort, final term boundary
  return chapterMonthEnd(cwaDateObj);
}

function chapterMonthEnd(d) {
  return new Date(d.getFullYear(), d.getMonth() + 1, 0, 23, 59, 59); // last day of that month
}

function chapterDateToStr(d) {
  if (!d) return '';
  return Utilities.formatDate(d, 'Etc/GMT', 'yyyy-MM-dd');
}

// Three-tier escalation for a chapter with no matching CWA yet, per Uday
// (2026-08-27): "a soft falling behind pace if the [planned] dates are
// missed, a stronger Overdue if the month is missed, a harsh warning if
// the term is likely to be missed." Boundaries, softest to harshest:
//   plannedDate (the literal Course Mapping date) -> monthEnd (end of
//   that month) -> termDeadline (Term end, or monthEnd again as fallback
//   when Term isn't usable -- see chapterEffectiveDeadline).
// When termDeadline === monthEnd (the fallback case), Overdue and Critical
// collapse into one jump straight from Behind Pace to Critical -- an
// honest consequence of not knowing the real Term for that chapter, not a
// bug; it's part of the incentive to get Term filled in from Term 2
// onward.
function chapterPendingStatus(now, plannedDate, monthEnd, termDeadline) {
  if (now < plannedDate) return { status: 'Pending', days: '' };
  if (now < monthEnd) return { status: 'Behind Pace', days: Math.round((now - plannedDate) / 86400000) };
  if (now < termDeadline) return { status: 'Overdue', days: Math.round((now - monthEnd) / 86400000) };
  return { status: 'Critical', days: Math.round((now - termDeadline) / 86400000) };
}

/** Cross-references planned Course Mapping chapters against actual CWA
 *  events, matched by CHAPTER NAME (not chronological order -- see file
 *  header). cwRecords/cmData are exactly what buildData() in Code.gs
 *  already builds -- this reuses them, no extra fetching. */
function buildChapterStatusRows(cwRecords, cmData, ay, now) {
  const deadlineSets = chapterBuildDeadlineSets(ay); // one CalendarApp pass for the whole run
  const teacherCanon  = chapterFetchRosterCanon();   // one Roster Proxy fetch for the whole run

  // dailyCWA[lmsIdx][classMapped][normSubject] -> list of {dateObj,dateStr,chapter,teacher,subject}
  const dailyCWA = {};
  cwRecords.forEach(r => {
    if (!chapterHasCWA(r.cwTag, r.subject)) return;
    const d = new Date(r.timeMs);
    if (d < ay.start || d > ay.end) return;
    if (!r.classMapped) return;
    const ns = chapterNormalizeSubject(r.subject);
    if (!dailyCWA[r.lmsIdx]) dailyCWA[r.lmsIdx] = {};
    if (!dailyCWA[r.lmsIdx][r.classMapped]) dailyCWA[r.lmsIdx][r.classMapped] = {};
    if (!dailyCWA[r.lmsIdx][r.classMapped][ns]) dailyCWA[r.lmsIdx][r.classMapped][ns] = [];
    dailyCWA[r.lmsIdx][r.classMapped][ns].push({
      dateObj: d, dateStr: r.dateStr, chapter: r.cwChapter, teacher: r.teacher, subject: r.subject,
    });
  });
  Object.keys(dailyCWA).forEach(li => Object.keys(dailyCWA[li]).forEach(cls =>
    Object.keys(dailyCWA[li][cls]).forEach(ns => dailyCWA[li][cls][ns].sort((a, b) => a.dateObj - b.dateObj))
  ));

  // teacherIndex[lmsIdx|classMapped|normSubject] -> Map(lowercased name ->
  // display name), for chapters that have no actual match yet (so a
  // Pending/Overdue row still shows who's responsible). Keyed by lowercase
  // so the same person logged under different daily-sheet casing ("ARPNA
  // SOOD" vs "Arpna Sood") collapses into ONE entry instead of listing
  // both -- see TEACHER NAMES v4 note at the top of this file. Display
  // casing prefers the roster's canonical name for that combo; falls back
  // to the first daily-sheet casing seen if the roster doesn't cover it.
  const teacherIndex = {};
  cwRecords.forEach(r => {
    const d = new Date(r.timeMs);
    if (d < ay.start || d > ay.end) return;
    if (!r.classMapped || !r.teacher) return;
    const key = r.lmsIdx + '|' + r.classMapped + '|' + chapterNormalizeSubject(r.subject);
    if (!teacherIndex[key]) teacherIndex[key] = new Map();
    const lower = r.teacher.trim().toLowerCase();
    if (!teacherIndex[key].has(lower)) teacherIndex[key].set(lower, teacherCanon[key] || r.teacher.trim());
  });
  function lookupTeacher(lmsIdx, cls, subj) {
    const ns = chapterNormalizeSubject(subj);
    const key = lmsIdx + '|' + cls + '|' + ns;
    if (teacherIndex[key]) return [...teacherIndex[key].values()].join(', ');
    const fKey = Object.keys(teacherIndex).find(k => {
      const parts = k.split('|');
      return +parts[0] === lmsIdx && parts[1] === cls && (parts[2] === ns || parts[2].includes(ns) || ns.includes(parts[2]));
    });
    return fKey ? [...teacherIndex[fKey].values()].join(', ') : '';
  }
  function getDailyArr(lmsIdx, cls, subj) {
    const byS = (dailyCWA[lmsIdx] || {})[cls];
    if (!byS) return [];
    const ns = chapterNormalizeSubject(subj);
    if (byS[ns]) return byS[ns];
    const key = Object.keys(byS).find(k => k === ns || k.includes(ns) || ns.includes(k));
    return key ? byS[key] : [];
  }

  const rows = [];
  const processedCombos = {}; // 'lmsIdx|class|normSubject' -> true, for the second pass below

  Object.keys(cmData).forEach(lmsIdxStr => {
    const lmsIdx = +lmsIdxStr;
    const lms = lmsIdx + 1;
    const schoolName = 'LMS ' + lms;
    const tabs = cmData[lmsIdxStr];
    Object.keys(tabs).forEach(cls => {
      const chapters = tabs[cls];
      if (!Array.isArray(chapters)) return;
      const deadlines = deadlineSets[lms][chapterIsBoardClass(cls) ? 'board' : 'regular'];
      const bySubj = {};
      chapters.forEach(ch => {
        // CWA-exempt subjects (General Knowledge, Computer) can never get a
        // CWA tag, so a planned chapter here would sit as perpetually
        // Pending/Overdue with no way to ever resolve. Excluded from
        // tracking entirely -- per Uday 2026-08-27, for this academic year;
        // a real system for these subjects is planned for next year, not
        // a permanent design decision made here.
        if (chapterIsSubjectExempt(ch.subject)) return;
        const ns = chapterNormalizeSubject(ch.subject);
        if (!bySubj[ns]) bySubj[ns] = [];
        bySubj[ns].push(ch);
      });
      Object.keys(bySubj).forEach(ns => {
        const chList = bySubj[ns];
        const planned = chList
          .filter(ch => ch.cwaDate)
          .map(ch => Object.assign({}, ch, { cwaDateObj: new Date(ch.cwaDate) }))
          .filter(ch => ch.cwaDateObj >= ay.start && ch.cwaDateObj <= ay.end)
          .sort((a, b) => (a.chapterNo || 0) - (b.chapterNo || 0));
        if (!planned.length) return;

        processedCombos[lmsIdx + '|' + cls + '|' + ns] = true;

        const actual = getDailyArr(lmsIdx, cls, chList[0].subject).slice();
        const teacher = lookupTeacher(lmsIdx, cls, chList[0].subject);
        const used = new Array(actual.length).fill(false);

        planned.forEach(ch => {
          const wantName = chapterNormalizeName(ch.chapterName);
          let matchIdx = -1;
          if (wantName) {
            for (let j = 0; j < actual.length; j++) {
              if (used[j]) continue;
              if (chapterNormalizeName(actual[j].chapter) === wantName) { matchIdx = j; break; }
            }
          }
          const monthEnd = chapterMonthEnd(ch.cwaDateObj);
          const termDeadline = chapterEffectiveDeadline(ch.term, ch.cwaDateObj, deadlines);
          let status, actualDateStr = '', daysVal = '';
          if (matchIdx >= 0) {
            used[matchIdx] = true;
            const ev = actual[matchIdx];
            actualDateStr = ev.dateStr;
            const daysLate = Math.round((ev.dateObj - ch.cwaDateObj) / 86400000);
            status = daysLate > 0 ? 'Late' : 'OnTime';
            daysVal = daysLate > 0 ? daysLate : 0;
          } else {
            const pending = chapterPendingStatus(now, ch.cwaDateObj, monthEnd, termDeadline);
            status = pending.status;
            daysVal = pending.days;
          }
          rows.push([
            schoolName, cls, ch.subject, ch.chapterNo || '', ch.chapterName || '', teacher,
            ch.term || '(blank)', ch.cwaDate, chapterDateToStr(termDeadline), actualDateStr,
            status, daysVal, now,
          ]);
        });

        // Actual CWA events for this subject that didn't match any planned
        // chapter by name -- genuinely unmapped, not guessed at.
        actual.forEach((ev, j) => {
          if (used[j]) return;
          rows.push([
            schoolName, cls, ev.subject || chList[0].subject, '', ev.chapter || '', ev.teacher || '',
            '', '', '', ev.dateStr, 'Unmapped', '', now,
          ]);
        });
      });
    });
  });

  // Second pass: actual CWA activity in a (school, class, subject) combo
  // that has NO Course Mapping data at all -- previously silently invisible
  // (skipped entirely, since the first pass only iterates cmData). This is
  // real signal: e.g. Class 6 Hindi has this gap at every campus.
  Object.keys(dailyCWA).forEach(lmsIdxStr => {
    const lmsIdx = +lmsIdxStr;
    const schoolName = 'LMS ' + (lmsIdx + 1);
    Object.keys(dailyCWA[lmsIdx]).forEach(cls => {
      if (cls.indexOf('Class') !== 0) return; // Montessori has no CWA by design
      Object.keys(dailyCWA[lmsIdx][cls]).forEach(ns => {
        const comboKey = lmsIdx + '|' + cls + '|' + ns;
        if (processedCombos[comboKey]) return; // already handled in the first pass
        dailyCWA[lmsIdx][cls][ns].forEach(ev => {
          rows.push([
            schoolName, cls, ev.subject || '', '', ev.chapter || '', ev.teacher || '',
            '', '', '', ev.dateStr, 'Unmapped', '', now,
          ]);
        });
      });
    });
  });

  return rows;
}

function getOrCreateTrackerSpreadsheet() {
  if (CHAPTER_TRACKER_SHEET_ID) {
    try { return SpreadsheetApp.openById(CHAPTER_TRACKER_SHEET_ID); } catch (e) { /* fall through */ }
  }
  const props = PropertiesService.getScriptProperties();
  const storedId = props.getProperty('chapter_tracker_sheet_id');
  if (storedId) {
    try { return SpreadsheetApp.openById(storedId); } catch (e) { /* fall through, recreate */ }
  }
  const ss = SpreadsheetApp.create('LMCS Chapter Completion Tracker');
  DriveApp.getFileById(ss.getId()).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  props.setProperty('chapter_tracker_sheet_id', ss.getId());
  Logger.log('Created Chapter Completion Tracker: ' + ss.getUrl());
  Logger.log('Sheet ID: ' + ss.getId());
  Logger.log('Paste this ID into CHAPTER_TRACKER_SHEET_ID at the top of this file, then redeploy.');
  return ss;
}

function getOrCreateYearTab(ss, label) {
  let sheet = ss.getSheetByName(label);
  if (!sheet) sheet = ss.insertSheet(label);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, CHAPTER_TRACKER_HEADERS.length).setValues([CHAPTER_TRACKER_HEADERS]);
  sheet.setFrozenRows(1);
  const junk = ss.getSheetByName('Sheet1');
  if (junk && ss.getSheets().length > 1) ss.deleteSheet(junk);
  return sheet;
}

/** Called from updateCache() in Code.gs, reusing the cwRecords/cmData it
 *  already built for the main dashboard cache -- no extra fetching. */
function updateChapterTracker(cwRecords, cmData) {
  const now = new Date();
  const ay = chapterAcademicYear(now);
  const rows = buildChapterStatusRows(cwRecords, cmData, ay, now);
  if (!rows.length) { Logger.log('Chapter Tracker: no rows built, skipping write'); return; }

  const ss = getOrCreateTrackerSpreadsheet();
  const sheet = getOrCreateYearTab(ss, ay.label);
  sheet.getRange(2, 1, rows.length, CHAPTER_TRACKER_HEADERS.length).setValues(rows);
  Logger.log('Chapter Tracker: wrote ' + rows.length + ' rows to tab ' + ay.label);
}

/** Manual entry point -- run this from the Apps Script editor any time to
 *  force an immediate rebuild (e.g. right after this v3 algorithm change).
 *  Fetches its own data rather than reusing a cached buildData() call. */
function updateChapterTrackerManual() {
  const data = buildData(); // from Code.gs, same project
  updateChapterTracker(data.cwRecords, data.cmData);
}
