// ═══════════════════════════════════════════════════════════════════
// LMCS Dashboard Data Proxy — Google Apps Script
//
// SETUP:
//   1. Go to script.google.com → New project → paste this file
//   2. Run clearOldCache() once to clear any stale properties
//   3. Run updateCache() once to build the first cache file in Drive
//   4. Deploy → New deployment → Web App
//      Execute as: Me  |  Who has access: Anyone
//   5. Copy the Web App URL into daily-progress.html  PROXY_URL = '...'
//   6. Add a time trigger: updateCache() → every 30 minutes
// ═══════════════════════════════════════════════════════════════════

const CW_SHEET_ID = '1z5hzpMRHMru6ulpYs2lsGZr7quAywwpQ_zySa3rvGUA';

// ── Payroll Sheet ID ─────────────────────────────────────────────
// Create a new Google Sheet, paste its ID here
const PAYROLL_SHEET_ID = '1GlyK5fu4WAwICVj0YWnILmL2U_rL8xmjNb5THPOXRhc';
const PAYROLL_TAB      = 'Employee_Master';

const PAYROLL_HEADERS = [
  'Timestamp','Employee Name','Employee Type','School','Designation',
  'Working Hours','EPF Enrolled','ESI Applicable','LMS Exp (yrs)','Other Exp (yrs)','Increments',
  'Kid 1 Grade','Kid 1 Fee','Kid 2 Grade','Kid 2 Fee',
  'Basic','DA','Add. Duty Allow.','Net Salary',
  'EPF Employer','PF Employer','ESI Employer','Total EPF','Total ESI',
  'Tuition Fee','Cost to Institution',
  'RRF Y1','RRF Y2','RRF Y3',
  'Net Bank Y1','Net Bank Y2','Net Bank Y3',
  'Status'
];

// ── Payroll doPost — receives submissions from the hiring dashboard ──
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.openById(PAYROLL_SHEET_ID);
    let   tab  = ss.getSheetByName(PAYROLL_TAB);

    // Create tab + header row on first use
    if (!tab) {
      tab = ss.insertSheet(PAYROLL_TAB);
      tab.appendRow(PAYROLL_HEADERS);
      tab.setFrozenRows(1);
      tab.getRange(1, 1, 1, PAYROLL_HEADERS.length)
         .setBackground('#1e3a5f').setFontColor('#ffffff').setFontWeight('bold');
    }

    tab.appendRow([
      data.timestamp,      data.name,          data.empType,       data.school,
      data.designation,    data.workingHours,   data.epfEnrolled,   data.esiApplicable,
      data.lmsExp,         data.otherExp,       data.increments,
      data.kid1Grade,      data.kid1Fee,        data.kid2Grade,     data.kid2Fee,
      data.basic,          data.da,             data.ada,           data.netSalary,
      data.epfEmployer,    data.pfEmployer,     data.esiEmployer,   data.totalEpf,
      data.totalEsi,       data.tuitionFee,     data.costToInstitution,
      data.rrfY1,          data.rrfY2,          data.rrfY3,
      data.netBankY1,      data.netBankY2,      data.netBankY3,
      data.status || 'Pending'
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

const CM_SCHOOLS = [
  { name: 'LMS 1', id: '1BorpCdePG7O6iZiwgGwNpWZBl7I1IGAT-OAh0fw2yMs' },
  { name: 'LMS 2', id: '1ejXKsnK3a2kNGEboTs-yQ2d3Tpk3xG5JRtm9DAmiYXs' },
  { name: 'LMS 3', id: '1MYTOPH_zFIK8xnQk3SsKIgsX5eT5mZgHaJpMgyoJhL8' },
  { name: 'LMS 4', id: '1nwPLQprQoQJ35xKDWSi2Y4YDzbvUPZ688u5QwOYmSxU' },
  { name: 'LMS 5', id: '172YKw1cMMcG5UZ0jjI8RzO2tGxnVWG3QVCEsDY6CZHg' },
  { name: 'LMS 6', id: '1PkZYbeVCdKeqCbbX2eXAwk7c8w7NAHUOSQKmO6nDZtE' },
];

const CM_TABS = [
  'Class 1','Class 2','Class 3','Class 4','Class 5','Class 6',
  'Class 7','Class 8','Class 9','Class 10','Class 11','Class 12'
];

const CACHE_TTL_MS = 30 * 60 * 1000; // 30 minutes

// ── Web App entry point ──────────────────────────────────────────
function doGet() {
  try {
    const json = getOrBuildCache();
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({ error: e.toString(), stack: e.stack }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Time trigger (every 30 min) ──────────────────────────────────
function updateCache() {
  const json = JSON.stringify(buildData());
  writeCacheFile(json);
  Logger.log('Cache updated: ' + json.length + ' bytes');
}

// ── Cache logic (Google Drive file) ─────────────────────────────
function getOrBuildCache() {
  const props  = PropertiesService.getScriptProperties();
  const ts     = parseInt(props.getProperty('cache_ts') || '0');
  const fileId = props.getProperty('cache_file_id');

  if (Date.now() - ts < CACHE_TTL_MS && fileId) {
    try {
      const content = DriveApp.getFileById(fileId).getBlob().getDataAsString('UTF-8');
      if (content && content.length > 20) {
        Logger.log('Served from cache: ' + content.length + ' bytes');
        return content;
      }
      Logger.log('Cache file empty, rebuilding');
    } catch(e) {
      Logger.log('Cache read failed: ' + e);
    }
  }

  Logger.log('Building fresh data...');
  const json = JSON.stringify(buildData());
  Logger.log('Built: ' + json.length + ' bytes');
  writeCacheFile(json);
  return json;
}

function writeCacheFile(json) {
  const props  = PropertiesService.getScriptProperties();
  const fileId = props.getProperty('cache_file_id');
  try {
    if (fileId) {
      DriveApp.getFileById(fileId).setContent(json);
      Logger.log('Updated existing cache file');
    } else {
      const f = DriveApp.createFile('lmcs_dashboard_cache.json', json, MimeType.PLAIN_TEXT);
      f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      props.setProperty('cache_file_id', f.getId());
      Logger.log('Created new cache file: ' + f.getId());
    }
    props.setProperty('cache_ts', String(Date.now()));
  } catch(e) {
    Logger.log('Cache write error: ' + e);
    throw e;
  }
}

// Run once to clear stale PropertiesService data
function clearOldCache() {
  PropertiesService.getScriptProperties().deleteAllProperties();
  Logger.log('All properties cleared');
}

// ── Build full dataset ───────────────────────────────────────────
function buildData() {
  const now       = new Date();
  const ayStartYr = now.getMonth() >= 3 ? now.getFullYear() : now.getFullYear() - 1;
  const ayStart   = new Date(ayStartYr, 3, 1);
  const year      = now.getFullYear();

  const cwUrl = 'https://docs.google.com/spreadsheets/d/' + CW_SHEET_ID +
    '/gviz/tq?tqx=out:json&tq=' +
    encodeURIComponent("SELECT * WHERE B CONTAINS '/" + year + "' LIMIT 100000");

  const cmJobs = CM_SCHOOLS.flatMap((school, idx) =>
    CM_TABS.map(tab => ({
      idx, tab,
      url: 'https://docs.google.com/spreadsheets/d/' + school.id +
           '/gviz/tq?tqx=out:json&sheet=' + encodeURIComponent(tab)
    }))
  );

  const allUrls   = [cwUrl, ...cmJobs.map(j => j.url)];
  const responses = UrlFetchApp.fetchAll(allUrls.map(url => ({ url, muteHttpExceptions: true })));

  const cwRecords = parseCW(parseGviz(responses[0].getContentText()), ayStart);

  const cmData = {};
  cmJobs.forEach((job, i) => {
    const resp = responses[i + 1];
    if (resp.getResponseCode() !== 200) return;
    if (!cmData[job.idx]) cmData[job.idx] = {};
    cmData[job.idx][job.tab] = parseCM(parseGviz(resp.getContentText()));
  });

  return { generated: now.toISOString(), ayStart: ayStartYr + '-04-01', cwRecords, cmData };
}

// ── Parse helpers ────────────────────────────────────────────────
function parseGviz(text) {
  try {
    const s = text.indexOf('{'), e = text.lastIndexOf('}');
    return JSON.parse(text.slice(s, e + 1)).table;
  } catch (_) { return null; }
}

const SCHOOL_KEYWORDS = ['kullu','kelheli','dunkhra','nerchowk','sayoli','jogindernagar'];
function schoolToIdx(name) {
  const n = (name || '').toLowerCase();
  for (let i = 0; i < SCHOOL_KEYWORDS.length; i++) if (n.includes(SCHOOL_KEYWORDS[i])) return i;
  if (n.includes('guru international')) return 5;
  return null;
}

const ROMAN = {I:1,II:2,III:3,IV:4,V:5,VI:6,VII:7,VIII:8,IX:9,X:10,XI:11,XII:12};
function mapClass(cls, section) {
  const c = (cls||'').trim().toUpperCase(), s = (section||'').trim().toUpperCase();
  if (c === 'M') {
    if (ROMAN[s]===1||s==='1') return 'M-I';
    if (ROMAN[s]===2||s==='2') return 'M-II';
    if (ROMAN[s]===3||s==='3') return 'M-III';
    return 'M-I';
  }
  if (ROMAN[c]) return 'Class ' + ROMAN[c];
  const n = parseInt(c);
  return (!isNaN(n) && n >= 1 && n <= 12) ? 'Class ' + n : null;
}

function parseCW(table, ayStart) {
  if (!table || !table.rows) return [];
  const records = [];
  for (const row of table.rows) {
    const cell = i => { const c = row.c?.[i]; return (!c||c.v==null)?'':String(c.f??c.v).trim(); };
    const lmsIdx = schoolToIdx(cell(0));
    if (lmsIdx === null) continue;
    const m = cell(1).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2}):\d{2}\s*(AM|PM)$/i);
    if (!m) continue;
    let [,mo,da,yr,hr,mi,ap] = m;
    mo=+mo; da=+da; yr=+yr; hr=+hr; mi=+mi;
    if (ap.toUpperCase()==='PM' && hr!==12) hr+=12;
    if (ap.toUpperCase()==='AM' && hr===12) hr=0;
    const dateObj = new Date(yr, mo-1, da, hr, mi);
    if (dateObj < ayStart) continue;
    const dateStr = yr+'-'+String(mo).padStart(2,'0')+'-'+String(da).padStart(2,'0');
    records.push({
      lmsIdx, dateStr, timeMs: dateObj.getTime(),
      teacher: cell(2), classMapped: mapClass(cell(3), cell(4)), subject: cell(5),
      cwBook: cell(6), cwChapter: cell(7), cwDesc: cell(8), cwTag: cell(9),
      hwChapter: cell(11), hwDesc: cell(12), hwTag: cell(13),
    });
  }
  records.sort((a,b) => b.timeMs - a.timeMs);
  return records;
}

function parseCM(table) {
  if (!table || !table.rows) return [];
  const cm = {SUBJECT:1,CHAPTER_NO:3,CHAPTER_NAME:4,TERM:5,START_DATE:7,CWA_DATE:10};
  if (table.cols) table.cols.forEach((col,i) => {
    const lbl = (col.label||col.id||'').toLowerCase().trim();
    if      (lbl==='subject')              cm.SUBJECT=i;
    else if (lbl==='chapter no.')          cm.CHAPTER_NO=i;
    else if (lbl==='chapter name')         cm.CHAPTER_NAME=i;
    else if (lbl==='term')                 cm.TERM=i;
    else if (lbl==='chapter start date')   cm.START_DATE=i;
    else if (lbl==='cwa date')             cm.CWA_DATE=i;
  });
  const chapters = [];
  for (const row of table.rows) {
    const raw = i => row.c?.[i]||null;
    const str = i => { const c=raw(i); return (!c||c.v==null)?'':String(c.f??c.v).trim(); };
    const subject = str(cm.SUBJECT);
    if (!subject || subject.toLowerCase()==='subject') continue;
    chapters.push({
      subject, chapterNo: raw(cm.CHAPTER_NO)?.v??null,
      chapterName: str(cm.CHAPTER_NAME), term: str(cm.TERM),
      startDate: gvizToDateStr(raw(cm.START_DATE)),
      cwaDate:   gvizToDateStr(raw(cm.CWA_DATE)),
    });
  }
  return chapters;
}

function gvizToDateStr(cell) {
  if (!cell||cell.v==null) return null;
  const m = String(cell.v).match(/Date\((\d+),(\d+),(\d+)/);
  if (!m) return null;
  return +m[1]+'-'+String(+m[2]+1).padStart(2,'0')+'-'+String(+m[3]).padStart(2,'0');
}
