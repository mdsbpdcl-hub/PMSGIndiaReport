/**
 * PMSG-MBY All States – automatic update (Google Apps Script, bound to the sheet)
 * ---------------------------------------------------------------------------
 * Tabs it maintains (created automatically):
 *   State-wise : one row per State/UT, always sorted by Installations (high → low),
 *                Rank, Free-electricity column, Total row, change since last update,
 *                Bihar highlighted, "Last updated" + source shown at the right.
 *   Month-wise : month progression table (when the office-PC script sends it).
 *   History    : one snapshot per state per day (for trend / Bihar tracking).
 *
 * Data sources (in order of preference):
 *   1. Office PC in India runs pmsg_portal_scraper.py → reads pmsuryaghar.gov.in
 *      → POSTs here (doPost). This is the OFFICIAL portal data.
 *   2. If no PC push in the last 20 hours, the daily trigger reads the public
 *      daily copy at solentra.in (unofficial mirror of the portal).
 *      (Apps Script cannot read pmsuryaghar.gov.in directly – it blocks non-Indian servers.)
 *
 * SETUP (once):
 *   1. In the sheet: Extensions → Apps Script → delete any code → paste this file → Save.
 *   2. Change CFG.TOKEN below to your own secret word (same word goes in the Python script).
 *   3. Run  installDailyTrigger  once (Run ▶) and allow permissions.
 *   4. Deploy → New deployment → type "Web app" → Execute as: Me,
 *      Who has access: Anyone → Deploy → copy the Web app URL into the Python script.
 *   5. Reload the sheet: a menu "PMSG Update" appears (Update now / Install daily trigger).
 */

const CFG = {
  TOKEN: 'CHANGE-ME-sbpdcl-2026',
  SHEET_STATE: 'State-wise',
  SHEET_MONTH: 'Month-wise',
  SHEET_HISTORY: 'History',
  MIRROR_URL: 'https://solentra.in/solar-dashboard.html',
  TZ: 'Asia/Kolkata',
  DAILY_HOUR: 10,          // 10 AM IST daily fallback run
  HIGHLIGHT_STATE: 'bihar'
};

// Free domestic electricity schemes (checked 18-Nov-2026) – edit here if a policy changes.
const FREE_POWER = [
  ['punjab', 'Yes - 300 units/month (domestic)'],
  ['delhi', 'Yes - 200 units/month; 50% subsidy 201-400 units'],
  ['jharkhand', 'Yes - 200 units/month (consumption up to 200 units)'],
  ['karnataka', 'Yes - up to 200 units/month (Gruha Jyothi)'],
  ['telangana', 'Yes - 200 units/month (Gruha Jyothi, eligible families)'],
  ['tamil nadu', 'Yes - 200 units per 2 months (use up to 500 units); others 100 units'],
  ['bihar', 'Yes - 125 units/month (all domestic, from Aug-2025)'],
  ['himachal', 'Yes - 125 units/month (non-taxpayers, max 2 meters)'],
  ['rajasthan', 'Linked to rooftop solar - 150 units/month after 1.1 kW RTS'],
  ['jammu', 'Linked to rooftop solar - 200 units/month for AAY families (rolling out)'],
  ['west bengal', 'Targeted - up to 75 units/month for poor (Hasir Alo); check current status']
];

const HEADERS = ['Rank', 'State/UT', 'Free Electricity (Domestic)', 'Applications', 'Installations',
  'Conversion %', 'Capacity (MW)', 'Avg System (kW)', 'Subsidy Released (Rs Cr)',
  'Installations added since last update'];

/* ============================ MENU & TRIGGERS ============================ */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('PMSG Update')
    .addItem('Update now (daily copy site)', 'updateFromMirror')
    .addItem('Install daily auto-update', 'installDailyTrigger')
    .addToUi();
}

function installDailyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'dailyJob')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailyJob').timeBased().everyDays(1)
    .atHour(CFG.DAILY_HOUR).inTimezone(CFG.TZ).create();
  toast_('Daily auto-update installed (around ' + CFG.DAILY_HOUR + ':00 IST).');
}

function dailyJob() {
  const last = Number(PropertiesService.getScriptProperties().getProperty('lastPortalPush') || 0);
  if (Date.now() - last < 20 * 3600 * 1000) return;      // office PC already sent official data
  updateFromMirror();
}

/* ============================ SOURCE 1: OFFICE PC ============================ */

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.token !== CFG.TOKEN) return json_({ ok: false, error: 'wrong token' });
    const out = { ok: true };
    if (body.state && body.state.length > 1) {
      out.state = writeState_(normaliseRows_(body.state), 'National Portal (pmsuryaghar.gov.in) – office PC');
      PropertiesService.getScriptProperties().setProperty('lastPortalPush', String(Date.now()));
    }
    if (body.month && body.month.length > 1) out.month = writeMonth_(body.month);
    return json_(out);
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doGet() { return json_({ ok: true, msg: 'PMSG sheet updater is running' }); }

/* ============================ SOURCE 2: DAILY COPY SITE ============================ */

function updateFromMirror() {
  const res = UrlFetchApp.fetch(CFG.MIRROR_URL, { muteHttpExceptions: true, followRedirects: true });
  if (res.getResponseCode() !== 200) throw new Error('Copy site returned HTTP ' + res.getResponseCode());
  const rows = htmlTableRows_(res.getContentText());
  if (rows.length < 5) throw new Error('No state table found on the copy site.');
  const n = writeState_(normaliseRows_(rows), 'solentra.in daily copy of National Portal (unofficial)');
  toast_('Updated ' + n + ' states.');
}

/* ============================ PARSING (pure functions) ============================ */

function htmlTableRows_(html) {
  // take the biggest <table> on the page
  const tables = html.match(/<table[\s\S]*?<\/table>/gi) || [];
  let best = '';
  tables.forEach(t => { if ((t.match(/<tr/gi) || []).length > (best.match(/<tr/gi) || []).length) best = t; });
  const trs = best.match(/<tr[\s\S]*?<\/tr>/gi) || [];
  return trs.map(tr => (tr.match(/<t[hd][^>]*>[\s\S]*?<\/t[hd]>/gi) || []).map(td => decode_(td.replace(/<[^>]+>/g, ' '))))
            .filter(r => r.length > 1);
}

function decode_(s) {
  return s.replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&#8377;|&#x20b9;|&rupee;/gi, '₹')
          .replace(/&#(\d+);/g, (m, d) => String.fromCharCode(d)).replace(/\s+/g, ' ').trim();
}

/** rows[0] = header (any wording), returns [{state, apps, inst, mw, cr}] */
function normaliseRows_(rows) {
  const h = rows[0].map(x => String(x).toLowerCase());
  const find = (re, not) => h.findIndex(x => re.test(x) && !(not && not.test(x)));
  const iState = find(/state|\but\b/);
  const iCap = find(/capacity/) >= 0 ? find(/capacity/) : find(/\bmw\b|\bgw\b/);
  const iInst = find(/install/, /capacity|\bmw\b|\bkw\b/);
  const iApps = find(/applica|registr/);
  const iCr = find(/subsidy|cfa|disburs|released/);
  if (iState < 0 || iInst < 0) throw new Error('Could not find State / Installations columns in: ' + rows[0].join(' | '));
  return rows.slice(1).map(r => ({
    state: String(r[iState] || '').trim(),
    apps: iApps >= 0 ? num_(r[iApps]) : null,
    inst: num_(r[iInst]),
    mw: iCap >= 0 ? toMW_(r[iCap], h[iCap]) : null,
    cr: iCr >= 0 ? toCr_(r[iCr], h[iCr]) : null
  })).filter(o => o.state && !/^(total|grand total|india|all india)/i.test(o.state) && o.inst !== null);
}

function num_(v) {
  if (v === null || v === undefined) return null;
  if (typeof v === 'number') return v;
  const s = String(v).replace(/[₹,\s]/g, '').replace(/(rs\.?|inr)/i, '');
  const m = s.match(/-?\d+(\.\d+)?/);
  return m ? parseFloat(m[0]) : null;
}

function toMW_(v, header) {
  const n = num_(v); if (n === null) return null;
  const s = String(v).toLowerCase();
  if (/gw/.test(s)) return n * 1000;
  if (/kw/.test(s) && !/mw/.test(s)) return n / 1000;
  if (/mw/.test(s)) return n;
  if (/\bkw|kwp/.test(header)) return n / 1000;
  if (/\bgw/.test(header)) return n * 1000;
  return n;                                   // assume MW
}

function toCr_(v, header) {
  const n = num_(v); if (n === null) return null;
  const s = String(v).toLowerCase();
  if (/cr/.test(s)) return n;
  if (/\bl\b|lakh|lac/.test(s)) return n / 100;
  if (/cr/.test(header)) return n;
  if (/lakh|lac/.test(header)) return n / 100;
  return (/₹/.test(s) || n >= 1e5) ? n / 1e7 : n;   // plain rupees → crore
}

function freePower_(state) {
  const s = state.toLowerCase();
  const hit = FREE_POWER.find(([k]) => s.indexOf(k) >= 0);
  return hit ? hit[1] : 'No';
}

/* ============================ WRITING ============================ */

function writeState_(list, source) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.SHEET_STATE) || ss.insertSheet(CFG.SHEET_STATE, 0);
  const prev = previousSnapshot_();                       // state -> installations (last earlier day)

  list.sort((a, b) => (b.inst || 0) - (a.inst || 0));
  const body = list.map((o, i) => {
    const conv = o.apps ? +(o.inst / o.apps * 100).toFixed(1) : '';
    const avg = (o.mw && o.inst) ? +(o.mw * 1000 / o.inst).toFixed(2) : '';
    const p = prev[o.state.toLowerCase()];
    return [i + 1, o.state, freePower_(o.state), o.apps ?? '', o.inst, conv,
            o.mw ?? '', avg, o.cr ?? '', p !== undefined ? o.inst - p : ''];
  });
  const sum = k => list.reduce((s, o) => s + (o[k] || 0), 0);
  const tApps = sum('apps'), tInst = sum('inst'), tMw = sum('mw'), tCr = sum('cr');
  const tAdded = body.reduce((s, r) => s + (typeof r[9] === 'number' ? r[9] : 0), 0);
  const total = ['', 'INDIA (Total)', '', tApps || '', tInst, tApps ? +(tInst / tApps * 100).toFixed(1) : '',
                 tMw || '', tMw ? +(tMw * 1000 / tInst).toFixed(2) : '', tCr || '', tAdded || ''];

  sh.clear();
  const all = [HEADERS].concat(body, [total]);
  sh.getRange(1, 1, all.length, HEADERS.length).setValues(all);

  // formatting
  const n = body.length;
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#e8f1ec').setWrap(true).setVerticalAlignment('middle');
  sh.getRange(2, 4, n + 1, 2).setNumberFormat('#,##0');
  sh.getRange(2, 6, n + 1, 1).setNumberFormat('0.0');
  sh.getRange(2, 7, n + 1, 1).setNumberFormat('#,##0.0');
  sh.getRange(2, 8, n + 1, 1).setNumberFormat('0.00');
  sh.getRange(2, 9, n + 1, 1).setNumberFormat('#,##0.00');
  sh.getRange(2, 10, n + 1, 1).setNumberFormat('+#,##0;-#,##0;0');
  sh.getRange(n + 2, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#f3f4f6');
  const hi = body.findIndex(r => r[1].toLowerCase().indexOf(CFG.HIGHLIGHT_STATE) >= 0);
  if (hi >= 0) sh.getRange(hi + 2, 1, 1, HEADERS.length).setBackground('#fff4cc').setFontWeight('bold');
  sh.getRange(1, 12, 3, 2).setValues([
    ['Last updated', Utilities.formatDate(new Date(), CFG.TZ, 'dd-MMM-yyyy HH:mm') + ' IST'],
    ['Source', source],
    ['Sorted by', 'Installations (high to low)']
  ]);
  sh.getRange(1, 12, 3, 1).setFontWeight('bold');
  sh.setColumnWidth(2, 230); sh.setColumnWidth(3, 330); sh.setColumnWidth(13, 330);
  sh.setColumnWidths(4, 7, 110);

  saveHistory_(list, source);
  return n;
}

function writeMonth_(rows) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.SHEET_MONTH) || ss.insertSheet(CFG.SHEET_MONTH);
  const w = Math.max.apply(null, rows.map(r => r.length));
  const data = rows.map((r, i) => {
    const row = r.concat(Array(w - r.length).fill(''));
    return i === 0 ? row : row.map((c, j) => { const x = num_(c); return (j > 0 && x !== null && /^[₹\d,.\s]+$/.test(String(c))) ? x : c; });
  });
  sh.clear();
  sh.getRange(1, 1, data.length, w).setValues(data);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, w).setFontWeight('bold').setBackground('#e8f1ec');
  sh.getRange(1, w + 2, 1, 2).setValues([['Last updated', Utilities.formatDate(new Date(), CFG.TZ, 'dd-MMM-yyyy HH:mm') + ' IST']]);
  return data.length - 1;
}

function historySheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CFG.SHEET_HISTORY);
  if (!sh) {
    sh = ss.insertSheet(CFG.SHEET_HISTORY);
    sh.appendRow(['Date', 'State/UT', 'Applications', 'Installations', 'Capacity (MW)', 'Subsidy (Rs Cr)', 'Source']);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#e8f1ec');
  }
  return sh;
}

function today_() { return Utilities.formatDate(new Date(), CFG.TZ, 'yyyy-MM-dd'); }

/** installations by state from the most recent day BEFORE today */
function previousSnapshot_() {
  const sh = historySheet_();
  const last = sh.getLastRow(); if (last < 2) return {};
  const v = sh.getRange(2, 1, last - 1, 4).getDisplayValues();
  const t = today_();
  const days = v.map(r => r[0]).filter(d => d && d < t).sort();
  if (!days.length) return {};
  const d = days[days.length - 1], m = {};
  v.forEach(r => { if (r[0] === d) m[r[1].toLowerCase()] = num_(r[3]); });
  return m;
}

/** replace today's snapshot (so re-runs on the same day don't duplicate) */
function saveHistory_(list, source) {
  const sh = historySheet_();
  const t = today_();
  const last = sh.getLastRow();
  if (last >= 2) {
    const dates = sh.getRange(2, 1, last - 1, 1).getDisplayValues().map(r => r[0]);
    for (let i = dates.length - 1; i >= 0; i--) if (dates[i] === t) sh.deleteRow(i + 2);
  }
  const rows = list.map(o => [t, o.state, o.apps ?? '', o.inst, o.mw ?? '', o.cr ?? '', source]);
  const start = sh.getLastRow() + 1;
  sh.getRange(start, 1, rows.length, 1).setNumberFormat('@');   // keep date as text yyyy-MM-dd
  sh.getRange(start, 1, rows.length, 7).setValues(rows);
}

/* ============================ helpers ============================ */

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function toast_(m) { try { SpreadsheetApp.getActive().toast(m, 'PMSG Update', 6); } catch (e) { Logger.log(m); } }
