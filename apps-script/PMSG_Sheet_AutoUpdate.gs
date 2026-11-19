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
 * Data source (fully online, no PC needed):
 *   Twice a day the script reads the public daily copy of the National Portal
 *   (solentra.in – unofficial). Apps Script cannot read pmsuryaghar.gov.in itself
 *   because the portal blocks servers outside India.
 *   If the copy site fails or changes, the last good figures stay in the sheet.
 *   (Optional: an office PC can push official portal data via doPost – see README.)
 *
 * SETUP (once):
 *   1. In the sheet: Extensions → Apps Script → delete any code → paste this file → Save.
 *   2. Choose  installDailyTrigger  in the function list → Run ▶ → allow permissions.
 *      This fills the State-wise tab immediately and sets the 9 AM / 6 PM updates.
 *   3. Reload the sheet: a menu "PMSG Update" appears (Update now / Install daily auto-update).
 *
 * DASHBOARD "FORCE UPDATE" BUTTON (once):
 *   4. Set CFG.TOKEN below to your password (keep it only here – not in GitHub).
 *   5. Deploy → New deployment → Select type: Web app → Execute as: Me,
 *      Who has access: Anyone → Deploy → open the Web app link once.
 *      That saves the link; the dashboard button then works.
 */

const CFG = {
  TOKEN: 'CHANGE-ME',      // ← your password for the dashboard "Force update" button
  SHEET_STATE: 'State-wise',
  SHEET_MONTH: 'Month-wise',
  SHEET_HISTORY: 'History',
  MIRROR_URL: 'https://solentra.in/data/pmsg-snapshot.json',   // daily copy of the portal (data file)
  TZ: 'Asia/Kolkata',
  RUN_HOURS: [9, 18],      // checks at 9 AM and 6 PM IST every day
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
  CFG.RUN_HOURS.forEach(h => ScriptApp.newTrigger('dailyJob').timeBased().everyDays(1)
    .atHour(h).inTimezone(CFG.TZ).create());
  toast_('Auto-update installed: every day around ' + CFG.RUN_HOURS.map(h => h + ':00').join(' and ') + ' IST.');
  updateFromMirror();                                     // fill the State-wise tab straight away
}

function dailyJob() {
  const last = Number(PropertiesService.getScriptProperties().getProperty('lastPortalPush') || 0);
  if (Date.now() - last < 8 * 3600 * 1000) return;       // office PC sent official data recently
  try { updateFromMirror(); }
  catch (e) { Logger.log('Update skipped, last good figures kept: ' + e); }  // keep last good data
}

/* ============================ WEB REQUESTS (dashboard + office PC) ============================ */
/* The dashboard "Force update" button and the Python program call this.
 * Every request must carry the password (CFG.TOKEN). 5 wrong tries lock it for 10 minutes. */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const cache = CacheService.getScriptCache();
    const bad = Number(cache.get('badTries') || 0);
    if (bad >= 5) return json_({ ok: false, error: 'locked' });
    if (body.token !== CFG.TOKEN) { cache.put('badTries', String(bad + 1), 600); return json_({ ok: false, error: 'password' }); }
    rememberUrl_();

    if (body.action === 'check') return json_({ ok: true });

    if (body.action === 'force') {                         // dashboard: update now from the daily copy
      if (cache.get('recentForce')) return json_({ ok: false, error: 'wait' });
      if (!lock.tryLock(30000)) return json_({ ok: false, error: 'busy' });
      const n = updateFromMirror();
      cache.put('recentForce', '1', 120);
      return json_({ ok: true, states: n, updated: Utilities.formatDate(new Date(), CFG.TZ, 'dd MMM yyyy, HH:mm') });
    }

    // office PC / Python program: official portal rows
    if (!lock.tryLock(30000)) return json_({ ok: false, error: 'busy' });
    const out = { ok: true };
    if (body.state && body.state.length > 1) {
      out.state = writeState_(normaliseRows_(body.state), 'National Portal (pmsuryaghar.gov.in) – official');
      PropertiesService.getScriptProperties().setProperty('lastPortalPush', String(Date.now()));
    }
    if (body.month && body.month.length > 1) out.month = writeMonth_(body.month);
    return json_(out);
  } catch (err) {
    Logger.log(err);
    return json_({ ok: false, error: 'failed' });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/* Open the web-app link once after deploying: it saves the link so the dashboard can find it. */
function doGet() {
  const u = rememberUrl_();
  return HtmlService.createHtmlOutput('<p style="font:16px sans-serif">PMSG update service is running.' +
    (u ? ' Update link saved – the dashboard button is now connected.' : '') + '</p>');
}

function rememberUrl_() {
  try {
    const u = ScriptApp.getService().getUrl();
    if (u && /\/exec$/.test(u)) {
      PropertiesService.getScriptProperties().setProperty('webAppUrl', u);
      const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_STATE);
      if (sh) sh.getRange(4, 12, 1, 2).setValues([['Update link', u]]);
      return u;
    }
  } catch (e) {}
  return '';
}

/* ============================ SOURCE 2: DAILY COPY SITE ============================ */

function updateFromMirror() {
  const res = UrlFetchApp.fetch(CFG.MIRROR_URL + '?t=' + Date.now(), { muteHttpExceptions: true, followRedirects: true });
  if (res.getResponseCode() !== 200) throw new Error('Copy site returned HTTP ' + res.getResponseCode());
  const snap = JSON.parse(res.getContentText());
  const list = snapshotToList_(snap);
  if (list.length < 30) throw new Error('Copy site data looks incomplete (' + list.length + ' states).');
  const when = snap.capturedAt ? new Date(snap.capturedAt) : null;
  const n = writeState_(list, 'National Portal (pmsuryaghar.gov.in) via daily copy at solentra.in (unofficial)', when);
  toast_('Updated ' + n + ' states.');
  return n;
}

/** snapshot JSON → [{state, apps, inst, mw, cr}]; skips the national row (stateId "NT") */
function snapshotToList_(snap) {
  return (snap.states || []).filter(s => s.stateId !== 'NT' && s.stateName).map(s => ({
    state: titleCase_(s.stateName),
    apps: Number(s.applications) || 0,
    inst: Number(s.installations) || 0,
    mw: Math.round((Number(s.capacityKW) || 0) / 10) / 100,       // kW → MW, 2 decimals
    cr: Math.round((Number(s.subsidyAmount) || 0) / 1e5) / 100     // ₹ → crore, 2 decimals
  }));
}

function titleCase_(name) {
  return String(name).trim().toLowerCase()
    .replace(/\b([a-z])/g, (m, c) => c.toUpperCase())
    .replace(/\b(And|Of)\b/g, w => w.toLowerCase())
    .replace(/^Nct /, 'NCT ');
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

function writeState_(list, source, dataTime) {
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
    ['Last updated', Utilities.formatDate(dataTime || new Date(), CFG.TZ, 'dd-MMM-yyyy HH:mm') + ' IST'],
    ['Source', source],
    ['Sorted by', 'Installations (high to low)']
  ]);
  const link = PropertiesService.getScriptProperties().getProperty('webAppUrl');
  if (link) sh.getRange(4, 12, 1, 2).setValues([['Update link', link]]);
  sh.getRange(1, 12, 4, 1).setFontWeight('bold');
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
