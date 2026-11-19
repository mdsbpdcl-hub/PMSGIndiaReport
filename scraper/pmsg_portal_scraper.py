"""
PM Surya Ghar portal scraper  (State/UT-wise + Month-wise)
-----------------------------------------------------------
Opens, in a real browser:
    https://pmsuryaghar.gov.in/#/state-ut-wise-progress
    https://pmsuryaghar.gov.in/#/month-progression
reads the table (all pages) or, if only a chart is shown, the chart data,
and saves one Excel (two sheets) + two CSVs with date-time stamp.

RUN ON A PC IN INDIA (the portal returns 403 to foreign / cloud IPs,
so Google Apps Script or cloud servers will NOT work).

One-time setup (Command Prompt):
    pip install playwright pandas openpyxl
    playwright install chromium

Run:
    python pmsg_state_scraper.py              # opens a visible browser window
    python pmsg_state_scraper.py --headless   # runs hidden (try this once it works)

Google Sheet: set SHEET_WEBAPP_URL and SHEET_TOKEN below; every run then updates
the "PMSG-MBY All States" sheet (State-wise sorted by installations, Month-wise, History).

Daily auto-run: Windows Task Scheduler -> Create Basic Task -> Daily ->
    Program: python   Arguments: "C:\\path\\pmsg_state_scraper.py" --headless
Please keep it to once or twice a day; it is a Government portal.
"""
import os
import sys
import re
import json
import urllib.request
from datetime import datetime
from pathlib import Path

import pandas as pd
from playwright.sync_api import sync_playwright

PAGES = {
    "State-wise": "https://pmsuryaghar.gov.in/#/state-ut-wise-progress",
    "Month-wise": "https://pmsuryaghar.gov.in/#/month-progression",
}
OUT_DIR = Path(__file__).resolve().parent / "PMSG_Data"

# ---- Google Sheet auto-update (PMSG-MBY All States) ----
# Paste the Web app URL from Apps Script (Deploy -> Web app) and the same TOKEN as in CFG.TOKEN.
# Read from environment (GitHub secrets) so nothing secret is stored in the public repository.
# For a manual run on a PC you may also type them here.
SHEET_WEBAPP_URL = os.environ.get("SHEET_WEBAPP_URL", "")   # e.g. "https://script.google.com/macros/s/AKfy.../exec"
SHEET_TOKEN = os.environ.get("SHEET_TOKEN", "")

# Same extraction logic as the console script, returns ALL pages as a list of rows.
COLLECT_JS = r"""
async () => {
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const clean = s => (s || '').replace(/\s+/g, ' ').trim();

  // ---- 1. Tables (HTML <table> or Angular-Material / ARIA grids) ----
  function grabRows() {
    let best = null;
    document.querySelectorAll('table').forEach(t => {
      if (t.rows.length >= 2 && (!best || t.rows.length > best.rows.length)) best = t;
    });
    if (best) return [...best.rows].map(r => [...r.cells].map(c => clean(c.innerText)));
    const rowSel = '[role="row"], mat-row, mat-header-row, .mat-row, .mat-header-row, .mat-mdc-row, .mat-mdc-header-row';
    const cellSel = '[role="cell"], [role="columnheader"], [role="gridcell"], mat-cell, mat-header-cell, .mat-mdc-cell, .mat-mdc-header-cell';
    return [...document.querySelectorAll(rowSel)]
      .map(r => [...r.querySelectorAll(cellSel)].map(c => clean(c.innerText)))
      .filter(r => r.length > 1);
  }

  // ---- 2. Charts (Highcharts / Chart.js / ECharts / ApexCharts) ----
  function chartRows() {
    const tables = [];
    const pivot = (labels, series) => {
      if (!labels || !labels.length || !series.length) return;
      tables.push([['Period', ...series.map(s => s.name)],
        ...labels.map((l, i) => [clean(String(l)), ...series.map(s => s.data[i] ?? '')])]);
    };
    const val = p => (p && typeof p === 'object') ? (p.y ?? p.value ?? p[1] ?? '') : p;
    try {
      (window.Highcharts?.charts || []).filter(Boolean).forEach(c => {
        const cats = c.xAxis?.[0]?.categories || c.series?.[0]?.points?.map(p => p.category || p.name);
        pivot(cats, c.series.map(s => ({ name: s.name, data: (s.points || s.data).map(p => p.y) })));
      });
    } catch (e) {}
    try {
      const inst = window.Chart?.instances;
      Object.values(inst || {}).forEach(c => pivot(c.data.labels,
        c.data.datasets.map(d => ({ name: d.label, data: d.data.map(val) }))));
    } catch (e) {}
    try {
      if (window.echarts) document.querySelectorAll('[_echarts_instance_]').forEach(el => {
        const o = window.echarts.getInstanceByDom(el)?.getOption(); if (!o) return;
        const ax = [].concat(o.xAxis || [])[0]?.data || [].concat(o.yAxis || [])[0]?.data;
        pivot(ax, (o.series || []).map(s => ({ name: s.name, data: (s.data || []).map(val) })));
      });
    } catch (e) {}
    try {
      (window.Apex?._chartInstances || []).forEach(ci => {
        const w = ci.chart?.w; if (!w) return;
        pivot(w.globals.categoryLabels?.length ? w.globals.categoryLabels : w.globals.labels,
          w.config.series.map(s => ({ name: s.name, data: (s.data || []).map(val) })));
      });
    } catch (e) {}
    return tables.sort((a, b) => b.length - a.length)[0] || [];
  }

  function nextButton() {
    const b = document.querySelector(
      'button[aria-label*="Next" i], .mat-paginator-navigation-next, .mat-mdc-paginator-navigation-next, ' +
      'li.next:not(.disabled) a, a[aria-label*="Next" i], .pagination-next:not(.disabled) a');
    if (!b) return null;
    const dis = b.disabled || b.getAttribute('aria-disabled') === 'true' ||
                b.classList.contains('disabled') || b.closest('.disabled');
    return dis ? null : b;
  }

  // Wait up to 45 s for a table or chart to appear
  let rows = [], source = 'table';
  for (let i = 0; i < 90; i++) {
    rows = grabRows();
    if (rows.length >= 3) break;
    const ch = chartRows();
    if (ch.length >= 3) { rows = ch; source = 'chart'; break; }
    await sleep(500);
  }
  if (rows.length < 3) return { source: 'none', rows: [] };
  if (source === 'chart') return { source, rows };

  const header = rows[0], seen = new Set(), out = [header];
  const add = rs => rs.slice(1).forEach(r => {
    const k = r.join('|');
    if (!seen.has(k) && k !== header.join('|')) { seen.add(k); out.push(r); }
  });
  add(rows);
  for (let g = 0; g < 50; g++) {
    const nb = nextButton(); if (!nb) break;
    const before = JSON.stringify(grabRows());
    nb.click();
    for (let i = 0; i < 20 && JSON.stringify(grabRows()) === before; i++) await sleep(300);
    add(grabRows());
  }
  return { source, rows: out };
}
"""


# Free domestic electricity schemes (as checked on 18-Nov-2026; update if a state changes its policy)
FREE_POWER = {
    "punjab": "Yes - 300 units/month (domestic)",
    "delhi": "Yes - 200 units/month; 50% subsidy 201-400 units",
    "jharkhand": "Yes - 200 units/month (consumption up to 200 units)",
    "karnataka": "Yes - up to 200 units/month (Gruha Jyothi)",
    "telangana": "Yes - 200 units/month (Gruha Jyothi, eligible families)",
    "tamil nadu": "Yes - 200 units per 2 months (use up to 500 units); others 100 units",
    "bihar": "Yes - 125 units/month (all domestic, from Aug-2025)",
    "himachal": "Yes - 125 units/month (non-taxpayers, max 2 meters)",
    "rajasthan": "Linked to rooftop solar - 150 units/month after 1.1 kW RTS",
    "jammu": "Linked to rooftop solar - 200 units/month for AAY families (rolling out)",
    "west bengal": "Targeted - up to 75 units/month for poor (Hasir Alo); check current status",
}


def add_free_power_column(df):
    state_col = next((c for c in df.columns if "state" in str(c).lower()), None)
    if state_col is None:
        return df
    def lookup(name):
        n = str(name).lower()
        if not n or n.startswith(("total", "india", "grand")):
            return ""
        return next((v for k, v in FREE_POWER.items() if k in n), "No")
    df.insert(list(df.columns).index(state_col) + 1, "Free Electricity (Domestic)", df[state_col].map(lookup))
    return df


def to_number(v):
    """'9,70,710' -> 970710 ; '3.25 GW' stays text ; '' -> None"""
    if v is None:
        return None
    s = str(v).replace(",", "").replace("₹", "").strip()
    return float(s) if re.fullmatch(r"-?\d+(\.\d+)?", s) else v


def to_frame(rows):
    header, data = rows[0], rows[1:]
    width = max(len(header), *(len(r) for r in data))
    header = [h or f"Col{i+1}" for i, h in enumerate(header)] + [f"Col{i+1}" for i in range(len(header), width)]
    data = [list(r) + [""] * (width - len(r)) for r in data]
    df = pd.DataFrame(data, columns=header)
    # DataFrame.map needs pandas 2.1+; older pandas has applymap
    return df.map(to_number) if hasattr(df, 'map') else df.applymap(to_number)


def push_to_sheet(raw):
    """Send raw table rows to the Google Sheet web app (it sorts, ranks and formats)."""
    if not SHEET_WEBAPP_URL:
        print("Google Sheet push skipped (SHEET_WEBAPP_URL not set).")
        return
    payload = {"token": SHEET_TOKEN,
               "state": raw.get("State-wise", []),
               "month": raw.get("Month-wise", [])}
    req = urllib.request.Request(SHEET_WEBAPP_URL, data=json.dumps(payload).encode("utf-8"),
                                 headers={"Content-Type": "application/json"}, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=120) as r:
            print("Google Sheet:", r.read().decode("utf-8", "ignore")[:300])
    except Exception as e:
        print("Google Sheet push FAILED:", e)


def main():
    headless = "--headless" in sys.argv
    frames, raw = {}, {}
    with sync_playwright() as p:
        try:  # use installed Google Chrome if available (looks like a normal user)
            browser = p.chromium.launch(channel="chrome", headless=headless)
        except Exception:
            browser = p.chromium.launch(headless=headless)
        page = browser.new_page(locale="en-IN", viewport={"width": 1400, "height": 900})
        for name, url in PAGES.items():
            resp = page.goto(url, wait_until="networkidle", timeout=90_000)
            if resp and resp.status == 403:
                sys.exit("403 Forbidden - run this from a PC on an Indian internet connection.")
            page.reload(wait_until="networkidle")  # hash routes: make sure the new view rendered
            res = page.evaluate(COLLECT_JS)
            if res["rows"]:
                raw[name] = [[str(c) for c in r] for r in res["rows"]]
                frames[name] = to_frame(res["rows"])
                if name == "State-wise":
                    frames[name] = add_free_power_column(frames[name])
                print(f"{name}: {len(frames[name])} rows (from {res['source']})")
            else:
                print(f"{name}: no table/chart data found (try without --headless)")
        browser.close()

    if not frames:
        sys.exit("Nothing captured.")
    OUT_DIR.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M")
    xlsx_path = OUT_DIR / f"PMSG_Portal_{stamp}.xlsx"
    with pd.ExcelWriter(xlsx_path) as xw:
        for name, df in frames.items():
            df.to_excel(xw, index=False, sheet_name=name)
            df.to_csv(OUT_DIR / f"PMSG_{name}_{stamp}.csv", index=False, encoding="utf-8-sig")
    print(f"Saved: {xlsx_path}")
    push_to_sheet(raw)
    for name, df in frames.items():
        print(f"\n--- {name} ---\n" + df.head(8).to_string(index=False))


if __name__ == "__main__":
    main()
