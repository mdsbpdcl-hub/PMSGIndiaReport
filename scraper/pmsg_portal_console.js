/*
  PM Surya Ghar portal  ->  CSV download   (works on BOTH pages)
    https://pmsuryaghar.gov.in/#/state-ut-wise-progress
    https://pmsuryaghar.gov.in/#/month-progression
  ------------------------------------------------------------------
  HOW TO USE (Chrome / Edge, on any PC in India):
   1. Open one of the pages above and wait till the table / chart loads.
   2. Press F12 -> "Console" tab.
   3. Paste this whole script, press Enter.
      (First time Chrome may ask you to type  allow pasting  -> type it, Enter, paste again.)
   4. A CSV downloads (PMSG_state-ut-wise-progress_YYYY-MM-DD.csv or
      PMSG_month-progression_YYYY-MM-DD.csv). Opens directly in Excel.
   Reads normal tables (all pages of a paged table) and, if the page
   only has a chart, the chart's underlying numbers.
*/
(async () => {
  const collect = async () => {
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
};
  const res = await collect();
  if (!res.rows.length) { console.error('No table or chart data found - let the page finish loading and run again.'); return; }
  const esc = v => { v = String(v ?? ''); return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; };
  const csv = '\ufeff' + res.rows.map(r => r.map(esc).join(',')).join('\r\n');
  const page = (location.hash.replace(/^#\/?/, '') || 'page').replace(/[^a-z0-9-]/gi, '_');
  const d = new Date();
  const stamp = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
  a.download = `PMSG_${page}_${stamp}.csv`;
  document.body.appendChild(a); a.click(); a.remove();
  console.table(res.rows.slice(0, 8));
  console.log(`Done (${res.source}): ${res.rows.length - 1} rows saved to ${a.download}`);
})();
