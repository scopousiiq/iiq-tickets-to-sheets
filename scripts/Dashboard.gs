/**
 * Dashboard — Native web-app dashboard for ticket + SLA analytics.
 *
 * Deployment model: published as an Apps Script Web App via doGet().
 * The script is container-bound, so SpreadsheetApp.getActiveSpreadsheet()
 * resolves to the district's sheet when deployed with "Execute as: Me".
 *
 * Data flow:
 *   viewer browser → doGet(e) → HtmlOutput (Dashboard.html)
 *                      ↓  google.script.run.getDashboardData()
 *   getDashboardData() reads TicketData + registered analytics sheets
 *                      ↓  returns { kpis, badges, categoryGroups }
 *   client renders header → KPI row → badge row → tab bar → active tab
 *
 * See scripts/ChartRegistry.gs for the declarative sheet→chart map.
 * See CLAUDE.md "Deploying the Dashboard (Web App)" for the playbook.
 */

/** Column indexes in TicketData used by the fixed KPI row. */
const DASH_COL = {
  CREATED: 4,            // E
  CLOSED_DATE: 7,        // H
  IS_CLOSED: 8,          // I  "Open" / "Closed"
  AGE_DAYS: 17,          // R
  RESOLUTION_BREACH: 34  // AI
};

/**
 * Web-app entry point. Called when a viewer loads the /exec URL.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('DashboardApp')
    .setTitle('iiQ Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Menu handler: shows the deployed dashboard URL (or deployment instructions
 * if not set yet) in a small HtmlService modal inside the sheet.
 */
function showDashboardUrl() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  let url = '';
  if (configSheet) {
    const values = configSheet.getDataRange().getValues();
    for (let i = 0; i < values.length; i++) {
      if (values[i][0] === 'DASHBOARD_URL') {
        url = String(values[i][1] || '').trim();
        break;
      }
    }
  }

  const html = url
    ? buildUrlModalHtml_(url)
    : buildDeploymentInstructionsHtml_();

  const output = HtmlService.createHtmlOutput(html)
    .setWidth(500)
    .setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(output, 'iiQ Dashboard');
}

function buildUrlModalHtml_(url) {
  const safeUrl = String(url).replace(/"/g, '&quot;');
  return '<style>' +
    'body{font-family:"Proxima Nova","Open Sans",system-ui,sans-serif;padding:18px;color:#374151;}' +
    'h2{color:#365c96;font-size:16px;margin:0 0 8px 0;border-bottom:2px solid #febb12;padding-bottom:6px;}' +
    'p{font-size:13px;margin:10px 0;}' +
    'input{width:100%;padding:8px;font-family:monospace;font-size:12px;border:1px solid #E5E7EB;border-radius:4px;}' +
    '.row{display:flex;gap:8px;margin-top:10px;}' +
    'button,a.btn{display:inline-block;padding:8px 14px;font-size:13px;font-weight:600;border-radius:6px;border:none;cursor:pointer;text-decoration:none;}' +
    '.copy{background:#365c96;color:#fff;}' +
    '.open{background:#22b2a3;color:#fff;}' +
    '.hint{color:#6B7280;font-size:11px;}' +
    '</style>' +
    '<h2>Dashboard URL</h2>' +
    '<p>Share this link with anyone in your domain:</p>' +
    '<input id="u" readonly value="' + safeUrl + '" onclick="this.select();">' +
    '<div class="row">' +
      '<button class="copy" onclick="document.getElementById(\'u\').select();document.execCommand(\'copy\');this.textContent=\'Copied!\';">Copy</button>' +
      '<a class="btn open" href="' + safeUrl + '" target="_blank">Open</a>' +
    '</div>' +
    '<p class="hint">To update the dashboard, redeploy via <strong>Deploy → Manage deployments → Edit → New version</strong> (keeps this URL stable).</p>';
}

function buildDeploymentInstructionsHtml_() {
  return '<style>' +
    'body{font-family:"Proxima Nova","Open Sans",system-ui,sans-serif;padding:18px;color:#374151;}' +
    'h2{color:#365c96;font-size:16px;margin:0 0 8px 0;border-bottom:2px solid #febb12;padding-bottom:6px;}' +
    'p{font-size:13px;margin:8px 0;}' +
    'ol{font-size:12px;line-height:1.6;padding-left:20px;}' +
    'li{margin-bottom:4px;}' +
    'code{background:#F9FAFB;border:1px solid #E5E7EB;padding:1px 4px;border-radius:3px;font-size:11px;}' +
    '</style>' +
    '<h2>Dashboard Not Deployed Yet</h2>' +
    '<p>Publish the dashboard as a web app, then paste the URL into the <code>DASHBOARD_URL</code> row in the Config sheet.</p>' +
    '<ol>' +
      '<li>Open <strong>Extensions → Apps Script</strong>.</li>' +
      '<li>Click <strong>Deploy → New deployment</strong>. Type: <strong>Web app</strong>.</li>' +
      '<li>Execute as: <strong>Me</strong> (the deployer).</li>' +
      '<li>Who has access: <strong>Anyone within your domain</strong>.</li>' +
      '<li>Click <strong>Deploy</strong>, authorize, and copy the <code>/exec</code> URL.</li>' +
      '<li>Paste the URL into the <strong>DASHBOARD_URL</strong> row in the Config sheet.</li>' +
      '<li>Re-open this menu to see the URL and a Copy button.</li>' +
      '<li>For future updates: <strong>Deploy → Manage deployments → Edit → New version</strong> (keeps the same URL).</li>' +
    '</ol>';
}

/**
 * Primary data provider. Called from Dashboard.html via google.script.run.
 *
 * Returns:
 *   { schoolYear, generatedAt, totalTickets,
 *     kpis:   { open, closed, avgResolutionDays, slaBreachPct },
 *     badges: [ { sheetName, label, color, count } ],
 *     categoryGroups: [
 *       { category, tabLabel, charts: [
 *         { sheetName, title, type, labels, datasets }
 *       ] }
 *     ] }
 *   or { error: string } on fatal failure.
 */
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ticketSheet = ss.getSheetByName('TicketData');
  if (!ticketSheet) {
    return { error: 'TicketData sheet not found. Run Setup first.' };
  }

  const lastRow = ticketSheet.getLastRow();
  const rows = lastRow >= 2
    ? ticketSheet.getRange(2, 1, lastRow - 1, TICKET_COLUMN_COUNT).getValues()
    : [];

  const kpis = computeKpis_(rows);

  const existingSheets = {};
  ss.getSheets().forEach(s => { existingSheets[s.getName()] = s; });

  const badges = [];
  const chartsByCategory = {};

  CHART_REGISTRY.forEach(entry => {
    const sheet = existingSheets[entry.sheetName];
    if (!sheet) return;

    if (entry.kpiBadge) {
      const count = countBadgeRows_(sheet, entry.kpiBadge.rowStart);
      badges.push({
        sheetName: entry.sheetName,
        label: entry.kpiBadge.label,
        color: entry.kpiBadge.color,
        count: count
      });
      return;
    }

    if (!entry.charts || entry.charts.length === 0) return;

    const values = sheet.getDataRange().getValues();
    entry.charts.forEach(spec => {
      const built = buildChartPayload_(values, spec);
      if (!built) return;
      if (!chartsByCategory[entry.category]) {
        chartsByCategory[entry.category] = {
          category: entry.category,
          tabLabel: entry.tabLabel || entry.category,
          charts: []
        };
      }
      chartsByCategory[entry.category].charts.push(Object.assign({
        sheetName: entry.sheetName,
        title: spec.title,
        type: spec.type
      }, built));
    });
  });

  const categoryGroups = [];
  CATEGORY_ORDER.forEach(cat => {
    if (chartsByCategory[cat]) categoryGroups.push(chartsByCategory[cat]);
  });

  const config = getConfig();

  return {
    schoolYear: config.schoolYear || '',
    generatedAt: new Date().toISOString(),
    totalTickets: rows.length,
    kpis: kpis,
    badges: badges,
    categoryGroups: categoryGroups
  };
}

/** Fixed KPI row — computed directly from TicketData rows. */
function computeKpis_(rows) {
  let open = 0;
  let closed = 0;
  let resSum = 0;
  let resCount = 0;
  let breachCount = 0;

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (r[DASH_COL.IS_CLOSED] === 'Closed') {
      closed++;
      const c = r[DASH_COL.CREATED];
      const cd = r[DASH_COL.CLOSED_DATE];
      if (c instanceof Date && cd instanceof Date) {
        const d = cd.getTime() - c.getTime();
        if (d >= 0) { resSum += d / 86400000; resCount++; }
      }
      const br = r[DASH_COL.RESOLUTION_BREACH];
      if (br === 1 || br === '1') breachCount++;
    } else {
      open++;
    }
  }

  return {
    open: open,
    closed: closed,
    avgResolutionDays: resCount > 0 ? Math.round((resSum / resCount) * 10) / 10 : 0,
    slaBreachPct: closed > 0 ? Math.round((breachCount / closed) * 1000) / 10 : 0
  };
}

/** Count non-empty label cells past rowStart, skipping section markers. */
function countBadgeRows_(sheet, rowStart) {
  const lastRow = sheet.getLastRow();
  if (lastRow < rowStart) return 0;
  const values = sheet.getRange(rowStart, 1, lastRow - rowStart + 1, 1).getValues();
  let count = 0;
  for (let i = 0; i < values.length; i++) {
    const v = values[i][0];
    if (v === '' || v === null || v === undefined) continue;
    const s = String(v);
    if (s.indexOf('---') === 0) continue;
    if (s.indexOf('No ') === 0) continue;       // "No reopened tickets found"
    count++;
  }
  return count;
}

/**
 * Build a Chart.js-ready payload from a sheet's raw values using a chart spec.
 * Returns null if no valid data rows remain after dirty-data filtering.
 */
function buildChartPayload_(values, spec) {
  const rowStart = spec.rowStart || 2;
  const startIdx = rowStart - 1;

  // 1. Slice the data window.
  let window;
  if (spec.rowMode === 'fixed') {
    window = values.slice(startIdx, startIdx + (spec.rowCount || 0));
  } else {
    window = [];
    for (let i = startIdx; i < values.length; i++) {
      const label = values[i][spec.labelCol];
      if (label === '' || label === null || label === undefined) break;
      const s = String(label);
      if (s.indexOf('---') === 0) break;
      window.push(values[i]);
    }
  }
  if (window.length === 0) return null;

  // 2. Resolve series — explicit or dynamic.
  let series = spec.series;
  if (!series && spec.dynamicSeries) {
    series = resolveDynamicSeries_(values, spec.dynamicSeries);
  }
  if (!series || series.length === 0) return null;

  // 3. Extract labels + numeric series, dropping rows with any non-finite value.
  const labels = [];
  const dataCols = series.map(() => []);

  for (let i = 0; i < window.length; i++) {
    const row = window[i];
    const label = row[spec.labelCol];
    const values_ = series.map(s => Number(row[s.col]));
    if (values_.some(v => !isFinite(v))) continue;
    labels.push(label instanceof Date ? label.toISOString().slice(0, 10) : String(label));
    values_.forEach((v, idx) => dataCols[idx].push(v));
  }

  if (labels.length === 0) return null;

  const datasets = series.map((s, idx) => ({
    label: s.header,
    color: s.color,
    data: dataCols[idx],
    percent: !!s.percent
  }));

  return { labels: labels, datasets: datasets };
}

/**
 * Dynamic series resolution — scans the header row for columns matching
 * a pattern and builds a series per match. Used for cross-tabs like
 * MonthlyVolumeByFA where FA columns are generated at setup time.
 */
function resolveDynamicSeries_(values, cfg) {
  if (!values[0]) return [];
  const headerRow = values[0];
  const palette = ['darkBlue', 'teal', 'gold', 'orange', 'purple', 'slate'];
  const series = [];
  const startCol = cfg.startCol || 1;

  for (let c = startCol; c < headerRow.length; c++) {
    const h = String(headerRow[c] || '');
    if (cfg.mode === 'headerSuffix' && h.endsWith(cfg.suffix)) {
      series.push({
        header: h,
        col: c,
        color: palette[series.length % palette.length]
      });
    }
  }
  return series;
}
