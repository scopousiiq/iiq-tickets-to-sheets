/**
 * ChartRegistry — declarative map of analytics sheets to chart specifications.
 *
 * The Dashboard web app iterates this registry, keeps entries whose sheetName
 * exists in the workbook, reads each sheet's data using the spec, and ships a
 * serializable payload to the client for Chart.js rendering.
 *
 * Entry shape:
 *   {
 *     sheetName:   string — must exactly match the sheet name created by Setup/OptionalMetrics
 *     category:    string — full category name (section/tab grouping)
 *     tabLabel:    string — short name shown on the tab bar
 *     charts:      array  — one or more chart specs (see Spec shape below)
 *     kpiBadge:    object — alternative to `charts` for list-only sheets
 *   }
 *
 * Spec shape:
 *   {
 *     title:    string — card header
 *     type:     'bar' | 'line' | 'stackedBar' | 'horizontalBar' | 'stackedHorizontalBar'
 *     labelCol: number — 0-indexed column for category/axis labels
 *     series:   [ { header: string, col: number, color: string } ]
 *     rowStart: number — 1-indexed first data row (row 1 is normally headers)
 *     rowMode:  'contiguous' — read until labelCol goes blank
 *               | 'fixed'    — take exactly rowCount rows from rowStart
 *     rowCount: number — required when rowMode='fixed'
 *   }
 *
 * Badge shape:
 *   { label: string, color: string, rowStart: number }
 *
 * Colors map to CSS tokens defined in Dashboard.html (IIQ.* palette).
 *
 * Sheet schemas verified against Setup.gs / OptionalMetrics.gs header arrays.
 * When adding a new analytics sheet: register it here or the dashboard won't
 * discover it.
 */

const CHART_REGISTRY = [

  // ───────────────────────────── Volume & Trends ─────────────────────────────
  {
    sheetName: 'MonthlyVolume',
    category: 'Volume & Trends',
    tabLabel: 'Volume',
    charts: [{
      title: 'Monthly Volume — Created vs Closed',
      type: 'bar',
      labelCol: 0,
      series: [
        { header: 'Created', col: 2, color: 'darkBlue' },
        { header: 'Closed',  col: 3, color: 'teal' }
      ],
      rowStart: 2,
      rowMode: 'contiguous'
    }, {
      title: 'Closure Rate — Monthly',
      type: 'line',
      labelCol: 0,
      series: [
        { header: 'Closure Rate', col: 5, color: 'gold', percent: true }
      ],
      rowStart: 2,
      rowMode: 'contiguous'
    }]
  },
  {
    sheetName: 'PerformanceTrends',
    category: 'Volume & Trends',
    tabLabel: 'Volume',
    charts: [{
      title: 'Performance Trends — Avg Resolution & Breach Rate',
      type: 'line',
      labelCol: 0,
      series: [
        { header: 'Avg Resolution (days)', col: 3, color: 'darkBlue' },
        { header: 'Breach Rate',           col: 5, color: 'orange', percent: true }
      ],
      rowStart: 2,
      rowMode: 'contiguous'
    }, {
      title: 'Backlog End-of-Month',
      type: 'bar',
      labelCol: 0,
      series: [
        { header: 'Backlog EOM',   col: 6, color: 'purple' },
        { header: '% Aged 30+ EOM', col: 7, color: 'gold', percent: true }
      ],
      rowStart: 2,
      rowMode: 'contiguous'
    }]
  },
  {
    sheetName: 'SeasonalComparison',
    category: 'Volume & Trends',
    tabLabel: 'Volume',
    charts: [{
      title: 'Seasonal Comparison — Tickets Created (YoY)',
      type: 'bar',
      labelCol: 0,
      series: [
        { header: 'Prior Year',   col: 1, color: 'slate' },
        { header: 'Current Year', col: 2, color: 'darkBlue' }
      ],
      rowStart: 2,
      rowMode: 'contiguous'
    }]
  },
  {
    sheetName: 'TemporalPatterns',
    category: 'Volume & Trends',
    tabLabel: 'Volume',
    charts: [{
      title: 'Tickets by Day of Week',
      type: 'bar',
      labelCol: 0,
      series: [
        { header: 'Created', col: 1, color: 'darkBlue' }
      ],
      rowStart: 2,
      rowMode: 'fixed',
      rowCount: 7
    }, {
      title: 'Tickets by Hour of Day',
      type: 'bar',
      labelCol: 3,   // "Time Range" column — human-readable
      series: [
        { header: 'Created', col: 1, color: 'teal' }
      ],
      rowStart: 12,
      rowMode: 'fixed',
      rowCount: 24
    }]
  },
  {
    sheetName: 'MonthlyVolumeByFA',
    category: 'Volume & Trends',
    tabLabel: 'Volume',
    charts: [{
      title: 'Monthly Volume by Functional Area',
      // Dynamic FA columns; client will auto-extract pairs ending in "Created"
      type: 'stackedBar',
      labelCol: 0,
      dynamicSeries: { mode: 'headerSuffix', suffix: 'Created', startCol: 2, color: 'darkBlue' },
      rowStart: 2,
      rowMode: 'contiguous'
    }]
  },

  // ───────────────────────────── Backlog & Quality ────────────────────────────
  {
    sheetName: 'BacklogAging',
    category: 'Backlog & Quality',
    tabLabel: 'Backlog',
    charts: [{
      title: 'Backlog Aging — Open Tickets',
      type: 'horizontalBar',
      labelCol: 0,
      series: [
        { header: 'Count', col: 1, color: 'orange' }
      ],
      rowStart: 2,
      rowMode: 'fixed',
      rowCount: 5          // 5 buckets; exclude TOTAL row
    }]
  },
  {
    sheetName: 'ResolutionAging',
    category: 'Backlog & Quality',
    tabLabel: 'Backlog',
    charts: [{
      title: 'Resolution Aging — Closed Tickets',
      type: 'horizontalBar',
      labelCol: 0,
      series: [
        { header: 'Count', col: 1, color: 'teal' }
      ],
      rowStart: 2,
      rowMode: 'fixed',
      rowCount: 5
    }]
  },
  {
    sheetName: 'BacklogAgingByFA',
    category: 'Backlog & Quality',
    tabLabel: 'Backlog',
    charts: [{
      title: 'Backlog Aging by Functional Area',
      type: 'stackedHorizontalBar',
      labelCol: 0,
      series: [
        { header: '0-7 days',   col: 1, color: 'teal' },
        { header: '8-14 days',  col: 2, color: 'darkBlue' },
        { header: '15-30 days', col: 3, color: 'gold' },
        { header: '30+ days',   col: 4, color: 'orange' }
      ],
      rowStart: 2,
      rowMode: 'contiguous'
    }]
  },
  {
    sheetName: 'BacklogAgingByTeam',
    category: 'Backlog & Quality',
    tabLabel: 'Backlog',
    charts: [{
      title: 'Backlog Aging by Team',
      type: 'stackedHorizontalBar',
      labelCol: 0,
      series: [
        { header: '0-7 days',   col: 1, color: 'teal' },
        { header: '8-14 days',  col: 2, color: 'darkBlue' },
        { header: '15-30 days', col: 3, color: 'gold' },
        { header: '30+ days',   col: 4, color: 'orange' }
      ],
      rowStart: 2,
      rowMode: 'contiguous'
    }]
  },
  {
    sheetName: 'BacklogAgingByLocationType',
    category: 'Backlog & Quality',
    tabLabel: 'Backlog',
    charts: [{
      title: 'Backlog Aging by Location Type',
      type: 'stackedHorizontalBar',
      labelCol: 0,
      series: [
        { header: '0-7 days',   col: 1, color: 'teal' },
        { header: '8-14 days',  col: 2, color: 'darkBlue' },
        { header: '15-30 days', col: 3, color: 'gold' },
        { header: '30+ days',   col: 4, color: 'orange' }
      ],
      rowStart: 2,
      rowMode: 'contiguous'
    }]
  },
  {
    sheetName: 'BacklogAgingByPriority',
    category: 'Backlog & Quality',
    tabLabel: 'Backlog',
    charts: [{
      title: 'Backlog Aging by Priority',
      type: 'stackedHorizontalBar',
      labelCol: 0,
      series: [
        { header: '0-7 days',   col: 1, color: 'teal' },
        { header: '8-14 days',  col: 2, color: 'darkBlue' },
        { header: '15-30 days', col: 3, color: 'gold' },
        { header: '30+ days',   col: 4, color: 'orange' }
      ],
      rowStart: 2,
      rowMode: 'contiguous'
    }]
  },
  {
    sheetName: 'StaleTickets',
    category: 'Backlog & Quality',
    tabLabel: 'Backlog',
    kpiBadge: { label: 'Stale Tickets', color: 'gold', rowStart: 2 }
  },
  {
    sheetName: 'ReopenRate',
    category: 'Backlog & Quality',
    tabLabel: 'Backlog',
    kpiBadge: { label: 'Reopened Tickets', color: 'purple', rowStart: 15 }
  },

  // ───────────────────────────── SLA & Response ──────────────────────────────
  {
    sheetName: 'SLACompliance',
    category: 'SLA & Response',
    tabLabel: 'SLA',
    charts: [{
      title: 'SLA Breach Rate — Monthly',
      type: 'line',
      labelCol: 0,
      series: [
        { header: 'Breach Rate', col: 4, color: 'orange', percent: true }
      ],
      rowStart: 2,
      rowMode: 'contiguous'
    }, {
      title: 'SLA Response vs Resolution (hrs)',
      type: 'bar',
      labelCol: 0,
      series: [
        { header: 'Avg Response',   col: 5, color: 'darkBlue' },
        { header: 'Avg Resolution', col: 6, color: 'teal' }
      ],
      rowStart: 2,
      rowMode: 'contiguous'
    }]
  },
  {
    sheetName: 'AtRiskResponse',
    category: 'SLA & Response',
    tabLabel: 'SLA',
    kpiBadge: { label: 'At-Risk Response', color: 'orange', rowStart: 2 }
  },
  {
    sheetName: 'AtRiskResolution',
    category: 'SLA & Response',
    tabLabel: 'SLA',
    kpiBadge: { label: 'At-Risk Resolution', color: 'red', rowStart: 2 }
  },
  {
    sheetName: 'FirstContactResolution',
    category: 'SLA & Response',
    tabLabel: 'SLA',
    charts: [{
      title: 'First Contact Resolution — Monthly',
      type: 'line',
      labelCol: 0,
      series: [
        { header: 'Same Day %', col: 4, color: 'teal',   percent: true },
        { header: '4-Hour %',   col: 6, color: 'darkBlue', percent: true }
      ],
      rowStart: 13,
      rowMode: 'contiguous'
    }]
  },
  {
    sheetName: 'ResponseDistribution',
    category: 'SLA & Response',
    tabLabel: 'SLA',
    charts: [{
      title: 'Response Time Distribution',
      type: 'bar',
      labelCol: 0,
      series: [
        { header: 'Response (hrs)',   col: 1, color: 'darkBlue' },
        { header: 'Resolution (hrs)', col: 2, color: 'teal' }
      ],
      rowStart: 13,
      rowMode: 'fixed',
      rowCount: 5
    }]
  },
  {
    sheetName: 'ResponseTrends',
    category: 'SLA & Response',
    tabLabel: 'SLA',
    charts: [{
      title: 'Response Time — Monthly',
      type: 'line',
      labelCol: 0,
      series: [
        { header: 'Avg Resp',    col: 3, color: 'darkBlue' },
        { header: 'Median Resp', col: 4, color: 'teal' },
        { header: '90th % Resp', col: 5, color: 'orange' }
      ],
      rowStart: 2,
      rowMode: 'contiguous'
    }]
  },
  {
    sheetName: 'QueueTimeAnalysis',
    category: 'SLA & Response',
    tabLabel: 'SLA',
    charts: [{
      title: 'Queue Time Distribution',
      type: 'horizontalBar',
      labelCol: 0,
      series: [
        { header: 'Count', col: 1, color: 'orange' }
      ],
      rowStart: 16,
      rowMode: 'fixed',
      rowCount: 7
    }]
  },
  {
    sheetName: 'QueueTimeByTeam',
    category: 'SLA & Response',
    tabLabel: 'SLA',
    charts: [{
      title: 'Queue Time by Team',
      type: 'horizontalBar',
      labelCol: 0,
      series: [
        { header: 'Avg Queue (hrs)', col: 1, color: 'orange' }
      ],
      rowStart: 2,
      rowMode: 'contiguous'
    }]
  },
  {
    sheetName: 'QueueTimeTrend',
    category: 'SLA & Response',
    tabLabel: 'SLA',
    charts: [{
      title: 'Queue Time — Monthly',
      type: 'line',
      labelCol: 0,
      series: [
        { header: 'Avg Queue',  col: 2, color: 'darkBlue' },
        { header: 'Median',     col: 3, color: 'teal' },
        { header: '90th Pctl',  col: 4, color: 'orange' }
      ],
      rowStart: 2,
      rowMode: 'contiguous'
    }]
  },

  // ───────────────────────────── Team & Staff ────────────────────────────────
  {
    sheetName: 'TeamWorkload',
    category: 'Team & Staff',
    tabLabel: 'Team',
    charts: [{
      title: 'Team Workload — Open & Aged 30+',
      type: 'horizontalBar',
      labelCol: 0,
      series: [
        { header: 'Open',     col: 2, color: 'darkBlue' },
        { header: 'Aged 30+', col: 5, color: 'red' }
      ],
      rowStart: 2,
      rowMode: 'contiguous'
    }]
  },
  {
    sheetName: 'TechnicianPerformance',
    category: 'Team & Staff',
    tabLabel: 'Team',
    charts: [{
      title: 'Technician Workload — Open Tickets',
      type: 'horizontalBar',
      labelCol: 0,
      series: [
        { header: 'Open',     col: 2, color: 'darkBlue' },
        { header: 'Aged 30+', col: 5, color: 'red' }
      ],
      rowStart: 2,
      rowMode: 'contiguous'
    }]
  },
  {
    sheetName: 'FunctionalAreaSummary',
    category: 'Team & Staff',
    tabLabel: 'Team',
    charts: [{
      title: 'Functional Area — Open & Aged',
      type: 'horizontalBar',
      labelCol: 0,
      series: [
        { header: 'Open',     col: 2, color: 'darkBlue' },
        { header: 'Aged 30+', col: 5, color: 'red' }
      ],
      rowStart: 2,
      rowMode: 'contiguous'
    }]
  },

  // ───────────────────────────── Location ────────────────────────────────────
  {
    sheetName: 'LocationBreakdown',
    category: 'Location',
    tabLabel: 'Location',
    charts: [{
      title: 'Location Breakdown — Open Tickets',
      type: 'horizontalBar',
      labelCol: 0,
      series: [
        { header: 'Open', col: 2, color: 'darkBlue' }
      ],
      rowStart: 2,
      rowMode: 'contiguous'
    }]
  },
  {
    sheetName: 'LocationTypeComparison',
    category: 'Location',
    tabLabel: 'Location',
    charts: [{
      title: 'Location Type — Open & Aged 30+',
      type: 'horizontalBar',
      labelCol: 0,
      series: [
        { header: 'Open',     col: 2, color: 'darkBlue' },
        { header: 'Aged 30+', col: 6, color: 'red' }
      ],
      rowStart: 2,
      rowMode: 'contiguous'
    }]
  },

  // ───────────────────────────── Issue & Requester ───────────────────────────
  {
    sheetName: 'IssueCategoryVolume',
    category: 'Issue & Requester',
    tabLabel: 'Issue',
    charts: [{
      title: 'Issue Category — Open Tickets',
      type: 'horizontalBar',
      labelCol: 0,
      series: [
        { header: 'Open', col: 1, color: 'darkBlue' }
      ],
      rowStart: 2,
      rowMode: 'contiguous'
    }]
  },
  {
    sheetName: 'PriorityAnalysis',
    category: 'Issue & Requester',
    tabLabel: 'Issue',
    charts: [{
      title: 'Priority — Open & Avg Resolution',
      type: 'bar',
      labelCol: 0,
      series: [
        { header: 'Open',               col: 1, color: 'darkBlue' },
        { header: 'Avg Resolution (d)', col: 4, color: 'orange' }
      ],
      rowStart: 2,
      rowMode: 'contiguous'
    }]
  },
  {
    sheetName: 'FrequentRequesters',
    category: 'Issue & Requester',
    tabLabel: 'Issue',
    charts: [{
      title: 'Frequent Requesters — Top Users by Ticket Count',
      type: 'horizontalBar',
      labelCol: 0,
      series: [
        { header: 'Total Tickets', col: 1, color: 'purple' }
      ],
      rowStart: 2,
      rowMode: 'contiguous'
    }]
  },

  // ───────────────────────────── Device ──────────────────────────────────────
  {
    sheetName: 'DeviceReliability',
    category: 'Device',
    tabLabel: 'Device',
    charts: [{
      title: 'Device Reliability — Tickets by Model',
      type: 'horizontalBar',
      labelCol: 0,
      series: [
        { header: 'Total',            col: 1, color: 'darkBlue' },
        { header: 'Avg Resolution (d)', col: 4, color: 'orange' }
      ],
      rowStart: 2,
      rowMode: 'contiguous'
    }]
  },
  {
    sheetName: 'DevicesByRole',
    category: 'Device',
    tabLabel: 'Device',
    charts: [{
      title: 'Devices by Role — Tickets by Model',
      type: 'horizontalBar',
      labelCol: 0,
      series: [
        { header: 'Total', col: 2, color: 'teal' }
      ],
      rowStart: 2,
      rowMode: 'contiguous'
    }]
  },
  {
    sheetName: 'FrequentFlyers',
    category: 'Device',
    tabLabel: 'Device',
    kpiBadge: { label: 'Frequent Flyers', color: 'purple', rowStart: 5 }
  }
];

/** Ordered category list — drives tab bar order, matches Menu.gs submenu taxonomy. */
const CATEGORY_ORDER = [
  'Volume & Trends',
  'Backlog & Quality',
  'SLA & Response',
  'Team & Staff',
  'Location',
  'Issue & Requester',
  'Device'
];
