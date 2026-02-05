/**
 * Analytics Sheets - All Available Metric Sheets
 *
 * All analytics sheets can be added individually via the iiQ Data > Add Analytics Sheet menu.
 * Adding a sheet will DELETE and RECREATE it if it already exists (clean slate).
 *
 * DEFAULT SHEETS (created by Setup Spreadsheet):
 * - MonthlyVolume: Ticket creation/closure by month
 * - BacklogAging: Current open ticket distribution by age
 * - TeamWorkload: Tickets by team with aging metrics
 * - SLACompliance: Monthly SLA breach rates
 * - PerformanceTrends: "Are we getting better?" trending
 * - AtRiskQueue: Tickets approaching SLA breach
 *
 * ADDITIONAL SHEETS (add via menu):
 * - StaleTickets: Tickets with no recent activity
 * - LocationBreakdown: Metrics by location
 * - FunctionalAreaSummary: Aggregated by functional area
 * - IssueCategoryVolume: What types of problems are we handling?
 * - PriorityAnalysis: Are high-priority tickets handled faster?
 * - FirstContactResolution: How many tickets resolved same-day?
 * - TechnicianPerformance: Individual technician workload and metrics
 * - SeasonalComparison: Year-over-year comparison by month
 * - LocationTypeComparison: Metrics by school type (Elementary/Middle/High)
 * - ReopenRate: Quality metric - are we truly resolving issues?
 * - FrequentRequesters: Who generates the most tickets?
 * - ResponseDistribution: Response time consistency analysis
 * - ResponseTrends: Monthly response/resolution time trends
 * - TemporalPatterns: When do tickets come in? (day/hour analysis)
 * - QueueTimeAnalysis: How long do tickets wait before being picked up?
 */

// ============================================================================
// MENU FUNCTIONS - Add Sheet Wrappers
// All analytics sheets can be added/recreated via menu
// ============================================================================

// --- Volume & Trends ---

/**
 * Add Monthly Volume sheet (deletes and recreates if exists)
 */
function addMonthlyVolumeSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupMonthlyVolumeSheet(ss);
  SpreadsheetApp.getUi().alert('Created', 'MonthlyVolume sheet has been created.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Add Performance Trends sheet (deletes and recreates if exists)
 */
function addPerformanceTrendsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupPerformanceTrendsSheet(ss);
  SpreadsheetApp.getUi().alert('Created', 'PerformanceTrends sheet has been created.', SpreadsheetApp.getUi().ButtonSet.OK);
}

// --- Backlog & Quality ---

/**
 * Add Backlog Aging sheet (deletes and recreates if exists)
 */
function addBacklogAgingSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupBacklogAgingSheet(ss);
  SpreadsheetApp.getUi().alert('Created', 'BacklogAging sheet has been created.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Add Stale Tickets sheet (deletes and recreates if exists)
 */
function addStaleTicketsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupStaleTicketsSheet(ss);
  SpreadsheetApp.getUi().alert('Created', 'StaleTickets sheet has been created.', SpreadsheetApp.getUi().ButtonSet.OK);
}

// --- SLA & Response ---

/**
 * Add SLA Compliance sheet (deletes and recreates if exists)
 */
function addSLAComplianceSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSLAComplianceSheet(ss);
  SpreadsheetApp.getUi().alert('Created', 'SLACompliance sheet has been created.\n\nData calculates automatically from TicketData via formulas.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Add At-Risk Response sheet (deletes and recreates if exists)
 */
function addAtRiskResponseSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupAtRiskResponseSheet(ss);
  SpreadsheetApp.getUi().alert('Created', 'AtRiskResponse sheet has been created.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Add At-Risk Resolution sheet (deletes and recreates if exists)
 */
function addAtRiskResolutionSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupAtRiskResolutionSheet(ss);
  SpreadsheetApp.getUi().alert('Created', 'AtRiskResolution sheet has been created.', SpreadsheetApp.getUi().ButtonSet.OK);
}

// --- Team & Staff ---

/**
 * Add Team Workload sheet (deletes and recreates if exists)
 */
function addTeamWorkloadSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupTeamWorkloadSheet(ss);
  SpreadsheetApp.getUi().alert('Created', 'TeamWorkload sheet has been created.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Add Functional Area Summary sheet (deletes and recreates if exists)
 */
function addFunctionalAreaSummarySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupFunctionalAreaSummarySheet(ss);
  SpreadsheetApp.getUi().alert('Created', 'FunctionalAreaSummary sheet has been created.\n\nNote: Requires TeamWorkload sheet and FunctionalArea column in Teams sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
}

// --- Location ---

/**
 * Add Location Breakdown sheet (deletes and recreates if exists)
 */
function addLocationBreakdownSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupLocationBreakdownSheet(ss);
  SpreadsheetApp.getUi().alert('Created', 'LocationBreakdown sheet has been created.', SpreadsheetApp.getUi().ButtonSet.OK);
}

// --- Additional Metrics (defined below) ---

/**
 * Add Issue Category Volume sheet (deletes and recreates if exists)
 */
function addIssueCategoryVolumeSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupIssueCategoryVolumeSheet(ss);
  SpreadsheetApp.getUi().alert('Created', 'IssueCategoryVolume sheet has been created.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Add Priority Analysis sheet (deletes and recreates if exists)
 */
function addPriorityAnalysisSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupPriorityAnalysisSheet(ss);
  SpreadsheetApp.getUi().alert('Created', 'PriorityAnalysis sheet has been created.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Add First Contact Resolution sheet (deletes and recreates if exists)
 */
function addFirstContactResolutionSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupFirstContactResolutionSheet(ss);
  SpreadsheetApp.getUi().alert('Created', 'FirstContactResolution sheet has been created.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Add Technician Performance sheet (deletes and recreates if exists)
 */
function addTechnicianPerformanceSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupTechnicianPerformanceSheet(ss);
  SpreadsheetApp.getUi().alert('Created', 'TechnicianPerformance sheet has been created.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Add Seasonal Comparison sheet (deletes and recreates if exists)
 */
function addSeasonalComparisonSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSeasonalComparisonSheet(ss);
  SpreadsheetApp.getUi().alert('Created', 'SeasonalComparison sheet has been created.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Add Location Type Comparison sheet (deletes and recreates if exists)
 */
function addLocationTypeComparisonSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupLocationTypeComparisonSheet(ss);
  SpreadsheetApp.getUi().alert('Created', 'LocationTypeComparison sheet has been created.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Add Reopen Rate sheet (deletes and recreates if exists)
 */
function addReopenRateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupReopenRateSheet(ss);
  SpreadsheetApp.getUi().alert('Created', 'ReopenRate sheet has been created.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Add Frequent Requesters sheet (deletes and recreates if exists)
 */
function addFrequentRequestersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupFrequentRequestersSheet(ss);
  SpreadsheetApp.getUi().alert('Created', 'FrequentRequesters sheet has been created.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Add Response Distribution sheet (deletes and recreates if exists)
 */
function addResponseDistributionSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupResponseDistributionSheet(ss);
  SpreadsheetApp.getUi().alert('Created', 'ResponseDistribution sheet has been created.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Add Response Trends sheet (deletes and recreates if exists)
 */
function addResponseTrendsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupResponseTrendsSheet(ss);
  SpreadsheetApp.getUi().alert('Created', 'ResponseTrends sheet has been created.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Add Queue Time Analysis sheet (deletes and recreates if exists)
 */
function addQueueTimeAnalysisSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupQueueTimeAnalysisSheet(ss);
  SpreadsheetApp.getUi().alert('Created', 'QueueTimeAnalysis sheet has been created.\n\nThis sheet analyzes the time tickets spend waiting before an agent starts working on them.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Add Queue Time by Team sheet (deletes and recreates if exists)
 */
function addQueueTimeByTeamSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupQueueTimeByTeamSheet(ss);
  SpreadsheetApp.getUi().alert('Created', 'QueueTimeByTeam sheet has been created.\n\nThis sheet shows queue time metrics broken down by team.\nIncludes diagnostics to help debug any formula issues.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Add Queue Time Trend sheet (deletes and recreates if exists)
 */
function addQueueTimeTrendSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupQueueTimeTrendSheet(ss);
  SpreadsheetApp.getUi().alert('Created', 'QueueTimeTrend sheet has been created.\n\nThis sheet shows monthly queue time trends with avg, median, and 90th percentile metrics.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Add Temporal Patterns sheet (deletes and recreates if exists)
 */
function addTemporalPatternsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupTemporalPatternsSheet(ss);
  SpreadsheetApp.getUi().alert('Created', 'TemporalPatterns sheet has been created.', SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Get the range of months to display based on actual ticket data.
 * Returns array of {month: 0-11, year: YYYY, monthName: 'January'} objects.
 *
 * Priority:
 * 1. TicketData date range (uses specified date column)
 * 2. Config year keys (TICKET_{YEAR}_*)
 * 3. Default: last 13 months
 *
 * @param {Spreadsheet} ss - The spreadsheet
 * @param {string} dateColumn - Column letter to use for date range (e.g., 'E' for CreatedDate, 'G' for ClosedDate)
 * @returns {Array} Array of month objects sorted oldest to newest
 */
function getMonthRangeFromData(ss, dateColumn) {
  const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
                      'July', 'August', 'September', 'October', 'November', 'December'];

  // Priority 1: Use school year dates from Config (authoritative for single-year spreadsheets)
  const config = getConfig();
  if (config.schoolYear) {
    const dates = getSchoolYearDates(config);
    if (dates) {
      return generateMonthRange(dates.startDate, dates.endDate, monthNames);
    }
  }

  // Priority 2: Fall back to date range from TicketData (if no school year configured)
  const ticketSheet = ss.getSheetByName('TicketData');
  if (ticketSheet && ticketSheet.getLastRow() > 1) {
    const colIndex = dateColumn.charCodeAt(0) - 64; // A=1, B=2, etc.
    const dateRange = ticketSheet.getRange(2, colIndex, ticketSheet.getLastRow() - 1, 1).getValues();

    let minDate = null;
    let maxDate = null;

    for (const row of dateRange) {
      const val = row[0];
      if (!val) continue;

      let date;
      if (val instanceof Date) {
        date = val;
      } else if (typeof val === 'string' && val.length >= 10) {
        date = new Date(val.substring(0, 10));
      }

      if (date && !isNaN(date.getTime())) {
        if (!minDate || date < minDate) minDate = date;
        if (!maxDate || date > maxDate) maxDate = date;
      }
    }

    if (minDate && maxDate) {
      return generateMonthRange(minDate, maxDate, monthNames);
    }
  }

  // Default: last 13 months
  const now = new Date();
  const maxDate = new Date(now.getFullYear(), now.getMonth(), 1);
  const minDate = new Date(now.getFullYear(), now.getMonth() - 12, 1);
  return generateMonthRange(minDate, maxDate, monthNames);
}

/**
 * Generate array of month objects between two dates
 */
function generateMonthRange(minDate, maxDate, monthNames) {
  const months = [];
  const startYear = minDate.getFullYear();
  const startMonth = minDate.getMonth();
  const endYear = maxDate.getFullYear();
  const endMonth = maxDate.getMonth();

  for (let y = startYear; y <= endYear; y++) {
    const mStart = (y === startYear) ? startMonth : 0;
    const mEnd = (y === endYear) ? endMonth : 11;

    for (let m = mStart; m <= mEnd; m++) {
      months.push({
        month: m,
        year: y,
        monthName: monthNames[m],
        monthNum: m + 1 // 1-indexed for formulas
      });
    }
  }

  return months;
}

// ============================================================================
// SETUP FUNCTIONS - Create individual metric sheets
// ============================================================================

/**
 * Setup IssueCategoryVolume sheet
 * Question: "What types of problems are we spending the most time on?"
 * Deletes existing sheet if present for clean slate
 */
function setupIssueCategoryVolumeSheet(ss) {
  deleteSheetIfExists(ss, 'IssueCategoryVolume');
  const sheet = ss.insertSheet('IssueCategoryVolume');

  // Headers - includes sort controls
  const headers = ['Issue Category', 'Open', 'Created (MTD)', 'Closed (MTD)', 'Avg Resolution (days)', 'Breach Rate', 'Last Refreshed', 'Sort Col#', 'Desc?'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Main formula - aggregates by IssueCategoryName (column Y), with dynamic sorting
  const mainFormula =
    '=LET(' +
    'cats, UNIQUE(FILTER(TicketData!Y2:Y, TicketData!Y2:Y<>"", TicketData!Y2:Y<>"IssueCategoryName")),' +
    'mtdStart, TEXT(DATE(YEAR(TODAY()),MONTH(TODAY()),1), "YYYY-MM-DD"),' +
    'mtdEnd, TEXT(DATE(YEAR(TODAY()),MONTH(TODAY())+1,1), "YYYY-MM-DD"),' +
    'col_a, cats,' +
    'col_b, BYROW(cats, LAMBDA(c, COUNTIFS(TicketData!Y:Y, c, TicketData!I:I, "No"))),' +
    'col_c, BYROW(cats, LAMBDA(c, COUNTIFS(TicketData!Y:Y, c, TicketData!E:E, ">="&mtdStart, TicketData!E:E, "<"&mtdEnd))),' +
    'col_d, BYROW(cats, LAMBDA(c, COUNTIFS(TicketData!Y:Y, c, TicketData!H:H, ">="&mtdStart, TicketData!H:H, "<"&mtdEnd))),' +
    'col_e, BYROW(cats, LAMBDA(c, IFERROR(AVERAGEIFS(TicketData!R:R, TicketData!Y:Y, c, TicketData!I:I, "Yes"), "N/A"))),' +
    'col_f, BYROW(cats, LAMBDA(c, LET(total, COUNTIFS(TicketData!Y:Y, c, TicketData!I:I, "Yes"), breached, COUNTIFS(TicketData!Y:Y, c, TicketData!I:I, "Yes", TicketData!AF:AF, TRUE)+COUNTIFS(TicketData!Y:Y, c, TicketData!I:I, "Yes", TicketData!AI:AI, TRUE), IF(total>0, breached/total, "N/A")))),' +
    'data, HSTACK(col_a, col_b, col_c, col_d, col_e, col_f),' +
    'SORT(data, $H$2, $I$2))';

  sheet.getRange('A2').setValue(mainFormula);
  sheet.getRange('G2').setValue('=IFERROR(VLOOKUP("LAST_REFRESH", Config!A:B, 2, FALSE), "")');

  // Default sort settings (column 2 = Open, descending)
  sheet.getRange('H2').setValue(2);
  sheet.getRange('I2').setValue('FALSE');

  // Format header
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#00897b')
    .setFontColor('white');

  // Format columns
  sheet.getRange('E:E').setNumberFormat('0.0');   // Avg Resolution
  sheet.getRange('F:F').setNumberFormat('0.0%');  // Breach Rate

  // Column widths
  sheet.setColumnWidth(1, 200);  // Issue Category
  sheet.setColumnWidth(7, 180);  // Last Refreshed
  sheet.setColumnWidth(8, 80);   // Sort Col#
  sheet.setColumnWidth(9, 60);   // Desc?

  sheet.setFrozenRows(1);

  // Add data validation for sort column
  const sortColRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['1', '2', '3', '4', '5', '6'], true)
    .setHelpText('1=Category, 2=Open, 3=Created, 4=Closed, 5=AvgRes, 6=Breach')
    .build();
  sheet.getRange('H2').setDataValidation(sortColRule);

  // Add data validation for sort order
  const sortOrderRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['FALSE', 'TRUE'], true)
    .setHelpText('FALSE=Descending, TRUE=Ascending')
    .build();
  sheet.getRange('I2').setDataValidation(sortOrderRule);

  // Add notes
  sheet.getRange('A1').setNote(
    'Issue Category Volume Analysis\n\n' +
    'Question: "What types of problems are we handling?"\n\n' +
    'Use this to:\n' +
    '- Identify systemic issues (e.g., Chromebook damage spikes)\n' +
    '- Justify budget for equipment replacement\n' +
    '- Target training for common user issues\n' +
    '- Prioritize automation opportunities\n\n' +
    'Use Sort Col# and Desc? to change sorting.'
  );
  sheet.getRange('H2').setNote('Sort column: 1=Category, 2=Open, 3=Created, 4=Closed, 5=AvgRes, 6=Breach');
  sheet.getRange('I2').setNote('FALSE=Descending (high to low), TRUE=Ascending (low to high)');

  return true;
}

/**
 * Setup PriorityAnalysis sheet
 * Question: "Are high-priority tickets being handled faster?"
 * Deletes existing sheet if present for clean slate
 */
function setupPriorityAnalysisSheet(ss) {
  deleteSheetIfExists(ss, 'PriorityAnalysis');
  const sheet = ss.insertSheet('PriorityAnalysis');

  // Headers
  const headers = ['Priority', 'Open', 'Created (MTD)', 'Closed (MTD)', 'Avg Resolution (days)', 'Avg Response (hrs)', 'Breach Rate', 'Last Refreshed'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Main formula - aggregates by Priority (column S)
  const mainFormula =
    '=LET(' +
    'pris, UNIQUE(FILTER(TicketData!S2:S, TicketData!S2:S<>"", TicketData!S2:S<>"Priority")),' +
    'mtdStart, TEXT(DATE(YEAR(TODAY()),MONTH(TODAY()),1), "YYYY-MM-DD"),' +
    'mtdEnd, TEXT(DATE(YEAR(TODAY()),MONTH(TODAY())+1,1), "YYYY-MM-DD"),' +
    'col_a, pris,' +
    'col_b, BYROW(pris, LAMBDA(p, COUNTIFS(TicketData!S:S, p, TicketData!I:I, "No"))),' +
    'col_c, BYROW(pris, LAMBDA(p, COUNTIFS(TicketData!S:S, p, TicketData!E:E, ">="&mtdStart, TicketData!E:E, "<"&mtdEnd))),' +
    'col_d, BYROW(pris, LAMBDA(p, COUNTIFS(TicketData!S:S, p, TicketData!H:H, ">="&mtdStart, TicketData!H:H, "<"&mtdEnd))),' +
    'col_e, BYROW(pris, LAMBDA(p, IFERROR(AVERAGEIFS(TicketData!R:R, TicketData!S:S, p, TicketData!I:I, "Yes"), "N/A"))),' +
    'col_f, BYROW(pris, LAMBDA(p, IFERROR(AVERAGEIFS(TicketData!AE:AE, TicketData!S:S, p, TicketData!I:I, "Yes", TicketData!AE:AE, ">0")/60, "N/A"))),' +
    'col_g, BYROW(pris, LAMBDA(p, LET(total, COUNTIFS(TicketData!S:S, p, TicketData!I:I, "Yes"), breached, COUNTIFS(TicketData!S:S, p, TicketData!I:I, "Yes", TicketData!AF:AF, TRUE)+COUNTIFS(TicketData!S:S, p, TicketData!I:I, "Yes", TicketData!AI:AI, TRUE), IF(total>0, breached/total, "N/A")))),' +
    'data, HSTACK(col_a, col_b, col_c, col_d, col_e, col_f, col_g),' +
    'SORT(data, 2, FALSE))';

  sheet.getRange('A2').setValue(mainFormula);
  sheet.getRange('H2').setValue('=IFERROR(VLOOKUP("LAST_REFRESH", Config!A:B, 2, FALSE), "")');

  // Format header
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#d32f2f')
    .setFontColor('white');

  // Format columns
  sheet.getRange('E:E').setNumberFormat('0.0');   // Avg Resolution
  sheet.getRange('F:F').setNumberFormat('0.0');   // Avg Response
  sheet.getRange('G:G').setNumberFormat('0.0%');  // Breach Rate

  // Column widths
  sheet.setColumnWidth(1, 120);  // Priority
  sheet.setColumnWidth(8, 180);  // Last Refreshed

  sheet.setFrozenRows(1);

  // Add note
  sheet.getRange('A1').setNote(
    'Priority Analysis\n\n' +
    'Question: "Are high-priority tickets being handled faster?"\n\n' +
    'Use this to:\n' +
    '- Validate priority classifications are meaningful\n' +
    '- Ensure critical issues get appropriate attention\n' +
    '- Identify if priority escalation is overused\n' +
    '- Set expectations for response times by priority'
  );

  return true;
}

/**
 * Setup FirstContactResolution sheet
 * Question: "What percentage of tickets are resolved quickly?"
 * Deletes existing sheet if present for clean slate
 */
function setupFirstContactResolutionSheet(ss) {
  deleteSheetIfExists(ss, 'FirstContactResolution');
  const sheet = ss.insertSheet('FirstContactResolution');

  // Headers for summary section
  const headers = ['Metric', 'Value', 'Description'];
  sheet.getRange(1, 1, 1, 3).setValues([headers]);

  // Define FCR thresholds
  const metrics = [
    ['', '', ''],
    ['--- Current Month ---', '', ''],
    ['Total Closed (MTD)',
     '=LET(mtdStart, TEXT(DATE(YEAR(TODAY()),MONTH(TODAY()),1), "YYYY-MM-DD"), mtdEnd, TEXT(DATE(YEAR(TODAY()),MONTH(TODAY())+1,1), "YYYY-MM-DD"), COUNTIFS(TicketData!H:H, ">="&mtdStart, TicketData!H:H, "<"&mtdEnd))',
     'Tickets closed this month'],
    ['Resolved Same Day (<24 hrs)',
     '=LET(mtdStart, TEXT(DATE(YEAR(TODAY()),MONTH(TODAY()),1), "YYYY-MM-DD"), mtdEnd, TEXT(DATE(YEAR(TODAY()),MONTH(TODAY())+1,1), "YYYY-MM-DD"), COUNTIFS(TicketData!H:H, ">="&mtdStart, TicketData!H:H, "<"&mtdEnd, TicketData!R:R, "<=1"))',
     'Closed within 1 day of creation'],
    ['Same Day Rate',
     '=IF(B4>0, B5/B4, "N/A")',
     'FCR target: 40-60%'],
    ['Resolved Within 4 Hours',
     '=LET(mtdStart, TEXT(DATE(YEAR(TODAY()),MONTH(TODAY()),1), "YYYY-MM-DD"), mtdEnd, TEXT(DATE(YEAR(TODAY()),MONTH(TODAY())+1,1), "YYYY-MM-DD"), COUNTIFS(TicketData!H:H, ">="&mtdStart, TicketData!H:H, "<"&mtdEnd, TicketData!AH:AH, "<=240", TicketData!AH:AH, ">0"))',
     'Closed within 4 hours (resolution SLA)'],
    ['4-Hour Resolution Rate',
     '=IF(B4>0, B7/B4, "N/A")',
     'Quick resolution target: 20-30%'],
    ['', '', ''],
    ['--- Historical Trend ---', '', ''],
  ];

  sheet.getRange(2, 1, metrics.length, 3).setValues(metrics);

  // Add monthly trend section starting at row 12
  const trendHeaders = ['Month', 'Year', 'Total Closed', 'Same Day', 'Same Day %', '4-Hour', '4-Hour %'];
  sheet.getRange(12, 1, 1, 7).setValues([trendHeaders]);

  // Get month range from actual data (uses ClosedDate column G)
  const monthRange = getMonthRangeFromData(ss, 'G');

  const trendRows = [];
  for (const period of monthRange) {
    const rowNum = trendRows.length + 13;
    const y = period.year;
    const monthNum = period.monthNum;

    trendRows.push([
      period.monthName,
      y,
      // Total closed - simplified formula using hardcoded year/month
      `=COUNTIFS(TicketData!H:H, ">="&TEXT(DATE(${y},${monthNum},1), "YYYY-MM-DD"), TicketData!H:H, "<"&TEXT(DATE(${y},${monthNum}+1,1), "YYYY-MM-DD"))`,
      // Same day
      `=COUNTIFS(TicketData!H:H, ">="&TEXT(DATE(${y},${monthNum},1), "YYYY-MM-DD"), TicketData!H:H, "<"&TEXT(DATE(${y},${monthNum}+1,1), "YYYY-MM-DD"), TicketData!R:R, "<=1")`,
      // Same day %
      `=IF(C${rowNum}>0, D${rowNum}/C${rowNum}, "N/A")`,
      // 4-hour
      `=COUNTIFS(TicketData!H:H, ">="&TEXT(DATE(${y},${monthNum},1), "YYYY-MM-DD"), TicketData!H:H, "<"&TEXT(DATE(${y},${monthNum}+1,1), "YYYY-MM-DD"), TicketData!AH:AH, "<=240", TicketData!AH:AH, ">0")`,
      // 4-hour %
      `=IF(C${rowNum}>0, F${rowNum}/C${rowNum}, "N/A")`
    ]);
  }

  if (trendRows.length > 0) {
    sheet.getRange(13, 1, trendRows.length, 7).setValues(trendRows);
  }
  const lastTrendRow = 12 + trendRows.length;

  // Format headers
  sheet.getRange(1, 1, 1, 3)
    .setFontWeight('bold')
    .setBackground('#1976d2')
    .setFontColor('white');

  sheet.getRange(12, 1, 1, 7)
    .setFontWeight('bold')
    .setBackground('#1976d2')
    .setFontColor('white');

  // Format section headers
  sheet.getRange('A3').setFontWeight('bold').setBackground('#e3f2fd');
  sheet.getRange('A11').setFontWeight('bold').setBackground('#e3f2fd');

  // Format percentages
  sheet.getRange('B6').setNumberFormat('0.0%');
  sheet.getRange('B9').setNumberFormat('0.0%');
  sheet.getRange(13, 5, trendRows.length, 1).setNumberFormat('0.0%');  // Same Day %
  sheet.getRange(13, 7, trendRows.length, 1).setNumberFormat('0.0%');  // 4-Hour %

  // Column widths
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 250);

  sheet.setFrozenRows(1);

  // Add note
  sheet.getRange('A1').setNote(
    'First Contact Resolution (FCR) Analysis\n\n' +
    'Question: "How many tickets are resolved quickly?"\n\n' +
    'Use this to:\n' +
    '- Measure service desk efficiency\n' +
    '- Identify training opportunities\n' +
    '- Track customer satisfaction proxy\n' +
    '- Justify staffing levels\n\n' +
    'Industry benchmarks:\n' +
    '- Same Day: 40-60% is good\n' +
    '- 4-Hour: 20-30% is good'
  );

  return true;
}

/**
 * Setup TechnicianPerformance sheet
 * Question: "How is workload distributed among staff?"
 * Deletes existing sheet if present for clean slate
 */
function setupTechnicianPerformanceSheet(ss) {
  deleteSheetIfExists(ss, 'TechnicianPerformance');
  const sheet = ss.insertSheet('TechnicianPerformance');

  // Headers - includes sort controls
  const headers = ['Technician', 'Team', 'Open', 'Created (MTD)', 'Closed (MTD)', 'Aged 30+', 'Avg Resolution (days)', 'Breach Rate', 'Last Refreshed', 'Sort Col#', 'Desc?'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Main formula - aggregates by OwnerName (column Q), with dynamic sorting
  const mainFormula =
    '=LET(' +
    'techs, UNIQUE(FILTER(TicketData!Q2:Q, TicketData!Q2:Q<>"", TicketData!Q2:Q<>"OwnerName")),' +
    'mtdStart, TEXT(DATE(YEAR(TODAY()),MONTH(TODAY()),1), "YYYY-MM-DD"),' +
    'mtdEnd, TEXT(DATE(YEAR(TODAY()),MONTH(TODAY())+1,1), "YYYY-MM-DD"),' +
    'col_a, techs,' +
    'col_b, BYROW(techs, LAMBDA(t, IFERROR(INDEX(TicketData!L:L, MATCH(t, TicketData!Q:Q, 0)), ""))),' +
    'col_c, BYROW(techs, LAMBDA(t, COUNTIFS(TicketData!Q:Q, t, TicketData!I:I, "No"))),' +
    'col_d, BYROW(techs, LAMBDA(t, COUNTIFS(TicketData!Q:Q, t, TicketData!E:E, ">="&mtdStart, TicketData!E:E, "<"&mtdEnd))),' +
    'col_e, BYROW(techs, LAMBDA(t, COUNTIFS(TicketData!Q:Q, t, TicketData!H:H, ">="&mtdStart, TicketData!H:H, "<"&mtdEnd))),' +
    'col_f, BYROW(techs, LAMBDA(t, COUNTIFS(TicketData!Q:Q, t, TicketData!I:I, "No", TicketData!R:R, ">=30"))),' +
    'col_g, BYROW(techs, LAMBDA(t, IFERROR(AVERAGEIFS(TicketData!R:R, TicketData!Q:Q, t, TicketData!I:I, "Yes"), "N/A"))),' +
    'col_h, BYROW(techs, LAMBDA(t, LET(total, COUNTIFS(TicketData!Q:Q, t, TicketData!I:I, "Yes"), breached, COUNTIFS(TicketData!Q:Q, t, TicketData!I:I, "Yes", TicketData!AF:AF, TRUE)+COUNTIFS(TicketData!Q:Q, t, TicketData!I:I, "Yes", TicketData!AI:AI, TRUE), IF(total>0, breached/total, "N/A")))),' +
    'data, HSTACK(col_a, col_b, col_c, col_d, col_e, col_f, col_g, col_h),' +
    'SORT(data, $J$2, $K$2))';

  sheet.getRange('A2').setValue(mainFormula);
  sheet.getRange('I2').setValue('=IFERROR(VLOOKUP("LAST_REFRESH", Config!A:B, 2, FALSE), "")');

  // Default sort settings (column 5 = Closed MTD, descending)
  sheet.getRange('J2').setValue(5);
  sheet.getRange('K2').setValue('FALSE');

  // Format header
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#7b1fa2')
    .setFontColor('white');

  // Format columns
  sheet.getRange('G:G').setNumberFormat('0.0');   // Avg Resolution
  sheet.getRange('H:H').setNumberFormat('0.0%');  // Breach Rate

  // Column widths
  sheet.setColumnWidth(1, 180);  // Technician
  sheet.setColumnWidth(2, 150);  // Team
  sheet.setColumnWidth(9, 180);  // Last Refreshed
  sheet.setColumnWidth(10, 80);  // Sort Col#
  sheet.setColumnWidth(11, 60);  // Desc?

  sheet.setFrozenRows(1);

  // Add data validation for sort column
  const sortColRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['1', '2', '3', '4', '5', '6', '7', '8'], true)
    .setHelpText('1=Tech, 2=Team, 3=Open, 4=Created, 5=Closed, 6=Aged, 7=AvgRes, 8=Breach')
    .build();
  sheet.getRange('J2').setDataValidation(sortColRule);

  // Add data validation for sort order
  const sortOrderRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['FALSE', 'TRUE'], true)
    .setHelpText('FALSE=Descending, TRUE=Ascending')
    .build();
  sheet.getRange('K2').setDataValidation(sortOrderRule);

  // Add notes
  sheet.getRange('A1').setNote(
    'Technician Performance Dashboard\n\n' +
    'Question: "How is workload distributed among staff?"\n\n' +
    'Use this to:\n' +
    '- Identify overloaded technicians\n' +
    '- Balance workload across team\n' +
    '- Inform performance reviews\n' +
    '- Plan training based on breach rates\n' +
    '- Justify additional staffing\n\n' +
    'Use Sort Col# and Desc? to change sorting.'
  );
  sheet.getRange('J2').setNote('Sort column: 1=Tech, 2=Team, 3=Open, 4=Created, 5=Closed, 6=Aged, 7=AvgRes, 8=Breach');
  sheet.getRange('K2').setNote('FALSE=Descending (high to low), TRUE=Ascending (low to high)');

  return true;
}

/**
 * Setup SeasonalComparison sheet
 * Question: "How does this month compare to the same month last year?"
 * Deletes existing sheet if present for clean slate
 */
function setupSeasonalComparisonSheet(ss) {
  deleteSheetIfExists(ss, 'SeasonalComparison');
  const sheet = ss.insertSheet('SeasonalComparison');

  // Get school year configuration
  const config = getConfig();
  const dates = getSchoolYearDates(config);

  // Use school year dates if available, otherwise fall back to calendar years
  let startYear, endYear;
  if (dates) {
    startYear = dates.startYear;
    endYear = dates.endYear;
  } else {
    const currentYear = new Date().getFullYear();
    startYear = currentYear - 1;
    endYear = currentYear;
  }

  // Headers - compare first half vs second half of school year
  const headers = ['Month', `Created (${startYear})`, `Created (${endYear})`, 'Change', 'Change %', `Closed (${startYear})`, `Closed (${endYear})`, 'Closed Change %'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const months = ['January', 'February', 'March', 'April', 'May', 'June',
                  'July', 'August', 'September', 'October', 'November', 'December'];

  const dataRows = [];
  for (let m = 0; m < 12; m++) {
    const rowNum = m + 2;
    // Determine which year this month belongs to based on school year
    // For a typical July-June school year: July-Dec = startYear, Jan-June = endYear
    const monthYear = m >= 6 ? startYear : endYear; // July (6) through Dec (11) = startYear

    dataRows.push([
      months[m],
      // Created in startYear (only shows data if month is in startYear portion)
      `=LET(m, ${m + 1}, COUNTIFS(TicketData!E:E, ">="&TEXT(DATE(${startYear},m,1), "YYYY-MM-DD"), TicketData!E:E, "<"&TEXT(DATE(${startYear},m+1,1), "YYYY-MM-DD")))`,
      // Created in endYear (only shows data if month is in endYear portion)
      `=LET(m, ${m + 1}, COUNTIFS(TicketData!E:E, ">="&TEXT(DATE(${endYear},m,1), "YYYY-MM-DD"), TicketData!E:E, "<"&TEXT(DATE(${endYear},m+1,1), "YYYY-MM-DD")))`,
      // Change
      `=B${rowNum}-C${rowNum}`,
      // Change %
      `=IF(C${rowNum}>0, (B${rowNum}-C${rowNum})/C${rowNum}, IF(B${rowNum}>0, "N/A (prev=0)", "N/A"))`,
      // Closed in startYear
      `=LET(m, ${m + 1}, COUNTIFS(TicketData!H:H, ">="&TEXT(DATE(${startYear},m,1), "YYYY-MM-DD"), TicketData!H:H, "<"&TEXT(DATE(${startYear},m+1,1), "YYYY-MM-DD")))`,
      // Closed in endYear
      `=LET(m, ${m + 1}, COUNTIFS(TicketData!H:H, ">="&TEXT(DATE(${endYear},m,1), "YYYY-MM-DD"), TicketData!H:H, "<"&TEXT(DATE(${endYear},m+1,1), "YYYY-MM-DD")))`,
      // Closed Change %
      `=IF(G${rowNum}>0, (F${rowNum}-G${rowNum})/G${rowNum}, IF(F${rowNum}>0, "N/A (prev=0)", "N/A"))`
    ]);
  }

  sheet.getRange(2, 1, dataRows.length, 8).setValues(dataRows);

  // Format header
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#0288d1')
    .setFontColor('white');

  // Format percentages
  sheet.getRange('E:E').setNumberFormat('+0.0%;-0.0%');
  sheet.getRange('H:H').setNumberFormat('+0.0%;-0.0%');

  // Conditional formatting for Change
  const changeRange = sheet.getRange('D2:D13');
  const positiveRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground('#fce8e6')
    .setRanges([changeRange])
    .build();
  const negativeRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setBackground('#e6f4ea')
    .setRanges([changeRange])
    .build();
  sheet.setConditionalFormatRules([positiveRule, negativeRule]);

  // Column widths
  sheet.setColumnWidth(1, 100);

  sheet.setFrozenRows(1);

  // Add note explaining the school year context
  const schoolYearStr = config.schoolYear || `${startYear}-${endYear}`;
  sheet.getRange('A1').setNote(
    `Seasonal Comparison for School Year ${schoolYearStr}\n\n` +
    'NOTE: This sheet compares months within the school year.\n' +
    `For a typical July-June school year:\n` +
    `- Column B shows ${startYear} data (July-December)\n` +
    `- Column C shows ${endYear} data (January-June)\n\n` +
    'For true year-over-year comparison, open both school year\n' +
    'spreadsheets and compare MonthlyVolume sheets manually.\n\n' +
    'K-12 peaks:\n' +
    '- August/September: Back-to-school\n' +
    '- March/April: State testing\n' +
    '- May/June: End-of-year\n\n' +
    'Green = decreased volume (good)\n' +
    'Red = increased volume (may need resources)'
  );

  return true;
}

/**
 * Setup LocationTypeComparison sheet
 * Question: "Do elementary schools generate more tickets than high schools?"
 * Deletes existing sheet if present for clean slate
 */
function setupLocationTypeComparisonSheet(ss) {
  deleteSheetIfExists(ss, 'LocationTypeComparison');
  const sheet = ss.insertSheet('LocationTypeComparison');

  // Headers
  const headers = ['Location Type', 'Locations', 'Open', 'Created (MTD)', 'Closed (MTD)', 'Avg per Location', 'Aged 30+', '% Aged', 'Last Refreshed'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Main formula - aggregates by LocationType (column O)
  const mainFormula =
    '=LET(' +
    'types, UNIQUE(FILTER(TicketData!O2:O, TicketData!O2:O<>"", TicketData!O2:O<>"LocationType")),' +
    'mtdStart, TEXT(DATE(YEAR(TODAY()),MONTH(TODAY()),1), "YYYY-MM-DD"),' +
    'mtdEnd, TEXT(DATE(YEAR(TODAY()),MONTH(TODAY())+1,1), "YYYY-MM-DD"),' +
    'col_a, types,' +
    'col_b, BYROW(types, LAMBDA(t, COUNTA(UNIQUE(FILTER(TicketData!N:N, TicketData!O:O=t))))),' +
    'col_c, BYROW(types, LAMBDA(t, COUNTIFS(TicketData!O:O, t, TicketData!I:I, "No"))),' +
    'col_d, BYROW(types, LAMBDA(t, COUNTIFS(TicketData!O:O, t, TicketData!E:E, ">="&mtdStart, TicketData!E:E, "<"&mtdEnd))),' +
    'col_e, BYROW(types, LAMBDA(t, COUNTIFS(TicketData!O:O, t, TicketData!H:H, ">="&mtdStart, TicketData!H:H, "<"&mtdEnd))),' +
    'col_f, BYROW(types, LAMBDA(t, LET(locs, COUNTA(UNIQUE(FILTER(TicketData!N:N, TicketData!O:O=t))), open, COUNTIFS(TicketData!O:O, t, TicketData!I:I, "No"), IF(locs>0, open/locs, 0)))),' +
    'col_g, BYROW(types, LAMBDA(t, COUNTIFS(TicketData!O:O, t, TicketData!I:I, "No", TicketData!R:R, ">=30"))),' +
    'col_h, BYROW(types, LAMBDA(t, LET(open, COUNTIFS(TicketData!O:O, t, TicketData!I:I, "No"), aged, COUNTIFS(TicketData!O:O, t, TicketData!I:I, "No", TicketData!R:R, ">=30"), IF(open>0, aged/open, 0)))),' +
    'data, HSTACK(col_a, col_b, col_c, col_d, col_e, col_f, col_g, col_h),' +
    'SORT(data, 3, FALSE))';

  sheet.getRange('A2').setValue(mainFormula);
  sheet.getRange('I2').setValue('=IFERROR(VLOOKUP("LAST_REFRESH", Config!A:B, 2, FALSE), "")');

  // Format header
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#388e3c')
    .setFontColor('white');

  // Format columns
  sheet.getRange('F:F').setNumberFormat('0.0');   // Avg per Location
  sheet.getRange('H:H').setNumberFormat('0.0%');  // % Aged

  // Column widths
  sheet.setColumnWidth(1, 150);  // Location Type
  sheet.setColumnWidth(9, 180);  // Last Refreshed

  sheet.setFrozenRows(1);

  // Add note
  sheet.getRange('A1').setNote(
    'Location Type Comparison\n\n' +
    'Question: "Which school types generate the most tickets?"\n\n' +
    'Use this to:\n' +
    '- Compare Elementary vs Middle vs High schools\n' +
    '- Identify systemic issues by school type\n' +
    '- Plan equipment refresh by school level\n' +
    '- Target training by school type\n\n' +
    '"Avg per Location" normalizes for different counts of schools.'
  );

  return true;
}

/**
 * Setup ReopenRate sheet
 * Question: "Are we truly resolving issues, or just closing tickets?"
 * Deletes existing sheet if present for clean slate
 */
function setupReopenRateSheet(ss) {
  deleteSheetIfExists(ss, 'ReopenRate');
  const sheet = ss.insertSheet('ReopenRate');

  // Headers
  const headers = ['Metric', 'Value', 'Description'];
  sheet.getRange(1, 1, 1, 3).setValues([headers]);

  // Explanation of reopened ticket detection
  const content = [
    ['', '', ''],
    ['--- Reopen Detection ---', '', ''],
    ['Currently Open with ClosedDate',
     '=COUNTIFS(TicketData!I:I, "No", TicketData!H:H, "<>")',
     'Tickets that were closed but are now open (reopened)'],
    ['Total Open Tickets',
     '=COUNTIFS(TicketData!I:I, "No")',
     'All currently open tickets'],
    ['Implied Reopen Rate',
     '=IF(B5>0, B4/B5, 0)',
     'Reopened / Total Open (lower is better, target <5%)'],
    ['', '', ''],
    ['--- Quality Indicators ---', '', ''],
    ['Tickets Closed Multiple Times',
     '=COUNTIFS(TicketData!I:I, "No", TicketData!H:H, "<>", TicketData!R:R, ">7")',
     'Reopened tickets open for 7+ days (significant rework)'],
    ['Avg Age of Reopened',
     '=IFERROR(AVERAGEIFS(TicketData!R:R, TicketData!I:I, "No", TicketData!H:H, "<>"), "N/A")',
     'How long reopened tickets have been open'],
    ['', '', ''],
    ['--- Reopened Tickets List ---', '', 'Sorted by age, oldest first'],
  ];

  sheet.getRange(2, 1, content.length, 3).setValues(content);

  // Add reopened ticket list headers at row 14
  const listHeaders = ['Ticket Number', 'Subject', 'Team', 'Original Close Date', 'Days Open', 'Status'];
  sheet.getRange(14, 1, 1, 6).setValues([listHeaders]);

  // Formula to list reopened tickets
  const listFormula =
    '=IFERROR(SORT(FILTER({TicketData!B2:B, LEFT(TicketData!C2:C,60), TicketData!L2:L, ' +
    'LEFT(TicketData!G2:G,10), TicketData!Q2:Q, TicketData!I2:I}, ' +
    '(TicketData!H2:H="No")*(TicketData!G2:G<>"")), 5, FALSE), "No reopened tickets found")';

  sheet.getRange('A15').setValue(listFormula);

  // Format headers
  sheet.getRange(1, 1, 1, 3)
    .setFontWeight('bold')
    .setBackground('#e65100')
    .setFontColor('white');

  sheet.getRange(14, 1, 1, 6)
    .setFontWeight('bold')
    .setBackground('#e65100')
    .setFontColor('white');

  // Format section headers
  sheet.getRange('A3').setFontWeight('bold').setBackground('#fff3e0');
  sheet.getRange('A8').setFontWeight('bold').setBackground('#fff3e0');
  sheet.getRange('A12').setFontWeight('bold').setBackground('#fff3e0');

  // Format percentage
  sheet.getRange('B6').setNumberFormat('0.0%');

  // Column widths
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 300);

  sheet.setFrozenRows(1);

  // Add note
  sheet.getRange('A1').setNote(
    'Reopen Rate Analysis\n\n' +
    'Question: "Are we truly resolving issues?"\n\n' +
    'Detection method: Tickets where IsClosed="No" but ClosedDate is not empty\n' +
    '(indicates ticket was closed then reopened)\n\n' +
    'Use this to:\n' +
    '- Identify quality issues in resolutions\n' +
    '- Find training opportunities\n' +
    '- Improve first-time fix rate\n\n' +
    'Target: <5% reopen rate\n' +
    'Industry benchmark: 3-7%'
  );

  return true;
}

/**
 * Setup FrequentRequesters sheet
 * Question: "Who generates the most tickets?"
 * Deletes existing sheet if present for clean slate
 */
function setupFrequentRequestersSheet(ss) {
  deleteSheetIfExists(ss, 'FrequentRequesters');
  const sheet = ss.insertSheet('FrequentRequesters');

  // Headers - includes sort controls
  const headers = ['Requester', 'Total Tickets', 'Open', 'Closed', 'Top Category', 'Avg Resolution (days)', 'Last Refreshed', 'Sort Col#', 'Desc?'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Main formula - top 50 requesters by RequesterName (column AC), with dynamic sorting
  const mainFormula =
    '=LET(' +
    'reqs, UNIQUE(FILTER(TicketData!AC2:AC, TicketData!AC2:AC<>"", TicketData!AC2:AC<>"RequesterName")),' +
    'col_a, reqs,' +
    'col_b, BYROW(reqs, LAMBDA(r, COUNTIF(TicketData!AC:AC, r))),' +
    'col_c, BYROW(reqs, LAMBDA(r, COUNTIFS(TicketData!AC:AC, r, TicketData!I:I, "No"))),' +
    'col_d, BYROW(reqs, LAMBDA(r, COUNTIFS(TicketData!AC:AC, r, TicketData!I:I, "Yes"))),' +
    'col_e, BYROW(reqs, LAMBDA(r, IFERROR(INDEX(SORT(UNIQUE(FILTER(TicketData!Y:Y, TicketData!AC:AC=r, TicketData!Y:Y<>"")), 1, FALSE), 1), ""))),' +
    'col_f, BYROW(reqs, LAMBDA(r, IFERROR(AVERAGEIFS(TicketData!R:R, TicketData!AC:AC, r, TicketData!I:I, "Yes"), "N/A"))),' +
    'sorted, SORT(HSTACK(col_a, col_b, col_c, col_d, col_e, col_f), $H$2, $I$2),' +
    'IFERROR(ARRAY_CONSTRAIN(sorted, 50, 6), sorted))';

  sheet.getRange('A2').setValue(mainFormula);
  sheet.getRange('G2').setValue('=IFERROR(VLOOKUP("LAST_REFRESH", Config!A:B, 2, FALSE), "")');

  // Default sort settings (column 2 = Total Tickets, descending)
  sheet.getRange('H2').setValue(2);
  sheet.getRange('I2').setValue('FALSE');

  // Format header
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#c2185b')
    .setFontColor('white');

  // Format columns
  sheet.getRange('F:F').setNumberFormat('0.0');  // Avg Resolution

  // Column widths
  sheet.setColumnWidth(1, 200);  // Requester
  sheet.setColumnWidth(5, 180);  // Top Category
  sheet.setColumnWidth(7, 180);  // Last Refreshed
  sheet.setColumnWidth(8, 80);   // Sort Col#
  sheet.setColumnWidth(9, 60);   // Desc?

  sheet.setFrozenRows(1);

  // Add data validation for sort column
  const sortColRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['1', '2', '3', '4', '5', '6'], true)
    .setHelpText('1=Requester, 2=Total, 3=Open, 4=Closed, 5=Category, 6=AvgRes')
    .build();
  sheet.getRange('H2').setDataValidation(sortColRule);

  // Add data validation for sort order
  const sortOrderRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['FALSE', 'TRUE'], true)
    .setHelpText('FALSE=Descending, TRUE=Ascending')
    .build();
  sheet.getRange('I2').setDataValidation(sortOrderRule);

  // Add notes
  sheet.getRange('A1').setNote(
    'Frequent Requesters Analysis (Top 50)\n\n' +
    'Question: "Who generates the most tickets?"\n\n' +
    'Use this to:\n' +
    '- Identify users who may need training\n' +
    '- Find equipment that needs replacement\n' +
    '- Discover power users who need dedicated support\n' +
    '- Plan proactive outreach\n\n' +
    'Use Sort Col# and Desc? to change sorting.'
  );
  sheet.getRange('H2').setNote('Sort column: 1=Requester, 2=Total, 3=Open, 4=Closed, 5=Category, 6=AvgRes');
  sheet.getRange('I2').setNote('FALSE=Descending (high to low), TRUE=Ascending (low to high)');

  return true;
}

/**
 * Setup ResponseDistribution sheet
 * Question: "How consistent is our response time?"
 * Deletes existing sheet if present for clean slate
 */
function setupResponseDistributionSheet(ss) {
  deleteSheetIfExists(ss, 'ResponseDistribution');
  const sheet = ss.insertSheet('ResponseDistribution');

  // Headers
  const headers = ['Metric', 'Response (hrs)', 'Resolution (hrs)', 'Description'];
  sheet.getRange(1, 1, 1, 4).setValues([headers]);

  // Stats content
  const content = [
    ['', '', '', ''],
    ['--- Response Time Distribution ---', '', '', ''],
    ['Average',
     '=IFERROR(AVERAGEIF(TicketData!AE:AE, ">0")/60, "N/A")',
     '=IFERROR(AVERAGEIF(TicketData!AH:AH, ">0")/60, "N/A")',
     'Mean time across all closed tickets with SLA data'],
    ['Median (50th %)',
     '=IFERROR(MEDIAN(FILTER(TicketData!AE:AE, TicketData!AE:AE>0))/60, "N/A")',
     '=IFERROR(MEDIAN(FILTER(TicketData!AH:AH, TicketData!AH:AH>0))/60, "N/A")',
     'Middle value - less affected by outliers'],
    ['25th Percentile',
     '=IFERROR(PERCENTILE(FILTER(TicketData!AE:AE, TicketData!AE:AE>0), 0.25)/60, "N/A")',
     '=IFERROR(PERCENTILE(FILTER(TicketData!AH:AH, TicketData!AH:AH>0), 0.25)/60, "N/A")',
     '25% of tickets faster than this'],
    ['75th Percentile',
     '=IFERROR(PERCENTILE(FILTER(TicketData!AE:AE, TicketData!AE:AE>0), 0.75)/60, "N/A")',
     '=IFERROR(PERCENTILE(FILTER(TicketData!AH:AH, TicketData!AH:AH>0), 0.75)/60, "N/A")',
     '75% of tickets faster than this'],
    ['90th Percentile',
     '=IFERROR(PERCENTILE(FILTER(TicketData!AE:AE, TicketData!AE:AE>0), 0.90)/60, "N/A")',
     '=IFERROR(PERCENTILE(FILTER(TicketData!AH:AH, TicketData!AH:AH>0), 0.90)/60, "N/A")',
     '90% of tickets faster than this (worst 10%)'],
    ['Minimum',
     '=IFERROR(MIN(FILTER(TicketData!AE:AE, TicketData!AE:AE>0))/60, "N/A")',
     '=IFERROR(MIN(FILTER(TicketData!AH:AH, TicketData!AH:AH>0))/60, "N/A")',
     'Fastest response/resolution'],
    ['Maximum',
     '=IFERROR(MAX(FILTER(TicketData!AE:AE, TicketData!AE:AE>0))/60, "N/A")',
     '=IFERROR(MAX(FILTER(TicketData!AH:AH, TicketData!AH:AH>0))/60, "N/A")',
     'Slowest response/resolution'],
    ['Std Deviation',
     '=IFERROR(STDEV(FILTER(TicketData!AE:AE, TicketData!AE:AE>0))/60, "N/A")',
     '=IFERROR(STDEV(FILTER(TicketData!AH:AH, TicketData!AH:AH>0))/60, "N/A")',
     'Consistency measure (lower = more consistent)'],
    ['', '', '', ''],
    ['--- Time Buckets (Response) ---', '', '', ''],
    ['< 1 hour',
     '=COUNTIFS(TicketData!AE:AE, ">0", TicketData!AE:AE, "<60")',
     '=COUNTIFS(TicketData!AE:AE, ">0", TicketData!AE:AE, "<60")/COUNTIF(TicketData!AE:AE, ">0")',
     'Count | % of total'],
    ['1-4 hours',
     '=COUNTIFS(TicketData!AE:AE, ">=60", TicketData!AE:AE, "<240")',
     '=COUNTIFS(TicketData!AE:AE, ">=60", TicketData!AE:AE, "<240")/COUNTIF(TicketData!AE:AE, ">0")',
     'Count | % of total'],
    ['4-8 hours',
     '=COUNTIFS(TicketData!AE:AE, ">=240", TicketData!AE:AE, "<480")',
     '=COUNTIFS(TicketData!AE:AE, ">=240", TicketData!AE:AE, "<480")/COUNTIF(TicketData!AE:AE, ">0")',
     'Count | % of total'],
    ['8-24 hours',
     '=COUNTIFS(TicketData!AE:AE, ">=480", TicketData!AE:AE, "<1440")',
     '=COUNTIFS(TicketData!AE:AE, ">=480", TicketData!AE:AE, "<1440")/COUNTIF(TicketData!AE:AE, ">0")',
     'Count | % of total'],
    ['> 24 hours',
     '=COUNTIF(TicketData!AE:AE, ">=1440")',
     '=COUNTIF(TicketData!AE:AE, ">=1440")/COUNTIF(TicketData!AE:AE, ">0")',
     'Count | % of total'],
  ];

  sheet.getRange(2, 1, content.length, 4).setValues(content);

  // Format header
  sheet.getRange(1, 1, 1, 4)
    .setFontWeight('bold')
    .setBackground('#5e35b1')
    .setFontColor('white');

  // Format section headers
  sheet.getRange('A3').setFontWeight('bold').setBackground('#ede7f6');
  sheet.getRange('A13').setFontWeight('bold').setBackground('#ede7f6');

  // Format numbers
  sheet.getRange('B4:C12').setNumberFormat('0.0');
  sheet.getRange('C14:C18').setNumberFormat('0.0%');

  // Column widths
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 130);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 280);

  sheet.setFrozenRows(1);

  // Add note
  sheet.getRange('A1').setNote(
    'Response Time Distribution Analysis\n\n' +
    'Question: "How consistent is our response time?"\n\n' +
    'Use this to:\n' +
    '- Understand variability in service\n' +
    '- Set realistic SLA targets\n' +
    '- Identify outliers affecting averages\n' +
    '- Compare median vs mean for skew\n\n' +
    'If median << average: outliers pulling average up\n' +
    'Low std deviation = consistent service'
  );

  return true;
}

/**
 * Setup TemporalPatterns sheet
 * Question: "When do tickets come in?"
 * Deletes existing sheet if present for clean slate
 */
function setupTemporalPatternsSheet(ss) {
  deleteSheetIfExists(ss, 'TemporalPatterns');
  const sheet = ss.insertSheet('TemporalPatterns');

  // Day of week section
  const dowHeaders = ['Day of Week', 'Tickets Created', '% of Total', 'Avg Created/Week'];
  sheet.getRange(1, 1, 1, 4).setValues([dowHeaders]);

  const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const dowData = [];
  for (let d = 0; d < 7; d++) {
    const rowNum = d + 2;
    dowData.push([
      days[d],
      // Count tickets created on this day of week (WEEKDAY returns 1=Sunday through 7=Saturday)
      `=COUNTIF(ARRAYFORMULA(WEEKDAY(DATEVALUE(LEFT(TicketData!E2:E, 10)))), ${d + 1})`,
      // % of total
      `=IF(SUM($B$2:$B$8)>0, B${rowNum}/SUM($B$2:$B$8), 0)`,
      // Avg per week (total / weeks of data)
      `=LET(dates, ARRAYFORMULA(DATEVALUE(LEFT(TicketData!E2:E,10))), totalDays, MAX(dates)-MIN(dates)+1, weeks, totalDays/7, IF(weeks>0, B${rowNum}/weeks, 0))`
    ]);
  }
  sheet.getRange(2, 1, dowData.length, 4).setValues(dowData);

  // Hour of day section
  sheet.getRange(10, 1).setValue('--- Hour of Day Distribution ---');
  const hourHeaders = ['Hour', 'Tickets', '% of Total', 'Time Range'];
  sheet.getRange(11, 1, 1, 4).setValues([hourHeaders]);

  const hourData = [];
  for (let h = 0; h < 24; h++) {
    const rowNum = h + 12;
    const nextHour = (h + 1) % 24;
    const timeRange = `${h.toString().padStart(2, '0')}:00 - ${nextHour.toString().padStart(2, '0')}:00`;
    hourData.push([
      h,
      // Count tickets created in this hour
      `=COUNTIF(ARRAYFORMULA(HOUR(TIMEVALUE(MID(TicketData!E2:E, 12, 8)))), ${h})`,
      // % of total
      `=IF(SUM($B$12:$B$35)>0, B${rowNum}/SUM($B$12:$B$35), 0)`,
      timeRange
    ]);
  }
  sheet.getRange(12, 1, hourData.length, 4).setValues(hourData);

  // Summary section
  sheet.getRange(37, 1).setValue('--- Key Insights ---');
  const summaryData = [
    ['Busiest Day', '=INDEX(A2:A8, MATCH(MAX(B2:B8), B2:B8, 0))', '', ''],
    ['Quietest Day', '=INDEX(A2:A8, MATCH(MIN(B2:B8), B2:B8, 0))', '', ''],
    ['Peak Hour', '=INDEX(D12:D35, MATCH(MAX(B12:B35), B12:B35, 0))', '', ''],
    ['Slowest Hour', '=INDEX(D12:D35, MATCH(MIN(FILTER(B12:B35, B12:B35>0)), FILTER(B12:B35, B12:B35>0), 0))', '', ''],
    ['Weekend %', '=IF(SUM(B2:B8)>0, (B2+B8)/SUM(B2:B8), 0)', '', '% of tickets created Sat+Sun'],
    ['After Hours %', '=IF(SUM(B12:B35)>0, (SUM(B12:B18)+SUM(B29:B35))/SUM(B12:B35), 0)', '', '% created before 7am or after 5pm'],
  ];
  sheet.getRange(38, 1, summaryData.length, 4).setValues(summaryData);

  // Format headers
  sheet.getRange(1, 1, 1, 4)
    .setFontWeight('bold')
    .setBackground('#00695c')
    .setFontColor('white');

  sheet.getRange(11, 1, 1, 4)
    .setFontWeight('bold')
    .setBackground('#00695c')
    .setFontColor('white');

  // Format section headers
  sheet.getRange('A10').setFontWeight('bold').setBackground('#e0f2f1');
  sheet.getRange('A37').setFontWeight('bold').setBackground('#e0f2f1');

  // Format percentages
  sheet.getRange('C2:C8').setNumberFormat('0.0%');
  sheet.getRange('C12:C35').setNumberFormat('0.0%');
  sheet.getRange('B42:B43').setNumberFormat('0.0%');

  // Format avg
  sheet.getRange('D2:D8').setNumberFormat('0.0');

  // Column widths
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(4, 150);

  sheet.setFrozenRows(1);

  // Add note
  sheet.getRange('A1').setNote(
    'Temporal Patterns Analysis\n\n' +
    'Question: "When do tickets come in?"\n\n' +
    'Use this to:\n' +
    '- Optimize staffing schedules\n' +
    '- Plan maintenance windows\n' +
    '- Set expectations for after-hours support\n' +
    '- Identify Monday morning surges\n\n' +
    'K-12 patterns:\n' +
    '- Monday morning spike (weekend issues reported)\n' +
    '- Friday afternoon submissions (low priority)\n' +
    '- Summer = lower volume opportunity for projects'
  );

  return true;
}

/**
 * Setup ResponseTrends sheet
 * Question: "Are our response times getting better and more consistent?"
 * Deletes existing sheet if present for clean slate
 */
function setupResponseTrendsSheet(ss) {
  deleteSheetIfExists(ss, 'ResponseTrends');
  const sheet = ss.insertSheet('ResponseTrends');

  // Headers
  const headers = [
    'Month', 'Year', 'Closed',
    'Avg Resp (hrs)', 'Median Resp (hrs)', '90th % Resp (hrs)',
    'Avg Resol (hrs)', 'Median Resol (hrs)', '90th % Resol (hrs)',
    '% Resp <4hrs', '% Resol <24hrs', 'Resp Std Dev'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Get month range from actual data (uses ClosedDate column H)
  const monthRange = getMonthRangeFromData(ss, 'H');

  const dataRows = [];
  for (const period of monthRange) {
    const y = period.year;
    const monthNum = period.monthNum;

    // Build date filter conditions used across formulas
    // dateFilter uses column H (ClosedDate) for COUNTIFS/AVERAGEIFS
    // arrayDateFilter uses column H (ClosedDate) for FILTER with array multiplication
    const dateFilter = `TicketData!H:H, ">="&TEXT(DATE(${y},${monthNum},1), "YYYY-MM-DD"), TicketData!H:H, "<"&TEXT(DATE(${y},${monthNum}+1,1), "YYYY-MM-DD")`;
    const arrayDateFilter = `(TicketData!H2:H>=DATE(${y},${monthNum},1))*(TicketData!H2:H<DATE(${y},${monthNum}+1,1))`;

    dataRows.push([
      period.monthName,
      y,
      // Closed count
      `=COUNTIFS(${dateFilter})`,
      // Avg Response (hrs) - column AE is ResponseActual in minutes
      `=IFERROR(AVERAGEIFS(TicketData!AE:AE, ${dateFilter}, TicketData!AE:AE, ">0")/60, "N/A")`,
      // Median Response (hrs) - column AE is ResponseActual in minutes
      `=IFERROR(MEDIAN(FILTER(TicketData!AE2:AE, ${arrayDateFilter}*(TicketData!AE2:AE>0)))/60, "N/A")`,
      // 90th percentile Response (hrs) - column AE is ResponseActual in minutes
      `=IFERROR(PERCENTILE(FILTER(TicketData!AE2:AE, ${arrayDateFilter}*(TicketData!AE2:AE>0)), 0.9)/60, "N/A")`,
      // Avg Resolution (hrs) - column AH is ResolutionActual in minutes
      `=IFERROR(AVERAGEIFS(TicketData!AH:AH, ${dateFilter}, TicketData!AH:AH, ">0")/60, "N/A")`,
      // Median Resolution (hrs) - column AH is ResolutionActual in minutes
      `=IFERROR(MEDIAN(FILTER(TicketData!AH2:AH, ${arrayDateFilter}*(TicketData!AH2:AH>0)))/60, "N/A")`,
      // 90th percentile Resolution (hrs) - column AH is ResolutionActual in minutes
      `=IFERROR(PERCENTILE(FILTER(TicketData!AH2:AH, ${arrayDateFilter}*(TicketData!AH2:AH>0)), 0.9)/60, "N/A")`,
      // % Responded within 4 hours (240 min)
      `=IFERROR(COUNTIFS(${dateFilter}, TicketData!AE:AE, ">0", TicketData!AE:AE, "<=240")/COUNTIFS(${dateFilter}, TicketData!AE:AE, ">0"), "N/A")`,
      // % Resolved within 24 hours (1440 min)
      `=IFERROR(COUNTIFS(${dateFilter}, TicketData!AH:AH, ">0", TicketData!AH:AH, "<=1440")/COUNTIFS(${dateFilter}, TicketData!AH:AH, ">0"), "N/A")`,
      // Response Std Dev (consistency measure) - column AE is ResponseActual
      `=IFERROR(STDEV(FILTER(TicketData!AE2:AE, ${arrayDateFilter}*(TicketData!AE2:AE>0)))/60, "N/A")`
    ]);
  }

  sheet.getRange(2, 1, dataRows.length, 12).setValues(dataRows);
  const lastDataRow = dataRows.length + 1;

  // Format header
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#5e35b1')
    .setFontColor('white');

  // Group headers with colors
  sheet.getRange(1, 4, 1, 3).setBackground('#7e57c2');  // Response metrics
  sheet.getRange(1, 7, 1, 3).setBackground('#9575cd');  // Resolution metrics
  sheet.getRange(1, 10, 1, 3).setBackground('#b39ddb'); // Targets/consistency

  // Format numbers
  sheet.getRange('D:I').setNumberFormat('0.0');    // Hours
  sheet.getRange('J:K').setNumberFormat('0.0%');   // Percentages
  sheet.getRange('L:L').setNumberFormat('0.0');    // Std Dev

  // Column widths
  sheet.setColumnWidth(1, 90);   // Month
  sheet.setColumnWidth(2, 50);   // Year
  sheet.setColumnWidth(3, 60);   // Closed
  for (let c = 4; c <= 12; c++) {
    sheet.setColumnWidth(c, 95);
  }

  sheet.setFrozenRows(1);

  // Conditional formatting - highlight improving trends (lower is better for response times)
  // Green for low values in response/resolution columns
  const timeRange = sheet.getRange(2, 4, dataRows.length, 6); // D2:I{lastRow}
  const lowRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(4)
    .setBackground('#e8f5e9')
    .setRanges([timeRange])
    .build();
  const highRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(24)
    .setBackground('#ffebee')
    .setRanges([timeRange])
    .build();
  sheet.setConditionalFormatRules([lowRule, highRule]);

  // Add note
  sheet.getRange('A1').setNote(
    'Response Time Trends Analysis\n\n' +
    'Question: "Are our response times getting better and more consistent?"\n\n' +
    'Key metrics:\n' +
    '- Avg vs Median: If avg >> median, outliers are skewing results\n' +
    '- 90th %: Worst 10% of tickets - shows your "bad days"\n' +
    '- Std Dev: Lower = more consistent service\n' +
    '- % <4hrs / <24hrs: Target achievement rates\n\n' +
    'What to look for:\n' +
    '- Downward trend in response/resolution times\n' +
    '- Decreasing std dev (more predictable)\n' +
    '- Increasing % meeting targets\n' +
    '- Gap between avg and median narrowing\n\n' +
    'Green = <4 hrs, Red = >24 hrs'
  );

  return true;
}

/**
 * Setup QueueTimeAnalysis sheet
 * Question: "How long do tickets wait before being picked up?"
 * Analyzes the time between CreatedDate and StartedDate
 *
 * Layout:
 * - Columns A-D: Summary stats, Distribution, Monthly Trend, Waiting Tickets (stacked vertically)
 *
 * Note: Queue Time by Team has been moved to its own sheet (QueueTimeByTeam)
 *
 * Deletes existing sheet if present for clean slate
 */
function setupQueueTimeAnalysisSheet(ss) {
  deleteSheetIfExists(ss, 'QueueTimeAnalysis');
  const sheet = ss.insertSheet('QueueTimeAnalysis');

  // ============================================================
  // SECTION 1: Overall Summary Stats (Columns A-C)
  // ============================================================
  const summaryHeaders = ['Metric', 'Value', 'Description'];
  sheet.getRange(1, 1, 1, 3).setValues([summaryHeaders]);

  // Queue time formula: (StartedDate - CreatedDate) in hours
  // Uses simple date subtraction which works with Date objects directly
  const qtimeFilter = '(TicketData!F2:F-TicketData!E2:E)*24, (TicketData!F2:F<>"")*(TicketData!E2:E<>"")';

  const summaryData = [
    ['', '', ''],
    ['--- Queue Time Summary ---', '', ''],
    ['Total Tickets with Queue Time',
     '=COUNTIFS(TicketData!F:F, "<>", TicketData!E:E, "<>")',
     'Tickets where both CreatedDate and StartedDate exist'],
    ['Tickets Never Started',
     '=COUNTIFS(TicketData!F:F, "", TicketData!I:I, "No")',
     'Open tickets with no StartedDate (still in queue)'],
    ['Currently in Queue',
     '=COUNTIFS(TicketData!F:F, "", TicketData!I:I, "No")',
     'Open tickets waiting to be picked up'],
    ['', '', ''],
    ['--- Queue Time Stats (hours) ---', '', ''],
    ['Average Queue Time',
     `=IFERROR(AVERAGE(FILTER(${qtimeFilter})), "N/A")`,
     'Mean time tickets wait before being picked up'],
    ['Median Queue Time',
     `=IFERROR(MEDIAN(FILTER(${qtimeFilter})), "N/A")`,
     'Middle value - less affected by outliers'],
    ['90th Percentile',
     `=IFERROR(PERCENTILE(FILTER(${qtimeFilter}), 0.9), "N/A")`,
     '90% of tickets picked up faster than this'],
    ['Maximum Queue Time',
     `=IFERROR(MAX(FILTER(${qtimeFilter})), "N/A")`,
     'Longest wait time'],
    ['Std Deviation',
     `=IFERROR(STDEV(FILTER(${qtimeFilter})), "N/A")`,
     'Consistency measure (lower = more consistent)'],
  ];

  sheet.getRange(2, 1, summaryData.length, 3).setValues(summaryData);

  // ============================================================
  // LEFT SIDE - SECTION 2: Queue Time Distribution Buckets
  // ============================================================
  const distRow = summaryData.length + 3;
  sheet.getRange(distRow, 1).setValue('--- Queue Time Distribution ---');

  const distHeaders = ['Time Bucket', 'Count', '% of Total', 'Cumulative %'];
  sheet.getRange(distRow + 1, 1, 1, 4).setValues([distHeaders]);

  // Queue time in hours using simple date subtraction (works with Date objects)
  const qtimeLet = 'LET(qtime, FILTER((TicketData!F2:F-TicketData!E2:E)*24, (TicketData!F2:F<>"")*(TicketData!E2:E<>"")),';

  const buckets = [
    ['< 15 minutes', `=IFERROR(${qtimeLet} COUNTIF(qtime, "<0.25")), 0)`, '', ''],
    ['15-60 minutes', `=IFERROR(${qtimeLet} SUMPRODUCT((qtime>=0.25)*(qtime<1))), 0)`, '', ''],
    ['1-4 hours', `=IFERROR(${qtimeLet} SUMPRODUCT((qtime>=1)*(qtime<4))), 0)`, '', ''],
    ['4-8 hours', `=IFERROR(${qtimeLet} SUMPRODUCT((qtime>=4)*(qtime<8))), 0)`, '', ''],
    ['8-24 hours', `=IFERROR(${qtimeLet} SUMPRODUCT((qtime>=8)*(qtime<24))), 0)`, '', ''],
    ['1-3 days', `=IFERROR(${qtimeLet} SUMPRODUCT((qtime>=24)*(qtime<72))), 0)`, '', ''],
    ['> 3 days', `=IFERROR(${qtimeLet} COUNTIF(qtime, ">=72")), 0)`, '', ''],
  ];

  // Add percentage and cumulative formulas
  for (let i = 0; i < buckets.length; i++) {
    const rowNum = distRow + 2 + i;
    const totalRef = `SUM($B$${distRow + 2}:$B$${distRow + 2 + buckets.length - 1})`;
    buckets[i][2] = `=IF(${totalRef}>0, B${rowNum}/${totalRef}, 0)`;
    buckets[i][3] = `=SUM($C$${distRow + 2}:C${rowNum})`;
  }

  sheet.getRange(distRow + 2, 1, buckets.length, 4).setValues(buckets);

  // ============================================================
  // LEFT SIDE - SECTION 3: Waiting Tickets
  // ============================================================
  // Note: Monthly Queue Time Trend is now in the separate QueueTimeTrend sheet
  const waitingRow = distRow + buckets.length + 4;
  sheet.getRange(waitingRow, 1).setValue('--- Tickets Still Waiting (Never Started) ---');

  const waitingHeaders = ['Ticket #', 'Subject', 'Team', 'Days'];
  sheet.getRange(waitingRow + 1, 1, 1, 4).setValues([waitingHeaders]);

  // List of open tickets with no StartedDate, sorted by age (constrained to fit left side)
  const waitingFormula =
    '=IFERROR(ARRAY_CONSTRAIN(SORT(FILTER({TicketData!B2:B, LEFT(TicketData!C2:C,30), TicketData!L2:L, TicketData!R2:R}, ' +
    '(TicketData!F2:F="")*(TicketData!I2:I="No")), 4, FALSE), 25, 4), "No tickets waiting")';

  sheet.getRange(waitingRow + 2, 1).setValue(waitingFormula);

  // ============================================================
  // FORMATTING
  // ============================================================

  // Format left side main header
  sheet.getRange(1, 1, 1, 3)
    .setFontWeight('bold')
    .setBackground('#0277bd')
    .setFontColor('white');

  // Format left side section headers
  sheet.getRange('A3').setFontWeight('bold').setBackground('#e1f5fe');
  sheet.getRange('A8').setFontWeight('bold').setBackground('#e1f5fe');
  sheet.getRange(distRow, 1).setFontWeight('bold').setBackground('#e1f5fe');
  sheet.getRange(waitingRow, 1).setFontWeight('bold').setBackground('#e1f5fe');

  // Format distribution header
  sheet.getRange(distRow + 1, 1, 1, 4)
    .setFontWeight('bold')
    .setBackground('#0277bd')
    .setFontColor('white');

  // Format waiting header
  sheet.getRange(waitingRow + 1, 1, 1, 4)
    .setFontWeight('bold')
    .setBackground('#0277bd')
    .setFontColor('white');

  // Format numbers
  sheet.getRange('B9:B13').setNumberFormat('0.0');  // Summary stats
  sheet.getRange(distRow + 2, 3, buckets.length, 2).setNumberFormat('0.0%');  // Distribution %

  // Column widths
  sheet.setColumnWidth(1, 180);  // Metric/bucket name
  sheet.setColumnWidth(2, 100);  // Value
  sheet.setColumnWidth(3, 200);  // Description or %
  sheet.setColumnWidth(4, 100);  // Cumulative %

  sheet.setFrozenRows(1);

  // Add note
  sheet.getRange('A1').setNote(
    'Queue Time Analysis\n\n' +
    'Question: "How long do tickets wait before being picked up?"\n\n' +
    'Queue Time = StartedDate - CreatedDate\n' +
    'This measures the time between ticket submission and when\n' +
    'an agent first starts working on it.\n\n' +
    'Use this to:\n' +
    '- Identify staffing gaps (long queues = understaffed)\n' +
    '- Set expectations for initial response\n' +
    '- Find tickets stuck in queue\n\n' +
    'Related sheets:\n' +
    '- QueueTimeByTeam: Team-by-team breakdown\n' +
    '- QueueTimeTrend: Monthly trends over time\n\n' +
    'Targets:\n' +
    '- < 1 hour: Excellent responsiveness\n' +
    '- 1-4 hours: Good for most tickets\n' +
    '- > 8 hours: May indicate capacity issues'
  );

  return true;
}

/**
 * Setup QueueTimeByTeam sheet
 * Question: "How does queue time vary by team?"
 * Shows queue time metrics broken down by team with diagnostic info
 * Deletes existing sheet if present for clean slate
 */
function setupQueueTimeByTeamSheet(ss) {
  deleteSheetIfExists(ss, 'QueueTimeByTeam');
  const sheet = ss.insertSheet('QueueTimeByTeam');

  // ============================================================
  // Headers and Sort Controls
  // ============================================================
  const headers = ['Team', 'Avg Queue (hrs)', 'Tickets'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Sort controls in column E-F
  sheet.getRange(1, 5).setValue('Sort by:');
  sheet.getRange(1, 6).setValue(2);  // Default: sort by column 2 (Avg Queue)
  sheet.getRange(2, 5).setValue('Ascending:');
  sheet.getRange(2, 6).setValue('TRUE');  // Default: ascending (fastest first)

  // ============================================================
  // Main Data - QUERY wrapped in SORT for user-controlled sorting
  // ============================================================
  // QUERY groups by Team and calculates AVG and COUNT
  // SORT wrapper uses E1 (column) and E2 (asc/desc) for dynamic sorting
  const queryFormula = `=SORT(QUERY({TicketData!L2:L, ARRAYFORMULA(IFERROR((TicketData!F2:F-TicketData!E2:E)*24, -1))},"SELECT Col1, AVG(Col2), COUNT(Col2) WHERE Col1 <> '' AND Col2 >= 0 GROUP BY Col1 LABEL Col1 '', AVG(Col2) '', COUNT(Col2) ''", 0), $F$1, $F$2)`;

  sheet.getRange(2, 1).setValue(queryFormula);

  // ============================================================
  // Sort Control Validation
  // ============================================================
  const sortColCell = sheet.getRange(1, 6);
  const sortOrderCell = sheet.getRange(2, 6);

  // Data validation for sort column
  const sortColRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['1', '2', '3'], true)
    .setHelpText('1=Team, 2=Avg Queue, 3=Tickets')
    .build();
  sortColCell.setDataValidation(sortColRule);

  // Data validation for sort order
  const sortOrderRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TRUE', 'FALSE'], true)
    .setHelpText('TRUE=Ascending, FALSE=Descending')
    .build();
  sortOrderCell.setDataValidation(sortOrderRule);

  // Notes for sort controls
  sortColCell.setNote('Sort column:\n1 = Team (alphabetical)\n2 = Avg Queue (hours)\n3 = Tickets (count)');
  sortOrderCell.setNote('TRUE = Ascending (A-Z, low to high)\nFALSE = Descending (Z-A, high to low)');

  // ============================================================
  // FORMATTING
  // ============================================================

  // Header row
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1976d2')
    .setFontColor('white');

  // Sort control labels
  sheet.getRange(1, 5, 2, 1)
    .setFontWeight('bold')
    .setHorizontalAlignment('right');
  sheet.getRange(1, 6, 2, 1)
    .setBackground('#e3f2fd')
    .setHorizontalAlignment('center');

  // Number formatting
  sheet.getRange(2, 2, 150, 1).setNumberFormat('0.0');    // Avg hours
  sheet.getRange(2, 3, 150, 1).setNumberFormat('#,##0'); // Ticket count

  // Column widths
  sheet.setColumnWidth(1, 200);  // Team name
  sheet.setColumnWidth(2, 120);  // Avg Queue
  sheet.setColumnWidth(3, 80);   // Tickets
  sheet.setColumnWidth(4, 30);   // Spacer
  sheet.setColumnWidth(5, 80);   // Sort label
  sheet.setColumnWidth(6, 60);   // Sort value

  sheet.setFrozenRows(1);

  // Conditional formatting for queue times
  const avgRange = sheet.getRange(2, 2, 150, 1);
  const fastRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(1)
    .setBackground('#c8e6c9')
    .setRanges([avgRange])
    .build();
  const mediumRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(1, 4)
    .setBackground('#fff9c4')
    .setRanges([avgRange])
    .build();
  const slowRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(8)
    .setBackground('#ffcdd2')
    .setRanges([avgRange])
    .build();
  sheet.setConditionalFormatRules([fastRule, mediumRule, slowRule]);

  // Add explanatory note
  sheet.getRange('A1').setNote(
    'Queue Time by Team\n\n' +
    'Queue Time = StartedDate - CreatedDate (in hours)\n\n' +
    'Use the Sort controls (columns E-F) to change sorting:\n' +
    '- Sort by: 1=Team, 2=Avg Queue, 3=Tickets\n' +
    '- Ascending: TRUE=low-to-high, FALSE=high-to-low\n\n' +
    'Color coding:\n' +
    '- Green: < 1 hour (excellent)\n' +
    '- Yellow: 1-4 hours (good)\n' +
    '- Red: > 8 hours (needs attention)\n\n' +
    'For Median (specific team):\n' +
    '=MEDIAN(FILTER((TicketData!F:F-TicketData!E:E)*24, TicketData!L:L="TeamName", TicketData!F:F<>""))'
  );

  return true;
}

/**
 * Setup QueueTimeTrend sheet
 * Question: "How is queue time trending over time?"
 * Shows monthly queue time trends with avg, median, and 90th percentile
 * Deletes existing sheet if present for clean slate
 */
function setupQueueTimeTrendSheet(ss) {
  deleteSheetIfExists(ss, 'QueueTimeTrend');
  const sheet = ss.insertSheet('QueueTimeTrend');

  // ============================================================
  // Headers and Sort Controls
  // ============================================================
  const headers = ['Month-Year', 'Tickets', 'Avg Queue (hrs)', 'Median (hrs)', '90th Pctl (hrs)'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Sort controls in column G-H
  sheet.getRange(1, 7).setValue('Sort by:');
  sheet.getRange(1, 8).setValue(1);  // Default: sort by column 1 (Month-Year)
  sheet.getRange(2, 7).setValue('Ascending:');
  sheet.getRange(2, 8).setValue('TRUE');  // Default: ascending (oldest first)

  // ============================================================
  // Main Data - QUERY for month grouping with SORT wrapper
  // ============================================================
  // QUERY groups by Month-Year (TEXT of StartedDate) and calculates COUNT and AVG
  // Median and 90th percentile added via separate formulas per row
  const queryFormula = `=SORT(QUERY({ARRAYFORMULA(IF(TicketData!F2:F<>"", TEXT(TicketData!F2:F, "YYYY-MM"), "")), ARRAYFORMULA(IFERROR((TicketData!F2:F-TicketData!E2:E)*24, -1))}, "SELECT Col1, COUNT(Col2), AVG(Col2) WHERE Col1 <> '' AND Col2 >= 0 GROUP BY Col1 LABEL Col1 '', COUNT(Col2) '', AVG(Col2) ''", 0), $H$1, $H$2)`;

  sheet.getRange(2, 1).setValue(queryFormula);

  // ============================================================
  // Median and 90th Percentile Formulas (columns D and E)
  // ============================================================
  // These reference the month-year in column A to calculate per-month stats
  // Uses ARRAYFORMULA to fill down automatically when QUERY expands
  const medianFormula = `=ARRAYFORMULA(IF(A2:A<>"", IFERROR(MAP(A2:A, LAMBDA(m, MEDIAN(FILTER((TicketData!F:F-TicketData!E:E)*24, TEXT(TicketData!F:F, "YYYY-MM")=m, TicketData!F:F<>"", TicketData!E:E<>"")))), ""), ""))`;
  const pctl90Formula = `=ARRAYFORMULA(IF(A2:A<>"", IFERROR(MAP(A2:A, LAMBDA(m, PERCENTILE(FILTER((TicketData!F:F-TicketData!E:E)*24, TEXT(TicketData!F:F, "YYYY-MM")=m, TicketData!F:F<>"", TicketData!E:E<>""), 0.9))), ""), ""))`;

  sheet.getRange(2, 4).setValue(medianFormula);
  sheet.getRange(2, 5).setValue(pctl90Formula);

  // ============================================================
  // Sort Control Validation
  // ============================================================
  const sortColCell = sheet.getRange(1, 8);
  const sortOrderCell = sheet.getRange(2, 8);

  // Data validation for sort column
  const sortColRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['1', '2', '3'], true)
    .setHelpText('1=Month-Year, 2=Tickets, 3=Avg Queue')
    .build();
  sortColCell.setDataValidation(sortColRule);

  // Data validation for sort order
  const sortOrderRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TRUE', 'FALSE'], true)
    .setHelpText('TRUE=Ascending, FALSE=Descending')
    .build();
  sortOrderCell.setDataValidation(sortOrderRule);

  // Notes for sort controls
  sortColCell.setNote('Sort column:\n1 = Month-Year (chronological)\n2 = Tickets (count)\n3 = Avg Queue (hours)');
  sortOrderCell.setNote('TRUE = Ascending (oldest first, low to high)\nFALSE = Descending (newest first, high to low)');

  // ============================================================
  // FORMATTING
  // ============================================================

  // Header row
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1976d2')
    .setFontColor('white');

  // Sort control labels
  sheet.getRange(1, 7, 2, 1)
    .setFontWeight('bold')
    .setHorizontalAlignment('right');
  sheet.getRange(1, 8, 2, 1)
    .setBackground('#e3f2fd')
    .setHorizontalAlignment('center');

  // Number formatting
  sheet.getRange(2, 2, 150, 1).setNumberFormat('#,##0');     // Ticket count
  sheet.getRange(2, 3, 150, 3).setNumberFormat('0.0');       // Queue time columns

  // Column widths
  sheet.setColumnWidth(1, 100);  // Month-Year
  sheet.setColumnWidth(2, 80);   // Tickets
  sheet.setColumnWidth(3, 120);  // Avg Queue
  sheet.setColumnWidth(4, 100);  // Median
  sheet.setColumnWidth(5, 110);  // 90th Pctl
  sheet.setColumnWidth(6, 30);   // Spacer
  sheet.setColumnWidth(7, 80);   // Sort label
  sheet.setColumnWidth(8, 60);   // Sort value

  sheet.setFrozenRows(1);

  // Conditional formatting for queue times (columns C-E)
  const queueRange = sheet.getRange(2, 3, 150, 3);
  const fastRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(1)
    .setBackground('#c8e6c9')
    .setRanges([queueRange])
    .build();
  const mediumRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(1, 4)
    .setBackground('#fff9c4')
    .setRanges([queueRange])
    .build();
  const slowRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(8)
    .setBackground('#ffcdd2')
    .setRanges([queueRange])
    .build();
  sheet.setConditionalFormatRules([fastRule, mediumRule, slowRule]);

  // Add explanatory note
  sheet.getRange('A1').setNote(
    'Queue Time Trend\n\n' +
    'Question: "How is queue time trending over time?"\n\n' +
    'Queue Time = StartedDate - CreatedDate (in hours)\n\n' +
    'Metrics:\n' +
    '- Avg: Mean queue time (affected by outliers)\n' +
    '- Median: Middle value (better for typical experience)\n' +
    '- 90th Pctl: 90% of tickets faster than this\n\n' +
    'Use the Sort controls (columns G-H) to change sorting:\n' +
    '- Sort by: 1=Month, 2=Tickets, 3=Avg\n' +
    '- Ascending: TRUE=oldest first, FALSE=newest first\n\n' +
    'Color coding:\n' +
    '- Green: < 1 hour (excellent)\n' +
    '- Yellow: 1-4 hours (good)\n' +
    '- Red: > 8 hours (needs attention)\n\n' +
    'Tip: Set Sort by=1 and Ascending=FALSE to see recent months first.'
  );

  return true;
}

// ============================================================================
// REGENERATE MONTHLY ANALYTICS SHEETS
// ============================================================================

/**
 * Regenerate all monthly-based analytics sheets using current school year config.
 *
 * This is useful when:
 * - School year config has changed
 * - You want all monthly sheets to show the full school year range
 * - Analytics sheets were created before data was loaded
 *
 * Regenerates the following sheets (if they exist):
 * - MonthlyVolume
 * - SLACompliance
 * - PerformanceTrends
 * - FirstContactResolution
 * - ResponseTrends
 * - SeasonalComparison
 */
function regenerateMonthlyAnalyticsSheets() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Verify school year is configured
  const config = getConfig();
  if (!config.schoolYear) {
    ui.alert('Configuration Missing',
      'No SCHOOL_YEAR configured.\n\n' +
      'Please set SCHOOL_YEAR in the Config sheet first (e.g., "2025-2026").',
      ui.ButtonSet.OK);
    return;
  }

  const dates = getSchoolYearDates(config);
  if (!dates) {
    ui.alert('Invalid Configuration',
      'Could not parse school year dates.\n\n' +
      'Please verify SCHOOL_YEAR format is "YYYY-YYYY" (e.g., "2025-2026").',
      ui.ButtonSet.OK);
    return;
  }

  // Format the date range for display
  const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
                      'July', 'August', 'September', 'October', 'November', 'December'];
  const startMonthName = monthNames[dates.startDate.getMonth()];
  const endMonthName = monthNames[dates.endDate.getMonth()];
  const dateRangeStr = `${startMonthName} ${dates.startYear} - ${endMonthName} ${dates.endYear}`;

  // List of monthly sheets and their setup functions
  const monthlySheets = [
    { name: 'MonthlyVolume', setup: setupMonthlyVolumeSheet },
    { name: 'SLACompliance', setup: setupSLAComplianceSheet },
    { name: 'PerformanceTrends', setup: setupPerformanceTrendsSheet },
    { name: 'FirstContactResolution', setup: setupFirstContactResolutionSheet },
    { name: 'ResponseTrends', setup: setupResponseTrendsSheet },
    { name: 'SeasonalComparison', setup: setupSeasonalComparisonSheet }
  ];

  // Find which sheets exist
  const existingSheets = monthlySheets.filter(s => ss.getSheetByName(s.name));

  if (existingSheets.length === 0) {
    ui.alert('No Monthly Sheets Found',
      'No monthly analytics sheets exist to regenerate.\n\n' +
      'Use "Add Analytics Sheet" menu to create them first.',
      ui.ButtonSet.OK);
    return;
  }

  // Confirm with user
  const sheetList = existingSheets.map(s => '• ' + s.name).join('\n');
  const response = ui.alert('Regenerate Monthly Analytics Sheets',
    `This will recreate the following ${existingSheets.length} sheet(s) using school year ${config.schoolYear}:\n\n` +
    sheetList + '\n\n' +
    `Date range: ${dateRangeStr}\n\n` +
    'Continue?',
    ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) {
    return;
  }

  // Regenerate each existing sheet
  let regenerated = 0;
  let errors = [];

  for (const sheet of existingSheets) {
    try {
      sheet.setup(ss);
      regenerated++;
    } catch (e) {
      errors.push(`${sheet.name}: ${e.message}`);
    }
  }

  // Report results
  if (errors.length === 0) {
    ui.alert('Regeneration Complete',
      `Successfully regenerated ${regenerated} monthly analytics sheet(s).\n\n` +
      `Date range: ${dateRangeStr}`,
      ui.ButtonSet.OK);
  } else {
    ui.alert('Regeneration Complete with Errors',
      `Regenerated: ${regenerated}\nErrors: ${errors.length}\n\n` +
      errors.join('\n'),
      ui.ButtonSet.OK);
  }
}
