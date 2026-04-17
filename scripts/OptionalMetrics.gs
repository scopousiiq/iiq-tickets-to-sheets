/**
 * Analytics Sheets - All Available Metric Sheets
 *
 * All analytics sheets can be added individually via the iiQ Data > Add Analytics Sheet menu.
 * Adding a sheet will DELETE and RECREATE it if it already exists (clean slate).
 *
 * DEFAULT SHEETS (created by Setup Spreadsheet):
 * - MonthlyVolume: Ticket creation/closure by month
 * - BacklogAging: Current open ticket distribution by age
 * - ResolutionAging: Days-to-resolve distribution for closed tickets
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
 * - DeviceReliability: Which device models generate the most tickets?
 * - FrequentFlyers: Users and devices with recurring issues (configurable filters)
 * - BacklogAgingByFA: Open ticket aging by Functional Area
 * - BacklogAgingByTeam: Open ticket aging by Team
 * - BacklogAgingByLocationType: Open ticket aging by Location Type
 * - BacklogAgingByPriority: Open ticket aging by Priority
 * - ResolutionAging: Days-to-resolve distribution for closed tickets
 * - MonthlyVolumeByFA: Monthly Created/Closed crossed with Functional Area
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

/**
 * Add Monthly Volume by Functional Area sheet (deletes and recreates if exists)
 */
function addMonthlyVolumeByFASheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupMonthlyVolumeByFASheet(ss);
  SpreadsheetApp.getUi().alert('Created', 'MonthlyVolumeByFA sheet has been created.\n\nNote: Requires Teams sheet with FunctionalArea column filled in.', SpreadsheetApp.getUi().ButtonSet.OK);
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

/**
 * Add Backlog Aging by Functional Area sheet (deletes and recreates if exists)
 */
function addBacklogAgingByFASheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupBacklogAgingByFASheet(ss);
  SpreadsheetApp.getUi().alert('Created', 'BacklogAgingByFA sheet has been created.\n\nNote: Requires Teams sheet with FunctionalArea column filled in.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Add Backlog Aging by Team sheet (deletes and recreates if exists)
 */
function addBacklogAgingByTeamSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupBacklogAgingByTeamSheet(ss);
  SpreadsheetApp.getUi().alert('Created', 'BacklogAgingByTeam sheet has been created.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Add Backlog Aging by Location Type sheet (deletes and recreates if exists)
 */
function addBacklogAgingByLocationTypeSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupBacklogAgingByLocationTypeSheet(ss);
  SpreadsheetApp.getUi().alert('Created', 'BacklogAgingByLocationType sheet has been created.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Add Backlog Aging by Priority sheet (deletes and recreates if exists)
 */
function addBacklogAgingByPrioritySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupBacklogAgingByPrioritySheet(ss);
  SpreadsheetApp.getUi().alert('Created', 'BacklogAgingByPriority sheet has been created.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Add Resolution Aging sheet (deletes and recreates if exists)
 */
function addResolutionAgingSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupResolutionAgingSheet(ss);
  SpreadsheetApp.getUi().alert('Created', 'ResolutionAging sheet has been created.\n\nShows days-to-resolve distribution for closed tickets.', SpreadsheetApp.getUi().ButtonSet.OK);
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

// --- Device ---

/**
 * Add Devices by Role sheet (deletes and recreates if exists)
 */
function addDevicesByRoleSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupDevicesByRoleSheet(ss);
  SpreadsheetApp.getUi().alert(
    'Created',
    'DevicesByRole sheet has been created.\n\n' +
    'Use the "User Role" dropdown (K2) to filter by Student, Staff, Agent, or All.\n\n' +
    'Note: Requires a Full Reload to populate the RequesterRole column on existing rows.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Add Device Reliability sheet (deletes and recreates if exists)
 */
function addDeviceReliabilitySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupDeviceReliabilitySheet(ss);
  SpreadsheetApp.getUi().alert('Created',
    'DeviceReliability sheet has been created.\n\n' +
    'Groups tickets by device model with issue category breakdown.\n' +
    'Select a model from K2 dropdown to see issue type percentages.',
    SpreadsheetApp.getUi().ButtonSet.OK);
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
    'mtdStart, DATE(YEAR(TODAY()),MONTH(TODAY()),1),' +
    'mtdEnd, DATE(YEAR(TODAY()),MONTH(TODAY())+1,1),' +
    'col_a, cats,' +
    'col_b, BYROW(cats, LAMBDA(c, COUNTIFS(TicketData!Y:Y, c, TicketData!I:I, "Open"))),' +
    'col_c, BYROW(cats, LAMBDA(c, COUNTIFS(TicketData!Y:Y, c, TicketData!E:E, ">="&mtdStart, TicketData!E:E, "<"&mtdEnd))),' +
    'col_d, BYROW(cats, LAMBDA(c, COUNTIFS(TicketData!Y:Y, c, TicketData!H:H, ">="&mtdStart, TicketData!H:H, "<"&mtdEnd))),' +
    'col_e, BYROW(cats, LAMBDA(c, IFERROR(AVERAGEIFS(TicketData!R:R, TicketData!Y:Y, c, TicketData!I:I, "Closed"), "N/A"))),' +
    'col_f, BYROW(cats, LAMBDA(c, LET(total, COUNTIFS(TicketData!Y:Y, c, TicketData!I:I, "Closed"), breached, COUNTIFS(TicketData!Y:Y, c, TicketData!I:I, "Closed", TicketData!AF:AF, 1)+COUNTIFS(TicketData!Y:Y, c, TicketData!I:I, "Closed", TicketData!AI:AI, 1), IF(total>0, breached/total, "N/A")))),' +
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
    'mtdStart, DATE(YEAR(TODAY()),MONTH(TODAY()),1),' +
    'mtdEnd, DATE(YEAR(TODAY()),MONTH(TODAY())+1,1),' +
    'col_a, pris,' +
    'col_b, BYROW(pris, LAMBDA(p, COUNTIFS(TicketData!S:S, p, TicketData!I:I, "Open"))),' +
    'col_c, BYROW(pris, LAMBDA(p, COUNTIFS(TicketData!S:S, p, TicketData!E:E, ">="&mtdStart, TicketData!E:E, "<"&mtdEnd))),' +
    'col_d, BYROW(pris, LAMBDA(p, COUNTIFS(TicketData!S:S, p, TicketData!H:H, ">="&mtdStart, TicketData!H:H, "<"&mtdEnd))),' +
    'col_e, BYROW(pris, LAMBDA(p, IFERROR(AVERAGEIFS(TicketData!R:R, TicketData!S:S, p, TicketData!I:I, "Closed"), "N/A"))),' +
    'col_f, BYROW(pris, LAMBDA(p, IFERROR(AVERAGEIFS(TicketData!AE:AE, TicketData!S:S, p, TicketData!I:I, "Closed", TicketData!AE:AE, ">0")/60, "N/A"))),' +
    'col_g, BYROW(pris, LAMBDA(p, LET(total, COUNTIFS(TicketData!S:S, p, TicketData!I:I, "Closed"), breached, COUNTIFS(TicketData!S:S, p, TicketData!I:I, "Closed", TicketData!AF:AF, 1)+COUNTIFS(TicketData!S:S, p, TicketData!I:I, "Closed", TicketData!AI:AI, 1), IF(total>0, breached/total, "N/A")))),' +
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
     '=LET(mtdStart, DATE(YEAR(TODAY()),MONTH(TODAY()),1), mtdEnd, DATE(YEAR(TODAY()),MONTH(TODAY())+1,1), COUNTIFS(TicketData!H:H, ">="&mtdStart, TicketData!H:H, "<"&mtdEnd))',
     'Tickets closed this month'],
    ['Resolved Same Day (<24 hrs)',
     '=LET(mtdStart, DATE(YEAR(TODAY()),MONTH(TODAY()),1), mtdEnd, DATE(YEAR(TODAY()),MONTH(TODAY())+1,1), COUNTIFS(TicketData!H:H, ">="&mtdStart, TicketData!H:H, "<"&mtdEnd, TicketData!R:R, "<=1"))',
     'Closed within 1 day of creation'],
    ['Same Day Rate',
     '=IF(B4>0, B5/B4, "N/A")',
     'FCR target: 40-60%'],
    ['Resolved Within 4 Hours',
     '=LET(mtdStart, DATE(YEAR(TODAY()),MONTH(TODAY()),1), mtdEnd, DATE(YEAR(TODAY()),MONTH(TODAY())+1,1), COUNTIFS(TicketData!H:H, ">="&mtdStart, TicketData!H:H, "<"&mtdEnd, TicketData!AH:AH, "<=240", TicketData!AH:AH, ">0"))',
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

  // Main formula - aggregates by AssignedToUserName (column AO = assigned technician), with dynamic sorting
  const mainFormula =
    '=LET(' +
    'techs, UNIQUE(FILTER(TicketData!AO2:AO, TicketData!AO2:AO<>"", TicketData!AO2:AO<>"AssignedToUserName")),' +
    'mtdStart, DATE(YEAR(TODAY()),MONTH(TODAY()),1),' +
    'mtdEnd, DATE(YEAR(TODAY()),MONTH(TODAY())+1,1),' +
    'col_a, techs,' +
    'col_b, BYROW(techs, LAMBDA(t, IFERROR(INDEX(TicketData!L:L, MATCH(t, TicketData!AO:AO, 0)), ""))),' +
    'col_c, BYROW(techs, LAMBDA(t, COUNTIFS(TicketData!AO:AO, t, TicketData!I:I, "Open"))),' +
    'col_d, BYROW(techs, LAMBDA(t, COUNTIFS(TicketData!AO:AO, t, TicketData!E:E, ">="&mtdStart, TicketData!E:E, "<"&mtdEnd))),' +
    'col_e, BYROW(techs, LAMBDA(t, COUNTIFS(TicketData!AO:AO, t, TicketData!H:H, ">="&mtdStart, TicketData!H:H, "<"&mtdEnd))),' +
    'col_f, BYROW(techs, LAMBDA(t, COUNTIFS(TicketData!AO:AO, t, TicketData!I:I, "Open", TicketData!R:R, ">=30"))),' +
    'col_g, BYROW(techs, LAMBDA(t, IFERROR(AVERAGEIFS(TicketData!R:R, TicketData!AO:AO, t, TicketData!I:I, "Closed"), "N/A"))),' +
    'col_h, BYROW(techs, LAMBDA(t, LET(total, COUNTIFS(TicketData!AO:AO, t, TicketData!I:I, "Closed"), breached, COUNTIFS(TicketData!AO:AO, t, TicketData!I:I, "Closed", TicketData!AF:AF, 1)+COUNTIFS(TicketData!AO:AO, t, TicketData!I:I, "Closed", TicketData!AI:AI, 1), IF(total>0, breached/total, "N/A")))),' +
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
    'mtdStart, DATE(YEAR(TODAY()),MONTH(TODAY()),1),' +
    'mtdEnd, DATE(YEAR(TODAY()),MONTH(TODAY())+1,1),' +
    'col_a, types,' +
    'col_b, BYROW(types, LAMBDA(t, COUNTA(UNIQUE(FILTER(TicketData!N:N, TicketData!O:O=t))))),' +
    'col_c, BYROW(types, LAMBDA(t, COUNTIFS(TicketData!O:O, t, TicketData!I:I, "Open"))),' +
    'col_d, BYROW(types, LAMBDA(t, COUNTIFS(TicketData!O:O, t, TicketData!E:E, ">="&mtdStart, TicketData!E:E, "<"&mtdEnd))),' +
    'col_e, BYROW(types, LAMBDA(t, COUNTIFS(TicketData!O:O, t, TicketData!H:H, ">="&mtdStart, TicketData!H:H, "<"&mtdEnd))),' +
    'col_f, BYROW(types, LAMBDA(t, LET(locs, COUNTA(UNIQUE(FILTER(TicketData!N:N, TicketData!O:O=t))), open, COUNTIFS(TicketData!O:O, t, TicketData!I:I, "Open"), IF(locs>0, open/locs, 0)))),' +
    'col_g, BYROW(types, LAMBDA(t, COUNTIFS(TicketData!O:O, t, TicketData!I:I, "Open", TicketData!R:R, ">=30"))),' +
    'col_h, BYROW(types, LAMBDA(t, LET(open, COUNTIFS(TicketData!O:O, t, TicketData!I:I, "Open"), aged, COUNTIFS(TicketData!O:O, t, TicketData!I:I, "Open", TicketData!R:R, ">=30"), IF(open>0, aged/open, 0)))),' +
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
     '=COUNTIFS(TicketData!I:I, "Open", TicketData!H:H, "<>")',
     'Tickets that were closed but are now open (reopened)'],
    ['Total Open Tickets',
     '=COUNTIFS(TicketData!I:I, "Open")',
     'All currently open tickets'],
    ['Implied Reopen Rate',
     '=IF(B5>0, B4/B5, 0)',
     'Reopened / Total Open (lower is better, target <5%)'],
    ['', '', ''],
    ['--- Quality Indicators ---', '', ''],
    ['Tickets Closed Multiple Times',
     '=COUNTIFS(TicketData!I:I, "Open", TicketData!H:H, "<>", TicketData!R:R, ">7")',
     'Reopened tickets open for 7+ days (significant rework)'],
    ['Avg Age of Reopened',
     '=IFERROR(AVERAGEIFS(TicketData!R:R, TicketData!I:I, "Open", TicketData!H:H, "<>"), "N/A")',
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
    'LEFT(TicketData!G2:G,10), TicketData!R2:R, TicketData!I2:I}, ' +
    '(TicketData!I2:I="Open")*(TicketData!H2:H<>"")), 5, FALSE), "No reopened tickets found")';

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
    'Detection method: Tickets where IsClosed="Open" but ClosedDate is not empty\n' +
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
    'col_c, BYROW(reqs, LAMBDA(r, COUNTIFS(TicketData!AC:AC, r, TicketData!I:I, "Open"))),' +
    'col_d, BYROW(reqs, LAMBDA(r, COUNTIFS(TicketData!AC:AC, r, TicketData!I:I, "Closed"))),' +
    'col_e, BYROW(reqs, LAMBDA(r, IFERROR(INDEX(SORT(UNIQUE(FILTER(TicketData!Y:Y, TicketData!AC:AC=r, TicketData!Y:Y<>"")), 1, FALSE), 1), ""))),' +
    'col_f, BYROW(reqs, LAMBDA(r, IFERROR(AVERAGEIFS(TicketData!R:R, TicketData!AC:AC, r, TicketData!I:I, "Closed"), "N/A"))),' +
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
     '=COUNTIFS(TicketData!F:F, "", TicketData!I:I, "Open")',
     'Open tickets with no StartedDate (still in queue)'],
    ['Currently in Queue',
     '=COUNTIFS(TicketData!F:F, "", TicketData!I:I, "Open")',
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
    '(TicketData!F2:F="")*(TicketData!I2:I="Open")), 4, FALSE), 25, 4), "No tickets waiting")';

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

/**
 * Setup DevicesByRole sheet — device model metrics filterable by requester role
 * Question: "Which device models are generating tickets for students vs staff?"
 * Deletes existing sheet if present for clean slate
 */
function setupDevicesByRoleSheet(ss) {
  deleteSheetIfExists(ss, 'DevicesByRole');
  const sheet = ss.insertSheet('DevicesByRole');

  // === ROLE FILTER CONTROL (row 1-2, columns K-L) ===
  sheet.getRange('K1').setValue('User Role');
  sheet.getRange('K1')
    .setFontWeight('bold')
    .setBackground('#1565c0')
    .setFontColor('white');
  sheet.getRange('K2').setValue('All');
  sheet.getRange('K1').setNote('Filter the device table by the role of the ticket requester (the "For" user)');

  const roleDropdownRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['All', 'Student', 'Staff', 'Agent', 'Guest'], true)
    .setAllowInvalid(false)
    .setHelpText('Filter by requester role: All, Student, Staff, Agent, or Guest')
    .build();
  sheet.getRange('K2').setDataValidation(roleDropdownRule);

  // === MAIN TABLE (row 1-2 headers, row 3+ data, columns A-I) ===
  const headers = ['Model Name', 'Category', 'Total Tickets', 'Open', 'Closed', 'Avg Resolution (days)', 'Breach Rate', 'Sort Col#', 'Desc?'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Main formula — aggregates by ModelName (AL), filtered by RequesterRole (AU)
  // When role = "All" every ticket counts; otherwise COUNTIFS adds the role criteria.
  const mainFormula =
    '=LET(' +
    'selectedRole, $K$2,' +
    'models, IF(selectedRole="All",' +
    '  UNIQUE(FILTER(TicketData!AL2:AL, TicketData!AL2:AL<>"")),' +
    '  UNIQUE(FILTER(TicketData!AL2:AL, TicketData!AL2:AL<>"", TicketData!AU2:AU=selectedRole))),' +
    'col_a, models,' +
    'col_b, BYROW(models, LAMBDA(m, IFERROR(INDEX(TicketData!AQ:AQ, MATCH(m, TicketData!AL:AL, 0)), ""))),' +
    'col_c, BYROW(models, LAMBDA(m, IF(selectedRole="All",' +
    '  COUNTIF(TicketData!AL:AL, m),' +
    '  COUNTIFS(TicketData!AL:AL, m, TicketData!AU:AU, selectedRole)))),' +
    'col_d, BYROW(models, LAMBDA(m, IF(selectedRole="All",' +
    '  COUNTIFS(TicketData!AL:AL, m, TicketData!I:I, "Open"),' +
    '  COUNTIFS(TicketData!AL:AL, m, TicketData!AU:AU, selectedRole, TicketData!I:I, "Open")))),' +
    'col_e, BYROW(models, LAMBDA(m, IF(selectedRole="All",' +
    '  COUNTIFS(TicketData!AL:AL, m, TicketData!I:I, "Closed"),' +
    '  COUNTIFS(TicketData!AL:AL, m, TicketData!AU:AU, selectedRole, TicketData!I:I, "Closed")))),' +
    'col_f, BYROW(models, LAMBDA(m, IFERROR(' +
    '  IF(selectedRole="All",' +
    '    AVERAGEIFS(TicketData!R:R, TicketData!AL:AL, m, TicketData!I:I, "Closed"),' +
    '    AVERAGEIFS(TicketData!R:R, TicketData!AL:AL, m, TicketData!AU:AU, selectedRole, TicketData!I:I, "Closed")), "N/A"))),' +
    'col_g, BYROW(models, LAMBDA(m, LET(' +
    '  total, IF(selectedRole="All",' +
    '    COUNTIFS(TicketData!AL:AL, m, TicketData!I:I, "Closed"),' +
    '    COUNTIFS(TicketData!AL:AL, m, TicketData!AU:AU, selectedRole, TicketData!I:I, "Closed")),' +
    '  breached, IF(selectedRole="All",' +
    '    COUNTIFS(TicketData!AL:AL, m, TicketData!I:I, "Closed", TicketData!AF:AF, 1)+COUNTIFS(TicketData!AL:AL, m, TicketData!I:I, "Closed", TicketData!AI:AI, 1),' +
    '    COUNTIFS(TicketData!AL:AL, m, TicketData!AU:AU, selectedRole, TicketData!I:I, "Closed", TicketData!AF:AF, 1)+COUNTIFS(TicketData!AL:AL, m, TicketData!AU:AU, selectedRole, TicketData!I:I, "Closed", TicketData!AI:AI, 1)),' +
    '  IF(total>0, breached/total, "N/A")))),' +
    'data, HSTACK(col_a, col_b, col_c, col_d, col_e, col_f, col_g),' +
    'SORT(data, $H$2, $I$2))';

  sheet.getRange('A2').setValue(mainFormula);

  // Default sort settings (column 3 = Total Tickets, descending)
  sheet.getRange('H2').setValue(3);
  sheet.getRange('I2').setValue('FALSE');

  // Format header
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1565c0')
    .setFontColor('white');

  // Format columns
  sheet.getRange('F:F').setNumberFormat('0.0');   // Avg Resolution
  sheet.getRange('G:G').setNumberFormat('0.0%');  // Breach Rate

  // Column widths
  sheet.setColumnWidth(1, 220);  // Model Name
  sheet.setColumnWidth(2, 130);  // Category
  sheet.setColumnWidth(8, 80);   // Sort Col#
  sheet.setColumnWidth(9, 60);   // Desc?
  sheet.setColumnWidth(10, 30);  // Spacer
  sheet.setColumnWidth(11, 100); // Role filter label/value

  // Sort column validation
  const sortColRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['1', '2', '3', '4', '5', '6', '7'], true)
    .setHelpText('1=Model, 2=Category, 3=Total, 4=Open, 5=Closed, 6=AvgRes, 7=Breach')
    .build();
  sheet.getRange('H2').setDataValidation(sortColRule);

  const sortOrderRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['FALSE', 'TRUE'], true)
    .setHelpText('FALSE=Descending, TRUE=Ascending')
    .build();
  sheet.getRange('I2').setDataValidation(sortOrderRule);

  sheet.setFrozenRows(1);

  sheet.getRange('A1').setNote(
    'Devices by Role\n\n' +
    'Question: "Which device models are generating tickets for students vs staff?"\n\n' +
    '• Use the User Role dropdown (K2) to switch between Student, Staff, Agent, or All\n' +
    '• Category = asset category (e.g., Chromebooks, Staff Laptops)\n' +
    '• Role is determined by the "For" user on each ticket (the person the ticket is for)\n' +
    '• Breach rate = % of closed tickets that breached SLA\n\n' +
    'IMPORTANT: Requires a Full Reload after updating to v1.3.4 to populate\n' +
    'the RequesterRole column on existing rows.'
  );
  sheet.getRange('H2').setNote('Sort column: 1=Model, 2=Category, 3=Total, 4=Open, 5=Closed, 6=AvgRes, 7=Breach');
  sheet.getRange('I2').setNote('FALSE=Descending (high to low), TRUE=Ascending (low to high)');

  return true;
}

/**
 * Setup DeviceReliability sheet with formulas
 * Groups tickets by device model to identify unreliable models
 * Deletes existing sheet if present for clean slate
 */
function setupDeviceReliabilitySheet(ss) {
  deleteSheetIfExists(ss, 'DeviceReliability');
  const sheet = ss.insertSheet('DeviceReliability');

  // === MAIN SECTION (Row 1-2 headers, Row 3+ data, Columns A-I) ===
  const headers = ['Model Name', 'Total Tickets', 'Open', 'Closed', 'Avg Resolution (days)', 'Breach Rate', 'Last Refreshed', 'Sort Col#', 'Desc?'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Main formula - aggregates by ModelName (column AL)
  const mainFormula =
    '=LET(' +
    'models, UNIQUE(FILTER(TicketData!AL2:AL, TicketData!AL2:AL<>"", TicketData!AL2:AL<>"ModelName")),' +
    'col_a, models,' +
    'col_b, BYROW(models, LAMBDA(m, COUNTIF(TicketData!AL:AL, m))),' +
    'col_c, BYROW(models, LAMBDA(m, COUNTIFS(TicketData!AL:AL, m, TicketData!I:I, "Open"))),' +
    'col_d, BYROW(models, LAMBDA(m, COUNTIFS(TicketData!AL:AL, m, TicketData!I:I, "Closed"))),' +
    'col_e, BYROW(models, LAMBDA(m, IFERROR(AVERAGEIFS(TicketData!R:R, TicketData!AL:AL, m, TicketData!I:I, "Closed"), "N/A"))),' +
    'col_f, BYROW(models, LAMBDA(m, LET(total, COUNTIFS(TicketData!AL:AL, m, TicketData!I:I, "Closed"), breached, COUNTIFS(TicketData!AL:AL, m, TicketData!I:I, "Closed", TicketData!AF:AF, 1)+COUNTIFS(TicketData!AL:AL, m, TicketData!I:I, "Closed", TicketData!AI:AI, 1), IF(total>0, breached/total, "N/A")))),' +
    'data, HSTACK(col_a, col_b, col_c, col_d, col_e, col_f),' +
    'SORT(data, $H$2, $I$2))';

  sheet.getRange('A2').setValue(mainFormula);
  sheet.getRange('G2').setValue('=IFERROR(VLOOKUP("LAST_REFRESH", Config!A:B, 2, FALSE), "")');

  // Default sort settings (column 2 = Total Tickets, descending)
  sheet.getRange('H2').setValue(2);
  sheet.getRange('I2').setValue('FALSE');

  // Format header
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#546e7a')
    .setFontColor('white');

  // Format columns
  sheet.getRange('E:E').setNumberFormat('0.0');   // Avg Resolution
  sheet.getRange('F:F').setNumberFormat('0.0%');  // Breach Rate

  // Add data validation for sort column
  const sortColRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['1', '2', '3', '4', '5', '6'], true)
    .setHelpText('1=Model, 2=Total, 3=Open, 4=Closed, 5=AvgRes, 6=Breach')
    .build();
  sheet.getRange('H2').setDataValidation(sortColRule);

  // Add data validation for sort order
  const sortOrderRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['FALSE', 'TRUE'], true)
    .setHelpText('FALSE=Descending, TRUE=Ascending')
    .build();
  sheet.getRange('I2').setDataValidation(sortOrderRule);

  // === MODEL DETAIL SECTION (Row 1-2 header, Row 3+ data, Columns K-N) ===
  // Control header
  sheet.getRange('K1').setValue('Selected Model');
  sheet.getRange('K1')
    .setFontWeight('bold')
    .setBackground('#ff9800')
    .setFontColor('white');
  sheet.getRange('K2').setValue(''); // Selected model (blank = none)
  sheet.getRange('K1').setNote('Select a model from the dropdown to see issue breakdown');

  // Issue breakdown headers (row 3)
  const detailHeaders = ['Issue Category', 'Count', '% of Model', 'Top Issue Types'];
  sheet.getRange(3, 11, 1, 4).setValues([detailHeaders]);
  sheet.getRange(3, 11, 1, 4)
    .setFontWeight('bold')
    .setBackground('#388e3c')
    .setFontColor('white');

  // Issue breakdown formula - shows issue categories and types for selected model
  // Columns: IssueCategoryName (Y), count, percentage, top issue types from IssueTypeName (AA)
  const detailFormula =
    '=LET(' +
    'selectedModel, $K$2,' +
    'IF(selectedModel="",' +
    '  {"Select a model from K2"},' +
    '  LET(' +
    '    totalForModel, COUNTIF(TicketData!AL:AL, selectedModel),' +
    '    categories, UNIQUE(FILTER(TicketData!Y:Y, TicketData!AL:AL=selectedModel, TicketData!Y:Y<>"")),' +
    '    counts, BYROW(categories, LAMBDA(c, COUNTIFS(TicketData!AL:AL, selectedModel, TicketData!Y:Y, c))),' +
    '    pcts, MAP(counts, LAMBDA(c, IFERROR(c/totalForModel, 0))),' +
    '    topTypes, BYROW(categories, LAMBDA(c, TEXTJOIN(", ", TRUE, ' +
    '      LET(' +
    '        types, UNIQUE(FILTER(TicketData!AA:AA, TicketData!AL:AL=selectedModel, TicketData!Y:Y=c, TicketData!AA:AA<>"")),' +
    '        typeCounts, BYROW(types, LAMBDA(t, COUNTIFS(TicketData!AL:AL, selectedModel, TicketData!Y:Y, c, TicketData!AA:AA, t))),' +
    '        sorted, SORT(HSTACK(types, typeCounts), 2, FALSE),' +
    '        topN, IF(ROWS(sorted)>3, CHOOSEROWS(sorted, 1, 2, 3), sorted),' +
    '        BYROW(topN, LAMBDA(row, INDEX(row, 1)&" ("&INDEX(row, 2)&")"))'  +
    '      )))),' +
    '    data, HSTACK(categories, counts, pcts, topTypes),' +
    '    IF(ROWS(categories)=0, {"No tickets for this model"}, SORT(data, 2, FALSE)))))';

  sheet.getRange('K4').setValue(detailFormula);

  // Model dropdown source (hidden in column P) - sorted by ticket count descending
  sheet.getRange('P1').setValue('ModelDropdownSource');
  const modelDropdownFormula =
    '=LET(' +
    'models, UNIQUE(FILTER(TicketData!AL2:AL, TicketData!AL2:AL<>"", TicketData!AL2:AL<>"ModelName")),' +
    'counts, BYROW(models, LAMBDA(m, COUNTIF(TicketData!AL:AL, m))),' +
    'SORT(models, counts, FALSE))';
  sheet.getRange('P2').setValue(modelDropdownFormula);
  sheet.hideColumns(16); // Hide column P

  // Create data validation rule for K2 using the dropdown source
  const modelDropdownRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sheet.getRange('P2:P1000'), true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('K2').setDataValidation(modelDropdownRule);

  // Format detail columns
  sheet.getRange('M:M').setNumberFormat('0.0%');  // % of Model

  // Column widths
  sheet.setColumnWidth(1, 200);  // Model Name
  sheet.setColumnWidth(7, 120);  // Last Refreshed
  sheet.setColumnWidth(8, 80);   // Sort Col#
  sheet.setColumnWidth(9, 60);   // Desc?
  sheet.setColumnWidth(10, 30);  // Spacer
  sheet.setColumnWidth(11, 200); // Selected Model / Issue Category
  sheet.setColumnWidth(12, 80);  // Count
  sheet.setColumnWidth(13, 100); // % of Model
  sheet.setColumnWidth(14, 350); // Top Issue Types

  sheet.setFrozenRows(1);

  // Add notes
  sheet.getRange('A1').setNote(
    'Device Reliability Analysis\n\n' +
    'Question: "Which device models generate the most tickets?"\n\n' +
    'MODEL SUMMARY (left):\n' +
    '- All models ranked by ticket count\n' +
    '- Use Sort Col# and Desc? to change sorting\n' +
    '- Breach rate = % of closed tickets that breached SLA\n\n' +
    'MODEL DETAIL (right):\n' +
    '- Select a model from K2 dropdown\n' +
    '- Shows issue category breakdown with percentages\n' +
    '- Top 3 issue types shown for each category\n\n' +
    'Use this to:\n' +
    '- Identify unreliable device models\n' +
    '- See what types of issues affect each model\n' +
    '- Justify replacement budgets with data'
  );
  sheet.getRange('H2').setNote('Sort column: 1=Model, 2=Total, 3=Open, 4=Closed, 5=AvgRes, 6=Breach');
  sheet.getRange('I2').setNote('FALSE=Descending (high to low), TRUE=Ascending (low to high)');

  return true;
}

/**
 * Setup BacklogAgingByFA sheet — Functional Areas as rows, age buckets as columns
 * Uses SUMPRODUCT because FA requires joining TeamName through Teams sheet
 */
function setupBacklogAgingByFASheet(ss) {
  deleteSheetIfExists(ss, 'BacklogAgingByFA');
  const sheet = ss.insertSheet('BacklogAgingByFA');

  // Get functional areas from Teams sheet
  const teamsSheet = ss.getSheetByName('Teams');
  if (!teamsSheet || teamsSheet.getLastRow() < 2) {
    sheet.getRange('A1').setValue('Teams sheet is empty or missing. Please load Teams first.');
    return true;
  }

  const faValues = teamsSheet.getRange(2, 3, teamsSheet.getLastRow() - 1, 1).getValues()
    .map(r => r[0])
    .filter(v => v && String(v).trim() !== '');
  const uniqueFAs = [...new Set(faValues)].sort();

  if (uniqueFAs.length === 0) {
    sheet.getRange('A1').setValue('No Functional Areas found. Please fill in the FunctionalArea column (C) in the Teams sheet.');
    return true;
  }

  // Build FA → teams mapping
  const teamData = teamsSheet.getRange(2, 2, teamsSheet.getLastRow() - 1, 2).getValues();
  const faToTeams = {};
  for (const [teamName, fa] of teamData) {
    if (fa && String(fa).trim() !== '') {
      const faKey = String(fa).trim();
      if (!faToTeams[faKey]) faToTeams[faKey] = [];
      if (teamName && String(teamName).trim() !== '') {
        faToTeams[faKey].push(String(teamName).trim());
      }
    }
  }

  // Headers: entity as rows, buckets as columns
  const headers = ['Functional Area', '0-7 days', '8-14 days', '15-30 days', '30+ days', 'Total'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const buckets = [
    { min: 0, max: 7 },
    { min: 8, max: 14 },
    { min: 15, max: 30 },
    { min: 31, max: null }
  ];

  const dataRows = [];
  for (const fa of uniqueFAs) {
    const rowNum = dataRows.length + 2;
    const teams = faToTeams[fa] || [];
    const row = [fa];

    if (teams.length === 0) {
      row.push(0, 0, 0, 0);
    } else {
      const teamCriteria = teams.map(t => '(TicketData!L:L="' + t.replace(/"/g, '""') + '")').join('+');
      for (const bucket of buckets) {
        if (bucket.max !== null) {
          row.push('=SUMPRODUCT((TicketData!I:I="Open")*(TicketData!R:R>=' + bucket.min + ')*(TicketData!R:R<=' + bucket.max + ')*((' + teamCriteria + ')>0))');
        } else {
          row.push('=SUMPRODUCT((TicketData!I:I="Open")*(TicketData!R:R>=' + bucket.min + ')*((' + teamCriteria + ')>0))');
        }
      }
    }

    row.push(`=SUM(B${rowNum}:E${rowNum})`);
    dataRows.push(row);
  }

  if (dataRows.length > 0) {
    sheet.getRange(2, 1, dataRows.length, 6).setValues(dataRows);
  }

  formatBacklogAgingCrossTab(sheet, headers, dataRows.length, '#ea4335',
    'Backlog Aging by Functional Area\n\n' +
    'Shows how long open tickets have been sitting, broken down by department.\n' +
    'Functional Areas are derived from the Teams sheet (column C).\n' +
    'FAs are listed alphabetically. Use Google Sheets built-in sort to re-sort.\n\n' +
    'To recreate: iiQ Data > Add Analytics Sheet > Backlog & Quality > Backlog Aging by FA');

  return true;
}

/**
 * Setup BacklogAgingByTeam sheet — Teams as rows, age buckets as columns
 * Uses a single LET formula since TeamName is a direct column in TicketData
 */
function setupBacklogAgingByTeamSheet(ss) {
  deleteSheetIfExists(ss, 'BacklogAgingByTeam');
  const sheet = ss.insertSheet('BacklogAgingByTeam');

  const headers = ['Team', '0-7 days', '8-14 days', '15-30 days', '30+ days', 'Total', 'Sort Col#', 'Desc?'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Single LET formula: entities as rows, age buckets as columns
  const mainFormula =
    '=LET(' +
    'teams, UNIQUE(FILTER(TicketData!L2:L, TicketData!L2:L<>"", TicketData!L2:L<>"TeamName")),' +
    'col_a, teams,' +
    'col_b, BYROW(teams, LAMBDA(t, COUNTIFS(TicketData!L:L, t, TicketData!I:I, "Open", TicketData!R:R, ">=0", TicketData!R:R, "<=7"))),' +
    'col_c, BYROW(teams, LAMBDA(t, COUNTIFS(TicketData!L:L, t, TicketData!I:I, "Open", TicketData!R:R, ">=8", TicketData!R:R, "<=14"))),' +
    'col_d, BYROW(teams, LAMBDA(t, COUNTIFS(TicketData!L:L, t, TicketData!I:I, "Open", TicketData!R:R, ">=15", TicketData!R:R, "<=30"))),' +
    'col_e, BYROW(teams, LAMBDA(t, COUNTIFS(TicketData!L:L, t, TicketData!I:I, "Open", TicketData!R:R, ">=31"))),' +
    'col_f, BYROW(teams, LAMBDA(t, COUNTIFS(TicketData!L:L, t, TicketData!I:I, "Open"))),' +
    'data, HSTACK(col_a, col_b, col_c, col_d, col_e, col_f),' +
    'SORT(data, $G$2, $H$2))';

  sheet.getRange('A2').setValue(mainFormula);

  // Default sort settings (column 6 = Total, descending)
  sheet.getRange('G2').setValue(6);
  sheet.getRange('H2').setValue('FALSE');

  formatBacklogAgingCrossTab(sheet, headers, -1, '#673ab7',
    'Backlog Aging by Team\n\n' +
    'Shows how long open tickets have been sitting, broken down by team.\n' +
    'Uses TeamName (column L) from TicketData.\n\n' +
    'To recreate: iiQ Data > Add Analytics Sheet > Backlog & Quality > Backlog Aging by Team');

  // Sort column validation
  const sortColRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['1', '2', '3', '4', '5', '6'], true)
    .setHelpText('1=Team, 2=0-7d, 3=8-14d, 4=15-30d, 5=30+, 6=Total')
    .build();
  sheet.getRange('G2').setDataValidation(sortColRule);

  const sortOrderRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['FALSE', 'TRUE'], true)
    .setHelpText('FALSE=Descending, TRUE=Ascending')
    .build();
  sheet.getRange('H2').setDataValidation(sortOrderRule);

  sheet.getRange('G2').setNote('Sort column: 1=Team, 2=0-7d, 3=8-14d, 4=15-30d, 5=30+, 6=Total');
  sheet.getRange('H2').setNote('FALSE=Descending (high to low), TRUE=Ascending (low to high)');

  return true;
}

/**
 * Setup BacklogAgingByLocationType sheet — Location types as rows, age buckets as columns
 * Uses a single LET formula since LocationType is a direct column in TicketData
 */
function setupBacklogAgingByLocationTypeSheet(ss) {
  deleteSheetIfExists(ss, 'BacklogAgingByLocationType');
  const sheet = ss.insertSheet('BacklogAgingByLocationType');

  const headers = ['Location Type', '0-7 days', '8-14 days', '15-30 days', '30+ days', 'Total', 'Sort Col#', 'Desc?'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const mainFormula =
    '=LET(' +
    'types, UNIQUE(FILTER(TicketData!O2:O, TicketData!O2:O<>"", TicketData!O2:O<>"LocationType")),' +
    'col_a, types,' +
    'col_b, BYROW(types, LAMBDA(t, COUNTIFS(TicketData!O:O, t, TicketData!I:I, "Open", TicketData!R:R, ">=0", TicketData!R:R, "<=7"))),' +
    'col_c, BYROW(types, LAMBDA(t, COUNTIFS(TicketData!O:O, t, TicketData!I:I, "Open", TicketData!R:R, ">=8", TicketData!R:R, "<=14"))),' +
    'col_d, BYROW(types, LAMBDA(t, COUNTIFS(TicketData!O:O, t, TicketData!I:I, "Open", TicketData!R:R, ">=15", TicketData!R:R, "<=30"))),' +
    'col_e, BYROW(types, LAMBDA(t, COUNTIFS(TicketData!O:O, t, TicketData!I:I, "Open", TicketData!R:R, ">=31"))),' +
    'col_f, BYROW(types, LAMBDA(t, COUNTIFS(TicketData!O:O, t, TicketData!I:I, "Open"))),' +
    'data, HSTACK(col_a, col_b, col_c, col_d, col_e, col_f),' +
    'SORT(data, $G$2, $H$2))';

  sheet.getRange('A2').setValue(mainFormula);

  // Default sort settings (column 6 = Total, descending)
  sheet.getRange('G2').setValue(6);
  sheet.getRange('H2').setValue('FALSE');

  formatBacklogAgingCrossTab(sheet, headers, -1, '#0f9d58',
    'Backlog Aging by Location Type\n\n' +
    'Shows how long open tickets have been sitting by school type (Elementary, Middle, High, etc.).\n' +
    'Uses LocationType (column O) from TicketData.\n\n' +
    'To recreate: iiQ Data > Add Analytics Sheet > Backlog & Quality > Backlog Aging by Location Type');

  const sortColRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['1', '2', '3', '4', '5', '6'], true)
    .setHelpText('1=Type, 2=0-7d, 3=8-14d, 4=15-30d, 5=30+, 6=Total')
    .build();
  sheet.getRange('G2').setDataValidation(sortColRule);

  const sortOrderRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['FALSE', 'TRUE'], true)
    .setHelpText('FALSE=Descending, TRUE=Ascending')
    .build();
  sheet.getRange('H2').setDataValidation(sortOrderRule);

  sheet.getRange('G2').setNote('Sort column: 1=Type, 2=0-7d, 3=8-14d, 4=15-30d, 5=30+, 6=Total');
  sheet.getRange('H2').setNote('FALSE=Descending (high to low), TRUE=Ascending (low to high)');

  return true;
}

/**
 * Setup BacklogAgingByPriority sheet — Priorities as rows, age buckets as columns
 * Uses a single LET formula since Priority is a direct column in TicketData
 */
function setupBacklogAgingByPrioritySheet(ss) {
  deleteSheetIfExists(ss, 'BacklogAgingByPriority');
  const sheet = ss.insertSheet('BacklogAgingByPriority');

  const headers = ['Priority', '0-7 days', '8-14 days', '15-30 days', '30+ days', 'Total', 'Sort Col#', 'Desc?'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const mainFormula =
    '=LET(' +
    'pris, UNIQUE(FILTER(TicketData!S2:S, TicketData!S2:S<>"", TicketData!S2:S<>"Priority")),' +
    'col_a, pris,' +
    'col_b, BYROW(pris, LAMBDA(p, COUNTIFS(TicketData!S:S, p, TicketData!I:I, "Open", TicketData!R:R, ">=0", TicketData!R:R, "<=7"))),' +
    'col_c, BYROW(pris, LAMBDA(p, COUNTIFS(TicketData!S:S, p, TicketData!I:I, "Open", TicketData!R:R, ">=8", TicketData!R:R, "<=14"))),' +
    'col_d, BYROW(pris, LAMBDA(p, COUNTIFS(TicketData!S:S, p, TicketData!I:I, "Open", TicketData!R:R, ">=15", TicketData!R:R, "<=30"))),' +
    'col_e, BYROW(pris, LAMBDA(p, COUNTIFS(TicketData!S:S, p, TicketData!I:I, "Open", TicketData!R:R, ">=31"))),' +
    'col_f, BYROW(pris, LAMBDA(p, COUNTIFS(TicketData!S:S, p, TicketData!I:I, "Open"))),' +
    'data, HSTACK(col_a, col_b, col_c, col_d, col_e, col_f),' +
    'SORT(data, $G$2, $H$2))';

  sheet.getRange('A2').setValue(mainFormula);

  // Default sort settings (column 6 = Total, descending)
  sheet.getRange('G2').setValue(6);
  sheet.getRange('H2').setValue('FALSE');

  formatBacklogAgingCrossTab(sheet, headers, -1, '#e65100',
    'Backlog Aging by Priority\n\n' +
    'Shows how long open tickets have been sitting by priority level.\n' +
    'Uses Priority (column S) from TicketData.\n\n' +
    'To recreate: iiQ Data > Add Analytics Sheet > Backlog & Quality > Backlog Aging by Priority');

  const sortColRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['1', '2', '3', '4', '5', '6'], true)
    .setHelpText('1=Priority, 2=0-7d, 3=8-14d, 4=15-30d, 5=30+, 6=Total')
    .build();
  sheet.getRange('G2').setDataValidation(sortColRule);

  const sortOrderRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['FALSE', 'TRUE'], true)
    .setHelpText('FALSE=Descending, TRUE=Ascending')
    .build();
  sheet.getRange('H2').setDataValidation(sortOrderRule);

  sheet.getRange('G2').setNote('Sort column: 1=Priority, 2=0-7d, 3=8-14d, 4=15-30d, 5=30+, 6=Total');
  sheet.getRange('H2').setNote('FALSE=Descending (high to low), TRUE=Ascending (low to high)');

  return true;
}

/**
 * Shared formatting for all BacklogAging cross-tab sheets.
 * @param {Sheet} sheet - The sheet to format
 * @param {string[]} headers - Header row values
 * @param {number} dataRowCount - Number of data rows (-1 for dynamic LET formulas)
 * @param {string} headerColor - Header background color hex
 * @param {string} noteText - Note to add to A1
 */
function formatBacklogAgingCrossTab(sheet, headers, dataRowCount, headerColor, noteText) {
  // Format header
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground(headerColor)
    .setFontColor('white');

  // Column widths
  sheet.setColumnWidth(1, 180);  // Entity name
  sheet.setColumnWidth(2, 80);   // 0-7 days
  sheet.setColumnWidth(3, 80);   // 8-14 days
  sheet.setColumnWidth(4, 85);   // 15-30 days
  sheet.setColumnWidth(5, 80);   // 30+ days
  sheet.setColumnWidth(6, 70);   // Total
  sheet.setColumnWidth(7, 80);   // Sort Col#
  sheet.setColumnWidth(8, 60);   // Desc?

  sheet.setFrozenRows(1);

  // Conditional formatting: highlight 30+ days cells with light red background
  // For dynamic sheets, use a large range that covers expected data
  const maxRows = (dataRowCount > 0) ? dataRowCount : 100;
  const agingRange = sheet.getRange(2, 5, maxRows, 1); // Column E = 30+ days
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground('#fce8e6')
    .setRanges([agingRange])
    .build();
  sheet.setConditionalFormatRules([rule]);

  // Add note
  sheet.getRange('A1').setNote(noteText);
}

/**
 * Setup MonthlyVolumeByFA sheet — cross-tab of months × functional areas
 * Rows: Months (from school year range)
 * Columns: One per Functional Area with Created/Closed sub-columns
 * Values: Ticket counts per month per FA
 */
function setupMonthlyVolumeByFASheet(ss) {
  deleteSheetIfExists(ss, 'MonthlyVolumeByFA');
  const sheet = ss.insertSheet('MonthlyVolumeByFA');

  // Get functional areas from Teams sheet
  const teamsSheet = ss.getSheetByName('Teams');
  if (!teamsSheet || teamsSheet.getLastRow() < 2) {
    sheet.getRange('A1').setValue('Teams sheet is empty or missing. Please load Teams first.');
    return true;
  }

  const faValues = teamsSheet.getRange(2, 3, teamsSheet.getLastRow() - 1, 1).getValues()
    .map(r => r[0])
    .filter(v => v && String(v).trim() !== '');
  const uniqueFAs = [...new Set(faValues)].sort();

  if (uniqueFAs.length === 0) {
    sheet.getRange('A1').setValue('No Functional Areas found. Please fill in the FunctionalArea column (C) in the Teams sheet.');
    return true;
  }

  // Build FA → teams mapping
  const teamData = teamsSheet.getRange(2, 2, teamsSheet.getLastRow() - 1, 2).getValues();
  const faToTeams = {};
  for (const [teamName, fa] of teamData) {
    if (fa && String(fa).trim() !== '') {
      const faKey = String(fa).trim();
      if (!faToTeams[faKey]) faToTeams[faKey] = [];
      if (teamName && String(teamName).trim() !== '') {
        faToTeams[faKey].push(String(teamName).trim());
      }
    }
  }

  // Build headers: Month | Year | FA1 Created | FA1 Closed | FA2 Created | FA2 Closed | ... | Total Created | Total Closed
  const headers = ['Month', 'Year'];
  for (const fa of uniqueFAs) {
    headers.push(fa + ' Created');
    headers.push(fa + ' Closed');
  }
  headers.push('Total Created', 'Total Closed');
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Get month range
  const monthRange = getMonthRangeFromData(ss, 'E');

  const dataRows = [];
  for (const period of monthRange) {
    const rowNum = dataRows.length + 2;
    const y = period.year;
    const monthNum = period.monthNum;

    const startDate = `TEXT(DATE(${y},${monthNum},1), "YYYY-MM-DD")`;
    const endDate = `TEXT(DATE(${y},${monthNum}+1,1), "YYYY-MM-DD")`;

    const row = [period.monthName, y];

    for (const fa of uniqueFAs) {
      const teams = faToTeams[fa] || [];
      if (teams.length === 0) {
        row.push(0, 0);
      } else {
        // Sum COUNTIFS per team — COUNTIFS handles date string comparison correctly
        // unlike SUMPRODUCT which fails when comparing date cells to TEXT() strings
        const createdParts = teams.map(t =>
          'COUNTIFS(TicketData!E:E,">="&' + startDate + ',TicketData!E:E,"<"&' + endDate + ',TicketData!L:L,"' + t.replace(/"/g, '""') + '")'
        );
        const closedParts = teams.map(t =>
          'COUNTIFS(TicketData!H:H,">="&' + startDate + ',TicketData!H:H,"<"&' + endDate + ',TicketData!L:L,"' + t.replace(/"/g, '""') + '")'
        );
        row.push('=' + createdParts.join('+'));
        row.push('=' + closedParts.join('+'));
      }
    }

    // Total Created and Total Closed
    const createdCols = [];
    const closedCols = [];
    for (let i = 0; i < uniqueFAs.length; i++) {
      const colIdx = 3 + (i * 2); // 1-indexed: col 3, 5, 7, ...
      createdCols.push(String.fromCharCode(64 + colIdx) + rowNum);
      closedCols.push(String.fromCharCode(64 + colIdx + 1) + rowNum);
    }
    row.push(`=${createdCols.join('+')}`);
    row.push(`=${closedCols.join('+')}`);

    dataRows.push(row);
  }

  if (dataRows.length > 0) {
    sheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
  }

  // Format header row
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#fbbc04')
    .setFontColor('black');

  // Alternate header colors for each FA pair
  const faColors = ['#e8f0fe', '#fce8e6', '#e6f4ea', '#fef7e0', '#f3e8fd', '#e8f5e9', '#fff3e0', '#e1f5fe'];
  for (let i = 0; i < uniqueFAs.length; i++) {
    const colStart = 3 + (i * 2);
    const color = faColors[i % faColors.length];
    sheet.getRange(1, colStart, 1, 2).setBackground(color);
  }

  // Column widths
  sheet.setColumnWidth(1, 90);  // Month
  sheet.setColumnWidth(2, 50);  // Year
  for (let c = 3; c <= headers.length; c++) {
    sheet.setColumnWidth(c, 85);
  }

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);

  // Add note
  sheet.getRange('A1').setNote(
    'Monthly Volume by Functional Area\n\n' +
    'Shows ticket throughput (Created/Closed) per department over time.\n' +
    'Functional Areas are derived from the Teams sheet (column C).\n\n' +
    'This is a monthly sheet — use "Regenerate All Monthly Sheets" if school year changes.\n\n' +
    'To update: iiQ Data > Add Analytics Sheet > Volume & Trends > Monthly Volume by FA'
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
 * - MonthlyVolumeByFA
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
    { name: 'SeasonalComparison', setup: setupSeasonalComparisonSheet },
    { name: 'MonthlyVolumeByFA', setup: setupMonthlyVolumeByFASheet }
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

// ============================================================================
// FREQUENT FLYERS - Users and Devices with recurring issues
// ============================================================================

/**
 * Add Frequent Flyers sheet (deletes and recreates if exists)
 */
function addFrequentFlyersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupFrequentFlyersSheet(ss);
  SpreadsheetApp.getUi().alert('Created',
    'FrequentFlyers sheet has been created.\n\n' +
    'Shows users with recurring tickets (left) and ticket details (right).\n' +
    'Select a requester from H2 dropdown to see their tickets.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Setup FrequentFlyers sheet - user frequent flyers with ticket detail drill-down
 * Question: "Which users have recurring issues?"
 *
 * Layout:
 * - Left: User frequent flyers table (filtered by min tickets, date range)
 * - Right: Ticket details for selected requester
 *
 * Uses:
 * - RequesterId/RequesterName (AB/AC) for user identification (the "on behalf of" user)
 *
 * Controls:
 * - Min Tickets: Minimum ticket count threshold (default 3)
 * - Date From/To: Optional date range filter
 * - Selected Requester: Dropdown to view ticket details
 */
function setupFrequentFlyersSheet(ss) {
  deleteSheetIfExists(ss, 'FrequentFlyers');
  const sheet = ss.insertSheet('FrequentFlyers');

  // === CONTROLS SECTION (Row 1-2) ===
  const controlHeaders = ['Min Tickets', 'Date From', 'Date To', '', '', '', '', 'Selected Requester'];
  sheet.getRange(1, 1, 1, 8).setValues([controlHeaders]);

  // Default control values
  sheet.getRange('A2').setValue(3);  // Min tickets threshold
  sheet.getRange('B2').setValue(''); // Date from (blank = all time)
  sheet.getRange('C2').setValue(''); // Date to (blank = today)
  sheet.getRange('H2').setValue(''); // Selected requester (blank = none)

  // Format controls header
  sheet.getRange(1, 1, 1, 3)
    .setFontWeight('bold')
    .setBackground('#ff9800')
    .setFontColor('white');
  sheet.getRange('H1')
    .setFontWeight('bold')
    .setBackground('#ff9800')
    .setFontColor('white');

  // Add notes to controls
  sheet.getRange('A1').setNote('Minimum ticket count to appear in results (default: 3)');
  sheet.getRange('B1').setNote('Start date (YYYY-MM-DD). Leave blank for all time.');
  sheet.getRange('C1').setNote('End date (YYYY-MM-DD). Leave blank for today.');
  sheet.getRange('H1').setNote('Select a requester from the list to see their ticket details');

  // === USER FREQUENT FLYERS SECTION (Row 4+, Columns A-G) ===
  const userHeaders = ['Requester Name', 'Location', 'Total Tickets', 'Open', 'Closed', 'Avg Resolution', 'Last Refreshed'];
  sheet.getRange(4, 1, 1, 7).setValues([userHeaders]);

  // User formula - filters by min tickets and optional date range
  // Uses RequesterName (AC) and includes their location from the For user
  const userFormula =
    '=LET(' +
    'minTix, $A$2,' +
    'dateFrom, $B$2,' +
    'dateTo, IF($C$2="", TEXT(TODAY(), "YYYY-MM-DD"), $C$2),' +
    'hasDateFilter, dateFrom<>"",' +
    // Get all requesters with their ticket counts
    'allReqs, UNIQUE(FILTER(TicketData!AC2:AC, TicketData!AC2:AC<>"")),' +
    // Count tickets per requester (with optional date filter)
    'counts, IF(hasDateFilter,' +
    '  BYROW(allReqs, LAMBDA(r, COUNTIFS(TicketData!AC:AC, r, TicketData!E:E, ">="&dateFrom, TicketData!E:E, "<="&dateTo))),' +
    '  BYROW(allReqs, LAMBDA(r, COUNTIF(TicketData!AC:AC, r)))),' +
    // Filter to those meeting threshold
    'filtered, FILTER(allReqs, counts>=minTix),' +
    'filteredCounts, FILTER(counts, counts>=minTix),' +
    // Build data columns
    'col_a, filtered,' +
    'col_b, BYROW(filtered, LAMBDA(r, IFERROR(INDEX(FILTER(TicketData!N:N, TicketData!AC:AC=r), 1), ""))),' +  // Location
    'col_c, filteredCounts,' +
    'col_d, IF(hasDateFilter,' +
    '  BYROW(filtered, LAMBDA(r, COUNTIFS(TicketData!AC:AC, r, TicketData!I:I, "Open", TicketData!E:E, ">="&dateFrom, TicketData!E:E, "<="&dateTo))),' +
    '  BYROW(filtered, LAMBDA(r, COUNTIFS(TicketData!AC:AC, r, TicketData!I:I, "Open")))),' +
    'col_e, IF(hasDateFilter,' +
    '  BYROW(filtered, LAMBDA(r, COUNTIFS(TicketData!AC:AC, r, TicketData!I:I, "Closed", TicketData!E:E, ">="&dateFrom, TicketData!E:E, "<="&dateTo))),' +
    '  BYROW(filtered, LAMBDA(r, COUNTIFS(TicketData!AC:AC, r, TicketData!I:I, "Closed")))),' +
    'col_f, BYROW(filtered, LAMBDA(r, IFERROR(AVERAGEIFS(TicketData!R:R, TicketData!AC:AC, r, TicketData!I:I, "Closed"), "N/A"))),' +
    'data, HSTACK(col_a, col_b, col_c, col_d, col_e, col_f),' +
    'IF(ROWS(filtered)=0, {"No users meet threshold"}, SORT(data, 3, FALSE)))';

  sheet.getRange('A5').setValue(userFormula);
  sheet.getRange('G5').setValue('=IFERROR(VLOOKUP("LAST_REFRESH", Config!A:B, 2, FALSE), "")');

  // Format user header
  sheet.getRange(4, 1, 1, 7)
    .setFontWeight('bold')
    .setBackground('#1976d2')
    .setFontColor('white');

  // === REQUESTER DETAIL SECTION (Row 4+, Columns I-N) ===
  const detailHeaders = ['Created Date', 'Asset Tag', 'Model', 'Ticket #', 'Issue Type', 'Issue Category'];
  sheet.getRange(4, 9, 1, 6).setValues([detailHeaders]);

  // Detail formula - shows tickets for selected requester within date range
  // Columns: CreatedDate (E), AssetTag (AK), ModelName (AL), TicketNumber (B), IssueTypeName (AA), IssueCategoryName (Y)
  const detailFormula =
    '=LET(' +
    'selectedReq, $H$2,' +
    'dateFrom, $B$2,' +
    'dateTo, IF($C$2="", TEXT(TODAY(), "YYYY-MM-DD"), $C$2),' +
    'hasDateFilter, dateFrom<>"",' +
    'IF(selectedReq="",' +
    '  {"Select a requester from H2"},' +
    '  IF(hasDateFilter,' +
    '    IFERROR(SORT(FILTER({TicketData!E:E, TicketData!AK:AK, TicketData!AL:AL, TicketData!B:B, TicketData!AA:AA, TicketData!Y:Y},' +
    '      TicketData!AC:AC=selectedReq, TicketData!E:E>=dateFrom, TicketData!E:E<=dateTo), 1, FALSE),' +
    '      {"No tickets found for this requester in date range"}),' +
    '    IFERROR(SORT(FILTER({TicketData!E:E, TicketData!AK:AK, TicketData!AL:AL, TicketData!B:B, TicketData!AA:AA, TicketData!Y:Y},' +
    '      TicketData!AC:AC=selectedReq), 1, FALSE),' +
    '      {"No tickets found for this requester"}))))';

  sheet.getRange('I5').setValue(detailFormula);

  // Format detail header
  sheet.getRange(4, 9, 1, 6)
    .setFontWeight('bold')
    .setBackground('#388e3c')
    .setFontColor('white');

  // Add dropdown for requester selection (populated from frequent flyers list)
  // This formula creates the dropdown options from the sorted frequent flyers
  const dropdownFormula =
    '=LET(' +
    'minTix, $A$2,' +
    'dateFrom, $B$2,' +
    'dateTo, IF($C$2="", TEXT(TODAY(), "YYYY-MM-DD"), $C$2),' +
    'hasDateFilter, dateFrom<>"",' +
    'allReqs, UNIQUE(FILTER(TicketData!AC2:AC, TicketData!AC2:AC<>"")),' +
    'counts, IF(hasDateFilter,' +
    '  BYROW(allReqs, LAMBDA(r, COUNTIFS(TicketData!AC:AC, r, TicketData!E:E, ">="&dateFrom, TicketData!E:E, "<="&dateTo))),' +
    '  BYROW(allReqs, LAMBDA(r, COUNTIF(TicketData!AC:AC, r)))),' +
    'filtered, FILTER(allReqs, counts>=minTix),' +
    'filteredCounts, FILTER(counts, counts>=minTix),' +
    'IF(ROWS(filtered)=0, "", SORT(filtered, filteredCounts, FALSE)))';

  // Put dropdown source formula in a hidden area (column P)
  sheet.getRange('P1').setValue('DropdownSource');
  sheet.getRange('P2').setValue(dropdownFormula);
  sheet.hideColumns(16); // Hide column P

  // Create data validation rule for H2 using the dropdown source
  const dropdownRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sheet.getRange('P2:P500'), true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('H2').setDataValidation(dropdownRule);

  // === FORMATTING ===

  // Section labels
  sheet.getRange('A3').setValue('USER FREQUENT FLYERS');
  sheet.getRange('I3').setValue('REQUESTER TICKET DETAILS');
  sheet.getRange('A3:G3').merge().setFontWeight('bold').setFontSize(11);
  sheet.getRange('I3:N3').merge().setFontWeight('bold').setFontSize(11);

  // Number formats
  sheet.getRange('F:F').setNumberFormat('0.0');  // Avg Resolution (users)
  sheet.getRange('I:I').setNumberFormat('yyyy-mm-dd');  // Created Date

  // Column widths
  sheet.setColumnWidth(1, 180);  // Requester Name
  sheet.setColumnWidth(2, 150);  // Location
  sheet.setColumnWidth(7, 120);  // Last Refreshed
  sheet.setColumnWidth(8, 180);  // Selected Requester
  sheet.setColumnWidth(9, 100);  // Created Date
  sheet.setColumnWidth(10, 120); // Asset Tag
  sheet.setColumnWidth(11, 180); // Model
  sheet.setColumnWidth(12, 80);  // Ticket #
  sheet.setColumnWidth(13, 150); // Issue Type
  sheet.setColumnWidth(14, 150); // Issue Category

  sheet.setFrozenRows(4);

  // Add main note
  sheet.getRange('A3').setNote(
    'User Frequent Flyers\n\n' +
    'Question: "Which users have recurring issues?"\n\n' +
    'USER FREQUENT FLYERS (left):\n' +
    '- Groups by "on behalf of" user (the person the ticket was submitted for)\n' +
    '- Use to identify students/staff who may need training or equipment replacement\n\n' +
    'REQUESTER TICKET DETAILS (right):\n' +
    '- Select a requester from H2 dropdown to see their tickets\n' +
    '- Shows ticket details within the date range filter\n\n' +
    'CONTROLS (row 2):\n' +
    '- Min Tickets: Only show users with this many+ tickets\n' +
    '- Date From/To: Filter to a specific date range\n' +
    '- Selected Requester: Pick from dropdown to see ticket details'
  );

  return true;
}
