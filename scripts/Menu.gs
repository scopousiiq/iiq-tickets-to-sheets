/**
 * iiQ Data Menu
 *
 * Provides menu access to data loader and analytics functions.
 *
 * Data Loading:
 * - Ticket Data includes consolidated SLA metrics (36 columns)
 * - SLA data is fetched per-batch during ticket loading
 *
 * Refresh Strategy:
 * - Initial Load: Use "Continue Loading" repeatedly until complete
 * - Ongoing: "Open Ticket Refresh" every 2 hours (updates open + recently closed)
 * - New Tickets: Automatically fetched with open refresh
 * - Weekly: "Full Reload" catches deletions and corrections
 *
 * Analytics Sheets:
 * - 7 default sheets created by Setup Spreadsheet
 * - 13 additional sheets available via "Add Analytics Sheet" menu
 * - All sheets can be deleted and recreated as needed
 * - Organized by category: Volume, Backlog, SLA, Team, Location, Issue
 * - Default sheets marked with ★ in menu
 */

/**
 * Create the iiQ Data menu when the spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('iiQ Data')
    .addSubMenu(ui.createMenu('Setup')
      .addItem('Setup Spreadsheet', 'setupSpreadsheet')
      .addItem('Verify Configuration', 'verifyConfiguration')
      .addSeparator()
      .addItem('Setup Automated Triggers', 'setupAutomatedTriggers')
      .addItem('View Trigger Status', 'viewTriggerStatus')
      .addItem('Remove Automated Triggers', 'removeAutomatedTriggers'))
    .addSeparator()
    .addItem('Refresh Teams', 'refreshTeams')
    .addSubMenu(ui.createMenu('Ticket Data')
      .addItem('Continue Loading (Initial)', 'refreshTicketDataContinue')
      .addSeparator()
      .addItem('Open Ticket Refresh (Start)', 'refreshOpenTicketsStart')
      .addItem('Open Ticket Refresh (Continue)', 'refreshOpenTicketsContinue')
      .addItem('Open Refresh Status', 'showOpenRefreshStatus')
      .addSeparator()
      .addItem('Full Reload', 'refreshTicketDataFull')
      .addItem('Show Status', 'showTicketDataStatus'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Add Analytics Sheet')
      .addSubMenu(ui.createMenu('Volume & Trends')
        .addItem('Monthly Volume ★', 'addMonthlyVolumeSheet')
        .addItem('Performance Trends ★', 'addPerformanceTrendsSheet')
        .addItem('Seasonal Comparison (YoY)', 'addSeasonalComparisonSheet')
        .addItem('Temporal Patterns', 'addTemporalPatternsSheet'))
      .addSubMenu(ui.createMenu('Backlog & Quality')
        .addItem('Backlog Aging ★', 'addBacklogAgingSheet')
        .addItem('Stale Tickets', 'addStaleTicketsSheet')
        .addItem('Reopen Rate', 'addReopenRateSheet'))
      .addSubMenu(ui.createMenu('SLA & Response')
        .addItem('SLA Compliance ★', 'addSLAComplianceSheet')
        .addItem('At-Risk Response ★', 'addAtRiskResponseSheet')
        .addItem('At-Risk Resolution ★', 'addAtRiskResolutionSheet')
        .addItem('First Contact Resolution', 'addFirstContactResolutionSheet')
        .addItem('Response Distribution', 'addResponseDistributionSheet')
        .addItem('Response Trends', 'addResponseTrendsSheet')
        .addItem('Queue Time Analysis', 'addQueueTimeAnalysisSheet'))
      .addSubMenu(ui.createMenu('Team & Staff')
        .addItem('Team Workload ★', 'addTeamWorkloadSheet')
        .addItem('Technician Performance', 'addTechnicianPerformanceSheet')
        .addItem('Functional Area Summary', 'addFunctionalAreaSummarySheet'))
      .addSubMenu(ui.createMenu('Location')
        .addItem('Location Breakdown', 'addLocationBreakdownSheet')
        .addItem('Location Type Comparison', 'addLocationTypeComparisonSheet'))
      .addSubMenu(ui.createMenu('Issue & Requester')
        .addItem('Issue Category Volume', 'addIssueCategoryVolumeSheet')
        .addItem('Priority Analysis', 'addPriorityAnalysisSheet')
        .addItem('Frequent Requesters', 'addFrequentRequestersSheet')))
    .addSubMenu(ui.createMenu('Daily Snapshot')
      .addItem('Capture Now', 'appendDailySnapshot')
      .addItem('Populate Historical Estimates', 'populateHistoricalSnapshots'))
    .addToUi();
}
