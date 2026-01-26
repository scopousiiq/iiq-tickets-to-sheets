/**
 * IIQ Data Menu
 *
 * Provides menu access to data loader and analytics functions.
 *
 * Data Loading:
 * - Ticket Data includes consolidated SLA metrics (35 columns)
 * - SLA data is fetched per-batch during ticket loading
 *
 * Refresh Strategy:
 * - Initial Load: Use "Continue Loading" repeatedly until complete
 * - Ongoing: "Open Ticket Refresh" every 2 hours (updates open + recently closed)
 * - New Tickets: Automatically fetched with open refresh
 * - Weekly: "Full Reload" catches deletions and corrections
 *
 * Analytics:
 * - SLA Compliance: Script-based, reads from consolidated TicketData
 * - Other analytics sheets are formula-based and auto-calculate
 */

/**
 * Create the IIQ Data menu when the spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('IIQ Data')
    .addSubMenu(ui.createMenu('Setup')
      .addItem('Setup Spreadsheet', 'setupSpreadsheet')
      .addItem('Verify Configuration', 'verifyConfiguration'))
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
    .addSubMenu(ui.createMenu('Analytics')
      .addItem('Refresh SLA Compliance', 'refreshSlaCompliance'))
    .addSubMenu(ui.createMenu('Daily Snapshot')
      .addItem('Capture Now', 'appendDailySnapshot')
      .addItem('Populate Historical Estimates', 'populateHistoricalSnapshots'))
    .addToUi();
}
