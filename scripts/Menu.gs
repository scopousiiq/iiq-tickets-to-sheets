/**
 * IIQ Data Menu
 *
 * Provides menu access to data loader functions.
 *
 * Note: Analytics sheets (MonthlyVolume, BacklogAging, TeamWorkload, etc.)
 * are formula-based and auto-calculate from TicketData/TicketSlaData.
 * No refresh functions needed for analytics.
 */

/**
 * Create the IIQ Data menu when the spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('IIQ Data')
    .addItem('Refresh Teams', 'refreshTeams')
    .addSeparator()
    .addSubMenu(ui.createMenu('Ticket Data')
      .addItem('Continue Loading', 'refreshTicketDataContinue')
      .addItem('Full Reload', 'refreshTicketDataFull')
      .addItem('Refresh Current Year', 'refreshTicketDataCurrentYear')
      .addItem('Show Status', 'showTicketDataStatus'))
    .addSubMenu(ui.createMenu('SLA Data')
      .addItem('Continue Loading', 'refreshSlaDataContinue')
      .addItem('Full Reload', 'refreshSlaDataFull')
      .addItem('Refresh Current Year', 'refreshSlaDataCurrentYear')
      .addItem('Show Status', 'showSlaDataStatus'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Daily Snapshot')
      .addItem('Capture Now', 'appendDailySnapshot')
      .addItem('Populate Historical Estimates', 'populateHistoricalSnapshots'))
    .addToUi();
}
