/**
 * Trigger-Safe Functions for Automated Updates
 *
 * These functions are designed for time-driven triggers.
 * They don't show UI dialogs (which fail in triggers) and log results instead.
 *
 * Recommended Trigger Schedule (after initial data load is complete):
 * 1. triggerTicketDataUpdate  - Daily 5:00 AM
 * 2. triggerSlaDataUpdate     - Daily 5:30 AM
 * 3. triggerDailySnapshot     - Daily 7:00 PM (captures backlog metrics for trending)
 *
 * Note: Analytics sheets (MonthlyVolume, BacklogAging, etc.) are now formula-based
 * and auto-calculate from TicketData/TicketSlaData. No refresh trigger needed.
 *
 * Setup: In Apps Script, go to Triggers (clock icon) and add time-driven triggers.
 */

/**
 * Trigger-safe function to update ticket data
 * Runs the loader without UI prompts, logs results
 */
function triggerTicketDataUpdate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TicketData');

  if (!sheet) {
    logOperation('Trigger', 'ERROR', 'TicketData sheet not found');
    return;
  }

  logOperation('Trigger', 'START', 'Ticket data update triggered');

  try {
    const result = runTicketDataLoader(sheet);
    logOperation('Trigger', 'TICKET_DATA',
      `${result.batchCount} batches, ${result.ticketCount} tickets, ` +
      `complete=${result.complete}, runtime=${(result.runtime/1000).toFixed(1)}s`);
  } catch (error) {
    logOperation('Trigger', 'ERROR', `Ticket data update failed: ${error.message}`);
  }
}

/**
 * Trigger-safe function to update SLA data
 * Runs the loader without UI prompts, logs results
 */
function triggerSlaDataUpdate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TicketSlaData');

  if (!sheet) {
    logOperation('Trigger', 'ERROR', 'TicketSlaData sheet not found');
    return;
  }

  logOperation('Trigger', 'START', 'SLA data update triggered');

  try {
    const result = runSlaDataLoader(sheet);
    logOperation('Trigger', 'SLA_DATA',
      `${result.batchCount} batches, ${result.slaCount} SLAs, ` +
      `complete=${result.complete}, runtime=${(result.runtime/1000).toFixed(1)}s`);
  } catch (error) {
    logOperation('Trigger', 'ERROR', `SLA data update failed: ${error.message}`);
  }
}

/**
 * Trigger-safe function to run a complete daily update
 * Runs ticket update, SLA update, then updates LAST_REFRESH timestamp
 * Use this if you want a single trigger to do everything
 *
 * Note: Analytics sheets are formula-based and auto-calculate.
 * Run triggerDailySnapshot separately at 7 PM for backlog trending.
 */
function triggerDailyUpdate() {
  logOperation('Trigger', 'START', 'Daily update triggered');

  triggerTicketDataUpdate();
  triggerSlaDataUpdate();
  updateLastRefresh();

  logOperation('Trigger', 'COMPLETE', 'Daily update finished');
}

/**
 * Trigger-safe function to continue bulk loading
 * Use this during initial data load to keep the loader running
 * Schedule every 10 minutes until historical data is complete
 */
function triggerBulkLoadContinue() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig();

  // Check if ticket data needs loading (all historical years must be complete)
  const ticketComplete = config.historicalYears.every(year => config[`ticket${year}Complete`]);
  if (!ticketComplete) {
    const sheet = ss.getSheetByName('TicketData');
    if (sheet) {
      logOperation('Trigger', 'BULK_LOAD', 'Continuing ticket data bulk load');
      const result = runTicketDataLoader(sheet);
      logOperation('Trigger', 'TICKET_DATA',
        `${result.batchCount} batches, ${result.ticketCount} tickets, complete=${result.complete}`);
    }
    return; // Focus on tickets first
  }

  // Check if SLA data needs loading (all SLA historical years must be complete)
  const slaComplete = config.slaHistoricalYears.every(year => config[`sla${year}Complete`]);
  if (!slaComplete) {
    const sheet = ss.getSheetByName('TicketSlaData');
    if (sheet) {
      logOperation('Trigger', 'BULK_LOAD', 'Continuing SLA data bulk load');
      const result = runSlaDataLoader(sheet);
      logOperation('Trigger', 'SLA_DATA',
        `${result.batchCount} batches, ${result.slaCount} SLAs, complete=${result.complete}`);
    }
    return;
  }

  logOperation('Trigger', 'BULK_LOAD', 'Historical data load complete - disable this trigger');
}
