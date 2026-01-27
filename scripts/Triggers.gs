/**
 * Trigger-Safe Functions for Automated Updates
 *
 * These functions are designed for time-driven triggers.
 * They don't show UI dialogs (which fail in triggers) and log results instead.
 *
 * =============================================================================
 * RECOMMENDED TRIGGER SCHEDULE (Optimized for Freshness + Efficiency)
 * =============================================================================
 *
 * | Function                  | Schedule         | Purpose                           |
 * |---------------------------|------------------|-----------------------------------|
 * | triggerDataContinue       | Every 10 min     | Continue any in-progress loading  |
 * | triggerOpenTicketRefresh  | Every 2 hours    | Start open ticket + SLA refresh   |
 * | triggerNewTickets         | Every 30 min     | Fetch newly created tickets       |
 * | triggerDailySnapshot      | Daily 7:00 PM    | Capture backlog metrics           |
 * | triggerWeeklyFullRefresh  | Weekly Sun 2 AM  | Full reload (catch deletions)     |
 *
 * =============================================================================
 * THE "KEEP THINGS MOVING" TRIGGER (triggerDataContinue)
 * =============================================================================
 *
 * This is the key trigger that ensures work gets completed:
 *
 * 1. If initial load is NOT complete → continues loading historical ticket data
 * 2. If initial load IS complete AND open refresh is in progress → continues it
 * 3. If both are complete → does nothing (no log noise)
 *
 * Safe to leave enabled permanently. It only does work when needed.
 *
 * =============================================================================
 * STRATEGY OVERVIEW
 * =============================================================================
 *
 * INITIAL LOAD:
 * - triggerDataContinue handles this automatically
 * - Loads historical years via pagination, current year via date windowing
 *
 * ONGOING OPERATIONS (after initial load):
 * 1. Open Ticket Refresh (triggered every 2 hours)
 *    - Fetches all open tickets from API
 *    - Fetches tickets closed in last 7 days
 *    - Updates existing rows IN PLACE (no delete/recreate)
 *    - If it times out, triggerDataContinue picks up where it left off
 *
 * 2. New Tickets (every 30 min)
 *    - Uses date windowing to fetch only new tickets since last check
 *    - Very fast - typically processes 0-50 tickets
 *
 * 3. Daily Snapshot (once per day)
 *    - Captures point-in-time backlog metrics
 *    - Used for trend analysis
 *
 * 4. Weekly Full Refresh (once per week)
 *    - Clears recent years and reloads from scratch
 *    - Catches: deleted tickets, data corrections, edge cases
 *    - Preserves old historical data (2+ years old)
 *
 * =============================================================================
 * DATA FRESHNESS
 * =============================================================================
 *
 * With this schedule:
 * - Open ticket SLA data: max 2 hours stale
 * - New tickets: appear within 30 minutes
 * - Status changes: captured within 2 hours
 * - Deletions/corrections: captured weekly
 *
 * Setup: In Apps Script, go to Triggers (clock icon) and add time-driven triggers.
 */

// =============================================================================
// PRIMARY TRIGGERS (Use these for ongoing operations)
// =============================================================================

/**
 * Refresh open tickets and recently closed tickets
 * Schedule: Every 2 hours
 *
 * This is the primary trigger for keeping data fresh:
 * - Updates all open tickets (status, SLA timers, assignments)
 * - Catches tickets that were recently closed
 * - Uses in-place updates for efficiency
 * - Appends new tickets not already in sheet
 *
 * NOTE: Skips if initial data load is not complete.
 * NOTE: For large districts, this may not complete in one run.
 * Use triggerDataContinue (every 10 min) to ensure completion.
 */
function triggerOpenTicketRefresh() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TicketData');
  const config = getConfig();

  if (!sheet) {
    logOperation('Trigger', 'ERROR', 'TicketData sheet not found');
    return;
  }

  // Only run if initial load is complete
  if (!isTicketLoadingComplete(config)) {
    logOperation('Trigger', 'SKIP', 'Open ticket refresh skipped - initial load not complete. Run triggerDataContinue to complete initial load.');
    return;
  }

  logOperation('Trigger', 'START', 'Open ticket refresh started');

  try {
    const result = runOpenTicketRefresh(sheet);
    logOperation('Trigger', 'OPEN_REFRESH',
      `Updated ${result.updatedCount}, appended ${result.appendedCount}, ` +
      `${result.ticketCount} total tickets, ${result.batchCount} batches, ` +
      `complete=${result.complete}, runtime=${(result.runtime/1000).toFixed(1)}s`);
  } catch (error) {
    logOperation('Trigger', 'ERROR', `Open ticket refresh failed: ${error.message}`);
  }
}

/**
 * @deprecated Use triggerDataContinue instead - it handles both initial load AND open refresh
 */
function triggerOpenRefreshContinue() {
  logOperation('Trigger', 'DEPRECATED', 'triggerOpenRefreshContinue called - use triggerDataContinue instead');
  triggerDataContinue();
}

/**
 * Fetch newly created tickets (incremental)
 * Schedule: Every 30 minutes
 *
 * Uses date windowing to fetch only tickets created since last check.
 * Very efficient - typically processes only a handful of tickets.
 */
function triggerNewTickets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TicketData');
  const config = getConfig();

  if (!sheet) {
    logOperation('Trigger', 'ERROR', 'TicketData sheet not found');
    return;
  }

  // Only run if initial load is complete
  if (!isTicketLoadingComplete(config)) {
    logOperation('Trigger', 'SKIP', 'New tickets check skipped - initial load not complete');
    return;
  }

  if (!config.currentYear) {
    logOperation('Trigger', 'SKIP', 'No current year configured for incremental updates');
    return;
  }

  logOperation('Trigger', 'START', 'New tickets check started');

  try {
    const result = runNewTicketsCheck(sheet, config);
    if (result.count > 0) {
      logOperation('Trigger', 'NEW_TICKETS',
        `Found and added ${result.count} new tickets, runtime=${(result.runtime/1000).toFixed(1)}s`);
    } else {
      logOperation('Trigger', 'NEW_TICKETS', 'No new tickets found');
    }
  } catch (error) {
    logOperation('Trigger', 'ERROR', `New tickets check failed: ${error.message}`);
  }
}

/**
 * Weekly full refresh - clears recent data and reloads
 * Schedule: Weekly (Sunday 2:00 AM recommended)
 *
 * This catches edge cases that incremental updates might miss:
 * - Deleted tickets
 * - Data corrections in iiQ
 * - Any sync issues
 *
 * Preserves historical data (2+ years old) to minimize reload time.
 */
function triggerWeeklyFullRefresh() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig();
  const currentYear = new Date().getFullYear();
  const ticketSheet = ss.getSheetByName('TicketData');

  // Preserve data from 2 years ago and older
  const historicalCutoff = currentYear - 2;

  logOperation('Trigger', 'START',
    `Weekly full refresh triggered. Preserving data from ${historicalCutoff} and earlier.`);

  try {
    // Step 1: Filter and rewrite TicketData (preserve historical years)
    if (ticketSheet) {
      const ticketResult = filterAndRewriteTicketData(ticketSheet, historicalCutoff);
      logOperation('Trigger', 'WEEKLY_RESET',
        `Preserved ${ticketResult.keptCount} historical rows, cleared ${ticketResult.clearedCount} recent rows`);
    }

    // Step 2: Reset progress only for recent ticket years
    config.historicalYears.forEach(year => {
      if (year > historicalCutoff) {
        resetYearProgress(year);
        logOperation('Trigger', 'PROGRESS_RESET', `Reset ticket progress for ${year}`);
      }
    });
    if (config.currentYear && config.currentYear > historicalCutoff) {
      resetCurrentYearProgress(config.currentYear);
      logOperation('Trigger', 'PROGRESS_RESET', `Reset ticket progress for ${config.currentYear} (current year)`);
    }

    // Step 3: Reset open refresh progress
    resetOpenRefreshProgress();

    logOperation('Trigger', 'WEEKLY_RESET',
      `Reset complete. Will reload years > ${historicalCutoff}. Run triggerDataContinue to reload.`);

    // Step 4: Start reloading immediately
    if (!ticketSheet) {
      logOperation('Trigger', 'ERROR', 'TicketData sheet not found - cannot reload');
      return;
    }
    const result = runTicketDataLoader(ticketSheet);
    logOperation('Trigger', 'WEEKLY_RELOAD',
      `Initial reload: ${result.batchCount} batches, ${result.ticketCount} tickets, ` +
      `complete=${result.complete}`);

  } catch (error) {
    logOperation('Trigger', 'ERROR', `Weekly full refresh failed: ${error.message}`);
  }
}

// =============================================================================
// INITIAL LOAD TRIGGER (Use for first-time setup or after weekly reset)
// =============================================================================

/**
 * Continue any in-progress data loading
 * Schedule: Every 10 minutes
 *
 * This is your "keep things moving" trigger. It handles two scenarios:
 *
 * 1. INITIAL LOAD NOT COMPLETE:
 *    - Continues loading historical ticket data
 *    - Run until all years show "Complete"
 *
 * 2. INITIAL LOAD COMPLETE + OPEN REFRESH IN PROGRESS:
 *    - Continues the open ticket refresh
 *    - Useful for large districts where 2-hour refresh doesn't finish in one run
 *
 * 3. BOTH COMPLETE:
 *    - Does nothing (minimal log to confirm it ran)
 *
 * You can leave this trigger enabled permanently - it only runs when needed.
 */
function triggerDataContinue() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig();

  // Priority 1: Check if initial ticket load needs continuing
  const ticketsComplete = isTicketLoadingComplete(config);
  if (!ticketsComplete) {
    let sheet = ss.getSheetByName('TicketData');
    if (!sheet) {
      sheet = createTicketSheet(ss);
    }

    logOperation('Trigger', 'START', 'Continuing ticket data loading (with SLA)');

    try {
      const result = runTicketDataLoader(sheet);
      logOperation('Trigger', 'TICKET_DATA',
        `${result.batchCount} batches, ${result.ticketCount} tickets, ` +
        `complete=${result.complete}, runtime=${(result.runtime/1000).toFixed(1)}s`);
    } catch (error) {
      logOperation('Trigger', 'ERROR', `Ticket data continue failed: ${error.message}`);
    }
    return;
  }

  // Priority 2: Check if open ticket refresh needs continuing
  // Read progress directly from config sheet for most accurate state
  const progress = getOpenRefreshProgress();
  const savedDate = config.openRefreshDate || '';
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  // Log diagnostic info about the state we found
  logOperation('Trigger', 'CHECK',
    `Open refresh state: date=${savedDate}, today=${today}, openPage=${progress.openPage}, ` +
    `openComplete=${progress.openComplete}, closedPage=${progress.closedPage}, ` +
    `closedComplete=${progress.closedComplete}`);

  // Check if there's an incomplete refresh to continue
  // Allow continuing yesterday's refresh if it didn't complete (handles day boundary edge case)
  const refreshInProgress = progress.openPage >= 0 || progress.closedPage >= 0;
  const refreshJustStarted = savedDate === today && progress.openPage < 0 && !progress.openComplete;
  const refreshIncomplete = !progress.openComplete || !progress.closedComplete;

  if ((refreshInProgress || refreshJustStarted) && refreshIncomplete) {
    const sheet = ss.getSheetByName('TicketData');
    if (!sheet) {
      logOperation('Trigger', 'ERROR', 'TicketData sheet not found');
      return;
    }

    logOperation('Trigger', 'CONTINUE',
      `Continuing open ticket refresh (savedDate=${savedDate}, today=${today})`);

    try {
      const result = runOpenTicketRefresh(sheet);
      logOperation('Trigger', 'OPEN_REFRESH',
        `Updated ${result.updatedCount}, appended ${result.appendedCount}, ` +
        `${result.ticketCount} total tickets, ${result.batchCount} batches, ` +
        `complete=${result.complete}, runtime=${(result.runtime/1000).toFixed(1)}s`);
    } catch (error) {
      logOperation('Trigger', 'ERROR', `Open ticket refresh continue failed: ${error.message}`);
    }
    return;
  }

  // Nothing to do - log why for diagnostics
  if (progress.openComplete && progress.closedComplete) {
    logOperation('Trigger', 'IDLE', `Open refresh already complete for ${savedDate}`);
  } else if (!savedDate) {
    logOperation('Trigger', 'IDLE', 'No open refresh has been started yet (waiting for triggerOpenTicketRefresh)');
  } else {
    logOperation('Trigger', 'IDLE',
      `Unexpected state - date=${savedDate}, inProgress=${refreshInProgress}, ` +
      `justStarted=${refreshJustStarted}, incomplete=${refreshIncomplete}`);
  }
}

// =============================================================================
// HELPER FUNCTIONS
// =============================================================================

/**
 * Check if all ticket loading is complete
 */
function isTicketLoadingComplete(config) {
  // Check historical years
  const historicalComplete = config.historicalYears.every(year => config[`ticket${year}Complete`]);
  if (!historicalComplete) return false;

  // Check current year (has data if lastFetch is set)
  if (config.currentYear) {
    const lastFetch = config[`ticket${config.currentYear}LastFetch`];
    if (!lastFetch) return false;
  }

  return true;
}

/**
 * Run incremental check for new tickets
 * Uses date windowing from last fetch timestamp
 */
function runNewTicketsCheck(sheet, config) {
  const startTime = Date.now();
  const currentYear = config.currentYear;
  const lastFetch = config[`ticket${currentYear}LastFetch`];
  const batchSize = config.ticketBatchSize;

  // Build date range from last fetch to now
  let startDate;
  if (lastFetch) {
    const lastDate = new Date(lastFetch);
    startDate = formatDateForApi(lastDate);
  } else {
    startDate = `01/01/${currentYear}`;
  }

  const endDate = formatDateForApi(new Date());
  const dateRange = `daterange:${startDate}-${endDate}`;

  // Fetch new tickets
  const response = searchTickets([
    { Facet: 'createddate', Value: dateRange }
  ], 0, batchSize, { field: 'TicketCreatedDate', direction: 'asc' });

  if (!response.Items || response.Items.length === 0) {
    return { count: 0, runtime: Date.now() - startTime };
  }

  // Filter to only tickets created AFTER last fetch
  let tickets = response.Items;
  if (lastFetch) {
    const lastFetchTime = new Date(lastFetch).getTime();
    tickets = tickets.filter(t => new Date(t.CreatedDate).getTime() > lastFetchTime);
  }

  if (tickets.length === 0) {
    // Update timestamp even if no new tickets
    const lastTicket = response.Items[response.Items.length - 1];
    updateConfigValue(`TICKET_${currentYear}_LAST_FETCH`, lastTicket.CreatedDate);
    return { count: 0, runtime: Date.now() - startTime };
  }

  // Fetch SLA data for new tickets
  const ticketIds = tickets.map(t => t.TicketId).filter(id => id);
  const slaMap = fetchSlaForTicketIds(ticketIds);

  // Write to sheet
  const now = new Date();
  tickets.sort((a, b) => new Date(a.CreatedDate) - new Date(b.CreatedDate));
  const rows = tickets.map(ticket => extractTicketRow(ticket, now, currentYear, slaMap));
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, rows.length, 35).setValues(rows);

  // Update last fetch timestamp
  const lastTicket = tickets[tickets.length - 1];
  updateConfigValue(`TICKET_${currentYear}_LAST_FETCH`, lastTicket.CreatedDate);

  return { count: tickets.length, runtime: Date.now() - startTime };
}

/**
 * Filter TicketData to keep only historical rows (year <= cutoff)
 * Uses filter-and-rewrite for efficiency (much faster than row deletion)
 */
function filterAndRewriteTicketData(sheet, historicalCutoff) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return { keptCount: 0, clearedCount: 0 };
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 35).getValues();
  const rowsToKeep = [];
  let clearedCount = 0;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const year = row[3]; // Column D

    if (year && year <= historicalCutoff) {
      rowsToKeep.push(row);
    } else {
      clearedCount++;
    }
  }

  sheet.getRange(2, 1, lastRow - 1, 35).clear();

  if (rowsToKeep.length > 0) {
    sheet.getRange(2, 1, rowsToKeep.length, 35).setValues(rowsToKeep);
  }

  return { keptCount: rowsToKeep.length, clearedCount: clearedCount };
}

/**
 * Create the TicketData sheet with headers (35 columns including SLA metrics)
 */
function createTicketSheet(ss) {
  const sheet = ss.insertSheet('TicketData');
  sheet.getRange(1, 1, 1, 35).setValues([[
    'TicketId', 'TicketNumber', 'Subject', 'Year',
    'CreatedDate', 'ModifiedDate', 'ClosedDate', 'IsClosed',
    'Status', 'TeamId', 'TeamName', 'LocationId', 'LocationName', 'LocationType',
    'OwnerId', 'OwnerName', 'AgeDays', 'Priority', 'IsPastDue', 'DueDate',
    'SlaId', 'SlaName', 'IssueCategoryId', 'IssueCategoryName',
    'IssueTypeId', 'IssueTypeName', 'RequesterId', 'RequesterName',
    'ResponseThreshold', 'ResponseActual', 'ResponseBreach',
    'ResolutionThreshold', 'ResolutionActual', 'ResolutionBreach', 'IsRunning'
  ]]);
  sheet.setFrozenRows(1);
  return sheet;
}

// Note: triggerDailySnapshot is defined in DailySnapshot.gs

// =============================================================================
// DEPRECATED FUNCTIONS - Kept for backward compatibility
// =============================================================================

/**
 * @deprecated Use triggerWeeklyFullRefresh instead
 */
function triggerDataReset() {
  logOperation('Trigger', 'DEPRECATED', 'triggerDataReset called - consider using triggerWeeklyFullRefresh');
  triggerWeeklyFullRefresh();
}

/**
 * @deprecated Use triggerDataContinue instead
 */
function triggerTicketDataReset() {
  logOperation('Trigger', 'DEPRECATED', 'triggerTicketDataReset called - use triggerDataContinue');
  triggerDataContinue();
}

/**
 * @deprecated Use triggerDataContinue instead
 */
function triggerTicketDataContinue() {
  logOperation('Trigger', 'DEPRECATED', 'triggerTicketDataContinue called - use triggerDataContinue');
  triggerDataContinue();
}

/**
 * @deprecated SLA data is now consolidated into TicketData
 */
function triggerSlaDataReset() {
  logOperation('Trigger', 'DEPRECATED', 'triggerSlaDataReset no longer needed - SLA data is now part of TicketData');
}

/**
 * @deprecated SLA data is now consolidated into TicketData
 */
function triggerSlaDataContinue() {
  logOperation('Trigger', 'DEPRECATED', 'triggerSlaDataContinue no longer needed - SLA data is now part of TicketData');
}

// =============================================================================
// TRIGGER MANAGEMENT (Setup/Remove automated triggers)
// =============================================================================

/**
 * Set up all recommended triggers for automated data refresh
 * Called from: iiQ Data > Setup > Setup Automated Triggers
 */
function setupAutomatedTriggers() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Setup Automated Triggers',
    'This will create the following time-driven triggers:\n\n' +
    '• triggerDataContinue - Every 10 minutes\n' +
    '  (Continues any in-progress loading)\n\n' +
    '• triggerOpenTicketRefresh - Every 2 hours\n' +
    '  (Refreshes open tickets and SLA data)\n\n' +
    '• triggerNewTickets - Every 30 minutes\n' +
    '  (Catches newly created tickets)\n\n' +
    '• triggerDailySnapshot - Daily at 7:00 PM\n' +
    '  (Captures backlog metrics for trending)\n\n' +
    '• triggerWeeklyFullRefresh - Weekly Sunday 2:00 AM\n' +
    '  (Full reload to catch deletions)\n\n' +
    'Any existing triggers for these functions will be replaced.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  try {
    const created = [];
    const errors = [];

    // Define trigger configurations
    const triggerConfigs = [
      {
        functionName: 'triggerDataContinue',
        type: 'everyMinutes',
        interval: 10,
        description: 'Every 10 minutes'
      },
      {
        functionName: 'triggerOpenTicketRefresh',
        type: 'everyHours',
        interval: 2,
        description: 'Every 2 hours'
      },
      {
        functionName: 'triggerNewTickets',
        type: 'everyMinutes',
        interval: 30,
        description: 'Every 30 minutes'
      },
      {
        functionName: 'triggerDailySnapshot',
        type: 'daily',
        hour: 19,
        description: 'Daily at 7:00 PM'
      },
      {
        functionName: 'triggerWeeklyFullRefresh',
        type: 'weekly',
        dayOfWeek: ScriptApp.WeekDay.SUNDAY,
        hour: 2,
        description: 'Weekly Sunday 2:00 AM'
      }
    ];

    // Remove existing triggers for these functions
    const existingTriggers = ScriptApp.getProjectTriggers();
    const functionNames = triggerConfigs.map(c => c.functionName);

    existingTriggers.forEach(trigger => {
      if (functionNames.includes(trigger.getHandlerFunction())) {
        ScriptApp.deleteTrigger(trigger);
      }
    });

    // Create new triggers
    triggerConfigs.forEach(config => {
      try {
        let trigger;

        switch (config.type) {
          case 'everyMinutes':
            trigger = ScriptApp.newTrigger(config.functionName)
              .timeBased()
              .everyMinutes(config.interval)
              .create();
            break;

          case 'everyHours':
            trigger = ScriptApp.newTrigger(config.functionName)
              .timeBased()
              .everyHours(config.interval)
              .create();
            break;

          case 'daily':
            trigger = ScriptApp.newTrigger(config.functionName)
              .timeBased()
              .everyDays(1)
              .atHour(config.hour)
              .create();
            break;

          case 'weekly':
            trigger = ScriptApp.newTrigger(config.functionName)
              .timeBased()
              .onWeekDay(config.dayOfWeek)
              .atHour(config.hour)
              .create();
            break;
        }

        created.push(`${config.functionName} (${config.description})`);
      } catch (e) {
        errors.push(`${config.functionName}: ${e.message}`);
      }
    });

    // Show results
    let message = '';
    if (created.length > 0) {
      message += 'Created triggers:\n• ' + created.join('\n• ');
    }
    if (errors.length > 0) {
      message += '\n\nErrors:\n• ' + errors.join('\n• ');
    }

    logOperation('Setup', 'TRIGGERS', `Created ${created.length} triggers, ${errors.length} errors`);

    ui.alert('Triggers Created', message, ui.ButtonSet.OK);

  } catch (error) {
    logOperation('Setup', 'ERROR', `Failed to create triggers: ${error.message}`);
    ui.alert('Error', 'Failed to create triggers: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Remove all automated triggers for this project
 * Called from: iiQ Data > Setup > Remove Automated Triggers
 */
function removeAutomatedTriggers() {
  const ui = SpreadsheetApp.getUi();

  const triggers = ScriptApp.getProjectTriggers();

  if (triggers.length === 0) {
    ui.alert('No Triggers', 'There are no triggers to remove.', ui.ButtonSet.OK);
    return;
  }

  // Build list of current triggers
  const triggerList = triggers.map(t =>
    `• ${t.getHandlerFunction()}`
  ).join('\n');

  const response = ui.alert(
    'Remove Automated Triggers',
    `This will remove ALL ${triggers.length} triggers for this project:\n\n${triggerList}\n\nContinue?`,
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  try {
    triggers.forEach(trigger => {
      ScriptApp.deleteTrigger(trigger);
    });

    logOperation('Setup', 'TRIGGERS', `Removed ${triggers.length} triggers`);
    ui.alert('Triggers Removed', `Successfully removed ${triggers.length} triggers.`, ui.ButtonSet.OK);

  } catch (error) {
    logOperation('Setup', 'ERROR', `Failed to remove triggers: ${error.message}`);
    ui.alert('Error', 'Failed to remove triggers: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Show current trigger status
 * Called from: iiQ Data > Setup > View Trigger Status
 */
function viewTriggerStatus() {
  const ui = SpreadsheetApp.getUi();
  const triggers = ScriptApp.getProjectTriggers();

  if (triggers.length === 0) {
    ui.alert('Trigger Status',
      'No automated triggers are configured.\n\n' +
      'Use "Setup Automated Triggers" to create the recommended triggers.',
      ui.ButtonSet.OK);
    return;
  }

  // Build status report
  const triggerInfo = triggers.map(t => {
    const funcName = t.getHandlerFunction();
    const type = t.getEventType();
    let schedule = '';

    // Try to describe the schedule
    if (type === ScriptApp.EventType.CLOCK) {
      const source = t.getTriggerSource();
      schedule = 'Time-based';
    }

    return `• ${funcName}\n  Type: ${schedule}`;
  }).join('\n\n');

  ui.alert('Trigger Status',
    `${triggers.length} trigger(s) configured:\n\n${triggerInfo}`,
    ui.ButtonSet.OK);
}
