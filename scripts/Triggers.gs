/**
 * Trigger-Safe Functions for Automated Updates
 *
 * These functions are designed for time-driven triggers.
 * They don't show UI dialogs (which fail in triggers) and log results instead.
 *
 * =============================================================================
 * SCHOOL YEAR MODEL
 * =============================================================================
 *
 * Each spreadsheet contains ONE school year's data (e.g., 2025-2026).
 * The school year is configured via SCHOOL_YEAR and SCHOOL_YEAR_START in Config.
 *
 * IMPORTANT: Historical vs Current School Year
 * - HISTORICAL school years (past): Data becomes STATIC once loaded completely.
 *   No ongoing refreshes are needed or performed.
 * - CURRENT school year (in progress): Data is actively refreshed to capture
 *   new tickets, status changes, SLA updates, and deletions.
 *
 * =============================================================================
 * RECOMMENDED TRIGGER SCHEDULE
 * =============================================================================
 *
 * | Function                  | Schedule         | Runs For                | Purpose                    |
 * |---------------------------|------------------|-------------------------|----------------------------|
 * | triggerDataContinue       | Every 10 min     | All                     | Continue in-progress load  |
 * | triggerOpenTicketRefresh  | Every 2 hours    | Current + hist w/ open  | Refresh open tickets + SLA |
 * | triggerNewTickets         | Every 30 min     | Current                 | Fetch newly created tickets|
 * | triggerDailySnapshot      | Daily 7:00 PM    | Current                 | Capture backlog metrics    |
 * | triggerWeeklyFullRefresh  | Weekly Sun 2 AM  | Current                 | Full reload for deletions  |
 *
 * For HISTORICAL school years, only triggerDataContinue is needed (to complete
 * initial load). Once loaded, no triggers do any work - they skip automatically.
 *
 * =============================================================================
 * THE "KEEP THINGS MOVING" TRIGGER (triggerDataContinue)
 * =============================================================================
 *
 * This is the key trigger that ensures work gets completed:
 *
 * 1. If initial load is NOT complete → continues loading school year ticket data
 * 2. If initial load IS complete AND open refresh is in progress → continues it
 * 3. If both are complete → does nothing (minimal logging)
 *
 * Safe to leave enabled permanently. It only does work when needed.
 *
 * =============================================================================
 * STRATEGY BY SCHOOL YEAR TYPE
 * =============================================================================
 *
 * HISTORICAL SCHOOL YEAR:
 * - Run triggerDataContinue until TICKET_COMPLETE = TRUE
 * - Open ticket refresh continues until all tickets are closed
 *   (handles tickets opened near end of school year)
 * - Once all tickets closed, data becomes STATIC
 * - Weekly refresh, new tickets, and daily snapshot skip automatically
 *
 * CURRENT SCHOOL YEAR:
 * - Initial load via triggerDataContinue
 * - Ongoing operations after initial load:
 *   1. Open Ticket Refresh (every 2 hours) - updates open tickets + SLA
 *   2. New Tickets (every 30 min) - catches newly created tickets
 *   3. Daily Snapshot (7 PM) - captures backlog metrics for trending
 *   4. Weekly Full Refresh (Sunday 2 AM) - catches deletions/corrections
 *
 * =============================================================================
 * DATA FRESHNESS (Current School Year Only)
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
 * For HISTORICAL school years: Continues running until all tickets are closed.
 * A ticket opened June 29th shouldn't be stuck "open" forever just because
 * July 1st arrived. Once all tickets in the sheet are closed, skips.
 *
 * NOTE: Skips if initial data load is not complete.
 * NOTE: For large districts, this may not complete in one run.
 * Use triggerDataContinue (every 10 min) to ensure completion.
 */
function triggerOpenTicketRefresh() {
  // SAFETY: Try to acquire lock - skip if another operation is running
  const lock = tryAcquireScriptLock();
  if (!lock) {
    logOperation('Trigger', 'SKIP', 'Open ticket refresh skipped - another operation is in progress');
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('TicketData');
    const config = getConfig();
    const reloadInProgress = getCurrentYearReloadInProgress();

    if (!sheet) {
      logOperation('Trigger', 'ERROR', 'TicketData sheet not found');
      return;
    }

    // For historical school years, only run if there are still open tickets
    if (!isSchoolYearCurrent(config)) {
      const openCount = countOpenTickets(sheet);
      if (openCount === 0) {
        logOperation('Trigger', 'SKIP',
          `Open ticket refresh skipped - school year ${config.schoolYear} is historical and all tickets are closed`);
        return;
      }
      logOperation('Trigger', 'INFO',
        `Historical school year ${config.schoolYear} still has ${openCount} open tickets - refreshing`);
    }

    if (reloadInProgress) {
      logOperation('Trigger', 'SKIP', 'Open ticket refresh skipped - current year reload in progress');
      return;
    }

    // Only run if initial load is complete
    if (!isTicketLoadingComplete(config)) {
      logOperation('Trigger', 'SKIP', 'Open ticket refresh skipped - initial load not complete. Run triggerDataContinue to complete initial load.');
      return;
    }

    // Reset progress to start a new cycle (each trigger run = new cycle)
    // This fetches tickets modified since the last successful run
    resetOpenRefreshProgress();

    logOperation('Trigger', 'START', 'Open ticket refresh started');

    try {
      const result = runOpenTicketRefresh(sheet);
      const skippedInfo = result.skippedCount ? `, skipped ${result.skippedCount}` : '';
      logOperation('Trigger', 'OPEN_REFRESH',
        `Updated ${result.updatedCount}, appended ${result.appendedCount}${skippedInfo}, ` +
        `${result.ticketCount} total tickets, ${result.batchCount} batches, ` +
        `complete=${result.complete}, runtime=${(result.runtime/1000).toFixed(1)}s`);
    } catch (error) {
      logOperation('Trigger', 'ERROR', `Open ticket refresh failed: ${error.message}`);
    }

  } finally {
    releaseScriptLock(lock);
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
 * Only runs for current school year (historical school years don't need incremental updates).
 */
function triggerNewTickets() {
  // SAFETY: Try to acquire lock - skip if another operation is running
  const lock = tryAcquireScriptLock();
  if (!lock) {
    logOperation('Trigger', 'SKIP', 'New tickets check skipped - another operation is in progress');
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('TicketData');
    const config = getConfig();
    const reloadInProgress = getCurrentYearReloadInProgress();

    if (!sheet) {
      logOperation('Trigger', 'ERROR', 'TicketData sheet not found');
      return;
    }

    if (reloadInProgress) {
      logOperation('Trigger', 'SKIP', 'New tickets check skipped - school year reload in progress');
      return;
    }

    // Only run if initial load is complete
    if (!isTicketLoadingComplete(config)) {
      logOperation('Trigger', 'SKIP', 'New tickets check skipped - initial load not complete');
      return;
    }

    // Only run for current school year
    if (!isSchoolYearCurrent(config)) {
      logOperation('Trigger', 'SKIP', 'New tickets check skipped - school year is historical');
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

  } finally {
    releaseScriptLock(lock);
  }
}

/**
 * Weekly full refresh - clears school year data and reloads
 * Schedule: Weekly (Sunday 2:00 AM recommended)
 *
 * ONLY runs for CURRENT school year. Historical school years are static
 * once loaded and should not be refreshed.
 *
 * This catches edge cases that incremental updates might miss:
 * - Deleted tickets
 * - Data corrections in iiQ
 * - Any sync issues
 */
function triggerWeeklyFullRefresh() {
  // SAFETY: Try to acquire lock - skip if another operation is running
  const lock = tryAcquireScriptLock();
  if (!lock) {
    logOperation('Trigger', 'SKIP', 'Weekly full refresh skipped - another operation is in progress');
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig();
    const ticketSheet = ss.getSheetByName('TicketData');

    if (!config.schoolYear) {
      logOperation('Trigger', 'ERROR', 'No SCHOOL_YEAR configured');
      return;
    }

    // Skip for historical school years - their data is static
    if (!isSchoolYearCurrent(config)) {
      logOperation('Trigger', 'SKIP',
        `Weekly refresh skipped - school year ${config.schoolYear} is historical (data is static)`);
      return;
    }

    logOperation('Trigger', 'START',
      `Weekly full refresh triggered for current school year ${config.schoolYear}.`);

    try {
      // Step 1: Clear all ticket data for this school year
      if (ticketSheet) {
        const lastRow = ticketSheet.getLastRow();
        if (lastRow > 1) {
          ticketSheet.getRange(2, 1, lastRow - 1, 36).clear();
          logOperation('Trigger', 'WEEKLY_RESET', `Cleared ${lastRow - 1} ticket rows`);
        }
      }

      // Step 2: Reset school year progress
      resetSchoolYearProgress();
      logOperation('Trigger', 'PROGRESS_RESET', `Reset ticket progress for ${config.schoolYear}`);

      // Step 3: Reset open refresh progress
      resetOpenRefreshProgress();

      // Step 4: Flag reload in progress
      setCurrentYearReloadInProgress(true);

      logOperation('Trigger', 'WEEKLY_RESET',
        `Reset complete. Reloading school year ${config.schoolYear}.`);

      // Step 5: Start reloading immediately
      if (!ticketSheet) {
        logOperation('Trigger', 'ERROR', 'TicketData sheet not found - cannot reload');
        return;
      }
      const result = runTicketDataLoader(ticketSheet);
      logOperation('Trigger', 'WEEKLY_RELOAD',
        `Initial reload: ${result.batchCount} batches, ${result.ticketCount} tickets, ` +
        `complete=${result.complete}`);
      if (result.complete) {
        setCurrentYearReloadInProgress(false);
      }

    } catch (error) {
      logOperation('Trigger', 'ERROR', `Weekly full refresh failed: ${error.message}`);
    }

  } finally {
    releaseScriptLock(lock);
  }
}

// =============================================================================
// INITIAL LOAD TRIGGER (Use for first-time setup or after a reset)
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
  // SAFETY: Try to acquire lock - skip if another operation is running
  const lock = tryAcquireScriptLock();
  if (!lock) {
    logOperation('Trigger', 'SKIP', 'Data continue skipped - another operation is in progress');
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig();

    // Priority 1: Check if initial ticket load needs continuing
    const reloadInProgress = getCurrentYearReloadInProgress();
    const ticketsComplete = reloadInProgress ? false : isTicketLoadingComplete(config);
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
        if (reloadInProgress && result.complete) {
          setCurrentYearReloadInProgress(false);
        }
      } catch (error) {
        logOperation('Trigger', 'ERROR', `Ticket data continue failed: ${error.message}`);
      }
      return;
    }

    // Priority 2: Check if open ticket refresh needs continuing
    const progress = getOpenRefreshProgress();

    // Log diagnostic info about the state we found
    logOperation('Trigger', 'CHECK',
      `Open refresh state: page=${progress.page}, complete=${progress.complete}, lastRun=${progress.lastRun || 'never'}`);

    // Check if there's an incomplete refresh to continue (page >= 0 means it started but didn't finish)
    const refreshInProgress = progress.page >= 0 && !progress.complete;

    if (refreshInProgress) {
      const sheet = ss.getSheetByName('TicketData');
      if (!sheet) {
        logOperation('Trigger', 'ERROR', 'TicketData sheet not found');
        return;
      }

      logOperation('Trigger', 'CONTINUE', 'Continuing open ticket refresh');

      try {
        const result = runOpenTicketRefresh(sheet);
        const skippedInfo = result.skippedCount ? `, skipped ${result.skippedCount}` : '';
        logOperation('Trigger', 'OPEN_REFRESH',
          `Updated ${result.updatedCount}, appended ${result.appendedCount}${skippedInfo}, ` +
          `${result.ticketCount} total tickets, ${result.batchCount} batches, ` +
          `complete=${result.complete}, runtime=${(result.runtime/1000).toFixed(1)}s`);
      } catch (error) {
        logOperation('Trigger', 'ERROR', `Open ticket refresh continue failed: ${error.message}`);
      }
      return;
    }

    // Nothing to do
    if (progress.complete) {
      logOperation('Trigger', 'IDLE', `Open refresh complete, last run: ${progress.lastRun || 'never'}`);
    } else if (!progress.lastRun && progress.page < 0) {
      logOperation('Trigger', 'IDLE', 'No open refresh has been started yet (waiting for triggerOpenTicketRefresh)');
    } else {
      logOperation('Trigger', 'IDLE', `Idle - page=${progress.page}, complete=${progress.complete}`);
    }

  } finally {
    releaseScriptLock(lock);
  }
}

// =============================================================================
// HELPER FUNCTIONS
// =============================================================================

/**
 * Check if school year ticket loading is complete
 */
function isTicketLoadingComplete(config) {
  // Check if pagination loading is complete
  if (!config.ticketComplete) return false;

  // For current school year, also need at least one lastFetch
  if (isSchoolYearCurrent(config)) {
    if (!config.ticketLastFetch) return false;
  }

  return true;
}

/**
 * Count open tickets in the TicketData sheet
 * Used to determine if historical school years still need refresh
 */
function countOpenTickets(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;

  // Column I is IsClosed (index 9, or column 9 in 1-indexed)
  const isClosedData = sheet.getRange(2, 9, lastRow - 1, 1).getValues();

  let openCount = 0;
  for (let i = 0; i < isClosedData.length; i++) {
    if (isClosedData[i][0] === 'No') {
      openCount++;
    }
  }

  return openCount;
}

/**
 * Persistent flag to keep current-year reload running across trigger invocations
 */
function getCurrentYearReloadInProgress() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty('CURRENT_YEAR_RELOAD_IN_PROGRESS') === 'TRUE';
}

function setCurrentYearReloadInProgress(inProgress) {
  const props = PropertiesService.getScriptProperties();
  if (inProgress) {
    props.setProperty('CURRENT_YEAR_RELOAD_IN_PROGRESS', 'TRUE');
  } else {
    props.deleteProperty('CURRENT_YEAR_RELOAD_IN_PROGRESS');
  }
}

/**
 * Run incremental check for new tickets
 * Uses date windowing from last fetch timestamp
 */
function runNewTicketsCheck(sheet, config) {
  const startTime = Date.now();
  const lastFetch = config.ticketLastFetch;
  const batchSize = config.ticketBatchSize;

  // Get school year dates
  const dates = getSchoolYearDates(config);
  if (!dates) {
    logOperation('Trigger', 'ERROR', 'Invalid school year configuration');
    return { count: 0, runtime: Date.now() - startTime };
  }

  // Build date range from last fetch to now
  let startDate;
  if (lastFetch) {
    const lastDate = new Date(lastFetch);
    startDate = formatDateForApi(lastDate);
  } else {
    startDate = formatDateForApi(dates.startDate);
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
    updateConfigValue('TICKET_LAST_FETCH', lastTicket.CreatedDate);
    return { count: 0, runtime: Date.now() - startTime };
  }

  // Fetch SLA data for new tickets
  const ticketIds = tickets.map(t => t.TicketId).filter(id => id);
  const slaMap = fetchSlaForTicketIds(ticketIds);

  // Write to sheet using school year string
  const now = new Date();
  tickets.sort((a, b) => new Date(a.CreatedDate) - new Date(b.CreatedDate));
  const rows = tickets.map(ticket => extractTicketRow(ticket, now, config.schoolYear, slaMap));
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, rows.length, 36).setValues(rows);

  // Update last fetch timestamp
  const lastTicket = tickets[tickets.length - 1];
  updateConfigValue('TICKET_LAST_FETCH', lastTicket.CreatedDate);

  return { count: tickets.length, runtime: Date.now() - startTime };
}

/**
 * @deprecated School year model uses single school year per spreadsheet
 * Filter TicketData to keep only historical rows (year <= cutoff)
 */
function filterAndRewriteTicketData(sheet, historicalCutoff) {
  // No longer used in school year model - each spreadsheet has one school year
  logOperation('Trigger', 'DEPRECATED', 'filterAndRewriteTicketData no longer used in school year model');
  return { keptCount: 0, clearedCount: 0 };
}

/**
 * Create the TicketData sheet with headers (36 columns including SLA metrics)
 */
function createTicketSheet(ss) {
  const sheet = ss.insertSheet('TicketData');
  sheet.getRange(1, 1, 1, 36).setValues([[
    'TicketId', 'TicketNumber', 'Subject', 'Year',
    'CreatedDate', 'StartedDate', 'ModifiedDate', 'ClosedDate', 'IsClosed',
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

  const config = getConfig();
  const isCurrent = isSchoolYearCurrent(config);
  const yearStatus = isCurrent ? 'CURRENT (all triggers active)' : 'HISTORICAL (only initial load needed)';

  const response = ui.alert(
    'Setup Automated Triggers',
    `School Year: ${config.schoolYear || 'Not configured'}\n` +
    `Status: ${yearStatus}\n\n` +
    'This will create the following time-driven triggers:\n\n' +
    '• triggerDataContinue - Every 10 minutes\n' +
    '  (Continues any in-progress loading) [ALL]\n\n' +
    '• triggerOpenTicketRefresh - Every 2 hours\n' +
    '  (Refreshes open tickets and SLA data) [CURRENT ONLY]\n\n' +
    '• triggerNewTickets - Every 30 minutes\n' +
    '  (Catches newly created tickets) [CURRENT ONLY]\n\n' +
    '• triggerDailySnapshot - Daily at 7:00 PM\n' +
    '  (Captures backlog metrics for trending) [CURRENT ONLY]\n\n' +
    '• triggerWeeklyFullRefresh - Weekly Sunday 2:00 AM\n' +
    '  (Full reload to catch deletions) [CURRENT ONLY]\n\n' +
    'Historical school years: Triggers skip automatically once data is loaded.\n\n' +
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
