/**
 * Ticket Data Export
 *
 * Comprehensive ticket data extraction with:
 * - School year model: each spreadsheet contains ONE school year's data
 * - Pagination-based loading for historical school years
 * - Date windowing for current school year incremental updates
 * - Auto-run until timeout with resume capability
 * - Single TicketData sheet for Power BI consumption
 * - Consolidated SLA metrics (fetched per-batch during ticket loading)
 *
 * School Year Configuration (via Config sheet):
 * - SCHOOL_YEAR: The school year (e.g., "2025-2026")
 * - SCHOOL_YEAR_START: Start date in MM-DD format (default "07-01" for July 1)
 * - Progress tracking: TICKET_LAST_PAGE, TICKET_COMPLETE, TICKET_LAST_FETCH
 *
 * SLA Data:
 * - SLA metrics are fetched for each batch of tickets (single API call per batch)
 * - Columns 30-36 (AD-AJ): ResponseThreshold, ResponseActual, ResponseBreach,
 *   ResolutionThreshold, ResolutionActual, ResolutionBreach, IsRunning
 */

// Safe runtime limit (5.5 minutes to allow for cleanup before 6 min Apps Script limit)
const MAX_RUNTIME_MS = 5.5 * 60 * 1000;

/**
 * Cached config row positions to avoid repeated reads during loading
 * Populated by cacheConfigRowPositions(), used by writeConfigValueDirect()
 */
let configRowCache_ = null;

/**
 * Main entry point - Continue loading ticket data
 * Runs until timeout or completion, then saves progress
 */
function refreshTicketDataContinue() {
  const ui = SpreadsheetApp.getUi();

  // SAFETY: Acquire lock to prevent concurrent operations
  const lock = acquireScriptLock();
  if (!lock) {
    showOperationBusyMessage('Continue Loading');
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('TicketData');

    if (!sheet) {
      ui.alert('Error', 'TicketData sheet not found. Please create it first.', ui.ButtonSet.OK);
      return;
    }

    const status = getTicketDataStatus();
    const response = ui.alert(
      'Continue Loading Ticket Data',
      `Current Status:\n${status.statusText}\n\n` +
      `This will run until timeout (~5.5 min) or completion.\n` +
      `Progress is saved after each batch.\n\n` +
      'Continue?',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) return;

    const result = runTicketDataLoader(sheet);

    ui.alert(
      result.complete ? 'Complete' : 'Paused',
      `Processed ${result.batchCount} batches, ${result.ticketCount} tickets.\n` +
      `${result.complete ? 'All data loaded!' : 'Run again to continue.'}\n\n` +
      `Runtime: ${(result.runtime / 1000).toFixed(1)} seconds`,
      ui.ButtonSet.OK
    );

    if (result.complete) {
      setCurrentYearReloadInProgress(false);
    }

  } finally {
    releaseScriptLock(lock);
  }
}

/**
 * Full reload - clears ALL data and resets ALL progress
 */
function refreshTicketDataFull() {
  const ui = SpreadsheetApp.getUi();

  // SAFETY: Require no triggers before destructive operation
  if (!requireNoTriggers('Full Reload')) {
    return;
  }

  // SAFETY: Acquire lock to prevent concurrent operations
  const lock = acquireScriptLock();
  if (!lock) {
    showOperationBusyMessage('Full Reload');
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('TicketData');

    if (!sheet) {
      ui.alert('Error', 'TicketData sheet not found.', ui.ButtonSet.OK);
      return;
    }

    const config = getConfig();
    if (!config.schoolYear) {
      ui.alert('Error', 'No SCHOOL_YEAR configured in Config sheet.', ui.ButtonSet.OK);
      return;
    }

    const response = ui.alert(
      'Full Reload',
      'This will:\n' +
      '1. Clear ALL ticket data\n' +
      `2. Reset progress for school year ${config.schoolYear}\n\n` +
      'This cannot be undone. Continue?',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) return;

    // Clear all data (keep header row)
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 36).clear();
    }

    // Reset school year progress
    resetSchoolYearProgress();
    setCurrentYearReloadInProgress(true);

    logOperation('Ticket Data', 'FULL_RESET', `Cleared all data and reset progress for ${config.schoolYear}`);
    ui.alert('Reset Complete', 'All data cleared. Use "Continue Loading" to start fresh.', ui.ButtonSet.OK);

  } finally {
    releaseScriptLock(lock);
  }
}

/**
 * Refresh school year data - for catching deletions
 */
function refreshTicketDataCurrentYear() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TicketData');

  if (!sheet) {
    ui.alert('Error', 'TicketData sheet not found.', ui.ButtonSet.OK);
    return;
  }

  const config = getConfig();
  if (!config.schoolYear) {
    ui.alert('Error', 'No SCHOOL_YEAR configured in Config sheet.', ui.ButtonSet.OK);
    return;
  }

  const response = ui.alert(
    `Refresh School Year (${config.schoolYear})`,
    'This will:\n' +
    `1. Delete all ${config.schoolYear} ticket rows\n` +
    `2. Reset progress\n` +
    `3. Re-fetch all tickets for this school year\n\n` +
    'Use this periodically to catch deleted tickets.\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  // Delete all rows where Year = schoolYear
  deleteRowsBySchoolYear(sheet, config.schoolYear);

  // Reset school year progress
  resetSchoolYearProgress();
  setCurrentYearReloadInProgress(true);

  logOperation('Ticket Data', 'RESET_SCHOOL_YEAR', `Cleared ${config.schoolYear} data, ready for re-fetch`);

  // Start loading
  const result = runTicketDataLoader(sheet);

  ui.alert(
    result.complete ? 'Complete' : 'Paused',
    `Processed ${result.batchCount} batches, ${result.ticketCount} tickets.\n` +
    `${result.complete ? 'School year refreshed!' : 'Run "Continue Loading" to finish.'}`,
    ui.ButtonSet.OK
  );

  if (result.complete) {
    setCurrentYearReloadInProgress(false);
  }
}

/**
 * Show current loading status
 */
function showTicketDataStatus() {
  const ui = SpreadsheetApp.getUi();
  const status = getTicketDataStatus();

  ui.alert(
    'Ticket Data Status',
    `Loading Progress:\n\n${status.statusText}\n\n` +
    `Total rows in sheet: ${status.totalRows}`,
    ui.ButtonSet.OK
  );
}

/**
 * Get human-readable status for the school year
 */
function getTicketDataStatus() {
  const config = getConfig();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TicketData');
  const totalRows = sheet ? Math.max(0, sheet.getLastRow() - 1) : 0;

  const statusLines = [];

  if (!config.schoolYear) {
    statusLines.push('• No school year configured');
    return {
      statusText: statusLines.join('\n'),
      totalRows: totalRows
    };
  }

  const isCurrent = isSchoolYearCurrent(config);
  const complete = config.ticketComplete;
  const lastPage = config.ticketLastPage;
  const totalPages = config.ticketTotalPages;
  const lastFetch = config.ticketLastFetch;

  let status;
  if (complete) {
    if (isCurrent && lastFetch) {
      status = `Complete (incremental last: ${lastFetch.substring(0, 10)})`;
    } else {
      status = 'Complete';
    }
  } else if (lastPage < 0) {
    status = 'Not started';
  } else if (totalPages >= 0) {
    // totalPages is last page INDEX (0-indexed), convert to 1-indexed for display
    status = `Page ${lastPage + 1} of ${totalPages + 1}`;
  } else {
    status = `Page ${lastPage + 1} (total unknown)`;
  }

  statusLines.push(`• ${config.schoolYear}${isCurrent ? ' (current)' : ' (historical)'}: ${status}`);

  return {
    statusText: statusLines.join('\n'),
    totalRows: totalRows
  };
}

/**
 * Main loader loop - runs until timeout or complete
 *
 * Performance optimization: Reads config once at start and caches row positions
 * to minimize sheet reads during the loop. Each read forces a flush which
 * triggers formula recalculation across all analytics sheets.
 */
function runTicketDataLoader(sheet) {
  const startTime = Date.now();
  let batchCount = 0;
  let ticketCount = 0;
  let complete = false;

  // Lock configuration on first run (SCHOOL_YEAR, PAGE_SIZE, BATCH_SIZE)
  // This prevents accidental changes during data loading
  try {
    lockConfig();
  } catch (e) {
    logOperation('Ticket Data', 'ERROR', e.message);
    return { batchCount: 0, ticketCount: 0, complete: false, error: e.message };
  }

  // Read config ONCE at start and cache row positions for direct writes
  // This avoids repeated reads that force flushes and trigger recalculation
  let config = getConfig();
  cacheConfigRowPositions();

  if (!config.schoolYear) {
    logOperation('Ticket Data', 'ERROR', 'No SCHOOL_YEAR configured');
    clearConfigCache();
    return { batchCount: 0, ticketCount: 0, complete: true, runtime: 0 };
  }

  const isCurrent = isSchoolYearCurrent(config);

  // Track progress in memory to avoid reads
  let ticketLastPage = config.ticketLastPage;
  let ticketComplete = config.ticketComplete;
  let ticketTotalPages = config.ticketTotalPages;
  let ticketLastFetch = config.ticketLastFetch;

  while (Date.now() - startTime < MAX_RUNTIME_MS) {
    // Check if initial pagination loading is complete
    if (!ticketComplete) {
      // Still doing initial pagination-based load
      // Pass in-memory progress to avoid re-reading config
      const result = processSchoolYearBatchOptimized(sheet, config, ticketLastPage, ticketTotalPages);
      batchCount++;
      ticketCount += result.count;

      // Update in-memory progress
      ticketLastPage = result.lastPage;
      if (result.totalPages !== undefined) {
        ticketTotalPages = result.totalPages;
      }

      // Write progress using direct writes (no reads) - use String() to prevent Sheets date formatting
      writeConfigValueDirect('TICKET_LAST_PAGE', String(ticketLastPage));
      if (result.totalPages !== undefined) {
        writeConfigValueDirect('TICKET_TOTAL_PAGES', String(ticketTotalPages));
      }

      if (result.complete) {
        ticketComplete = true;
        writeConfigValueDirect('TICKET_COMPLETE', 'TRUE');
        logOperation('Ticket Data', 'COMPLETE', `School year ${config.schoolYear} pagination loading finished`);

        // If historical school year, we're done
        if (!isCurrent) {
          complete = true;
          // Flush before breaking to ensure progress is saved
          flushWrites();
          break;
        }

        // For current school year, set TICKET_LAST_FETCH to last ticket's CreatedDate
        // so incremental mode starts from where pagination left off
        if (result.lastTicketCreatedDate) {
          ticketLastFetch = result.lastTicketCreatedDate;
          writeConfigValueDirect('TICKET_LAST_FETCH', ticketLastFetch);
          logOperation('Ticket Data', 'INFO', `Set incremental baseline to ${ticketLastFetch}`);
        }
      }

      // Flush once per batch - this is when recalculation happens
      flushWrites();

    } else if (isCurrent) {
      // Pagination complete, do incremental updates for current school year
      const result = processCurrentSchoolYearBatchOptimized(sheet, config, ticketLastFetch);
      batchCount++;
      ticketCount += result.count;

      // Update in-memory progress
      if (result.lastFetch) {
        ticketLastFetch = result.lastFetch;
        writeConfigValueDirect('TICKET_LAST_FETCH', ticketLastFetch);
      }

      // Flush once per batch
      flushWrites();

      if (!result.hasMore) {
        complete = true;
        break;
      }
    } else {
      // Historical school year, pagination complete - nothing more to do
      complete = true;
      break;
    }

    // Throttle between batches (use cached config value)
    Utilities.sleep(config.throttleMs || 1000);
  }

  // Clean up cache
  clearConfigCache();

  const runtime = Date.now() - startTime;
  logOperation('Ticket Data', complete ? 'COMPLETE' : 'PAUSED',
    `${batchCount} batches, ${ticketCount} tickets, ${(runtime/1000).toFixed(1)}s`);

  // Update last refresh timestamp so analytics sheets show current data
  if (ticketCount > 0) {
    updateLastRefresh();
  }

  return { batchCount, ticketCount, complete, runtime };
}

/**
 * Process one batch for the school year using pagination
 */
function processSchoolYearBatch(sheet, config) {
  const batchSize = config.ticketBatchSize;
  const lastPage = config.ticketLastPage;
  const nextPage = lastPage + 1;

  // Get school year date range
  const dates = getSchoolYearDates(config);
  if (!dates) {
    logOperation('Ticket Data', 'ERROR', `Invalid school year format: ${config.schoolYear}`);
    return { count: 0, complete: true };
  }

  // Build date range for API (MM/DD/YYYY format)
  const startDate = formatDateForApi(dates.startDate);
  const endDate = formatDateForApi(dates.endDate);
  const dateRange = `daterange:${startDate}-${endDate}`;

  logOperation('Ticket Data', 'BATCH', `School year ${config.schoolYear}, page ${nextPage}`);
  logOperation('Ticket Data', 'DEBUG', `Date filter: ${dateRange} | schoolYearStart config: "${config.schoolYearStart}"`);

  // Fetch page
  const response = searchTickets([
    { Facet: 'createddate', Value: dateRange }
  ], nextPage, batchSize, { field: 'TicketCreatedDate', direction: 'asc' });

  // On first page, calculate total pages and store last page index
  if (nextPage === 0 && response.Paging) {
    // Handle both TotalRows and Total field names (API may use either)
    const totalRows = response.Paging.TotalRows || response.Paging.Total || 0;
    const totalPages = Math.ceil(Number(totalRows) / batchSize);
    const lastPageIndex = totalPages - 1;
    // Write as string to prevent Sheets from auto-formatting as date
    // Store last page INDEX (0-indexed) so LAST_PAGE and TOTAL_PAGES match when complete
    updateConfigValue('TICKET_TOTAL_PAGES', String(lastPageIndex));
    logOperation('Ticket Data', 'INFO', `School year ${config.schoolYear}: ${totalRows} tickets, ${totalPages} pages (last index: ${lastPageIndex})`);
  }

  if (!response.Items || response.Items.length === 0) {
    // No more data - mark complete
    logOperation('Ticket Data', 'COMPLETE', `School year ${config.schoolYear} pagination finished`);
    return { count: 0, complete: true };
  }

  // Fetch SLA data for this batch of tickets
  const ticketIds = response.Items.map(t => t.TicketId).filter(id => id);
  const slaMap = fetchSlaForTicketIds(ticketIds);
  logOperation('Ticket Data', 'SLA_FETCH', `Fetched SLA for ${slaMap.size} of ${ticketIds.length} tickets`);

  // Write tickets to sheet (with merged SLA data) - use school year string for Year column
  const now = new Date();
  const rows = response.Items.map(ticket => extractTicketRow(ticket, now, config.schoolYear, slaMap));
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, rows.length, 36).setValues(rows);

  // Update progress
  updateConfigValue('TICKET_LAST_PAGE', nextPage);

  // Check if complete (totalPages now stores last page INDEX, not count)
  // Use -1 as sentinel for "unknown", so check explicitly
  let lastPageIndex = config.ticketTotalPages;
  if (lastPageIndex < 0 && response.Paging) {
    lastPageIndex = Math.ceil((response.Paging.TotalRows || response.Paging.Total || 0) / batchSize) - 1;
  }

  const isComplete = nextPage >= lastPageIndex;

  logOperation('Ticket Data', 'BATCH', `School year ${config.schoolYear} page ${nextPage}: wrote ${rows.length} tickets`);
  return { count: rows.length, complete: isComplete };
}

/**
 * Process one batch for current school year using date windowing (incremental)
 */
function processCurrentSchoolYearBatch(sheet, config) {
  const batchSize = config.ticketBatchSize;
  const lastFetch = config.ticketLastFetch;

  // Get school year date range
  const dates = getSchoolYearDates(config);
  if (!dates) {
    logOperation('Ticket Data', 'ERROR', `Invalid school year format: ${config.schoolYear}`);
    return { count: 0, hasMore: false };
  }

  // Determine start date for incremental fetch
  let startDate;
  if (lastFetch) {
    // Start from the date of last fetch
    const lastDate = new Date(lastFetch);
    startDate = formatDateForApi(lastDate);
  } else {
    // Fresh start - beginning of school year
    startDate = formatDateForApi(dates.startDate);
  }

  const endDate = formatDateForApi(new Date());
  const dateRange = `daterange:${startDate}-${endDate}`;

  logOperation('Ticket Data', 'BATCH', `School year ${config.schoolYear} incremental, from ${startDate}`);

  // Fetch tickets
  const response = searchTickets([
    { Facet: 'createddate', Value: dateRange }
  ], 0, batchSize, { field: 'TicketCreatedDate', direction: 'asc' });

  if (!response.Items || response.Items.length === 0) {
    logOperation('Ticket Data', 'COMPLETE', `School year ${config.schoolYear} up to date`);
    return { count: 0, hasMore: false };
  }

  // Filter out already-fetched tickets (by timestamp comparison)
  let tickets = response.Items;
  if (lastFetch) {
    const lastFetchTime = new Date(lastFetch).getTime();
    tickets = tickets.filter(t => new Date(t.CreatedDate).getTime() > lastFetchTime);
  }

  if (tickets.length === 0) {
    // All tickets in batch already fetched, but there may be more
    const hasMore = response.Paging && response.Paging.TotalRows > batchSize;
    if (hasMore) {
      // Advance the window
      const lastTicket = response.Items[response.Items.length - 1];
      updateConfigValue('TICKET_LAST_FETCH', lastTicket.CreatedDate);
    }
    return { count: 0, hasMore: hasMore };
  }

  // Sort to ensure order
  tickets.sort((a, b) => new Date(a.CreatedDate) - new Date(b.CreatedDate));

  // Fetch SLA data for this batch of tickets
  const ticketIds = tickets.map(t => t.TicketId).filter(id => id);
  const slaMap = fetchSlaForTicketIds(ticketIds);
  logOperation('Ticket Data', 'SLA_FETCH', `Fetched SLA for ${slaMap.size} of ${ticketIds.length} tickets`);

  // Write to sheet (with merged SLA data) - use school year string
  const now = new Date();
  const rows = tickets.map(ticket => extractTicketRow(ticket, now, config.schoolYear, slaMap));
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, rows.length, 36).setValues(rows);

  // Update last fetch timestamp
  const lastTicket = tickets[tickets.length - 1];
  updateConfigValue('TICKET_LAST_FETCH', lastTicket.CreatedDate);

  // Check if more pages
  const totalRows = response.Paging ? (response.Paging.TotalRows || response.Paging.Total || 0) : 0;
  const hasMore = totalRows > batchSize;

  logOperation('Ticket Data', 'BATCH', `School year ${config.schoolYear}: wrote ${rows.length} tickets`);
  return { count: rows.length, hasMore: hasMore };
}

/**
 * Optimized version of processSchoolYearBatch that accepts in-memory progress
 * Avoids reading config during the loop to minimize flush/recalculation cycles
 *
 * @param {Sheet} sheet - TicketData sheet
 * @param {Object} config - Config object (read once at start)
 * @param {number} lastPage - Last completed page (in-memory)
 * @param {number} totalPages - Total pages (in-memory, -1 if unknown)
 * @returns {Object} - { count, complete, lastPage, totalPages }
 */
function processSchoolYearBatchOptimized(sheet, config, lastPage, totalPages) {
  const batchSize = config.ticketBatchSize;
  const nextPage = lastPage + 1;

  // Get school year date range
  const dates = getSchoolYearDates(config);
  if (!dates) {
    logOperation('Ticket Data', 'ERROR', `Invalid school year format: ${config.schoolYear}`);
    return { count: 0, complete: true, lastPage: lastPage, totalPages: totalPages };
  }

  // Build date range for API (MM/DD/YYYY format)
  const startDate = formatDateForApi(dates.startDate);
  const endDate = formatDateForApi(dates.endDate);
  const dateRange = `daterange:${startDate}-${endDate}`;

  logOperation('Ticket Data', 'BATCH', `School year ${config.schoolYear}, page ${nextPage}`);

  // Fetch page
  const response = searchTickets([
    { Facet: 'createddate', Value: dateRange }
  ], nextPage, batchSize, { field: 'TicketCreatedDate', direction: 'asc' });

  // On first page, calculate total pages
  let newTotalPages = totalPages;
  if (nextPage === 0 && response.Paging) {
    const totalRows = response.Paging.TotalRows || response.Paging.Total || 0;
    const pageCount = Math.ceil(Number(totalRows) / batchSize);
    newTotalPages = pageCount - 1; // Last page INDEX (0-indexed)
    logOperation('Ticket Data', 'INFO', `School year ${config.schoolYear}: ${totalRows} tickets, ${pageCount} pages (last index: ${newTotalPages})`);
  }

  if (!response.Items || response.Items.length === 0) {
    logOperation('Ticket Data', 'COMPLETE', `School year ${config.schoolYear} pagination finished`);
    return { count: 0, complete: true, lastPage: nextPage, totalPages: newTotalPages };
  }

  // Fetch SLA data for this batch of tickets
  const ticketIds = response.Items.map(t => t.TicketId).filter(id => id);
  const slaMap = fetchSlaForTicketIds(ticketIds);
  logOperation('Ticket Data', 'SLA_FETCH', `Fetched SLA for ${slaMap.size} of ${ticketIds.length} tickets`);

  // Write tickets to sheet (with merged SLA data)
  const now = new Date();
  const rows = response.Items.map(ticket => extractTicketRow(ticket, now, config.schoolYear, slaMap));
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, rows.length, 36).setValues(rows);

  // Check if complete
  let lastPageIndex = newTotalPages;
  if (lastPageIndex < 0 && response.Paging) {
    lastPageIndex = Math.ceil((response.Paging.TotalRows || response.Paging.Total || 0) / batchSize) - 1;
  }
  const isComplete = nextPage >= lastPageIndex;

  // Get the last ticket's CreatedDate for setting TICKET_LAST_FETCH when pagination completes
  const lastTicket = response.Items[response.Items.length - 1];
  const lastTicketCreatedDate = lastTicket ? lastTicket.CreatedDate : null;

  logOperation('Ticket Data', 'BATCH', `School year ${config.schoolYear} page ${nextPage}: wrote ${rows.length} tickets`);
  return {
    count: rows.length,
    complete: isComplete,
    lastPage: nextPage,
    totalPages: newTotalPages !== totalPages ? newTotalPages : undefined,
    lastTicketCreatedDate: lastTicketCreatedDate
  };
}

/**
 * Optimized version of processCurrentSchoolYearBatch that accepts in-memory progress
 * Avoids reading config during the loop to minimize flush/recalculation cycles
 *
 * @param {Sheet} sheet - TicketData sheet
 * @param {Object} config - Config object (read once at start)
 * @param {string} lastFetch - Last fetch timestamp (in-memory)
 * @returns {Object} - { count, hasMore, lastFetch }
 */
function processCurrentSchoolYearBatchOptimized(sheet, config, lastFetch) {
  const batchSize = config.ticketBatchSize;

  // Get school year date range
  const dates = getSchoolYearDates(config);
  if (!dates) {
    logOperation('Ticket Data', 'ERROR', `Invalid school year format: ${config.schoolYear}`);
    return { count: 0, hasMore: false, lastFetch: lastFetch };
  }

  // Determine start date for incremental fetch
  let startDate;
  if (lastFetch) {
    const lastDate = new Date(lastFetch);
    startDate = formatDateForApi(lastDate);
  } else {
    startDate = formatDateForApi(dates.startDate);
  }

  const endDate = formatDateForApi(new Date());
  const dateRange = `daterange:${startDate}-${endDate}`;

  logOperation('Ticket Data', 'BATCH', `School year ${config.schoolYear} incremental, from ${startDate}`);

  // Fetch tickets
  const response = searchTickets([
    { Facet: 'createddate', Value: dateRange }
  ], 0, batchSize, { field: 'TicketCreatedDate', direction: 'asc' });

  if (!response.Items || response.Items.length === 0) {
    logOperation('Ticket Data', 'COMPLETE', `School year ${config.schoolYear} up to date`);
    return { count: 0, hasMore: false, lastFetch: lastFetch };
  }

  // Filter out already-fetched tickets (by timestamp comparison)
  let tickets = response.Items;
  if (lastFetch) {
    const lastFetchTime = new Date(lastFetch).getTime();
    tickets = tickets.filter(t => new Date(t.CreatedDate).getTime() > lastFetchTime);
  }

  if (tickets.length === 0) {
    // All tickets in batch already fetched, but there may be more
    const hasMore = response.Paging && response.Paging.TotalRows > batchSize;
    let newLastFetch = lastFetch;
    if (hasMore) {
      // Advance the window
      const lastTicket = response.Items[response.Items.length - 1];
      newLastFetch = lastTicket.CreatedDate;
    }
    return { count: 0, hasMore: hasMore, lastFetch: newLastFetch };
  }

  // Sort to ensure order
  tickets.sort((a, b) => new Date(a.CreatedDate) - new Date(b.CreatedDate));

  // Fetch SLA data for this batch of tickets
  const ticketIds = tickets.map(t => t.TicketId).filter(id => id);
  const slaMap = fetchSlaForTicketIds(ticketIds);
  logOperation('Ticket Data', 'SLA_FETCH', `Fetched SLA for ${slaMap.size} of ${ticketIds.length} tickets`);

  // Write to sheet (with merged SLA data)
  const now = new Date();
  const rows = tickets.map(ticket => extractTicketRow(ticket, now, config.schoolYear, slaMap));
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, rows.length, 36).setValues(rows);

  // Get last fetch timestamp from written tickets
  const lastTicket = tickets[tickets.length - 1];
  const newLastFetch = lastTicket.CreatedDate;

  // Check if more pages
  const totalRows = response.Paging ? (response.Paging.TotalRows || response.Paging.Total || 0) : 0;
  const hasMore = totalRows > batchSize;

  logOperation('Ticket Data', 'BATCH', `School year ${config.schoolYear}: wrote ${rows.length} tickets`);
  return { count: rows.length, hasMore: hasMore, lastFetch: newLastFetch };
}

/**
 * Extract a row of data from a ticket object
 * Returns 36 columns for comprehensive analytics (including SLA metrics)
 *
 * @param {Object} ticket - Ticket object from API
 * @param {Date} now - Current timestamp for age calculation
 * @param {number} year - Year for the ticket
 * @param {Map} slaMap - Optional map of TicketId -> SLA metrics
 */
function extractTicketRow(ticket, now, year, slaMap) {
  const createdDate = ticket.CreatedDate ? new Date(ticket.CreatedDate) : null;
  const closedDate = ticket.ClosedDate ? new Date(ticket.ClosedDate) : null;

  // Calculate age in days
  let ageDays = 0;
  if (createdDate) {
    if (closedDate) {
      ageDays = Math.floor((closedDate - createdDate) / (1000 * 60 * 60 * 24));
    } else {
      ageDays = Math.floor((now - createdDate) / (1000 * 60 * 60 * 24));
    }
  }

  // Extract location type name
  let locationType = '';
  if (ticket.Location && ticket.Location.LocationType) {
    locationType = ticket.Location.LocationType.Name || '';
  }

  // Extract Issue details
  const issue = ticket.Issue || {};

  // Extract SLA details from ticket (basic info)
  const ticketSla = ticket.Sla || {};

  // Get detailed SLA metrics from slaMap (if available)
  const slaMetrics = slaMap ? slaMap.get(ticket.TicketId) : null;

  // Use SlaName from detailed SLA data if available, fallback to ticket's Sla object
  const slaName = slaMetrics ? slaMetrics.slaName : (ticketSla.Name || ticketSla.SlaName || '');

  return [
    // A-D: Core identifiers
    ticket.TicketId || '',
    ticket.TicketNumber || '',
    (ticket.Subject || '').substring(0, 200),
    year,
    // E-I: Dates and status
    ticket.CreatedDate || '',
    ticket.StartedDate || '',
    ticket.ModifiedDate || '',
    ticket.ClosedDate || '',
    ticket.IsClosed ? 'Yes' : 'No',
    // J: Workflow status
    ticket.WorkflowStep ? (ticket.WorkflowStep.StatusName || ticket.WorkflowStep.Name || '') : '',
    // K-L: Team
    ticket.AssignedToTeam ? (ticket.AssignedToTeam.TeamId || '') : '',
    ticket.AssignedToTeam ? (ticket.AssignedToTeam.TeamName || '') : '',
    // M-O: Location
    ticket.Location ? (ticket.Location.LocationId || '') : '',
    ticket.Location ? (ticket.Location.Name || '') : '',
    locationType,
    // P-Q: Owner
    ticket.Owner ? (ticket.Owner.UserId || '') : '',
    ticket.Owner ? (ticket.Owner.Name || '') : '',
    // R: Age
    ageDays,
    // S-U: Priority and Due Date
    ticket.Priority || '',
    ticket.IsPastDue ? 'Yes' : 'No',
    ticket.DueDate || '',
    // V-W: SLA (basic)
    ticketSla.SlaId || '',
    slaName,
    // X-AA: Issue Category and Type
    issue.IssueCategoryId || '',
    issue.IssueCategoryName || '',
    issue.IssueTypeId || '',
    issue.Name || '', // Issue type name is in Name field
    // AB-AC: Requester (For)
    ticket.For ? (ticket.For.UserId || '') : '',
    ticket.For ? (ticket.For.Name || '') : '',
    // AD-AJ: SLA Metrics (from detailed SLA API)
    slaMetrics ? slaMetrics.responseThreshold : '',
    slaMetrics ? slaMetrics.responseActual : '',
    slaMetrics ? slaMetrics.responseBreach : '',
    slaMetrics ? slaMetrics.resolutionThreshold : '',
    slaMetrics ? slaMetrics.resolutionActual : '',
    slaMetrics ? slaMetrics.resolutionBreach : '',
    slaMetrics ? slaMetrics.isRunning : ''
  ];
}

/**
 * Fetch SLA data for a batch of ticket IDs
 * Makes a single API call with all ticket IDs
 *
 * @param {Array} ticketIds - Array of TicketId values
 * @returns {Map} - Map of TicketId -> SLA metrics object
 */
function fetchSlaForTicketIds(ticketIds) {
  const slaMap = new Map();

  if (!ticketIds || ticketIds.length === 0) {
    return slaMap;
  }

  const config = getConfig();

  // Build endpoint - fetch all tickets in one call
  const endpoint = `/v1.0/tickets/slas?$p=0&$s=${ticketIds.length}`;

  // Build filter with TicketIds - each ticket ID is a separate filter entry
  // Multiple filters with same Facet act as OR condition
  const payload = {
    Filters: ticketIds.map(id => ({
      Facet: 'Ticket',
      Id: id
    }))
  };

  try {
    const response = makeApiRequest(endpoint, 'POST', payload);

    if (!response || !response.Items) {
      return slaMap;
    }

    // Process each SLA item
    for (const item of response.Items) {
      const ticketId = item.TicketId;
      if (!ticketId) continue;

      // Skip tickets with no SLA assigned
      if (!item.Sla || !item.SlaTimes || item.SlaTimes.length === 0) {
        continue;
      }

      const sla = item.Sla || {};
      const slaTimes = item.SlaTimes || [];
      const metrics = sla.Metrics || [];

      // Find response metric and resolution metric by Name field
      const responseMetric = metrics.find(m => m.Name === 'Response Time') || {};
      const resolutionMetric = metrics.find(m => m.Name === 'Resolution Time') || {};

      // Find actual times from SlaTimes by Name field
      const responseTime = slaTimes.find(t => t.Name === 'Response Time') || {};
      const resolutionTime = slaTimes.find(t => t.Name === 'Resolution Time') || {};

      const responseThreshold = responseMetric.ValueInMinutes || null;
      const responseActual = responseTime.LogMinutes || null;
      const resolutionThreshold = resolutionMetric.ValueInMinutes || null;
      const resolutionActual = resolutionTime.LogMinutes || null;

      // Calculate breach status
      const responseBreach = (responseActual !== null && responseThreshold !== null && responseActual > responseThreshold);
      const resolutionBreach = (resolutionActual !== null && resolutionThreshold !== null && resolutionActual > resolutionThreshold);

      // Check if SLA is still running
      const isRunning = slaTimes.some(t => t.IsRunning === true);

      slaMap.set(ticketId, {
        slaName: sla.SlaName || sla.Name || '',
        responseThreshold: responseThreshold,
        responseActual: responseActual,
        responseBreach: responseBreach,
        resolutionThreshold: resolutionThreshold,
        resolutionActual: resolutionActual,
        resolutionBreach: resolutionBreach,
        isRunning: isRunning
      });
    }
  } catch (error) {
    logOperation('Ticket Data', 'SLA_ERROR', `Failed to fetch SLA data: ${error.message}`);
    // Return empty map - tickets will still be written without SLA data
  }

  return slaMap;
}

/**
 * Format a date for the API (MM/DD/YYYY)
 */
function formatDateForApi(date) {
  const d = date instanceof Date ? date : new Date(date);
  return `${(d.getMonth() + 1).toString().padStart(2, '0')}/${d.getDate().toString().padStart(2, '0')}/${d.getFullYear()}`;
}

/**
 * Cache Config sheet row positions for fast direct writes during loading
 * Call this once at the start of a loading session to avoid repeated reads
 *
 * @returns {Object} - { sheet, rowMap: Map<key, rowNum> }
 */
function cacheConfigRowPositions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Config');
  const data = sheet.getDataRange().getValues();

  const rowMap = new Map();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0]) {
      rowMap.set(data[i][0], i + 1); // 1-indexed row number
    }
  }

  configRowCache_ = { sheet, rowMap, data };
  return configRowCache_;
}

/**
 * Write a config value directly without reading the sheet
 * Requires cacheConfigRowPositions() to have been called first
 * Falls back to updateConfigValue() if cache miss (new key)
 *
 * @param {string} key - Config key to write
 * @param {any} value - Value to write
 */
function writeConfigValueDirect(key, value) {
  if (!configRowCache_) {
    // No cache - fall back to regular update
    updateConfigValue(key, value);
    return;
  }

  const rowNum = configRowCache_.rowMap.get(key);
  if (rowNum) {
    // Direct write to known row - no read required
    configRowCache_.sheet.getRange(rowNum, 2).setValue(value);
  } else {
    // Key doesn't exist in cache - use regular update (will read)
    // This should be rare during normal loading
    updateConfigValue(key, value);
    // Invalidate cache since sheet structure changed
    configRowCache_ = null;
  }
}

/**
 * Flush all pending writes and trigger recalculation
 * Call this at controlled points (e.g., end of each batch) to minimize recalc frequency
 */
function flushWrites() {
  SpreadsheetApp.flush();
}

/**
 * Clear the config row cache
 * Call this at the end of a loading session or if config structure changes
 */
function clearConfigCache() {
  configRowCache_ = null;
}

/**
 * Update a config value by key name
 * If the key doesn't exist, inserts it in a logical position near related keys
 * Note: This reads the sheet - use writeConfigValueDirect() in tight loops
 */
function updateConfigValue(key, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Config');
  const data = sheet.getDataRange().getValues();

  // Check if key already exists
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }

  // Key not found - find the best insertion point
  const insertRow = findConfigInsertPosition(key, data);

  if (insertRow <= data.length) {
    // Insert a new row at the position
    sheet.insertRowAfter(insertRow);
    sheet.getRange(insertRow + 1, 1, 1, 2).setValues([[key, value]]);
  } else {
    // Append to end
    sheet.getRange(insertRow, 1, 1, 2).setValues([[key, value]]);
  }
}

/**
 * Find the best row position to insert a new config key
 * Groups related keys together in logical order
 *
 * @param {string} key - The config key to insert
 * @param {Array} data - Current config data (2D array)
 * @returns {number} - Row number to insert after (1-indexed)
 */
function findConfigInsertPosition(key, data) {
  // Check for ticket progress keys (simplified - no year suffix)
  const ticketProgressKeys = ['TICKET_TOTAL_PAGES', 'TICKET_LAST_PAGE', 'TICKET_COMPLETE', 'TICKET_LAST_FETCH'];
  if (ticketProgressKeys.includes(key)) {
    return findTicketProgressInsertPosition(key, data);
  }

  // Check for open refresh keys
  const openRefreshMatch = key.match(/^OPEN_REFRESH_(.+)$/);
  if (openRefreshMatch) {
    return findOpenRefreshInsertPosition(key, data);
  }

  // For other keys, append to end
  return data.length + 1;
}

/**
 * Find insertion position for TICKET_* progress keys
 * Order: TOTAL_PAGES, LAST_PAGE, COMPLETE, LAST_FETCH
 */
function findTicketProgressInsertPosition(key, data) {
  const keyOrder = ['TICKET_TOTAL_PAGES', 'TICKET_LAST_PAGE', 'TICKET_COMPLETE', 'TICKET_LAST_FETCH'];
  const targetOrder = keyOrder.indexOf(key);

  let ticketRows = [];
  let lastTicketRow = 0;

  // Scan for existing TICKET_* progress rows
  for (let i = 0; i < data.length; i++) {
    const rowKey = data[i][0];
    if (!rowKey) continue;

    if (keyOrder.includes(String(rowKey))) {
      const order = keyOrder.indexOf(String(rowKey));
      ticketRows.push({ row: i + 1, order: order });
      lastTicketRow = i + 1;
    }
  }

  // If we have existing ticket rows, find the right position
  if (ticketRows.length > 0) {
    ticketRows.sort((a, b) => a.order - b.order);

    for (let i = 0; i < ticketRows.length; i++) {
      if (ticketRows[i].order > targetOrder) {
        // Insert before this row
        return ticketRows[i].row - 1;
      }
    }
    // Insert after the last ticket row
    return ticketRows[ticketRows.length - 1].row;
  }

  // No TICKET_* rows exist yet - find a good starting position
  // Look for the end of the settings section (after STALE_DAYS, SLA_RISK_PERCENT, etc.)
  const settingsKeys = ['PAGE_SIZE', 'THROTTLE_MS', 'TICKET_BATCH_SIZE', 'STALE_DAYS', 'SLA_RISK_PERCENT', 'SCHOOL_YEAR', 'SCHOOL_YEAR_START'];
  let lastSettingsRow = 0;

  for (let i = 0; i < data.length; i++) {
    if (settingsKeys.includes(data[i][0])) {
      lastSettingsRow = i + 1;
    }
  }

  return lastSettingsRow > 0 ? lastSettingsRow : data.length + 1;
}

/**
 * Find insertion position for OPEN_REFRESH_* keys
 * Groups all OPEN_REFRESH keys together
 */
function findOpenRefreshInsertPosition(key, data) {
  let lastOpenRefreshRow = 0;
  let firstOpenRefreshRow = 0;

  for (let i = 0; i < data.length; i++) {
    const rowKey = data[i][0];
    if (rowKey && String(rowKey).startsWith('OPEN_REFRESH_')) {
      if (firstOpenRefreshRow === 0) firstOpenRefreshRow = i + 1;
      lastOpenRefreshRow = i + 1;
    }
  }

  if (lastOpenRefreshRow > 0) {
    // Add after the last OPEN_REFRESH row
    return lastOpenRefreshRow;
  }

  // No OPEN_REFRESH rows yet - add after TICKET rows
  let lastTicketRow = 0;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] && String(data[i][0]).startsWith('TICKET_')) {
      lastTicketRow = i + 1;
    }
  }

  return lastTicketRow > 0 ? lastTicketRow : data.length + 1;
}

/**
 * Reset progress for school year loading
 */
function resetSchoolYearProgress() {
  updateConfigValue('TICKET_TOTAL_PAGES', '');
  updateConfigValue('TICKET_LAST_PAGE', '-1');  // String to prevent Sheets date formatting
  updateConfigValue('TICKET_COMPLETE', 'FALSE');
  updateConfigValue('TICKET_LAST_FETCH', '');
}

/**
 * @deprecated Use resetSchoolYearProgress instead
 */
function resetYearProgress(year) {
  resetSchoolYearProgress();
}

/**
 * @deprecated Use resetSchoolYearProgress instead
 */
function resetCurrentYearProgress(year) {
  updateConfigValue('TICKET_LAST_FETCH', '');
}

/**
 * Delete all rows from sheet where Year column matches the school year
 */
function deleteRowsBySchoolYear(sheet, schoolYear) {
  const data = sheet.getDataRange().getValues();
  const yearCol = 3; // Column D (0-indexed = 3) is Year

  // Find rows to delete (from bottom up to preserve indices)
  const rowsToDelete = [];
  for (let i = data.length - 1; i >= 1; i--) { // Skip header
    if (data[i][yearCol] === schoolYear) {
      rowsToDelete.push(i + 1); // Convert to 1-indexed
    }
  }

  // Delete in batches (more efficient)
  if (rowsToDelete.length > 0) {
    // Group consecutive rows for batch deletion
    let start = rowsToDelete[0];
    let count = 1;

    for (let i = 1; i < rowsToDelete.length; i++) {
      if (rowsToDelete[i] === start - count) {
        count++;
      } else {
        sheet.deleteRows(start - count + 1, count);
        start = rowsToDelete[i];
        count = 1;
      }
    }
    sheet.deleteRows(start - count + 1, count);

    logOperation('Ticket Data', 'DELETE', `Deleted ${rowsToDelete.length} rows for school year ${schoolYear}`);
  }

  return rowsToDelete.length;
}

/**
 * @deprecated Use deleteRowsBySchoolYear instead
 */
function deleteRowsByYear(sheet, year) {
  return deleteRowsBySchoolYear(sheet, year);
}

// =============================================================================
// DEPRECATED: Open Ticket Refresh Functions
// =============================================================================
// These functions are no longer used. The new approach is:
// - triggerTicketDataReset: Deletes all ticket data at midnight
// - triggerTicketDataContinue: Reloads all tickets every 10 minutes
// This is simpler and handles all ticket changes (status, owner, deleted, etc.)
// =============================================================================

/**
 * Menu: Start a fresh open ticket refresh (resets progress)
 */
function refreshOpenTicketsStart() {
  const ui = SpreadsheetApp.getUi();

  // SAFETY: Acquire lock to prevent concurrent operations
  const lock = acquireScriptLock();
  if (!lock) {
    showOperationBusyMessage('Open Ticket Refresh (Start)');
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('TicketData');

    if (!sheet) {
      ui.alert('Error', 'TicketData sheet not found.', ui.ButtonSet.OK);
      return;
    }

    const config = getConfig();
    const isHistorical = !isSchoolYearCurrent(config);
    const staleDays = config.staleDays || 7;

    // Warning for historical school years
    if (isHistorical) {
      const openCount = countOpenTickets(sheet);
      if (openCount === 0) {
        ui.alert(
          'Historical School Year - No Open Tickets',
          `School year ${config.schoolYear} is historical and has no open tickets.\n\n` +
          `There is nothing to refresh. The data is static.`,
          ui.ButtonSet.OK
        );
        return;
      }

      const response = ui.alert(
        'Historical School Year Warning',
        `School year ${config.schoolYear} is historical.\n\n` +
        `This sheet still has ${openCount} open ticket(s) that may need updates.\n\n` +
        `The refresh will:\n` +
        `• UPDATE existing tickets with status/SLA changes\n` +
        `• NOT add any new tickets (they belong to a different school year)\n\n` +
        `Continue?`,
        ui.ButtonSet.YES_NO
      );
      if (response !== ui.Button.YES) return;
    } else {
      const response = ui.alert(
        'Start Open Ticket Refresh',
        `This will fetch tickets modified since last refresh:\n\n` +
        `• Update existing tickets with status/SLA changes\n` +
        `• Add any new tickets not yet in the sheet\n\n` +
        `Progress is saved after each batch.\n` +
        `Use "Continue" if this times out.\n\n` +
        'Start fresh?',
        ui.ButtonSet.YES_NO
      );
      if (response !== ui.Button.YES) return;
    }

    // Reset progress to start fresh
    resetOpenRefreshProgress();

    logOperation('Ticket Data', 'START', `Open ticket refresh started fresh (historical=${isHistorical})`);

    const result = runOpenTicketRefresh(sheet);

    const skippedInfo = result.skippedCount ? `\nSkipped ${result.skippedCount} tickets (different school year).` : '';
    ui.alert(
      result.complete ? 'Complete' : 'Paused',
      `Updated ${result.updatedCount} existing rows.\n` +
      `Appended ${result.appendedCount} new tickets.${skippedInfo}\n` +
      `Processed ${result.ticketCount} tickets in ${result.batchCount} batches.\n` +
      `${result.complete ? 'Refresh complete!' : 'Use "Continue" to resume.'}\n\n` +
      `Runtime: ${(result.runtime / 1000).toFixed(1)} seconds`,
      ui.ButtonSet.OK
    );

  } finally {
    releaseScriptLock(lock);
  }
}

/**
 * Menu: Continue an in-progress open ticket refresh
 */
function refreshOpenTicketsContinue() {
  const ui = SpreadsheetApp.getUi();

  // SAFETY: Acquire lock to prevent concurrent operations
  const lock = acquireScriptLock();
  if (!lock) {
    showOperationBusyMessage('Open Ticket Refresh (Continue)');
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('TicketData');

    if (!sheet) {
      ui.alert('Error', 'TicketData sheet not found.', ui.ButtonSet.OK);
      return;
    }

    const config = getConfig();
    const isHistorical = !isSchoolYearCurrent(config);
    const status = getOpenRefreshStatusText();

    // Info message for historical school years
    let historicalNote = '';
    if (isHistorical) {
      historicalNote = `\n\nNote: Historical school year - only updating existing tickets.`;
    }

    const response = ui.alert(
      'Continue Open Ticket Refresh',
      `Current Status:\n${status}${historicalNote}\n\n` +
      `This will continue from where it left off.\n\n` +
      'Continue?',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) return;

    logOperation('Ticket Data', 'CONTINUE', `Open ticket refresh continued (historical=${isHistorical})`);

    const result = runOpenTicketRefresh(sheet);

    const skippedInfo = result.skippedCount ? `\nSkipped ${result.skippedCount} tickets (different school year).` : '';
    ui.alert(
      result.complete ? 'Complete' : 'Paused',
      `Updated ${result.updatedCount} existing rows.\n` +
      `Appended ${result.appendedCount} new tickets.${skippedInfo}\n` +
      `Processed ${result.ticketCount} tickets in ${result.batchCount} batches.\n` +
      `${result.complete ? 'Refresh complete!' : 'Run again to continue.'}\n\n` +
      `Runtime: ${(result.runtime / 1000).toFixed(1)} seconds`,
      ui.ButtonSet.OK
    );

  } finally {
    releaseScriptLock(lock);
  }
}

/**
 * Menu: Show open refresh status
 */
function showOpenRefreshStatus() {
  const ui = SpreadsheetApp.getUi();
  const status = getOpenRefreshStatusText();

  ui.alert('Open Ticket Refresh Status', status, ui.ButtonSet.OK);
}

/**
 * Get human-readable status for open ticket refresh
 */
function getOpenRefreshStatusText() {
  const progress = getOpenRefreshProgress();

  const lines = [];

  // Last run
  if (progress.lastRun) {
    const lastRunDate = new Date(progress.lastRun);
    lines.push(`Last refresh: ${lastRunDate.toLocaleString()}`);
  } else {
    lines.push('Never refreshed (will look back STALE_DAYS on first run)');
  }

  // Current progress
  if (progress.complete) {
    lines.push(`Status: COMPLETE (ready for next cycle)`);
  } else if (progress.page >= 0) {
    lines.push(`Status: IN PROGRESS (page ${progress.page + 1})`);
  } else {
    lines.push(`Status: NOT STARTED`);
  }

  return lines.join('\n');
}

/**
 * Legacy function - redirects to Start Fresh
 * @deprecated Use refreshOpenTicketsStart or refreshOpenTicketsContinue
 */
function refreshOpenTickets() {
  refreshOpenTicketsStart();
}

/**
 * Clear school year data and reset progress
 * Prevents duplicates when reloading after a partial wipe.
 * Also unlocks the school year configuration so it can be changed.
 */
function clearYearDataAndResetProgress() {
  const ui = SpreadsheetApp.getUi();

  // SAFETY: Require no triggers before destructive operation
  if (!requireNoTriggers('Clear Data + Reset Progress')) {
    return;
  }

  // SAFETY: Acquire lock to prevent concurrent operations
  const lock = acquireScriptLock();
  if (!lock) {
    showOperationBusyMessage('Clear Data + Reset Progress');
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('TicketData');

    if (!sheet) {
      ui.alert('Error', 'TicketData sheet not found.', ui.ButtonSet.OK);
      return;
    }

    const config = getConfig();
    if (!config.schoolYear) {
      ui.alert('No School Year Configured', 'No SCHOOL_YEAR found in Config sheet.', ui.ButtonSet.OK);
      return;
    }

    const configLockStatus = checkConfigLock(config);
    let lockWarning = '';
    if (configLockStatus.locked) {
      lockWarning = `\n3. UNLOCK configuration (currently locked):\n` +
        `   - SCHOOL_YEAR: ${config.schoolYearLoaded}\n` +
        `   - PAGE_SIZE: ${config.pageSizeLoaded}\n` +
        `   - BATCH_SIZE: ${config.batchSizeLoaded}\n` +
        `   After clearing, you can change these values in Config sheet.\n`;
    }

    const response = ui.alert(
      `Clear School Year Data`,
      `This will:\n` +
        `1. Delete all TicketData rows for ${config.schoolYear}\n` +
        `2. Reset progress keys in Config` +
        lockWarning + `\n\n` +
        `Continue?`,
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) return;

    const deleted = deleteRowsBySchoolYear(sheet, config.schoolYear);
    resetSchoolYearProgress();
    unlockConfig();  // Unlock so config values can be changed
    setCurrentYearReloadInProgress(true);

    logOperation('Ticket Data', 'CLEAR_SCHOOL_YEAR', `Cleared ${deleted} rows for ${config.schoolYear}, reset progress, and unlocked configuration`);

    ui.alert(
      'Clear Complete',
      `Deleted ${deleted} rows for ${config.schoolYear}.\n` +
        `Progress reset and configuration unlocked.\n\n` +
        `You can now change SCHOOL_YEAR, PAGE_SIZE, or BATCH_SIZE in Config if needed.\n` +
        `Run "Continue Loading" to reload.`,
      ui.ButtonSet.OK
    );

  } finally {
    releaseScriptLock(lock);
  }
}

/**
 * Core logic for refreshing recently modified tickets
 * Uses ModifiedDate filter to only fetch tickets that changed since last refresh
 *
 * This is much more efficient than fetching ALL open tickets:
 * - Only processes tickets that actually changed
 * - Single phase instead of separate open/closed phases
 * - Captures status changes, assignments, closures, and SLA updates
 *
 * Progress is saved to Config sheet after each batch:
 * - OPEN_REFRESH_LAST_RUN: Timestamp of last successful refresh start
 * - OPEN_REFRESH_PAGE: Current page (-1 = not started)
 * - OPEN_REFRESH_COMPLETE: TRUE when refresh is done
 */
function runOpenTicketRefresh(sheet) {
  const startTime = Date.now();
  const config = getConfig();

  // Get last refresh timestamp - if none, use STALE_DAYS as fallback window
  let lastRefreshTime = config.openRefreshLastRun;
  if (!lastRefreshTime) {
    // First run - look back STALE_DAYS
    const staleDays = config.staleDays || 7;
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - staleDays);
    lastRefreshTime = cutoff.toISOString();
    logOperation('Ticket Data', 'REFRESH_INIT', `First run - looking back ${staleDays} days to ${lastRefreshTime}`);
  }

  // Record this refresh start time (will be saved as lastRun when complete)
  const thisRefreshTime = new Date().toISOString();

  // Get current progress
  const progress = getOpenRefreshProgress();

  // Check if already complete for this cycle
  if (progress.complete) {
    logOperation('Ticket Data', 'REFRESH_SKIP', 'Open ticket refresh already complete for this cycle');
    return { batchCount: 0, ticketCount: 0, updatedCount: 0, appendedCount: 0, complete: true, runtime: 0 };
  }

  // Build TicketId -> row number map
  logOperation('Ticket Data', 'REFRESH_MAP', 'Building TicketId lookup map...');
  const ticketIdToRow = buildTicketIdMap(sheet);
  logOperation('Ticket Data', 'REFRESH_MAP', `Map built with ${ticketIdToRow.size} tickets`);

  // Setup
  let batchCount = 0;
  let ticketCount = 0;
  let updatedCount = 0;
  let appendedCount = 0;
  let skippedCount = 0;
  const batchSize = config.ticketBatchSize;
  const now = new Date();

  // For historical school years, only update existing tickets (don't append tickets from other school years)
  const isHistorical = !isSchoolYearCurrent(config);
  const batchOptions = isHistorical ? { updateOnly: true } : {};

  if (isHistorical) {
    logOperation('Ticket Data', 'REFRESH_MODE', 'Historical school year - update only, no appends');
  }

  // Build date filter for modified tickets
  const modifiedDate = new Date(lastRefreshTime);
  const dateFilter = `date>=${formatDateForApi(modifiedDate)}`;

  let page = progress.page + 1;
  let totalTickets = 0;
  let complete = false;

  logOperation('Ticket Data', 'REFRESH_START',
    `Fetching tickets modified since ${lastRefreshTime}, starting page ${page}`);

  while (Date.now() - startTime < MAX_RUNTIME_MS && !complete) {
    logOperation('Ticket Data', 'REFRESH_MODIFIED', `Fetching modified tickets, page ${page}`);

    const response = searchTickets([
      { Facet: 'modifieddate', Value: dateFilter }
    ], page, batchSize, { field: 'TicketModifiedDate', direction: 'asc' });

    batchCount++;

    if (page === 0) {
      totalTickets = response.Paging ? (response.Paging.TotalRows || response.Paging.Total || 0) : 0;
      logOperation('Ticket Data', 'REFRESH_MODIFIED', `API reports ${totalTickets} tickets modified since last refresh`);
    }

    if (!response.Items || response.Items.length === 0) {
      complete = true;
      break;
    }

    // Process this batch
    const result = processTicketBatch(sheet, response.Items, ticketIdToRow, now, batchOptions);
    updatedCount += result.updated;
    appendedCount += result.appended;
    skippedCount += result.skipped || 0;
    ticketCount += response.Items.length;

    // Save progress after each batch (string to prevent Sheets date formatting)
    updateConfigValue('OPEN_REFRESH_PAGE', String(page));

    const skipInfo = result.skipped ? `, ${result.skipped} skipped` : '';
    logOperation('Ticket Data', 'REFRESH_MODIFIED',
      `Page ${page}: ${response.Items.length} tickets (${result.updated} updated, ${result.appended} new${skipInfo})`);

    // Check if complete
    const totalExpectedPages = Math.ceil(totalTickets / batchSize);
    if (page >= totalExpectedPages - 1 || response.Items.length < batchSize) {
      complete = true;
    } else {
      page++;
      Utilities.sleep(config.throttleMs);
    }
  }

  // If complete, save the refresh timestamp and reset progress for next cycle
  if (complete) {
    updateConfigValue('OPEN_REFRESH_LAST_RUN', thisRefreshTime);
    updateConfigValue('OPEN_REFRESH_COMPLETE', 'TRUE');
    updateConfigValue('OPEN_REFRESH_PAGE', '-1');  // String to prevent Sheets date formatting
  }

  const runtime = Date.now() - startTime;

  logOperation('Ticket Data', complete ? 'REFRESH_COMPLETE' : 'REFRESH_PAUSED',
    `${updatedCount} updated, ${appendedCount} appended, ${ticketCount} total in ${batchCount} batches, ${(runtime/1000).toFixed(1)}s`);

  // Update last refresh timestamp so analytics sheets show current data
  if (ticketCount > 0) {
    updateLastRefresh();
  }

  return { batchCount, ticketCount, updatedCount, appendedCount, skippedCount, complete, runtime };
}

/**
 * Get current open refresh progress from Config
 */
function getOpenRefreshProgress() {
  const config = getConfig();
  return {
    page: config.openRefreshPage !== undefined ? config.openRefreshPage : -1,
    complete: config.openRefreshComplete === true,
    lastRun: config.openRefreshLastRun || null
  };
}

/**
 * Reset open refresh progress (for next cycle)
 */
function resetOpenRefreshProgress() {
  updateConfigValue('OPEN_REFRESH_PAGE', '-1');  // String to prevent Sheets date formatting
  updateConfigValue('OPEN_REFRESH_COMPLETE', 'FALSE');
  // Note: Don't reset OPEN_REFRESH_LAST_RUN - we need it for the next cycle's date filter
}

/**
 * Build a map of TicketId -> row number for fast lookups
 */
function buildTicketIdMap(sheet) {
  const ticketIds = sheet.getRange('A2:A' + sheet.getLastRow()).getValues();
  const map = new Map();

  for (let i = 0; i < ticketIds.length; i++) {
    const ticketId = ticketIds[i][0];
    if (ticketId) {
      map.set(ticketId, i + 2); // Row number (1-indexed, +1 for header, +1 for 0-index)
    }
  }

  return map;
}

/**
 * Process a batch of tickets - update existing rows or append new ones
 * Fetches SLA data for the batch and merges it into ticket rows
 *
 * @param {Sheet} sheet - TicketData sheet
 * @param {Array} tickets - Array of ticket objects from API
 * @param {Map} ticketIdToRow - Map of TicketId -> row number
 * @param {Date} now - Current timestamp for age calculation
 * @param {Object} options - Optional settings
 * @param {boolean} options.updateOnly - If true, only update existing tickets, don't append new ones
 */
function processTicketBatch(sheet, tickets, ticketIdToRow, now, options) {
  const updateOnly = options && options.updateOnly;
  let updated = 0;
  let appended = 0;
  let skipped = 0;
  const rowsToAppend = [];

  // Get school year from config for new rows
  const config = getConfig();
  const schoolYear = config.schoolYear || '';

  // Fetch SLA data for this batch
  const ticketIds = tickets.map(t => t.TicketId).filter(id => id);
  const slaMap = fetchSlaForTicketIds(ticketIds);

  for (const ticket of tickets) {
    const ticketId = ticket.TicketId;
    const rowData = extractTicketRow(ticket, now, schoolYear, slaMap);

    const existingRow = ticketIdToRow.get(ticketId);

    if (existingRow) {
      // Update existing row
      sheet.getRange(existingRow, 1, 1, 36).setValues([rowData]);
      updated++;
    } else if (updateOnly) {
      // Skip - don't append tickets not already in sheet (e.g., from different school year)
      skipped++;
    } else {
      // Queue for append
      rowsToAppend.push(rowData);
      appended++;
    }
  }

  // Batch append new rows
  if (rowsToAppend.length > 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, rowsToAppend.length, 36).setValues(rowsToAppend);

    // Update the map with new rows (for subsequent batches)
    for (let i = 0; i < rowsToAppend.length; i++) {
      ticketIdToRow.set(rowsToAppend[i][0], lastRow + 1 + i);
    }
  }

  return { updated, appended, skipped };
}
