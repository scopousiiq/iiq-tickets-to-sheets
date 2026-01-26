/**
 * Ticket Data Export
 *
 * Comprehensive ticket data extraction with:
 * - Year-based pagination for historical data (auto-discovered from Config)
 * - Date windowing for current year incremental updates
 * - Auto-run until timeout with resume capability
 * - Single TicketData sheet for Power BI consumption
 * - Consolidated SLA metrics (fetched per-batch during ticket loading)
 *
 * Year Configuration (via Config sheet):
 * - Historical years: Add TICKET_{YEAR}_LAST_PAGE, TICKET_{YEAR}_COMPLETE rows
 * - Current year: Add TICKET_{YEAR}_LAST_FETCH row
 * - Years are auto-discovered - no code changes needed
 *
 * SLA Data:
 * - SLA metrics are fetched for each batch of tickets (single API call per batch)
 * - Columns 29-35: ResponseThreshold, ResponseActual, ResponseBreach,
 *   ResolutionThreshold, ResolutionActual, ResolutionBreach, IsRunning
 */

// Safe runtime limit (5.5 minutes to allow for cleanup before 6 min Apps Script limit)
const MAX_RUNTIME_MS = 5.5 * 60 * 1000;

/**
 * Main entry point - Continue loading ticket data
 * Runs until timeout or completion, then saves progress
 */
function refreshTicketDataContinue() {
  const ui = SpreadsheetApp.getUi();
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
}

/**
 * Full reload - clears ALL data and resets ALL progress
 */
function refreshTicketDataFull() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TicketData');

  if (!sheet) {
    ui.alert('Error', 'TicketData sheet not found.', ui.ButtonSet.OK);
    return;
  }

  const config = getConfig();
  const years = [...config.historicalYears];
  if (config.currentYear) years.push(config.currentYear);

  const response = ui.alert(
    'Full Reload',
    'This will:\n' +
    '1. Clear ALL ticket data\n' +
    `2. Reset progress for ALL years (${years.join(', ')})\n\n` +
    'This cannot be undone. Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  // Clear all data (keep header row)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 35).clear();
  }

  // Reset all year progress
  config.historicalYears.forEach(year => resetYearProgress(year));
  if (config.currentYear) {
    resetCurrentYearProgress(config.currentYear);
  }

  logOperation('Ticket Data', 'FULL_RESET', 'Cleared all data and reset all progress');
  ui.alert('Reset Complete', 'All data cleared. Use "Continue Loading" to start fresh.', ui.ButtonSet.OK);
}

/**
 * Refresh current year only - for catching deletions
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
  if (!config.currentYear) {
    ui.alert('Error', 'No current year configured. Add a TICKET_{YEAR}_LAST_FETCH row to Config.', ui.ButtonSet.OK);
    return;
  }

  const currentYear = config.currentYear;

  const response = ui.alert(
    `Refresh Current Year (${currentYear})`,
    'This will:\n' +
    `1. Delete all ${currentYear} ticket rows\n` +
    `2. Reset ${currentYear} progress\n` +
    `3. Re-fetch all ${currentYear} tickets\n\n` +
    'Use this periodically to catch deleted tickets.\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  // Delete all rows where Year = currentYear
  deleteRowsByYear(sheet, currentYear);

  // Reset current year progress
  resetCurrentYearProgress(currentYear);

  logOperation('Ticket Data', 'RESET_CURRENT', `Cleared ${currentYear} data, ready for re-fetch`);

  // Start loading
  const result = runTicketDataLoader(sheet);

  ui.alert(
    result.complete ? 'Complete' : 'Paused',
    `Processed ${result.batchCount} batches, ${result.ticketCount} tickets.\n` +
    `${result.complete ? 'Current year refreshed!' : 'Run "Continue Loading" to finish.'}`,
    ui.ButtonSet.OK
  );
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
 * Get human-readable status for each year
 */
function getTicketDataStatus() {
  const config = getConfig();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TicketData');
  const totalRows = sheet ? Math.max(0, sheet.getLastRow() - 1) : 0;

  const statusLines = [];

  // Status for each historical year
  config.historicalYears.forEach(year => {
    const complete = config[`ticket${year}Complete`];
    const lastPage = config[`ticket${year}LastPage`];
    const totalPages = config[`ticket${year}TotalPages`];

    let status;
    if (complete) {
      status = 'Complete';
    } else if (lastPage < 0) {
      status = 'Not started';
    } else if (totalPages > 0) {
      status = `Page ${lastPage + 1} of ${totalPages}`;
    } else {
      status = `Page ${lastPage + 1} (total unknown)`;
    }
    statusLines.push(`• ${year}: ${status}`);
  });

  // Status for current year
  if (config.currentYear) {
    const lastFetch = config[`ticket${config.currentYear}LastFetch`];
    let status;
    if (!lastFetch) {
      status = 'Not started';
    } else {
      status = `Incremental (last: ${lastFetch.substring(0, 10)})`;
    }
    statusLines.push(`• ${config.currentYear} (current): ${status}`);
  }

  return {
    statusText: statusLines.join('\n'),
    totalRows: totalRows
  };
}

/**
 * Main loader loop - runs until timeout or complete
 */
function runTicketDataLoader(sheet) {
  const startTime = Date.now();
  let batchCount = 0;
  let ticketCount = 0;
  let complete = false;

  while (Date.now() - startTime < MAX_RUNTIME_MS) {
    const config = getConfig();

    // Find next work to do
    // 1. Check historical years in order
    let processed = false;
    for (const year of config.historicalYears) {
      if (!config[`ticket${year}Complete`]) {
        const result = processHistoricalYearBatch(sheet, year, config);
        batchCount++;
        ticketCount += result.count;
        processed = true;
        break; // Process one batch, then re-check
      }
    }

    // 2. If all historical complete, do current year
    if (!processed && config.currentYear) {
      const result = processCurrentYearBatch(sheet, config);
      batchCount++;
      ticketCount += result.count;

      if (!result.hasMore) {
        complete = true;
        break;
      }
    } else if (!processed && !config.currentYear) {
      // No current year configured, and all historical complete
      complete = true;
      break;
    }

    // Throttle between batches
    Utilities.sleep(getThrottleMs());
  }

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
 * Process one batch for a historical year using pagination
 */
function processHistoricalYearBatch(sheet, year, config) {
  const batchSize = config.ticketBatchSize;
  const lastPage = config[`ticket${year}LastPage`];
  const nextPage = lastPage + 1;

  // Build date range for the year
  const startDate = `01/01/${year}`;
  const endDate = `12/31/${year}`;
  const dateRange = `daterange:${startDate}-${endDate}`;

  logOperation('Ticket Data', 'BATCH', `Year ${year}, page ${nextPage}`);

  // Fetch page
  const response = searchTickets([
    { Facet: 'createddate', Value: dateRange }
  ], nextPage, batchSize, { field: 'TicketCreatedDate', direction: 'asc' });

  // On first page, calculate total pages
  if (nextPage === 0 && response.Paging) {
    // Handle both TotalRows and Total field names (API may use either)
    const totalRows = response.Paging.TotalRows || response.Paging.Total || 0;
    const totalPages = Math.ceil(Number(totalRows) / batchSize);
    // Write as string to prevent Sheets from auto-formatting as date
    updateConfigValue(`TICKET_${year}_TOTAL_PAGES`, String(totalPages));
    logOperation('Ticket Data', 'INFO', `Year ${year}: ${totalRows} tickets, ${totalPages} pages`);
  }

  if (!response.Items || response.Items.length === 0) {
    // No more data - mark complete
    updateConfigValue(`TICKET_${year}_COMPLETE`, 'TRUE');
    logOperation('Ticket Data', 'COMPLETE', `Year ${year} finished`);
    return { count: 0 };
  }

  // Fetch SLA data for this batch of tickets
  const ticketIds = response.Items.map(t => t.TicketId).filter(id => id);
  const slaMap = fetchSlaForTicketIds(ticketIds);
  logOperation('Ticket Data', 'SLA_FETCH', `Fetched SLA for ${slaMap.size} of ${ticketIds.length} tickets`);

  // Write tickets to sheet (with merged SLA data)
  const now = new Date();
  const rows = response.Items.map(ticket => extractTicketRow(ticket, now, year, slaMap));
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, rows.length, 35).setValues(rows);

  // Update progress
  updateConfigValue(`TICKET_${year}_LAST_PAGE`, nextPage);

  // Check if complete
  const totalPages = config[`ticket${year}TotalPages`] ||
    (response.Paging ? Math.ceil(response.Paging.TotalRows / batchSize) : 0);

  if (nextPage >= totalPages - 1) {
    updateConfigValue(`TICKET_${year}_COMPLETE`, 'TRUE');
    logOperation('Ticket Data', 'COMPLETE', `Year ${year} finished`);
  }

  logOperation('Ticket Data', 'BATCH', `Year ${year} page ${nextPage}: wrote ${rows.length} tickets`);
  return { count: rows.length };
}

/**
 * Process one batch for current year using date windowing
 */
function processCurrentYearBatch(sheet, config) {
  const batchSize = config.ticketBatchSize;
  const currentYear = config.currentYear;
  const lastFetch = config[`ticket${currentYear}LastFetch`];

  // Determine start date
  let startDate;
  if (lastFetch) {
    // Start from the date of last fetch
    const lastDate = new Date(lastFetch);
    startDate = formatDateForApi(lastDate);
  } else {
    // Fresh start - beginning of current year
    startDate = `01/01/${currentYear}`;
  }

  const endDate = formatDateForApi(new Date());
  const dateRange = `daterange:${startDate}-${endDate}`;

  logOperation('Ticket Data', 'BATCH', `Year ${currentYear}, from ${startDate}`);

  // Fetch tickets
  const response = searchTickets([
    { Facet: 'createddate', Value: dateRange }
  ], 0, batchSize, { field: 'TicketCreatedDate', direction: 'asc' });

  if (!response.Items || response.Items.length === 0) {
    logOperation('Ticket Data', 'COMPLETE', `Year ${currentYear} up to date`);
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
      updateConfigValue(`TICKET_${currentYear}_LAST_FETCH`, lastTicket.CreatedDate);
    }
    return { count: 0, hasMore: hasMore };
  }

  // Sort to ensure order
  tickets.sort((a, b) => new Date(a.CreatedDate) - new Date(b.CreatedDate));

  // Fetch SLA data for this batch of tickets
  const ticketIds = tickets.map(t => t.TicketId).filter(id => id);
  const slaMap = fetchSlaForTicketIds(ticketIds);
  logOperation('Ticket Data', 'SLA_FETCH', `Fetched SLA for ${slaMap.size} of ${ticketIds.length} tickets`);

  // Write to sheet (with merged SLA data)
  const now = new Date();
  const rows = tickets.map(ticket => extractTicketRow(ticket, now, currentYear, slaMap));
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, rows.length, 35).setValues(rows);

  // Update last fetch timestamp
  const lastTicket = tickets[tickets.length - 1];
  updateConfigValue(`TICKET_${currentYear}_LAST_FETCH`, lastTicket.CreatedDate);

  // Check if more pages
  const totalRows = response.Paging ? (response.Paging.TotalRows || response.Paging.Total || 0) : 0;
  const hasMore = totalRows > batchSize;

  logOperation('Ticket Data', 'BATCH', `Year ${currentYear}: wrote ${rows.length} tickets`);
  return { count: rows.length, hasMore: hasMore };
}

/**
 * Extract a row of data from a ticket object
 * Returns 35 columns for comprehensive analytics (including SLA metrics)
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
    // E-H: Dates and status
    ticket.CreatedDate || '',
    ticket.ModifiedDate || '',
    ticket.ClosedDate || '',
    ticket.IsClosed ? 'Yes' : 'No',
    // I: Workflow status
    ticket.WorkflowStep ? (ticket.WorkflowStep.StatusName || ticket.WorkflowStep.Name || '') : '',
    // J-K: Team
    ticket.AssignedToTeam ? (ticket.AssignedToTeam.TeamId || '') : '',
    ticket.AssignedToTeam ? (ticket.AssignedToTeam.TeamName || '') : '',
    // L-N: Location
    ticket.Location ? (ticket.Location.LocationId || '') : '',
    ticket.Location ? (ticket.Location.Name || '') : '',
    locationType,
    // O-P: Owner
    ticket.Owner ? (ticket.Owner.UserId || '') : '',
    ticket.Owner ? (ticket.Owner.Name || '') : '',
    // Q: Age
    ageDays,
    // R-T: Priority and Due Date
    ticket.Priority || '',
    ticket.IsPastDue ? 'Yes' : 'No',
    ticket.DueDate || '',
    // U-V: SLA (basic)
    ticketSla.SlaId || '',
    slaName,
    // W-Z: Issue Category and Type
    issue.IssueCategoryId || '',
    issue.IssueCategoryName || '',
    issue.IssueTypeId || '',
    issue.Name || '', // Issue type name is in Name field
    // AA-AB: Requester (For)
    ticket.For ? (ticket.For.UserId || '') : '',
    ticket.For ? (ticket.For.Name || '') : '',
    // AC-AI: SLA Metrics (from detailed SLA API)
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
 * Update a config value by key name
 */
function updateConfigValue(key, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Config');
  const data = sheet.getDataRange().getValues();

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }

  // If not found, add it
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, 2).setValues([[key, value]]);
}

/**
 * Reset progress for a historical year
 */
function resetYearProgress(year) {
  updateConfigValue(`TICKET_${year}_TOTAL_PAGES`, '');
  updateConfigValue(`TICKET_${year}_LAST_PAGE`, -1);
  updateConfigValue(`TICKET_${year}_COMPLETE`, 'FALSE');
}

/**
 * Reset progress for current year
 */
function resetCurrentYearProgress(year) {
  updateConfigValue(`TICKET_${year}_LAST_FETCH`, '');
}

/**
 * Delete all rows from sheet where Year column matches
 */
function deleteRowsByYear(sheet, year) {
  const data = sheet.getDataRange().getValues();
  const yearCol = 3; // Column D (0-indexed = 3) is Year

  // Find rows to delete (from bottom up to preserve indices)
  const rowsToDelete = [];
  for (let i = data.length - 1; i >= 1; i--) { // Skip header
    if (data[i][yearCol] === year) {
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

    logOperation('Ticket Data', 'DELETE', `Deleted ${rowsToDelete.length} rows for year ${year}`);
  }
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
 * @deprecated Use triggerTicketDataReset + triggerTicketDataContinue instead
 * Menu: Start a fresh open ticket refresh (resets progress)
 */
function refreshOpenTicketsStart() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TicketData');

  if (!sheet) {
    ui.alert('Error', 'TicketData sheet not found.', ui.ButtonSet.OK);
    return;
  }

  const config = getConfig();
  const staleDays = config.staleDays || 7;

  const response = ui.alert(
    'Start Open Ticket Refresh',
    `This will reset progress and start a fresh refresh:\n\n` +
    `1. Fetch all open tickets from API\n` +
    `2. Fetch tickets closed in last ${staleDays} days\n` +
    `3. Update existing rows or append new ones\n\n` +
    `Progress is saved after each batch.\n` +
    `Use "Continue" if this times out.\n\n` +
    'Start fresh?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  // Reset progress to start fresh
  resetOpenRefreshProgress();
  updateConfigValue('OPEN_REFRESH_DATE', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));

  logOperation('Ticket Data', 'START', 'Open/closed ticket refresh started fresh');

  const result = runOpenTicketRefresh(sheet);

  ui.alert(
    result.complete ? 'Complete' : 'Paused',
    `Updated ${result.updatedCount} existing rows.\n` +
    `Appended ${result.appendedCount} new tickets.\n` +
    `Processed ${result.ticketCount} tickets in ${result.batchCount} batches.\n` +
    `${result.complete ? 'Refresh complete!' : 'Use "Continue" to resume.'}\n\n` +
    `Runtime: ${(result.runtime / 1000).toFixed(1)} seconds`,
    ui.ButtonSet.OK
  );
}

/**
 * Menu: Continue an in-progress open ticket refresh
 */
function refreshOpenTicketsContinue() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TicketData');

  if (!sheet) {
    ui.alert('Error', 'TicketData sheet not found.', ui.ButtonSet.OK);
    return;
  }

  const status = getOpenRefreshStatusText();

  const response = ui.alert(
    'Continue Open Ticket Refresh',
    `Current Status:\n${status}\n\n` +
    `This will continue from where it left off.\n\n` +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  logOperation('Ticket Data', 'CONTINUE', 'Open/closed ticket refresh continued');

  const result = runOpenTicketRefresh(sheet);

  ui.alert(
    result.complete ? 'Complete' : 'Paused',
    `Updated ${result.updatedCount} existing rows.\n` +
    `Appended ${result.appendedCount} new tickets.\n` +
    `Processed ${result.ticketCount} tickets in ${result.batchCount} batches.\n` +
    `${result.complete ? 'Refresh complete!' : 'Run again to continue.'}\n\n` +
    `Runtime: ${(result.runtime / 1000).toFixed(1)} seconds`,
    ui.ButtonSet.OK
  );
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
  const config = getConfig();
  const progress = getOpenRefreshProgress();

  const lines = [];

  // Date
  if (config.openRefreshDate) {
    lines.push(`Refresh date: ${config.openRefreshDate}`);
  } else {
    lines.push('No refresh in progress');
  }

  // Open tickets phase
  if (progress.openComplete) {
    lines.push(`Open tickets: Complete`);
  } else if (progress.openPage >= 0) {
    lines.push(`Open tickets: Page ${progress.openPage + 1} (in progress)`);
  } else {
    lines.push(`Open tickets: Not started`);
  }

  // Closed tickets phase
  if (progress.closedComplete) {
    lines.push(`Recently closed: Complete`);
  } else if (progress.closedPage >= 0) {
    lines.push(`Recently closed: Page ${progress.closedPage + 1} (in progress)`);
  } else {
    lines.push(`Recently closed: Not started`);
  }

  // Overall status
  if (progress.openComplete && progress.closedComplete) {
    lines.push(`\nStatus: COMPLETE`);
  } else {
    lines.push(`\nStatus: IN PROGRESS`);
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
 * Core logic for refreshing open and recently closed tickets
 * Uses in-place updates instead of delete+append for efficiency
 *
 * Progress is saved to Config sheet after each batch:
 * - OPEN_REFRESH_DATE: Date of current refresh (resets if new day)
 * - OPEN_REFRESH_OPEN_PAGE: Current page for open tickets (-1 = not started)
 * - OPEN_REFRESH_OPEN_COMPLETE: TRUE when open phase done
 * - OPEN_REFRESH_CLOSED_PAGE: Current page for closed tickets (-1 = not started)
 * - OPEN_REFRESH_CLOSED_COMPLETE: TRUE when closed phase done
 */
function runOpenTicketRefresh(sheet) {
  const startTime = Date.now();
  const config = getConfig();
  const staleDays = config.staleDays || 7;

  // Check/reset progress for new day
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const savedDate = config.openRefreshDate || '';

  if (savedDate !== today) {
    // New day - reset progress
    logOperation('Ticket Data', 'REFRESH_RESET', `New day (${today}), resetting progress`);
    resetOpenRefreshProgress();
    updateConfigValue('OPEN_REFRESH_DATE', today);
  }

  // Re-read config after potential reset
  const progress = getOpenRefreshProgress();

  // Check if already complete
  if (progress.openComplete && progress.closedComplete) {
    logOperation('Ticket Data', 'REFRESH_SKIP', 'Open ticket refresh already complete for today');
    return { batchCount: 0, ticketCount: 0, updatedCount: 0, appendedCount: 0, complete: true, runtime: 0 };
  }

  logOperation('Ticket Data', 'REFRESH_START',
    `Resuming refresh: openPage=${progress.openPage}, openComplete=${progress.openComplete}, ` +
    `closedPage=${progress.closedPage}, closedComplete=${progress.closedComplete}`);

  // Step 1: Build TicketId -> row number map
  logOperation('Ticket Data', 'REFRESH_MAP', 'Building TicketId lookup map...');
  const ticketIdToRow = buildTicketIdMap(sheet);
  logOperation('Ticket Data', 'REFRESH_MAP', `Map built with ${ticketIdToRow.size} tickets`);

  // Step 2: Fetch and process tickets
  let batchCount = 0;
  let ticketCount = 0;
  let updatedCount = 0;
  let appendedCount = 0;
  const batchSize = config.ticketBatchSize;
  const now = new Date();

  // Phase 1: Fetch open tickets (resume from saved page)
  let openComplete = progress.openComplete;
  let openPage = progress.openPage + 1; // Start from next page
  let openTotal = 0;
  let openProcessed = 0;

  while (Date.now() - startTime < MAX_RUNTIME_MS && !openComplete) {
    logOperation('Ticket Data', 'REFRESH_OPEN', `Fetching open tickets, page ${openPage}`);

    const response = searchTickets([
      { Facet: 'ticketstate', Value: 'open' }
    ], openPage, batchSize, { field: 'TicketCreatedDate', direction: 'asc' });

    batchCount++;

    if (openPage === 0 || openTotal === 0) {
      openTotal = response.Paging ? (response.Paging.TotalRows || response.Paging.Total || 0) : 0;
      logOperation('Ticket Data', 'REFRESH_OPEN', `API reports ${openTotal} open tickets`);
    }

    if (!response.Items || response.Items.length === 0) {
      openComplete = true;
      updateConfigValue('OPEN_REFRESH_OPEN_COMPLETE', 'TRUE');
      break;
    }

    // Process this batch
    const result = processTicketBatch(sheet, response.Items, ticketIdToRow, now);
    updatedCount += result.updated;
    appendedCount += result.appended;
    ticketCount += response.Items.length;
    openProcessed += response.Items.length;

    // Save progress after each batch
    updateConfigValue('OPEN_REFRESH_OPEN_PAGE', openPage);

    logOperation('Ticket Data', 'REFRESH_OPEN',
      `Page ${openPage}: ${response.Items.length} tickets (${result.updated} updated, ${result.appended} new)`);

    // Check if phase complete
    const totalExpectedPages = Math.ceil(openTotal / batchSize);
    if (openPage >= totalExpectedPages - 1 || response.Items.length < batchSize) {
      openComplete = true;
      updateConfigValue('OPEN_REFRESH_OPEN_COMPLETE', 'TRUE');
    } else {
      openPage++;
      Utilities.sleep(config.throttleMs);
    }
  }

  // Phase 2: Fetch recently closed tickets (resume from saved page)
  let closedComplete = progress.closedComplete;
  let closedPage = progress.closedPage + 1; // Start from next page
  let closedTotal = 0;
  let closedProcessed = 0;

  if (openComplete && !closedComplete && Date.now() - startTime < MAX_RUNTIME_MS) {
    // Build date filter for recently closed
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - staleDays);
    const dateFilter = `date>=${formatDateForApi(cutoffDate)}`;

    while (Date.now() - startTime < MAX_RUNTIME_MS && !closedComplete) {
      logOperation('Ticket Data', 'REFRESH_CLOSED', `Fetching recently closed tickets, page ${closedPage}`);

      const response = searchTickets([
        { Facet: 'ticketstate', Value: 'closed' },
        { Facet: 'closeddate', Value: dateFilter }
      ], closedPage, batchSize, { field: 'TicketClosedDate', direction: 'desc' });

      batchCount++;

      if (closedPage === 0 || closedTotal === 0) {
        closedTotal = response.Paging ? (response.Paging.TotalRows || response.Paging.Total || 0) : 0;
        logOperation('Ticket Data', 'REFRESH_CLOSED', `API reports ${closedTotal} recently closed tickets`);
      }

      if (!response.Items || response.Items.length === 0) {
        closedComplete = true;
        updateConfigValue('OPEN_REFRESH_CLOSED_COMPLETE', 'TRUE');
        break;
      }

      // Process this batch
      const result = processTicketBatch(sheet, response.Items, ticketIdToRow, now);
      updatedCount += result.updated;
      appendedCount += result.appended;
      closedProcessed += response.Items.length;
      ticketCount += response.Items.length;

      // Save progress after each batch
      updateConfigValue('OPEN_REFRESH_CLOSED_PAGE', closedPage);

      logOperation('Ticket Data', 'REFRESH_CLOSED',
        `Page ${closedPage}: ${response.Items.length} tickets (${result.updated} updated, ${result.appended} new)`);

      // Check if phase complete
      const totalExpectedPages = Math.ceil(closedTotal / batchSize);
      if (closedPage >= totalExpectedPages - 1 || response.Items.length < batchSize) {
        closedComplete = true;
        updateConfigValue('OPEN_REFRESH_CLOSED_COMPLETE', 'TRUE');
      } else {
        closedPage++;
        Utilities.sleep(config.throttleMs);
      }
    }
  }

  const complete = openComplete && closedComplete;
  const runtime = Date.now() - startTime;

  logOperation('Ticket Data', complete ? 'REFRESH_COMPLETE' : 'REFRESH_PAUSED',
    `${updatedCount} updated, ${appendedCount} appended, ${ticketCount} total in ${batchCount} batches, ${(runtime/1000).toFixed(1)}s`);

  // Update last refresh timestamp so analytics sheets show current data
  if (ticketCount > 0) {
    updateLastRefresh();
  }

  return { batchCount, ticketCount, updatedCount, appendedCount, complete, runtime };
}

/**
 * Get current open refresh progress from Config
 */
function getOpenRefreshProgress() {
  const config = getConfig();
  return {
    openPage: config.openRefreshOpenPage !== undefined ? config.openRefreshOpenPage : -1,
    openComplete: config.openRefreshOpenComplete === true,
    closedPage: config.openRefreshClosedPage !== undefined ? config.openRefreshClosedPage : -1,
    closedComplete: config.openRefreshClosedComplete === true
  };
}

/**
 * Reset open refresh progress (for new day or manual reset)
 */
function resetOpenRefreshProgress() {
  updateConfigValue('OPEN_REFRESH_OPEN_PAGE', -1);
  updateConfigValue('OPEN_REFRESH_OPEN_COMPLETE', 'FALSE');
  updateConfigValue('OPEN_REFRESH_CLOSED_PAGE', -1);
  updateConfigValue('OPEN_REFRESH_CLOSED_COMPLETE', 'FALSE');
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
 */
function processTicketBatch(sheet, tickets, ticketIdToRow, now) {
  let updated = 0;
  let appended = 0;
  const rowsToAppend = [];

  // Fetch SLA data for this batch
  const ticketIds = tickets.map(t => t.TicketId).filter(id => id);
  const slaMap = fetchSlaForTicketIds(ticketIds);

  for (const ticket of tickets) {
    const ticketId = ticket.TicketId;
    const createdDate = ticket.CreatedDate ? new Date(ticket.CreatedDate) : null;
    const year = createdDate ? createdDate.getFullYear() : new Date().getFullYear();
    const rowData = extractTicketRow(ticket, now, year, slaMap);

    const existingRow = ticketIdToRow.get(ticketId);

    if (existingRow) {
      // Update existing row
      sheet.getRange(existingRow, 1, 1, 35).setValues([rowData]);
      updated++;
    } else {
      // Queue for append
      rowsToAppend.push(rowData);
      appended++;
    }
  }

  // Batch append new rows
  if (rowsToAppend.length > 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, rowsToAppend.length, 35).setValues(rowsToAppend);

    // Update the map with new rows (for subsequent batches)
    for (let i = 0; i < rowsToAppend.length; i++) {
      ticketIdToRow.set(rowsToAppend[i][0], lastRow + 1 + i);
    }
  }

  return { updated, appended };
}
