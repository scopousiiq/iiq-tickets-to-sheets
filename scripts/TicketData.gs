/**
 * Ticket Data Export
 *
 * Comprehensive ticket data extraction with:
 * - Year-based pagination for historical data (auto-discovered from Config)
 * - Date windowing for current year incremental updates
 * - Auto-run until timeout with resume capability
 * - Single TicketData sheet for Power BI consumption
 *
 * Year Configuration (via Config sheet):
 * - Historical years: Add TICKET_{YEAR}_LAST_PAGE, TICKET_{YEAR}_COMPLETE rows
 * - Current year: Add TICKET_{YEAR}_LAST_FETCH row
 * - Years are auto-discovered - no code changes needed
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
    sheet.getRange(2, 1, lastRow - 1, 28).clear();
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

  // Write tickets to sheet
  const now = new Date();
  const rows = response.Items.map(ticket => extractTicketRow(ticket, now, year));
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, rows.length, 28).setValues(rows);

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

  // Write to sheet
  const now = new Date();
  const rows = tickets.map(ticket => extractTicketRow(ticket, now, currentYear));
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, rows.length, 28).setValues(rows);

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
 * Returns 28 columns for comprehensive analytics
 */
function extractTicketRow(ticket, now, year) {
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

  // Extract SLA details
  const sla = ticket.Sla || {};

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
    // U-V: SLA
    sla.SlaId || '',
    sla.Name || sla.SlaName || '',
    // W-Z: Issue Category and Type
    issue.IssueCategoryId || '',
    issue.IssueCategoryName || '',
    issue.IssueTypeId || '',
    issue.Name || '', // Issue type name is in Name field
    // AA-AB: Requester (For)
    ticket.For ? (ticket.For.UserId || '') : '',
    ticket.For ? (ticket.For.Name || '') : ''
  ];
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
