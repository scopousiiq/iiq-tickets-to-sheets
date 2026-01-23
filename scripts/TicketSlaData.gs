/**
 * Ticket SLA Data Export
 *
 * Fetches SLA timing data from /tickets/slas endpoint.
 * Uses same year-based pagination strategy as TicketData.
 * API is capped at 100 records per page.
 *
 * Year Configuration (via Config sheet):
 * - Historical years: Add SLA_{YEAR}_LAST_PAGE, SLA_{YEAR}_COMPLETE rows
 * - Current year: Add SLA_{YEAR}_LAST_FETCH row
 * - Years are auto-discovered - no code changes needed
 *
 * Menu Options:
 * - Continue Loading: Resume from where it left off
 * - Full Reload: Clear all data and start fresh
 * - Refresh Current Year: Re-fetch current year data (handles deletions)
 * - Show Status: Display progress for each year
 */

const SLA_MAX_RUNTIME_MS = 5.5 * 60 * 1000; // 5.5 minutes safe limit

/**
 * Continue loading SLA data from where it left off
 */
function refreshSlaDataContinue() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('TicketSlaData');

  if (!sheet) {
    sheet = ss.insertSheet('TicketSlaData');
    sheet.getRange(1, 1, 1, 11).setValues([[
      'TicketId', 'Year', 'SlaId', 'SlaName',
      'ResponseThreshold', 'ResponseActual', 'ResponseBreached',
      'ResolutionThreshold', 'ResolutionActual', 'ResolutionBreached',
      'IsRunning'
    ]]);
    sheet.setFrozenRows(1);
  }

  // Show current status before starting
  const status = getSlaDataStatus();
  const response = ui.alert(
    'Continue Loading SLA Data',
    `Current Status:\n${status.statusText}\n\n` +
    `This will run until timeout (~5.5 min) or completion.\n` +
    `Progress is saved after each batch.\n\n` +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  logOperation('SLA Data', 'START', 'Continue loading SLA data');

  const result = runSlaDataLoader(sheet);

  ui.alert(
    result.complete ? 'Complete' : 'Paused',
    `Processed ${result.batchCount} batches, ${result.slaCount} SLA records.\n` +
    `${result.complete ? 'All data loaded!' : 'Run again to continue.'}\n\n` +
    `Runtime: ${(result.runtime / 1000).toFixed(1)} seconds\n` +
    `Total rows: ${sheet.getLastRow() - 1}`,
    ui.ButtonSet.OK
  );
}

/**
 * Get SLA data loading status
 * Note: SLA endpoint doesn't return Paging, so we only track page number
 */
function getSlaDataStatus() {
  const config = getConfig();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TicketSlaData');
  const totalRows = sheet ? Math.max(0, sheet.getLastRow() - 1) : 0;

  const statusLines = [];

  // Status for each historical year
  config.historicalYears.forEach(year => {
    const complete = config[`sla${year}Complete`];
    const lastPage = config[`sla${year}LastPage`];

    let status;
    if (complete) {
      status = 'Complete';
    } else if (lastPage < 0) {
      status = 'Not started';
    } else {
      status = `Page ${lastPage + 1} (in progress)`;
    }
    statusLines.push(`• ${year}: ${status}`);
  });

  // Status for current year
  if (config.currentYear) {
    const lastFetch = config[`sla${config.currentYear}LastFetch`];
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
 * Full reload - clear all SLA data and start fresh
 */
function refreshSlaDataFull() {
  const ui = SpreadsheetApp.getUi();
  const config = getConfig();
  const years = [...config.historicalYears];
  if (config.currentYear) years.push(config.currentYear);

  const response = ui.alert(
    'Full SLA Data Reload',
    `This will delete ALL SLA data and reload from scratch.\n` +
    `Years to reset: ${years.join(', ')}\n\n` +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('TicketSlaData');

  if (sheet) {
    // Clear all data except header
    if (sheet.getLastRow() > 1) {
      sheet.deleteRows(2, sheet.getLastRow() - 1);
    }
  } else {
    sheet = ss.insertSheet('TicketSlaData');
    sheet.getRange(1, 1, 1, 11).setValues([[
      'TicketId', 'Year', 'SlaId', 'SlaName',
      'ResponseThreshold', 'ResponseActual', 'ResponseBreached',
      'ResolutionThreshold', 'ResolutionActual', 'ResolutionBreached',
      'IsRunning'
    ]]);
    sheet.setFrozenRows(1);
  }

  // Reset all progress
  config.historicalYears.forEach(year => resetSlaYearProgress(year));
  if (config.currentYear) {
    resetSlaCurrentYearProgress(config.currentYear);
  }

  logOperation('SLA Data', 'RESET', 'Full reload initiated');

  refreshSlaDataContinue();
}

/**
 * Refresh current year only (handles deletions)
 */
function refreshSlaDataCurrentYear() {
  const ui = SpreadsheetApp.getUi();
  const config = getConfig();

  if (!config.currentYear) {
    ui.alert('Error', 'No current year configured. Add a SLA_{YEAR}_LAST_FETCH row to Config.', ui.ButtonSet.OK);
    return;
  }

  const currentYear = config.currentYear;

  const response = ui.alert(
    `Refresh Current Year SLA Data (${currentYear})`,
    `This will delete all ${currentYear} SLA data and reload it. Continue?`,
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TicketSlaData');

  if (!sheet) {
    ui.alert('Error', 'TicketSlaData sheet not found. Run Continue Loading first.', ui.ButtonSet.OK);
    return;
  }

  // Delete current year rows
  deleteSlaRowsByYear(sheet, currentYear);
  resetSlaCurrentYearProgress(currentYear);

  logOperation('SLA Data', 'RESET', `Reset ${currentYear} for refresh`);

  refreshSlaDataContinue();
}

/**
 * Show loading status for SLA data
 */
function showSlaDataStatus() {
  const status = getSlaDataStatus();

  SpreadsheetApp.getUi().alert(
    'SLA Data Status',
    `Total rows: ${status.totalRows}\n\n${status.statusText}`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Main loader loop - runs until timeout or complete
 * Matches TicketData loader pattern exactly
 */
function runSlaDataLoader(sheet) {
  const startTime = Date.now();
  let batchCount = 0;
  let slaCount = 0;
  let complete = false;

  while (Date.now() - startTime < SLA_MAX_RUNTIME_MS) {
    // Reload config fresh each iteration to get updated progress
    const config = getConfig();

    // Find next work to do
    // 1. Check historical years in order
    let processed = false;
    for (const year of config.historicalYears) {
      if (!config[`sla${year}Complete`]) {
        const result = processSlaHistoricalYearBatch(sheet, year, config);
        batchCount++;
        slaCount += result.count;
        processed = true;
        break; // Process one batch, then re-check
      }
    }

    // 2. If all historical complete, do current year
    if (!processed && config.currentYear) {
      const result = processSlaCurrentYearBatch(sheet, config);
      batchCount++;
      slaCount += result.count;

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
    Utilities.sleep(config.throttleMs);
  }

  const runtime = Date.now() - startTime;
  logOperation('SLA Data', complete ? 'COMPLETE' : 'PAUSED',
    `${batchCount} batches, ${slaCount} SLA records, ${(runtime/1000).toFixed(1)}s`);

  return { batchCount, slaCount, complete, runtime };
}

/**
 * Process one batch for a historical year using pagination
 * Note: SLA endpoint does NOT return Paging data, so we iterate until
 * we get fewer items than batch size (indicating last page)
 */
function processSlaHistoricalYearBatch(sheet, year, config) {
  const batchSize = config.slaBatchSize; // Capped at 100 by API
  const lastPage = config[`sla${year}LastPage`];
  const nextPage = lastPage + 1;

  // Build date range for the year
  const startDate = `01/01/${year}`;
  const endDate = `12/31/${year}`;
  const dateRange = `daterange:${startDate}-${endDate}`;

  logOperation('SLA Data', 'BATCH', `Year ${year}, page ${nextPage}`);

  // Fetch page
  const response = listTicketSlas([
    { Facet: 'createddate', Value: dateRange }
  ], nextPage, batchSize, { field: 'TicketCreatedDate', direction: 'asc' });

  if (!response || !response.Items || response.Items.length === 0) {
    // No more data - mark complete
    updateConfigValue(`SLA_${year}_COMPLETE`, 'TRUE');
    SpreadsheetApp.flush();
    logOperation('SLA Data', 'COMPLETE', `Year ${year} finished (no more data on page ${nextPage})`);
    return { count: 0 };
  }

  // Write SLA records to sheet
  const rows = response.Items.map(item => extractSlaRow(item, year));
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, rows.length, 11).setValues(rows);

  // Update progress - track page number for resume capability
  updateConfigValue(`SLA_${year}_LAST_PAGE`, nextPage);

  // Check if complete: fewer items than batch size means last page
  // (SLA endpoint doesn't return Paging metadata)
  if (rows.length < batchSize) {
    updateConfigValue(`SLA_${year}_COMPLETE`, 'TRUE');
    logOperation('SLA Data', 'COMPLETE', `Year ${year} finished (page ${nextPage}, ${rows.length} items < ${batchSize})`);
  }

  SpreadsheetApp.flush();

  logOperation('SLA Data', 'BATCH', `Year ${year} page ${nextPage}: wrote ${rows.length} SLA records`);
  return { count: rows.length };
}

/**
 * Process one batch for current year using date windowing
 * Matches TicketData pattern exactly
 */
function processSlaCurrentYearBatch(sheet, config) {
  const batchSize = config.slaBatchSize;
  const currentYear = config.currentYear;
  const lastFetch = config[`sla${currentYear}LastFetch`];

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

  logOperation('SLA Data', 'BATCH', `Year ${currentYear}, from ${startDate}`);

  // Fetch one batch of SLA records
  const response = listTicketSlas([
    { Facet: 'createddate', Value: dateRange }
  ], 0, batchSize, { field: 'TicketCreatedDate', direction: 'asc' });

  if (!response || !response.Items || response.Items.length === 0) {
    logOperation('SLA Data', 'COMPLETE', `Year ${currentYear} up to date`);
    return { count: 0, hasMore: false };
  }

  // Filter out already-fetched records (by timestamp comparison)
  let items = response.Items;
  if (lastFetch) {
    const lastFetchTime = new Date(lastFetch).getTime();
    items = items.filter(item => {
      const createdDate = item.CreatedDate || (item.Ticket && item.Ticket.CreatedDate);
      return createdDate && new Date(createdDate).getTime() > lastFetchTime;
    });
  }

  if (items.length === 0) {
    // All items in batch already fetched, but there may be more if we got a full batch
    const hasMore = response.Items.length >= batchSize;
    if (hasMore) {
      // Advance the window using the last item from response
      const lastItem = response.Items[response.Items.length - 1];
      const lastTimestamp = lastItem.CreatedDate || (lastItem.Ticket && lastItem.Ticket.CreatedDate);
      if (lastTimestamp) {
        updateConfigValue(`SLA_${currentYear}_LAST_FETCH`, lastTimestamp);
        SpreadsheetApp.flush();
      }
    }
    return { count: 0, hasMore: hasMore };
  }

  // Sort to ensure order by created date
  items.sort((a, b) => {
    const aDate = a.CreatedDate || (a.Ticket && a.Ticket.CreatedDate) || '';
    const bDate = b.CreatedDate || (b.Ticket && b.Ticket.CreatedDate) || '';
    return new Date(aDate) - new Date(bDate);
  });

  // Write to sheet
  const rows = items.map(item => extractSlaRow(item, currentYear));
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, rows.length, 11).setValues(rows);

  // Update last fetch timestamp
  const lastItem = items[items.length - 1];
  const lastTimestamp = lastItem.CreatedDate || (lastItem.Ticket && lastItem.Ticket.CreatedDate);
  if (lastTimestamp) {
    updateConfigValue(`SLA_${currentYear}_LAST_FETCH`, lastTimestamp);
  }

  SpreadsheetApp.flush();

  // Check if more pages - if we got a full batch from API, there may be more
  // (SLA endpoint doesn't return Paging metadata)
  const hasMore = response.Items.length >= batchSize;

  logOperation('SLA Data', 'BATCH', `Year ${currentYear}: wrote ${rows.length} SLA records`);
  return { count: rows.length, hasMore: hasMore };
}

/**
 * Extract a row from SLA response item
 */
function extractSlaRow(item, year) {
  const ticket = item.Ticket || {};
  const sla = item.Sla || {};
  const slaTimes = item.SlaTimes || [];
  const metrics = sla.Metrics || [];

  // Find response metric (MetricType = 1) and resolution metric (MetricType = 2)
  const responseMetric = metrics.find(m => m.MetricType === 1) || {};
  const resolutionMetric = metrics.find(m => m.MetricType === 2) || {};

  // Find actual times from SlaTimes
  const responseTime = slaTimes.find(t => t.MetricType === 1) || {};
  const resolutionTime = slaTimes.find(t => t.MetricType === 2) || {};

  const responseThreshold = responseMetric.ValueInMinutes || null;
  const responseActual = responseTime.LogMinutes || null;
  const resolutionThreshold = resolutionMetric.ValueInMinutes || null;
  const resolutionActual = resolutionTime.LogMinutes || null;

  // Calculate breach status
  const responseBreached = (responseActual !== null && responseThreshold !== null && responseActual > responseThreshold);
  const resolutionBreached = (resolutionActual !== null && resolutionThreshold !== null && resolutionActual > resolutionThreshold);

  // Check if SLA is still running (ticket not closed)
  const isRunning = item.IsRunning === true || !ticket.IsClosed;

  return [
    ticket.TicketId || item.TicketId || '',
    year,
    sla.SlaId || '',
    sla.SlaName || sla.Name || '',
    responseThreshold,
    responseActual,
    responseBreached,
    resolutionThreshold,
    resolutionActual,
    resolutionBreached,
    isRunning
  ];
}

/**
 * Call the /tickets/slas endpoint
 */
function listTicketSlas(filters, page, pageSize, sortOptions) {
  const config = getConfig();
  const size = pageSize || 100;

  let endpoint = `/v1.0/tickets/slas?$p=${page || 0}&$s=${size}`;

  // Add sorting - default to TicketCreatedDate asc for consistent pagination
  if (sortOptions && sortOptions.field) {
    const sortExpr = `${sortOptions.field} ${sortOptions.direction || 'asc'}`;
    endpoint += `&$o=${encodeURIComponent(sortExpr)}`;
  }

  const payload = {
    Filters: filters || []
  };

  return makeApiRequest(endpoint, 'POST', payload);
}

/**
 * Reset progress for a historical year (SLA)
 */
function resetSlaYearProgress(year) {
  updateConfigValue(`SLA_${year}_LAST_PAGE`, -1);
  updateConfigValue(`SLA_${year}_COMPLETE`, 'FALSE');
}

/**
 * Reset progress for current year (SLA)
 */
function resetSlaCurrentYearProgress(year) {
  updateConfigValue(`SLA_${year}_LAST_FETCH`, '');
}

/**
 * Delete all rows from sheet where Year column matches (SLA)
 */
function deleteSlaRowsByYear(sheet, year) {
  const data = sheet.getDataRange().getValues();
  const yearCol = 1; // Column B (0-indexed = 1) is Year

  // Find rows to delete (from bottom up to preserve indices)
  const rowsToDelete = [];
  for (let i = data.length - 1; i >= 1; i--) { // Skip header
    if (data[i][yearCol] === year) {
      rowsToDelete.push(i + 1); // Convert to 1-indexed
    }
  }

  // Delete in batches
  if (rowsToDelete.length > 0) {
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

    logOperation('SLA Data', 'DELETE', `Deleted ${rowsToDelete.length} rows for year ${year}`);
  }
}
