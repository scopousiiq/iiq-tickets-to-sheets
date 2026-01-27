/**
 * iiQ API Configuration
 * Reads settings from the Config sheet
 *
 * Year Discovery (Tickets):
 * - Historical years are auto-detected from TICKET_{YEAR}_LAST_PAGE rows
 * - Current year is auto-detected from TICKET_{YEAR}_LAST_FETCH row
 * - To add a year: add TICKET_{YEAR}_LAST_PAGE, TICKET_{YEAR}_COMPLETE rows
 * - To remove a year: delete those rows (and optionally clear data)
 *
 * SLA Data:
 * - SLA metrics are consolidated into TicketData (columns 30-36)
 * - SLA is fetched per-batch during ticket loading, no separate SLA loading phase
 */

function getConfig() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  const allData = sheet.getDataRange().getValues();
  const data = allData.slice(1); // Skip header row

  const rawConfig = {};
  data.forEach(row => {
    if (row[0]) {
      rawConfig[row[0]] = row[1];
    }
  });

  // Helper to preserve string values (for ISO timestamps)
  function getStringValue(val) {
    if (!val) return null;
    if (val === '') return null;
    return String(val).trim() || null;
  }

  // Helper to get integer value with default
  function getIntValue(val, defaultVal) {
    if (val === '' || val === null || val === undefined) return defaultVal;
    const parsed = parseInt(val);
    return isNaN(parsed) ? defaultVal : parsed;
  }

  // Helper to get boolean value
  function getBoolValue(val) {
    if (val === true || val === 'TRUE' || val === 'true') return true;
    return false;
  }

  // Helper to get date as yyyy-MM-dd string (handles Date objects from Sheets)
  function getDateString(val) {
    if (!val) return '';
    if (val instanceof Date) {
      // Format Date object as yyyy-MM-dd
      const year = val.getFullYear();
      const month = String(val.getMonth() + 1).padStart(2, '0');
      const day = String(val.getDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    }
    // If it's already a string in yyyy-MM-dd format, return as-is
    const str = String(val).trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
      return str;
    }
    // Try to parse other date formats
    const parsed = new Date(str);
    if (!isNaN(parsed.getTime())) {
      const year = parsed.getFullYear();
      const month = String(parsed.getMonth() + 1).padStart(2, '0');
      const day = String(parsed.getDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    }
    return '';
  }

  // Auto-discover years from config keys for tickets
  const ticketYears = discoverYearsFromConfig(rawConfig, 'TICKET');

  // Normalize base URL: strip trailing slash, ensure /api suffix
  let baseUrl = rawConfig['API_BASE_URL'] || '';
  if (baseUrl) {
    baseUrl = baseUrl.replace(/\/+$/, ''); // Remove trailing slashes
    if (!baseUrl.endsWith('/api')) {
      baseUrl = baseUrl + '/api';
    }
  }

  // Build the config object with base settings
  const config = {
    baseUrl: baseUrl,
    bearerToken: rawConfig['BEARER_TOKEN'] || '',
    siteId: rawConfig['SITE_ID'] || '',
    pageSize: getIntValue(rawConfig['PAGE_SIZE'], 100),
    throttleMs: getIntValue(rawConfig['THROTTLE_MS'], 1000),
    staleDays: getIntValue(rawConfig['STALE_DAYS'], 7),
    slaRiskPercent: getIntValue(rawConfig['SLA_RISK_PERCENT'], 75),
    ticketBatchSize: getIntValue(rawConfig['TICKET_BATCH_SIZE'], 2000),
    // Discovered years for tickets
    historicalYears: ticketYears.historical,
    currentYear: ticketYears.current
  };

  // Dynamically add ticket tracking properties for each historical year
  ticketYears.historical.forEach(year => {
    config[`ticket${year}TotalPages`] = getIntValue(rawConfig[`TICKET_${year}_TOTAL_PAGES`], 0);
    config[`ticket${year}LastPage`] = getIntValue(rawConfig[`TICKET_${year}_LAST_PAGE`], -1);
    config[`ticket${year}Complete`] = getBoolValue(rawConfig[`TICKET_${year}_COMPLETE`]);
  });

  // Add current year ticket tracking (date windowing)
  if (ticketYears.current) {
    config[`ticket${ticketYears.current}LastFetch`] = getStringValue(rawConfig[`TICKET_${ticketYears.current}_LAST_FETCH`]);
  }

  // Open refresh progress tracking
  config.openRefreshDate = getDateString(rawConfig['OPEN_REFRESH_DATE']);
  config.openRefreshOpenPage = getIntValue(rawConfig['OPEN_REFRESH_OPEN_PAGE'], -1);
  config.openRefreshOpenComplete = getBoolValue(rawConfig['OPEN_REFRESH_OPEN_COMPLETE']);
  config.openRefreshClosedPage = getIntValue(rawConfig['OPEN_REFRESH_CLOSED_PAGE'], -1);
  config.openRefreshClosedComplete = getBoolValue(rawConfig['OPEN_REFRESH_CLOSED_COMPLETE']);

  return config;
}

/**
 * @deprecated SLA loading now uses open/closed priority instead of years
 */
function discoverSlaYearsFromConfig(rawConfig) {
  return [];
}

/**
 * Discover configured years by scanning config keys (for tickets)
 * - {PREFIX}_{YEAR}_LAST_PAGE → historical year (pagination-based)
 * - {PREFIX}_{YEAR}_LAST_FETCH → current year (date windowing)
 *
 * @param {Object} rawConfig - Raw config key-value pairs
 * @param {string} prefix - Key prefix ('TICKET')
 * @returns {Object} - { historical: [2023, 2024, ...], current: 2026 }
 */
function discoverYearsFromConfig(rawConfig, prefix) {
  const historicalYears = [];
  let currentYear = null;

  // Regex patterns to extract years (dynamic prefix)
  const lastPagePattern = new RegExp(`^${prefix}_(\\d{4})_LAST_PAGE$`);
  const lastFetchPattern = new RegExp(`^${prefix}_(\\d{4})_LAST_FETCH$`);

  Object.keys(rawConfig).forEach(key => {
    // Check for historical year (pagination)
    const pageMatch = key.match(lastPagePattern);
    if (pageMatch) {
      historicalYears.push(parseInt(pageMatch[1]));
      return;
    }

    // Check for current year (date windowing)
    const fetchMatch = key.match(lastFetchPattern);
    if (fetchMatch) {
      currentYear = parseInt(fetchMatch[1]);
    }
  });

  // Sort historical years ascending
  historicalYears.sort((a, b) => a - b);

  return {
    historical: historicalYears,
    current: currentYear
  };
}

/**
 * Get just the discovered years (useful for status display)
 * @param {string} prefix - Optional: 'TICKET' or 'SLA'. If omitted, returns ticket years for backward compatibility.
 */
function getConfiguredYears(prefix) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  const allData = sheet.getDataRange().getValues();
  const data = allData.slice(1); // Skip header row

  const rawConfig = {};
  data.forEach(row => {
    if (row[0]) {
      rawConfig[row[0]] = row[1];
    }
  });

  return discoverYearsFromConfig(rawConfig, prefix || 'TICKET');
}

function updateLastRefresh() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!sheet) return;

  const now = new Date().toISOString();

  // Find LAST_REFRESH row and update (search full config range)
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === 'LAST_REFRESH') {
      sheet.getRange(i + 1, 2).setValue(now);
      return;
    }
  }

  // If not found, append it
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, 2).setValues([['LAST_REFRESH', now]]);
}

function logOperation(operation, status, details) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Logs');
  if (!sheet) {
    // Logs sheet doesn't exist - fail silently to avoid breaking operations
    console.log(`[${operation}] ${status}: ${details}`);
    return;
  }

  sheet.appendRow([
    new Date().toISOString(),
    operation,
    status,
    details
  ]);

  // Keep only last 500 log entries
  const lastRow = sheet.getLastRow();
  if (lastRow > 501) {
    sheet.deleteRows(2, lastRow - 501);
  }
}
