/**
 * IIQ API Configuration
 * Reads settings from the Config sheet
 *
 * Year Discovery:
 * - Historical years are auto-detected from TICKET_{YEAR}_LAST_PAGE rows
 * - Current year is auto-detected from TICKET_{YEAR}_LAST_FETCH row
 * - To add a year: add TICKET_{YEAR}_LAST_PAGE, TICKET_{YEAR}_COMPLETE rows
 * - To remove a year: delete those rows (and optionally clear data)
 */

function getConfig() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  const data = sheet.getRange('A2:B50').getValues(); // Extended range for dynamic years

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

  // Auto-discover years from config keys
  const discoveredYears = discoverYearsFromConfig(rawConfig);

  // Build the config object with base settings
  const config = {
    baseUrl: rawConfig['API_BASE_URL'] || '',
    bearerToken: rawConfig['BEARER_TOKEN'] || '',
    siteId: rawConfig['SITE_ID'] || '',
    pageSize: getIntValue(rawConfig['PAGE_SIZE'], 100),
    throttleMs: getIntValue(rawConfig['THROTTLE_MS'], 1000),
    staleDays: getIntValue(rawConfig['STALE_DAYS'], 7),
    slaRiskPercent: getIntValue(rawConfig['SLA_RISK_PERCENT'], 75),
    ticketBatchSize: getIntValue(rawConfig['TICKET_BATCH_SIZE'], 2000),
    slaBatchSize: getIntValue(rawConfig['SLA_BATCH_SIZE'], 100),
    // Discovered years
    historicalYears: discoveredYears.historical,
    currentYear: discoveredYears.current
  };

  // Dynamically add ticket tracking properties for each historical year
  discoveredYears.historical.forEach(year => {
    config[`ticket${year}TotalPages`] = getIntValue(rawConfig[`TICKET_${year}_TOTAL_PAGES`], 0);
    config[`ticket${year}LastPage`] = getIntValue(rawConfig[`TICKET_${year}_LAST_PAGE`], -1);
    config[`ticket${year}Complete`] = getBoolValue(rawConfig[`TICKET_${year}_COMPLETE`]);
  });

  // Add current year tracking (date windowing)
  if (discoveredYears.current) {
    config[`ticket${discoveredYears.current}LastFetch`] = getStringValue(rawConfig[`TICKET_${discoveredYears.current}_LAST_FETCH`]);
  }

  // Dynamically add SLA tracking properties for each historical year
  discoveredYears.historical.forEach(year => {
    config[`sla${year}LastPage`] = getIntValue(rawConfig[`SLA_${year}_LAST_PAGE`], -1);
    config[`sla${year}Complete`] = getBoolValue(rawConfig[`SLA_${year}_COMPLETE`]);
  });

  // Add current year SLA tracking
  if (discoveredYears.current) {
    config[`sla${discoveredYears.current}LastFetch`] = getStringValue(rawConfig[`SLA_${discoveredYears.current}_LAST_FETCH`]);
  }

  return config;
}

/**
 * Discover configured years by scanning config keys
 * - TICKET_{YEAR}_LAST_PAGE → historical year (pagination-based)
 * - TICKET_{YEAR}_LAST_FETCH → current year (date windowing)
 *
 * @param {Object} rawConfig - Raw config key-value pairs
 * @returns {Object} - { historical: [2023, 2024, ...], current: 2026 }
 */
function discoverYearsFromConfig(rawConfig) {
  const historicalYears = [];
  let currentYear = null;

  // Regex patterns to extract years
  const lastPagePattern = /^TICKET_(\d{4})_LAST_PAGE$/;
  const lastFetchPattern = /^TICKET_(\d{4})_LAST_FETCH$/;

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
 */
function getConfiguredYears() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  const data = sheet.getRange('A2:B50').getValues();

  const rawConfig = {};
  data.forEach(row => {
    if (row[0]) {
      rawConfig[row[0]] = row[1];
    }
  });

  return discoverYearsFromConfig(rawConfig);
}

function updateLastRefresh() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  const now = new Date().toISOString();

  // Find LAST_REFRESH row and update
  const data = sheet.getRange('A2:A20').getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === 'LAST_REFRESH') {
      sheet.getRange(i + 2, 2).setValue(now);
      break;
    }
  }
}

function logOperation(operation, status, details) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Logs');
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
