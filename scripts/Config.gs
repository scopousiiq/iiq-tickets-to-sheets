/**
 * iiQ API Configuration
 * Reads settings from the Config sheet
 *
 * School Year Model:
 * - Each spreadsheet contains ONE school year's data (e.g., 2025-2026)
 * - School year is configured via SCHOOL_YEAR (e.g., "2025-2026")
 * - School year start date via SCHOOL_YEAR_START (MM-DD format, default "07-01")
 * - Progress tracking uses simplified keys without year suffix
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

  // Helper to get MM-DD string (handles Date objects from Sheets for SCHOOL_YEAR_START)
  function getMonthDayString(val, defaultVal) {
    if (!val || val === '') return defaultVal;
    if (val instanceof Date) {
      // Format Date object as MM-DD (ignore year)
      const month = String(val.getMonth() + 1).padStart(2, '0');
      const day = String(val.getDate()).padStart(2, '0');
      return `${month}-${day}`;
    }
    // If it's already a string in MM-DD format, return as-is
    const str = String(val).trim();
    if (/^\d{2}-\d{2}$/.test(str)) {
      return str;
    }
    // Try to parse as a date string and extract MM-DD
    const parsed = new Date(str);
    if (!isNaN(parsed.getTime())) {
      const month = String(parsed.getMonth() + 1).padStart(2, '0');
      const day = String(parsed.getDate()).padStart(2, '0');
      return `${month}-${day}`;
    }
    return defaultVal;
  }

  // Normalize base URL: strip trailing slash, ensure /api suffix
  let baseUrl = rawConfig['API_BASE_URL'] || '';
  if (baseUrl) {
    baseUrl = baseUrl.replace(/\/+$/, ''); // Remove trailing slashes
    if (!baseUrl.endsWith('/api')) {
      baseUrl = baseUrl + '/api';
    }
  }

  // Parse school year configuration
  const schoolYear = getStringValue(rawConfig['SCHOOL_YEAR']) || '';
  const schoolYearStart = getMonthDayString(rawConfig['SCHOOL_YEAR_START'], '07-01');

  // Locked configuration values (set when loading starts)
  const schoolYearLoaded = getStringValue(rawConfig['SCHOOL_YEAR_LOADED']) || '';
  const pageSizeLoaded = getIntValue(rawConfig['PAGE_SIZE_LOADED'], 0) || null;
  const batchSizeLoaded = getIntValue(rawConfig['BATCH_SIZE_LOADED'], 0) || null;

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
    // School year configuration
    schoolYear: schoolYear,
    schoolYearStart: schoolYearStart,
    // Locked values (set when data loading starts, cleared by "Clear Data + Reset")
    schoolYearLoaded: schoolYearLoaded,
    pageSizeLoaded: pageSizeLoaded,
    batchSizeLoaded: batchSizeLoaded
  };

  // Simplified progress tracking (no year suffix)
  config.ticketTotalPages = getIntValue(rawConfig['TICKET_TOTAL_PAGES'], -1);
  config.ticketLastPage = getIntValue(rawConfig['TICKET_LAST_PAGE'], -1);
  config.ticketComplete = getBoolValue(rawConfig['TICKET_COMPLETE']);
  config.ticketLastFetch = getStringValue(rawConfig['TICKET_LAST_FETCH']);

  // Open refresh progress tracking (simplified - uses ModifiedDate filter)
  config.openRefreshLastRun = getStringValue(rawConfig['OPEN_REFRESH_LAST_RUN']);
  config.openRefreshPage = getIntValue(rawConfig['OPEN_REFRESH_PAGE'], -1);
  config.openRefreshComplete = getBoolValue(rawConfig['OPEN_REFRESH_COMPLETE']);

  return config;
}

/**
 * Calculate school year start and end dates from config
 *
 * @param {Object} config - Config object with schoolYear and schoolYearStart
 * @returns {Object} - { startDate, endDate, startYear, endYear } or null if invalid
 */
function getSchoolYearDates(config) {
  if (!config.schoolYear) return null;

  // Parse school year (e.g., "2025-2026")
  const yearMatch = config.schoolYear.match(/^(\d{4})-(\d{4})$/);
  if (!yearMatch) return null;

  const startYear = parseInt(yearMatch[1], 10);
  const endYear = parseInt(yearMatch[2], 10);

  // Validate consecutive years
  if (endYear !== startYear + 1) return null;

  // Parse start date (MM-DD format, default "07-01")
  const startParts = (config.schoolYearStart || '07-01').split('-');
  const startMonth = parseInt(startParts[0], 10) - 1; // 0-indexed for Date constructor
  const startDay = parseInt(startParts[1], 10);

  // School year starts on startYear and ends day before next start
  // Note: new Date(year, month, 0) gives last day of previous month
  // So for startDay=1 (July 1), new Date(2026, 6, 0) = June 30, 2026
  const startDate = new Date(startYear, startMonth, startDay);
  const endDate = new Date(endYear, startMonth, startDay - 1);

  return { startDate, endDate, startYear, endYear };
}

/**
 * Check if the configured school year is currently in progress
 * (i.e., today falls within the school year date range)
 *
 * @param {Object} config - Config object with schoolYear
 * @returns {boolean} - true if school year is current, false if historical or invalid
 */
function isSchoolYearCurrent(config) {
  const dates = getSchoolYearDates(config);
  if (!dates) return false;

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  return today >= dates.startDate && today <= dates.endDate;
}

/**
 * Check if the configuration is locked
 * Config is locked when data loading has started (SCHOOL_YEAR_LOADED is set)
 *
 * @param {Object} config - Config object
 * @returns {boolean} - true if locked, false if can be changed
 */
function isConfigLocked(config) {
  return config.schoolYearLoaded && config.schoolYearLoaded !== '';
}

/**
 * @deprecated Use isConfigLocked instead
 */
function isSchoolYearLocked(config) {
  return isConfigLocked(config);
}

/**
 * Check if the current config matches the locked config values
 *
 * @param {Object} config - Config object
 * @returns {Object} - { locked, matches, mismatches[] } or { locked: false } if not locked
 */
function checkConfigLock(config) {
  if (!isConfigLocked(config)) {
    return { locked: false, matches: true, mismatches: [] };
  }

  const mismatches = [];

  // Check school year
  if (config.schoolYear !== config.schoolYearLoaded) {
    mismatches.push({
      key: 'SCHOOL_YEAR',
      current: config.schoolYear,
      locked: config.schoolYearLoaded
    });
  }

  // Check page size
  if (config.pageSizeLoaded && config.pageSize !== config.pageSizeLoaded) {
    mismatches.push({
      key: 'PAGE_SIZE',
      current: config.pageSize,
      locked: config.pageSizeLoaded
    });
  }

  // Check batch size
  if (config.batchSizeLoaded && config.ticketBatchSize !== config.batchSizeLoaded) {
    mismatches.push({
      key: 'TICKET_BATCH_SIZE',
      current: config.ticketBatchSize,
      locked: config.batchSizeLoaded
    });
  }

  return {
    locked: true,
    matches: mismatches.length === 0,
    mismatches: mismatches,
    loadedYear: config.schoolYearLoaded  // For backward compatibility
  };
}

/**
 * @deprecated Use checkConfigLock instead
 */
function checkSchoolYearLock(config) {
  const result = checkConfigLock(config);
  return {
    locked: result.locked,
    matches: result.matches,
    loadedYear: result.loadedYear
  };
}

/**
 * Lock the configuration by storing current values
 * Called when data loading starts
 * Locks: SCHOOL_YEAR, PAGE_SIZE, TICKET_BATCH_SIZE
 */
function lockConfig() {
  const config = getConfig();
  if (!config.schoolYear) {
    throw new Error('Cannot lock config: SCHOOL_YEAR not configured');
  }

  // Check for mismatches if already locked
  const lockStatus = checkConfigLock(config);
  if (lockStatus.locked && !lockStatus.matches) {
    const mismatchDetails = lockStatus.mismatches.map(m =>
      `${m.key}: config has "${m.current}" but was locked to "${m.locked}"`
    ).join('\n');
    throw new Error(
      `Configuration mismatch detected!\n${mismatchDetails}\n\n` +
      `Use "Clear Data + Reset Progress" to unlock and change these values.`
    );
  }

  // If already locked with same values, nothing to do
  if (lockStatus.locked && lockStatus.matches) {
    return;
  }

  // Lock all values
  setConfigValue('SCHOOL_YEAR_LOADED', config.schoolYear);
  setConfigValue('PAGE_SIZE_LOADED', config.pageSize);
  setConfigValue('BATCH_SIZE_LOADED', config.ticketBatchSize);
  logOperation('Config', 'LOCKED',
    `Configuration locked: SCHOOL_YEAR=${config.schoolYear}, PAGE_SIZE=${config.pageSize}, BATCH_SIZE=${config.ticketBatchSize}`);
}

/**
 * @deprecated Use lockConfig instead
 */
function lockSchoolYear() {
  lockConfig();
}

/**
 * Unlock the configuration
 * Called when data is cleared for a full reload
 */
function unlockConfig() {
  setConfigValue('SCHOOL_YEAR_LOADED', '');
  setConfigValue('PAGE_SIZE_LOADED', '');
  setConfigValue('BATCH_SIZE_LOADED', '');
  logOperation('Config', 'UNLOCKED', 'Configuration unlocked (SCHOOL_YEAR, PAGE_SIZE, BATCH_SIZE)');
}

/**
 * @deprecated Use unlockConfig instead
 */
function unlockSchoolYear() {
  unlockConfig();
}

/**
 * Set a config value by key
 *
 * @param {string} key - Config key to set
 * @param {any} value - Value to set
 */
function setConfigValue(key, value) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }

  // If not found, append it
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, 2).setValues([[key, value]]);
}

/**
 * @deprecated School year model uses SCHOOL_YEAR config instead
 */
function discoverSlaYearsFromConfig(rawConfig) {
  return [];
}

/**
 * @deprecated School year model uses single SCHOOL_YEAR config instead of year-based keys
 */
function discoverYearsFromConfig(rawConfig, prefix) {
  // Return empty for backward compatibility
  return {
    historical: [],
    current: null
  };
}

/**
 * @deprecated School year model uses single SCHOOL_YEAR config
 */
function getConfiguredYears(prefix) {
  return {
    historical: [],
    current: null
  };
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

// =============================================================================
// CONCURRENCY CONTROL - LockService helpers
// =============================================================================

/**
 * Acquire the script lock for long-running operations.
 * Prevents concurrent execution of data loading, refreshing, or destructive operations.
 *
 * @param {number} waitTimeMs - Time to wait for lock (default 2000ms for menu items)
 * @returns {GoogleAppsScript.Lock.Lock|null} - Lock object if acquired, null if failed
 */
function acquireScriptLock(waitTimeMs = 2000) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(waitTimeMs);
    return lock;
  } catch (e) {
    return null;
  }
}

/**
 * Try to acquire script lock without waiting (for triggers that should skip if busy)
 *
 * @param {number} waitTimeMs - Brief wait time (default 1000ms)
 * @returns {GoogleAppsScript.Lock.Lock|null} - Lock object if acquired, null if busy
 */
function tryAcquireScriptLock(waitTimeMs = 1000) {
  const lock = LockService.getScriptLock();
  const acquired = lock.tryLock(waitTimeMs);
  return acquired ? lock : null;
}

/**
 * Release the script lock
 *
 * @param {GoogleAppsScript.Lock.Lock} lock - Lock object to release
 */
function releaseScriptLock(lock) {
  if (lock) {
    try {
      lock.releaseLock();
    } catch (e) {
      // Lock may have already been released or expired - ignore
    }
  }
}

// =============================================================================
// TRIGGER SAFETY - Prevent destructive operations while triggers are active
// =============================================================================

/**
 * Check if any automated triggers are installed
 *
 * @returns {Object} - { hasTriggers: boolean, count: number, triggers: string[] }
 */
function checkForTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  const triggerNames = triggers.map(t => t.getHandlerFunction());
  return {
    hasTriggers: triggers.length > 0,
    count: triggers.length,
    triggers: triggerNames
  };
}

/**
 * Require that no triggers are installed before proceeding with destructive operations.
 * Shows an error dialog if triggers exist.
 *
 * @param {string} operationName - Name of the operation for error message
 * @returns {boolean} - true if safe to proceed (no triggers), false if blocked
 */
function requireNoTriggers(operationName) {
  const ui = SpreadsheetApp.getUi();
  const triggerStatus = checkForTriggers();

  if (triggerStatus.hasTriggers) {
    ui.alert(
      'Triggers Must Be Removed',
      `Cannot run "${operationName}" while automated triggers are installed.\n\n` +
      `${triggerStatus.count} trigger(s) found:\n• ${triggerStatus.triggers.join('\n• ')}\n\n` +
      `To proceed:\n` +
      `1. Go to iiQ Data > Setup > Remove Automated Triggers\n` +
      `2. Run "${operationName}" again\n` +
      `3. Optionally re-add triggers via Setup > Setup Automated Triggers`,
      ui.ButtonSet.OK
    );
    return false;
  }

  return true;
}

/**
 * Show a "busy" message when lock acquisition fails (for menu items)
 *
 * @param {string} operationName - Name of the operation that couldn't start
 */
function showOperationBusyMessage(operationName) {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Operation In Progress',
    `Cannot start "${operationName}" because another operation is currently running.\n\n` +
    `Please wait for the current operation to complete and try again.`,
    ui.ButtonSet.OK
  );
}
