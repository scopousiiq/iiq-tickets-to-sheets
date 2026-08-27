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
 * - Device/asset info from first asset in columns 37-39 (AK-AM)
 * - Assigned technician in columns 40-41 (AN-AO)
 * - Asset identifiers in columns 42-43 (AP-AQ)
 * - Custom field values in columns 44-46 (AR-AT)
 * - Location custom field values in columns 48-52 (AV-AZ)
 * - SLA is fetched per-batch during ticket loading, no separate SLA loading phase
 */

/** Current script version — update when releasing new versions */
const SCRIPT_VERSION = '1.8.1';

/**
 * How many Location custom field slots the sheet exposes.
 * Higher than the 3 ticket slots because location values are resolved once per
 * run from a single paginated call — extra slots cost columns, not API calls.
 * Districts observed in the wild define up to 8 location fields (region, site
 * group, org unit, hardware area), so 3 would be limiting.
 */
const LOCATION_CUSTOM_FIELD_COUNT = 5;

/** Reusable empty value row, so tickets with no location match allocate nothing. */
const LOCATION_CUSTOM_FIELD_BLANK_ROW = new Array(LOCATION_CUSTOM_FIELD_COUNT).fill('');

/**
 * How long a cached location custom field editor type stays trusted, in days.
 *
 * Editor types change only when an admin edits the field definition in iiQ,
 * which is rare, so this is deliberately long — it exists to bound staleness,
 * not to poll. "Refresh Custom Fields" rewrites the cache immediately, and
 * "Clear Data + Reset Progress" drops it, so a district never has to wait out
 * the window to correct a wrong type.
 */
const LOCATION_CF_TYPE_CACHE_DAYS = 30;

/** Number of columns in TicketData sheet (41 base + 2 asset ID + 3 custom field slots + 1 requester role + 5 location custom field slots) */
const TICKET_COLUMN_COUNT = 52;

/**
 * Telemetry Master /exec URL (iiQ-owned). Maintainer-managed — districts
 * never see or edit it. Empty string disables telemetry entirely for this
 * build. See iiq-sheets-telemetry/CLIENT_SETUP.md for rotation procedure.
 */
const TELEMETRY_URL = 'https://script.google.com/macros/s/AKfycbyaPAkUWjAqkYgX01WhJlNGNQuZACDtQ_6zNVUEbHD73RDaM5uWq7IwwqwD54mP9qXYZA/exec';

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

  // Product configuration (Ticketing, Facilities, or HRSD)
  const module = getStringValue(rawConfig['MODULE']) || 'Ticketing';
  const MODULE_PRODUCT_IDS = {
    'Ticketing': '88df910c-91aa-e711-80c2-0004ffa00010',
    'Facilities': '88df910c-91aa-e711-80c2-0004ffa00020',
    'HRSD': '88df910c-91aa-e711-80c2-0004ffa00060'
  };

  // Locked configuration values (set when loading starts)
  const schoolYearLoaded = getStringValue(rawConfig['SCHOOL_YEAR_LOADED']) || '';
  const pageSizeLoaded = getIntValue(rawConfig['PAGE_SIZE_LOADED'], 0) || null;
  const batchSizeLoaded = getIntValue(rawConfig['BATCH_SIZE_LOADED'], 0) || null;
  const moduleLoaded = getStringValue(rawConfig['MODULE_LOADED']) || '';

  // Build the config object with base settings
  const config = {
    baseUrl: baseUrl,
    bearerToken: rawConfig['BEARER_TOKEN'] || '',
    siteId: rawConfig['SITE_ID'] || '',
    module: module,
    moduleProductId: MODULE_PRODUCT_IDS[module] || MODULE_PRODUCT_IDS['Ticketing'],
    pageSize: getIntValue(rawConfig['PAGE_SIZE'], 100),
    throttleMs: getIntValue(rawConfig['THROTTLE_MS'], 1000),
    staleDays: getIntValue(rawConfig['STALE_DAYS'], 7),
    slaRiskPercent: getIntValue(rawConfig['SLA_RISK_PERCENT'], 75),
    ticketBatchSize: getIntValue(rawConfig['TICKET_BATCH_SIZE'], 2000),
    // School year configuration
    schoolYear: schoolYear,
    schoolYearStart: schoolYearStart,
    // Custom field columns (optional, up to 3)
    customField1: getStringValue(rawConfig['CUSTOM_FIELD_1']) || '',
    customField2: getStringValue(rawConfig['CUSTOM_FIELD_2']) || '',
    customField3: getStringValue(rawConfig['CUSTOM_FIELD_3']) || '',
    customField1Id: getStringValue(rawConfig['CUSTOM_FIELD_1_ID']) || '',
    customField2Id: getStringValue(rawConfig['CUSTOM_FIELD_2_ID']) || '',
    customField3Id: getStringValue(rawConfig['CUSTOM_FIELD_3_ID']) || '',
    // Locked values (set when data loading starts, cleared by "Clear Data + Reset")
    schoolYearLoaded: schoolYearLoaded,
    pageSizeLoaded: pageSizeLoaded,
    batchSizeLoaded: batchSizeLoaded,
    moduleLoaded: moduleLoaded,
    customField1Loaded: getStringValue(rawConfig['CUSTOM_FIELD_1_LOADED']) || '',
    customField2Loaded: getStringValue(rawConfig['CUSTOM_FIELD_2_LOADED']) || '',
    customField3Loaded: getStringValue(rawConfig['CUSTOM_FIELD_3_LOADED']) || ''
  };

  // Location custom fields (optional). Kept in parallel arrays rather than
  // numbered properties — there are 5 slots x 3 keys, and every consumer
  // iterates them. Indexes are 0-based; Config keys are 1-based.
  config.locationCustomFields = [];
  config.locationCustomFieldIds = [];
  config.locationCustomFieldsLoaded = [];
  for (let n = 1; n <= LOCATION_CUSTOM_FIELD_COUNT; n++) {
    config.locationCustomFields.push(getStringValue(rawConfig['LOCATION_CUSTOM_FIELD_' + n]) || '');
    config.locationCustomFieldIds.push(getStringValue(rawConfig['LOCATION_CUSTOM_FIELD_' + n + '_ID']) || '');
    config.locationCustomFieldsLoaded.push(getStringValue(rawConfig['LOCATION_CUSTOM_FIELD_' + n + '_LOADED']) || '');
  }

  // Editor type cache — one JSON cell keyed by CustomFieldTypeId, so swapping a
  // configured field invalidates its own entry without touching the others.
  config.locationCfTypeCache = parseLocationCfTypeCache(
    getStringValue(rawConfig['LOCATION_CF_TYPE_CACHE']) || ''
  );

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

  // Check product
  if (config.moduleLoaded && config.module !== config.moduleLoaded) {
    mismatches.push({
      key: 'MODULE',
      current: config.module,
      locked: config.moduleLoaded
    });
  }

  // Check custom fields — no truthy guard: blank→non-blank must be caught
  if (config.customField1 !== config.customField1Loaded) {
    mismatches.push({
      key: 'CUSTOM_FIELD_1',
      current: config.customField1 || '(blank)',
      locked: config.customField1Loaded || '(blank)'
    });
  }
  if (config.customField2 !== config.customField2Loaded) {
    mismatches.push({
      key: 'CUSTOM_FIELD_2',
      current: config.customField2 || '(blank)',
      locked: config.customField2Loaded || '(blank)'
    });
  }
  if (config.customField3 !== config.customField3Loaded) {
    mismatches.push({
      key: 'CUSTOM_FIELD_3',
      current: config.customField3 || '(blank)',
      locked: config.customField3Loaded || '(blank)'
    });
  }

  // Location custom fields — same blank-sensitive comparison as the ticket ones
  const locNames = config.locationCustomFields || [];
  const locLocked = config.locationCustomFieldsLoaded || [];
  for (let i = 0; i < LOCATION_CUSTOM_FIELD_COUNT; i++) {
    if ((locNames[i] || '') !== (locLocked[i] || '')) {
      mismatches.push({
        key: 'LOCATION_CUSTOM_FIELD_' + (i + 1),
        current: locNames[i] || '(blank)',
        locked: locLocked[i] || '(blank)'
      });
    }
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
  setConfigValue('MODULE_LOADED', config.module);
  setConfigValue('CUSTOM_FIELD_1_LOADED', config.customField1);
  setConfigValue('CUSTOM_FIELD_2_LOADED', config.customField2);
  setConfigValue('CUSTOM_FIELD_3_LOADED', config.customField3);
  for (let i = 0; i < LOCATION_CUSTOM_FIELD_COUNT; i++) {
    setConfigValue('LOCATION_CUSTOM_FIELD_' + (i + 1) + '_LOADED', config.locationCustomFields[i] || '');
  }
  logOperation('Config', 'LOCKED',
    `Configuration locked: SCHOOL_YEAR=${config.schoolYear}, PAGE_SIZE=${config.pageSize}, BATCH_SIZE=${config.ticketBatchSize}, MODULE=${config.module}`);
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
  setConfigValue('MODULE_LOADED', '');
  // Clear custom field locks AND resolved IDs (forces re-resolution on next load)
  setConfigValue('CUSTOM_FIELD_1_LOADED', '');
  setConfigValue('CUSTOM_FIELD_2_LOADED', '');
  setConfigValue('CUSTOM_FIELD_3_LOADED', '');
  setConfigValue('CUSTOM_FIELD_1_ID', '');
  setConfigValue('CUSTOM_FIELD_2_ID', '');
  setConfigValue('CUSTOM_FIELD_3_ID', '');
  for (let i = 0; i < LOCATION_CUSTOM_FIELD_COUNT; i++) {
    setConfigValue('LOCATION_CUSTOM_FIELD_' + (i + 1) + '_LOADED', '');
    setConfigValue('LOCATION_CUSTOM_FIELD_' + (i + 1) + '_ID', '');
  }
  setConfigValue('LOCATION_CF_TYPE_CACHE', '');
  logOperation('Config', 'UNLOCKED', 'Configuration unlocked (SCHOOL_YEAR, PAGE_SIZE, BATCH_SIZE, MODULE, CUSTOM_FIELDS, LOCATION_CUSTOM_FIELDS)');
}

/**
 * @deprecated Use unlockConfig instead
 */
function unlockSchoolYear() {
  unlockConfig();
}

// =============================================================================
// CUSTOM FIELD RESOLUTION
// =============================================================================

/**
 * Resolve custom field names to CustomFieldTypeId UUIDs via the discovery API.
 * Only calls the API if at least one field needs resolution (name set, ID empty or NOT_FOUND).
 * Writes resolved IDs back to Config sheet via setConfigValue().
 *
 * @param {Object} config - Config object from getConfig()
 * @returns {Object} - { resolved: number, notFound: string[] }
 */
function resolveCustomFieldIds(config) {
  const fields = [
    { name: config.customField1, id: config.customField1Id, key: 'CUSTOM_FIELD_1_ID', label: '1' },
    { name: config.customField2, id: config.customField2Id, key: 'CUSTOM_FIELD_2_ID', label: '2' },
    { name: config.customField3, id: config.customField3Id, key: 'CUSTOM_FIELD_3_ID', label: '3' }
  ];

  // Check which fields need resolution
  const pending = fields.filter(f => f.name && (!f.id || f.id === 'NOT_FOUND'));
  if (pending.length === 0) {
    return { resolved: 0, notFound: [] };
  }

  // A pasted CustomFieldTypeId is used as-is — no lookup, no API call
  let resolvedFromGuid = 0;
  const needsResolution = [];
  for (const field of pending) {
    if (looksLikeCustomFieldGuid_(field.name)) {
      const guid = String(field.name).trim();
      setConfigValue(field.key, guid);
      logOperation('CustomFields', 'RESOLVED', `Custom Field ${field.label} set directly from CustomFieldTypeId ${guid}`);
      resolvedFromGuid++;
    } else {
      needsResolution.push(field);
    }
  }
  if (needsResolution.length === 0) {
    return { resolved: resolvedFromGuid, notFound: [] };
  }

  // Call discovery endpoint once, only for entries given as names
  logOperation('CustomFields', 'RESOLVING', `Resolving ${needsResolution.length} custom field name(s)`);
  let definitions;
  try {
    definitions = getTicketCustomFieldDefinitions();
  } catch (e) {
    logOperation('CustomFields', 'ERROR', 'Failed to fetch custom field definitions: ' + e.message);
    return { resolved: resolvedFromGuid, notFound: needsResolution.map(f => f.name) };
  }

  // Build lowercase name → CustomFieldTypeId map
  const nameToId = new Map();
  for (const def of definitions) {
    const typeName = def.CustomFieldType && def.CustomFieldType.Name;
    if (typeName && def.CustomFieldTypeId) {
      nameToId.set(typeName.trim().toLowerCase(), def.CustomFieldTypeId);
    }
  }

  let resolved = resolvedFromGuid;
  const notFound = [];

  for (const field of needsResolution) {
    const lookupKey = field.name.trim().toLowerCase();
    const uuid = nameToId.get(lookupKey);
    if (uuid) {
      setConfigValue(field.key, uuid);
      logOperation('CustomFields', 'RESOLVED', `Custom Field ${field.label} "${field.name}" → ${uuid}`);
      resolved++;
    } else {
      setConfigValue(field.key, 'NOT_FOUND');
      logOperation('CustomFields', 'WARNING', `Custom Field ${field.label} "${field.name}" not found in district`);
      notFound.push(field.name);
    }
  }

  return { resolved, notFound };
}

/**
 * Does this config value look like a CustomFieldTypeId rather than a field name?
 *
 * Custom field slots accept either. A pasted GUID is used directly, which is the
 * unambiguous option: field names are not unique (a district can have two
 * different field types sharing a display name), and the discovery endpoints
 * return one row per field/filter-set mapping, so name lookup is both ambiguous
 * and unnecessary when the district already has the id in front of them.
 *
 * @param {string} value - Config cell value
 * @returns {boolean}
 */
function looksLikeCustomFieldGuid_(value) {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(String(value || '').trim());
}

// =============================================================================
// LOCATION CUSTOM FIELD EDITOR TYPE CACHE
// =============================================================================
//
// Building the location value index needs each field's EditorType, to know
// whether its raw value needs resolving to a display name. That type used to
// cost a /custom-fields/for/location call on every script execution, purely to
// re-learn something that almost never changes. It is cached in one Config cell
// instead, so a steady-state run makes no definitions call at all.

/**
 * Parse the LOCATION_CF_TYPE_CACHE cell.
 * Shape: {"cached":"<ISO date>","types":{"<CustomFieldTypeId>":<editorType>}}
 * Any malformed or empty value is treated as "no cache" rather than an error —
 * a bad cell should cost one API call, never a failed load.
 *
 * @param {string} raw - Cell contents
 * @returns {Object|null} - { cached: string, types: Object } or null
 */
function parseLocationCfTypeCache(raw) {
  if (!raw) return null;
  try {
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== 'object' || !parsed.types) return null;
    return { cached: parsed.cached || '', types: parsed.types };
  } catch (e) {
    return null;
  }
}

/**
 * Is the cached editor type usable for every configured location field?
 *
 * Requires an entry for each configured, resolved field AND a cache timestamp
 * inside the TTL. Keying on CustomFieldTypeId means pointing a slot at a
 * different field misses the cache on its own, without invalidating the rest.
 *
 * @param {Object} config - Config object from getConfig()
 * @returns {boolean}
 */
function isLocationCfTypeCacheFresh(config) {
  const cache = config.locationCfTypeCache;
  if (!cache || !cache.cached) return false;

  const cachedAt = new Date(cache.cached);
  if (isNaN(cachedAt.getTime())) return false;
  const ageDays = (Date.now() - cachedAt.getTime()) / (1000 * 60 * 60 * 24);
  if (ageDays > LOCATION_CF_TYPE_CACHE_DAYS || ageDays < 0) return false;

  const ids = config.locationCustomFieldIds || [];
  for (let i = 0; i < LOCATION_CUSTOM_FIELD_COUNT; i++) {
    const id = ids[i];
    if (!id || id === 'NOT_FOUND') continue;
    if (cache.types[id] === undefined || cache.types[id] === null) return false;
  }
  return true;
}

/**
 * Write the editor type cache from a set of location custom field definitions.
 * Called wherever definitions were already fetched for another reason, so the
 * cache refresh itself never costs an extra API call.
 *
 * Caches every definition the district has, not just the configured ones, so
 * selecting a different field later is already a cache hit.
 *
 * @param {Array} definitions - CustomFieldDetail objects from getLocationCustomFieldDefinitions()
 */
function writeLocationCfTypeCache(definitions) {
  if (!definitions || !definitions.length) return;

  const types = {};
  for (const def of definitions) {
    const id = def.CustomFieldTypeId;
    if (!id) continue;
    const ct = def.CustomFieldType || {};
    const editorType = def.EditorTypeId || ct.EditorType || 0;
    types[id] = editorType;
  }
  if (!Object.keys(types).length) return;

  const payload = JSON.stringify({ cached: new Date().toISOString(), types: types });
  setConfigValue('LOCATION_CF_TYPE_CACHE', payload);
  logOperation('LocationCustomFields', 'CACHED',
    `Cached editor types for ${Object.keys(types).length} location custom field(s), valid ${LOCATION_CF_TYPE_CACHE_DAYS} days`);
}

/**
 * Resolve LOCATION_CUSTOM_FIELD_* names to CustomFieldTypeId UUIDs.
 * Mirrors resolveCustomFieldIds but queries the location custom field
 * definitions endpoint, so the two namespaces resolve independently and a
 * name present on both entities can never cross over.
 *
 * @param {Object} config - Config object from getConfig()
 * @returns {Object} - { resolved: number, notFound: string[] }
 */
function resolveLocationCustomFieldIds(config) {
  const names = config.locationCustomFields || [];
  const ids = config.locationCustomFieldIds || [];

  const pending = [];
  for (let i = 0; i < LOCATION_CUSTOM_FIELD_COUNT; i++) {
    const name = names[i] || '';
    const id = ids[i] || '';
    if (name && (!id || id === 'NOT_FOUND')) {
      pending.push({ name: name, key: 'LOCATION_CUSTOM_FIELD_' + (i + 1) + '_ID', label: String(i + 1) });
    }
  }
  if (pending.length === 0) {
    return { resolved: 0, notFound: [] };
  }

  // A pasted CustomFieldTypeId is used as-is — no lookup, no API call
  let resolvedFromGuid = 0;
  const needsResolution = [];
  for (const field of pending) {
    if (looksLikeCustomFieldGuid_(field.name)) {
      const guid = String(field.name).trim();
      setConfigValue(field.key, guid);
      logOperation('LocationCustomFields', 'RESOLVED', `Location Custom Field ${field.label} set directly from CustomFieldTypeId ${guid}`);
      resolvedFromGuid++;
    } else {
      needsResolution.push(field);
    }
  }
  if (needsResolution.length === 0) {
    return { resolved: resolvedFromGuid, notFound: [] };
  }

  logOperation('LocationCustomFields', 'RESOLVING', `Resolving ${needsResolution.length} location custom field name(s)`);
  let definitions;
  try {
    definitions = getLocationCustomFieldDefinitions();
  } catch (e) {
    logOperation('LocationCustomFields', 'ERROR', 'Failed to fetch location custom field definitions: ' + e.message);
    return { resolved: resolvedFromGuid, notFound: needsResolution.map(f => f.name) };
  }

  const nameToId = new Map();
  for (const def of definitions) {
    const typeName = def.CustomFieldType && def.CustomFieldType.Name;
    if (typeName && def.CustomFieldTypeId) {
      nameToId.set(typeName.trim().toLowerCase(), def.CustomFieldTypeId);
    }
  }

  // Free cache fill: definitions are already in hand
  writeLocationCfTypeCache(definitions);

  let resolved = resolvedFromGuid;
  const notFound = [];

  for (const field of needsResolution) {
    const uuid = nameToId.get(field.name.trim().toLowerCase());
    if (uuid) {
      setConfigValue(field.key, uuid);
      logOperation('LocationCustomFields', 'RESOLVED', `Location Custom Field ${field.label} "${field.name}" → ${uuid}`);
      resolved++;
    } else {
      setConfigValue(field.key, 'NOT_FOUND');
      logOperation('LocationCustomFields', 'WARNING', `Location Custom Field ${field.label} "${field.name}" not found in district`);
      notFound.push(field.name);
    }
  }

  return { resolved, notFound };
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
// VERSION CHECK - Remote update detection via GitHub
// =============================================================================

/**
 * Check for script updates from GitHub.
 * Fetches remote version.json and updates Config sheet display.
 * Fails silently on any error — version check must never break anything.
 */
function checkForUpdates() {
  try {
    const REMOTE_VERSION_URL =
      'https://raw.githubusercontent.com/scopousiiq/iiq-tickets-to-sheets/main/version.json';

    const response = UrlFetchApp.fetch(REMOTE_VERSION_URL, {
      muteHttpExceptions: true,
      followRedirects: true
    });

    if (response.getResponseCode() !== 200) {
      logOperation('VersionCheck', 'WARN',
        'Could not reach GitHub (HTTP ' + response.getResponseCode() + ')');
      return;
    }

    const remote = JSON.parse(response.getContentText());
    const remoteVersion = remote.version;

    if (!remoteVersion) {
      logOperation('VersionCheck', 'WARN', 'Remote version.json missing version field');
      return;
    }

    const updateAvailable = isNewerVersion(remoteVersion, SCRIPT_VERSION);

    // Update Config sheet rows
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Config');
    if (!sheet) return;

    setConfigValue('SCRIPT_VERSION', SCRIPT_VERSION);
    setConfigValue('VERSION_CHECK_DATE', new Date().toISOString().split('T')[0]);

    // Set LATEST_VERSION with status text and background color
    if (updateAvailable) {
      setConfigValue('LATEST_VERSION', remoteVersion + '  \u2190 update available');
    } else {
      setConfigValue('LATEST_VERSION', remoteVersion + '  (up to date)');
    }

    // Apply background color to the LATEST_VERSION value cell
    const data = sheet.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === 'LATEST_VERSION') {
        const cell = sheet.getRange(i + 1, 2);
        if (updateAvailable) {
          cell.setBackground('#fff2cc'); // light yellow
        } else {
          cell.setBackground('#d9ead3'); // light green
        }
        break;
      }
    }

    // Log result
    if (updateAvailable) {
      logOperation('VersionCheck', 'UPDATE_AVAILABLE',
        'v' + remoteVersion + ' available (current: v' + SCRIPT_VERSION + '). ' +
        (remote.releaseUrl || '') + ' \u2014 ' + (remote.message || ''));
    } else {
      logOperation('VersionCheck', 'CURRENT', 'v' + SCRIPT_VERSION + ' is up to date');
    }

  } catch (e) {
    logOperation('VersionCheck', 'ERROR', 'Version check failed: ' + e.message);
  }
}

/**
 * Compare two semver strings (e.g., "1.2.0" vs "1.3.0").
 * @param {string} remoteVer - The remote version string
 * @param {string} localVer - The local version string
 * @return {boolean} True if remoteVer is newer than localVer
 */
function isNewerVersion(remoteVer, localVer) {
  const remote = remoteVer.split('.').map(Number);
  const local = localVer.split('.').map(Number);
  for (let i = 0; i < 3; i++) {
    const r = remote[i] || 0;
    const l = local[i] || 0;
    if (r > l) return true;
    if (r < l) return false;
  }
  return false;
}

/**
 * Check if version check is stale (>24 hours since last check).
 * Used as a guard to avoid checking on every trigger invocation.
 * @return {boolean} True if a version check should be performed
 */
function isVersionCheckStale() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
    if (!sheet) return true;
    const data = sheet.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === 'VERSION_CHECK_DATE') {
        const val = data[i][1];
        if (!val) return true;
        const lastCheck = new Date(val);
        if (isNaN(lastCheck.getTime())) return true;
        const hoursSince = (Date.now() - lastCheck.getTime()) / (1000 * 60 * 60);
        return hoursSince > 24;
      }
    }
    return true; // No check date row found
  } catch (e) {
    return false; // On error, don't trigger a check
  }
}

/**
 * Menu wrapper for Check for Updates.
 * Shows a toast (non-blocking) to confirm the action completed.
 */
function menuCheckForUpdates() {
  checkForUpdates();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Version check complete. See Config sheet for results.',
    'Version Check', 5);
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
