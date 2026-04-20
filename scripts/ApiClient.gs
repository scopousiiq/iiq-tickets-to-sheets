/**
 * iiQ API Client
 * Handles HTTP requests to the Incident IQ API
 * Includes retry logic with exponential backoff for rate limiting
 */

const MAX_RETRIES = 3;
const BACKOFF_BASE_MS = 2000;

/**
 * Cached config for batch operations.
 * Set via setApiConfig() before a batch, cleared via clearApiConfig() after.
 * When set, makeApiRequest/searchTickets skip reading the Config sheet on every call.
 */
let apiConfig_ = null;

/**
 * Cache the config object for the duration of a batch operation.
 * Call this once before a loop of API requests to avoid repeated Config sheet reads.
 *
 * @param {Object} config - Config object from getConfig()
 */
function setApiConfig(config) {
  apiConfig_ = config;
}

/**
 * Clear the cached API config. Call after batch operations complete.
 */
function clearApiConfig() {
  apiConfig_ = null;
}

/**
 * Get the current API config — uses cache if available, otherwise reads from sheet.
 */
function getApiConfig_() {
  return apiConfig_ || getConfig();
}

/**
 * Get throttle delay from config (with fallback)
 */
function getThrottleMs() {
  try {
    const config = getApiConfig_();
    return config.throttleMs || 1000;
  } catch (e) {
    return 1000; // Default fallback
  }
}

function makeApiRequest(endpoint, method, payload, retryCount) {
  const config = getApiConfig_();
  retryCount = retryCount || 0;

  if (!config.baseUrl || !config.bearerToken) {
    throw new Error('API configuration missing. Check Config sheet.');
  }

  const url = config.baseUrl + endpoint;

  const options = {
    method: method || 'GET',
    headers: {
      'Authorization': 'Bearer ' + config.bearerToken,
      'Content-Type': 'application/json',
      'Accept': 'application/json'
    },
    muteHttpExceptions: true
  };

  if (config.siteId) {
    options.headers['SiteId'] = config.siteId;
  }

  if (config.moduleProductId) {
    options.headers['ProductId'] = config.moduleProductId;
  }

  if (payload) {
    options.payload = JSON.stringify(payload);
  }

  try {
    const startTime = Date.now();
    const response = UrlFetchApp.fetch(url, options);
    const elapsedMs = Date.now() - startTime;
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode >= 200 && responseCode < 300) {
      logOperation('API Request', 'SUCCESS', `${endpoint} (${elapsedMs}ms)`);
      // Throttle after successful request to avoid hitting limits
      Utilities.sleep(Math.floor(getThrottleMs() * 0.5));
      return JSON.parse(responseText);
    } else if (responseCode === 429 && retryCount < MAX_RETRIES) {
      // Rate limited - exponential backoff and retry
      const backoffMs = BACKOFF_BASE_MS * Math.pow(2, retryCount);
      logOperation('API Request', 'RATE_LIMITED', `${endpoint} - Retrying in ${backoffMs}ms (attempt ${retryCount + 1}/${MAX_RETRIES})`);
      Utilities.sleep(backoffMs);
      return makeApiRequest(endpoint, method, payload, retryCount + 1);
    } else if (responseCode === 503 && retryCount < MAX_RETRIES) {
      // Service unavailable - retry with backoff
      const backoffMs = BACKOFF_BASE_MS * Math.pow(2, retryCount);
      logOperation('API Request', 'SERVICE_UNAVAILABLE', `${endpoint} - Retrying in ${backoffMs}ms (attempt ${retryCount + 1}/${MAX_RETRIES})`);
      Utilities.sleep(backoffMs);
      return makeApiRequest(endpoint, method, payload, retryCount + 1);
    } else {
      logOperation('API Request', 'ERROR', `${endpoint} - HTTP ${responseCode}: ${responseText.substring(0, 200)}`);
      throw new Error(`API Error ${responseCode}: ${responseText.substring(0, 200)}`);
    }
  } catch (e) {
    if (e.message && e.message.includes('API Error')) {
      throw e; // Re-throw API errors
    }
    // Network errors - retry with backoff
    if (retryCount < MAX_RETRIES) {
      const backoffMs = BACKOFF_BASE_MS * Math.pow(2, retryCount);
      logOperation('API Request', 'NETWORK_ERROR', `${endpoint} - ${e.message} - Retrying in ${backoffMs}ms`);
      Utilities.sleep(backoffMs);
      return makeApiRequest(endpoint, method, payload, retryCount + 1);
    }
    logOperation('API Request', 'EXCEPTION', `${endpoint} - ${e.message}`);
    throw e;
  }
}

/**
 * Search tickets with filters
 * @param {Array} filters - Array of filter objects [{Facet, Value/Id}]
 * @param {number} page - Page index (0-based)
 * @param {number} pageSize - Optional page size override (defaults to config.pageSize)
 * @param {Object} sortOptions - Optional sort options {field, direction}
 * @returns {Object} - API response with Items and Paging
 *
 * Query parameters:
 *   $p - Zero-based page index
 *   $s - Page size (number of records)
 *   $o - Sort expression (e.g., "TicketCreatedDate asc" or "TicketPriority desc")
 *   $d - Legacy sort direction override ("asc", "desc", "0", "1")
 *
 * Available sort fields: TicketCreatedDate, TicketModifiedDate, TicketClosedDate,
 *   TicketDueDate, TicketId, TicketNumber, TicketSubject, TicketPriority, etc.
 */
function searchTickets(filters, page, pageSize, sortOptions) {
  const size = pageSize || getApiConfig_().pageSize;

  // Build endpoint with $p and $s query parameters
  let endpoint = `/v1.0/tickets?$p=${page || 0}&$s=${size}`;

  // Add sorting via $o query parameter (format: "FieldName direction")
  if (sortOptions && sortOptions.field) {
    const sortExpr = `${sortOptions.field} ${sortOptions.direction || 'asc'}`;
    endpoint += `&$o=${encodeURIComponent(sortExpr)}`;
  }

  const payload = {
    Filters: filters || [],
    Includes: ['Assets']
  };

  return makeApiRequest(endpoint, 'POST', payload);
}

/**
 * Get ticket count for filters (without fetching all data)
 * @param {Array} filters - Array of filter objects
 * @returns {number} - Total count of matching tickets
 */
function getTicketCount(filters) {
  const endpoint = '/v1.0/tickets?$p=0&$s=1'; // Minimal page size

  const payload = {
    Filters: filters || []
  };

  const response = makeApiRequest(endpoint, 'POST', payload);
  return response.Paging ? response.Paging.TotalRows : 0;
}

/**
 * Get all teams
 * Note: Despite /teams/all being documented as "without pagination",
 * the API may apply a default limit. We add $s=1000 to ensure all teams are returned.
 * @returns {Array} - Array of team objects
 */
function getAllTeams() {
  const endpoint = '/v1.0/teams/all?$s=1000';
  const response = makeApiRequest(endpoint, 'GET', null);
  return response.Items || response || [];
}

/**
 * Get all custom field definitions for tickets
 * POST /v1.0/custom-fields/for/ticket with empty body returns all ticket custom fields.
 * Each item has CustomFieldTypeId (UUID) and CustomFieldType.Name (display name).
 *
 * @returns {Array} - Array of CustomFieldDetail objects
 */
function getTicketCustomFieldDefinitions() {
  const endpoint = '/v1.0/custom-fields/for/ticket';
  const response = makeApiRequest(endpoint, 'POST', {});
  return response.Items || [];
}

/**
 * Get all locations for the district.
 * Paginates through locations endpoint. Prefers site-scoped endpoint (returns
 * all district locations) over user-scoped endpoint (returns only locations
 * the bearer token's user has access to — may be capped to a subset).
 * Used to build a lookup map for IiqLocation-type custom fields.
 *
 * @returns {Array} - Array of Location objects
 */
function getAllLocations() {
  const config = getApiConfig_();
  const base = config.siteId
    ? `/v1.0/locations/all/${config.siteId}`
    : '/v1.0/locations/all';
  const all = [];
  const pageSize = 500;
  let page = 0;
  while (true) {
    const endpoint = `${base}?$p=${page}&$s=${pageSize}`;
    const response = makeApiRequest(endpoint, 'GET', null);
    const items = response.Items || [];
    all.push.apply(all, items);
    const totalRows = response.Paging ? (response.Paging.TotalRows || response.Paging.Total || 0) : 0;
    if (items.length === 0 || all.length >= totalRows || items.length < pageSize) break;
    page++;
    if (page > 100) break; // safety
  }
  return all;
}
