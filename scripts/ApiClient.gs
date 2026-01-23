/**
 * IIQ API Client
 * Handles HTTP requests to the IncidentIQ API
 * Includes retry logic with exponential backoff for rate limiting
 */

const MAX_RETRIES = 3;
const BACKOFF_BASE_MS = 2000;

/**
 * Get throttle delay from config (with fallback)
 */
function getThrottleMs() {
  try {
    const config = getConfig();
    return config.throttleMs || 1000;
  } catch (e) {
    return 1000; // Default fallback
  }
}

function makeApiRequest(endpoint, method, payload, retryCount) {
  const config = getConfig();
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

  if (payload) {
    options.payload = JSON.stringify(payload);
  }

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode >= 200 && responseCode < 300) {
      logOperation('API Request', 'SUCCESS', endpoint);
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
  const config = getConfig();
  const size = pageSize || config.pageSize;

  // Build endpoint with $p and $s query parameters
  let endpoint = `/v1.0/tickets?$p=${page || 0}&$s=${size}`;

  // Add sorting via $o query parameter (format: "FieldName direction")
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
