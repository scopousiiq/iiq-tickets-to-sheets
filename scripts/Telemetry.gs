/**
 * iiQ Sheets Telemetry — Client
 *
 * Sends one ping per successful trigger refresh to the iiQ-owned Telemetry
 * Master, and enforces the policy that automated polling requires telemetry
 * opt-in.
 *
 * Three functions are called from project code:
 *
 *   reportTelemetry()                    — called at the TAIL of a trigger-
 *                                          fired function, after the refresh
 *                                          succeeds. Sends one ping.
 *   enforceTelemetryGate()               — called at the HEAD of every
 *                                          trigger-fired function. Returns
 *                                          false (and uninstalls all CLOCK
 *                                          triggers) if TELEMETRY_ENABLED is
 *                                          not TRUE.
 *   assertTelemetryEnabledForTriggers()  — called by any function about to
 *                                          install a time-based trigger.
 *                                          Throws a user-readable error if
 *                                          telemetry is off.
 *
 * Deployment model:
 *   - TELEMETRY_URL lives in Config.gs as a maintainer-managed constant
 *     (hardcoded, pushed via clasp). Districts don't see or edit it.
 *   - TELEMETRY_ENABLED lives in the project's Config sheet. Setup Spreadsheet
 *     stamps it TRUE for new installs. Districts can set it FALSE to opt out;
 *     doing so also disables automated polling (per policy).
 *
 * reportTelemetry() gating (silent no-op if any is true):
 *   - TELEMETRY_URL constant missing or empty
 *   - Config TELEMETRY_ENABLED != TRUE
 *   - No time-based trigger installed (install is dormant)
 *   - API_BASE_URL missing or can't be parsed to a hostname
 *
 * reportTelemetry() errors are caught and logged — telemetry must never break
 * the parent trigger.
 */

// ---- Per-project constants ----

const TELEMETRY_PROJECT = 'iiq-tickets-to-sheets';
const TELEMETRY_PRIMARY_SHEET = 'TicketData';
const TELEMETRY_CANONICAL_ANALYTICS = [
  // Default analytics (created by Setup Spreadsheet)
  'MonthlyVolume',
  'BacklogAging',
  'ResolutionAging',
  'TeamWorkload',
  'SLACompliance',
  'PerformanceTrends',
  'AtRiskResponse',
  'AtRiskResolution',
  // Optional analytics (added via iiQ Data > Add Analytics Sheet)
  'SeasonalComparison',
  'TemporalPatterns',
  'MonthlyVolumeByFA',
  'BacklogAgingByFA',
  'BacklogAgingByTeam',
  'BacklogAgingByLocationType',
  'BacklogAgingByPriority',
  'StaleTickets',
  'ReopenRate',
  'FirstContactResolution',
  'ResponseDistribution',
  'ResponseTrends',
  'QueueTimeAnalysis',
  'QueueTimeByTeam',
  'QueueTimeTrend',
  'TechnicianPerformance',
  'FunctionalAreaSummary',
  'LocationBreakdown',
  'LocationTypeComparison',
  'IssueCategoryVolume',
  'PriorityAnalysis',
  'FrequentRequesters',
  'DeviceReliability',
  'DevicesByRole',
  'FrequentFlyers'
];

// ---- Wire protocol constant: do NOT edit without coordinating a server change ----

const TELEMETRY_SCHEMA_VERSION = 1;


/**
 * Send one telemetry ping. Call at the tail of a trigger-fired function,
 * after the refresh completes successfully. Best-effort: never throws.
 */
function reportTelemetry() {
  try {
    const url = (typeof TELEMETRY_URL === 'string' ? TELEMETRY_URL : '').trim();
    if (!url) return;

    const cfg = telemetryReadConfig_();
    if (String(cfg.TELEMETRY_ENABLED || '').toUpperCase() !== 'TRUE') return;

    const hasTimeTrigger = ScriptApp.getProjectTriggers().some(function (t) {
      return t.getEventType() === ScriptApp.EventType.CLOCK;
    });
    if (!hasTimeTrigger) return;

    const instanceUrl = telemetryExtractHostname_(cfg.API_BASE_URL);
    if (!instanceUrl) return;

    const ss = SpreadsheetApp.getActive();
    const primary = ss.getSheetByName(TELEMETRY_PRIMARY_SHEET);
    const rowCount = primary ? Math.max(0, primary.getLastRow() - 1) : 0;

    const presentSheets = ss.getSheets().map(function (s) { return s.getName(); });
    const analyticsSheets = TELEMETRY_CANONICAL_ANALYTICS.filter(function (name) {
      return presentSheets.indexOf(name) !== -1;
    });

    const payload = {
      schemaVersion: TELEMETRY_SCHEMA_VERSION,
      installId: telemetryGetOrCreateInstallId_(),
      project: TELEMETRY_PROJECT,
      version: typeof SCRIPT_VERSION === 'string' ? SCRIPT_VERSION : '',
      instanceUrl: instanceUrl,
      installedAt: telemetryGetOrStampInstalledAt_(),
      scriptTimeZone: Session.getScriptTimeZone(),
      sentAt: new Date().toISOString(),
      rowCount: rowCount,
      primarySheet: TELEMETRY_PRIMARY_SHEET,
      analyticsSheets: analyticsSheets
    };

    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      followRedirects: true
    });
  } catch (err) {
    console.warn('telemetry failed: ' + (err && err.message || err));
  }
}


/**
 * Runtime policy gate. Call as the FIRST line of every trigger-fired function.
 *
 *     function triggerDataContinue() {
 *       if (!enforceTelemetryGate()) return;
 *       // ... refresh work ...
 *       reportTelemetry();
 *     }
 *
 * Returns true if TELEMETRY_ENABLED is TRUE in the Config sheet.
 *
 * Returns false and DELETES every time-based (CLOCK) trigger in the project
 * if TELEMETRY_ENABLED is anything else. Districts that turn telemetry off
 * forfeit automated polling until they re-enable TELEMETRY_ENABLED and re-run
 * Setup Automated Triggers.
 *
 * Edit / open / form-submit triggers are not touched.
 *
 * Fail-closed on error: if the Config sheet can't be read, returns false
 * without uninstalling — the next fire will retry.
 */
function enforceTelemetryGate() {
  try {
    const cfg = telemetryReadConfig_();
    const enabled = String(cfg.TELEMETRY_ENABLED || '').toUpperCase() === 'TRUE';
    if (enabled) return true;

    let removed = 0;
    ScriptApp.getProjectTriggers().forEach(function (t) {
      if (t.getEventType() === ScriptApp.EventType.CLOCK) {
        ScriptApp.deleteTrigger(t);
        removed++;
      }
    });
    console.warn(
      'Telemetry disabled (TELEMETRY_ENABLED != TRUE). Uninstalled ' +
      removed + ' time-based trigger(s). Automated polling is only available ' +
      'to installs with telemetry enabled. Set TELEMETRY_ENABLED=TRUE in ' +
      'Config and re-run Setup Automated Triggers to reinstall.'
    );
    try {
      logOperation('Telemetry', 'GATE_OFF',
        'TELEMETRY_ENABLED!=TRUE; uninstalled ' + removed + ' time-based trigger(s)');
    } catch (_) {}
    return false;
  } catch (err) {
    console.error('telemetry gate error: ' + (err && err.message || err));
    return false;
  }
}


/**
 * Install-time policy gate. Call at the top of any function about to create
 * a time-based trigger, BEFORE the ScriptApp.newTrigger() chain.
 *
 *     function setupAutomatedTriggers() {
 *       assertTelemetryEnabledForTriggers();
 *       ScriptApp.newTrigger('triggerDataContinue')
 *         .timeBased().everyMinutes(10).create();
 *     }
 *
 * Throws a user-readable error if telemetry is off.
 */
function assertTelemetryEnabledForTriggers() {
  const cfg = telemetryReadConfig_();
  const enabled = String(cfg.TELEMETRY_ENABLED || '').toUpperCase() === 'TRUE';
  if (!enabled) {
    throw new Error(
      'Cannot install automated triggers: TELEMETRY_ENABLED must be TRUE ' +
      'in the Config sheet. Automated API access in iiQ sheet projects ' +
      'requires telemetry opt-in.'
    );
  }
}


// Reads Config sheet (Key in col A, Value in col B) into a plain object.
// Self-contained so Telemetry.gs can drop into any project.
function telemetryReadConfig_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Config');
  if (!sheet) return {};
  const last = sheet.getLastRow();
  if (last < 1) return {};
  const values = sheet.getRange(1, 1, last, 2).getValues();
  const out = {};
  for (let i = 0; i < values.length; i++) {
    const key = String(values[i][0] || '').trim();
    if (key) out[key] = values[i][1];
  }
  return out;
}

function telemetryGetOrCreateInstallId_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('TELEMETRY_INSTALL_ID');
  if (!id) {
    id = Utilities.getUuid();
    props.setProperty('TELEMETRY_INSTALL_ID', id);
  }
  return id;
}

function telemetryGetOrStampInstalledAt_() {
  const props = PropertiesService.getScriptProperties();
  let ts = props.getProperty('TELEMETRY_INSTALLED_AT');
  if (!ts) {
    ts = new Date().toISOString();
    props.setProperty('TELEMETRY_INSTALLED_AT', ts);
  }
  return ts;
}

// Strips scheme, path, query, and fragment to yield a bare hostname.
// "https://demo.incidentiq.com/api/v1.0" -> "demo.incidentiq.com"
function telemetryExtractHostname_(urlLike) {
  if (!urlLike) return '';
  let s = String(urlLike).trim().replace(/^https?:\/\//i, '');
  s = s.split('/')[0].split('?')[0].split('#')[0];
  return s.toLowerCase();
}
