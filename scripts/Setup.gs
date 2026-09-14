/**
 * Setup Script - Initial Spreadsheet Configuration
 *
 * Creates required data sheets and default analytics sheets.
 * Running setup provides a CLEAN SLATE - existing sheets are overwritten.
 *
 * Data Sheets (always created):
 * - Instructions: Setup and usage guide
 * - Config: API settings and progress tracking
 * - TicketData: Main data (52 columns including SLA metrics, device/asset, assigned technician, ticket custom fields, and location custom fields)
 * - Teams: Team directory with Functional Area mapping
 * - DailySnapshot: Daily backlog metrics for trending
 * - Logs: Operation logs
 *
 * Default Analytics Sheets (created by setup):
 * - MonthlyVolume: Ticket volume by month
 * - BacklogAging: Current backlog distribution
 * - TeamWorkload: Tickets by team
 * - SLACompliance: Monthly SLA metrics
 * - PerformanceTrends: "Are we getting better?" trending
 * - AtRiskResponse: Tickets approaching Response SLA breach
 * - AtRiskResolution: Tickets approaching Resolution SLA breach
 *
 * Additional analytics sheets can be added via:
 * iiQ Data > Add Analytics Sheet menu
 *
 * Usage: Run setupSpreadsheet() from the iiQ Data > Setup menu
 */

/**
 * Delete a sheet if it exists
 * @param {Spreadsheet} ss - The spreadsheet
 * @param {string} sheetName - Name of the sheet to delete
 */
function deleteSheetIfExists(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    ss.deleteSheet(sheet);
  }
}

/**
 * Calculate the default school year based on today's date (July-June)
 * @return {string} School year string (e.g., "2025-2026")
 */
function getDefaultSchoolYear() {
  const now = new Date();
  const currentMonth = now.getMonth(); // 0-indexed
  const currentYear = now.getFullYear();
  const startYear = currentMonth >= 6 ? currentYear : currentYear - 1; // July = 6
  return `${startYear}-${startYear + 1}`;
}

/**
 * Main setup function - creates all sheets and configurations
 * WARNING: This overwrites existing sheets for a clean slate!
 */
function setupSpreadsheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const response = ui.alert(
    'Setup iiQ Data Spreadsheet',
    '⚠️ WARNING: This will DELETE and RECREATE the following sheets:\n\n' +
    'DATA SHEETS:\n' +
    '- Instructions (setup guide)\n' +
    '- Config (API settings) - CREDENTIALS WILL BE LOST!\n' +
    '- TicketData (52 columns) - ALL DATA WILL BE LOST!\n' +
    '- Teams (directory) - ALL DATA WILL BE LOST!\n' +
    '- DailySnapshot (trending) - ALL DATA WILL BE LOST!\n' +
    '- Logs (operations)\n\n' +
    'DEFAULT ANALYTICS:\n' +
    '- MonthlyVolume\n' +
    '- BacklogAging\n' +
    '- TeamWorkload\n' +
    '- SLACompliance\n' +
    '- PerformanceTrends\n' +
    '- AtRiskResponse\n' +
    '- AtRiskResolution\n\n' +
    'This provides a CLEAN SLATE for the spreadsheet.\n\n' +
    'Are you sure you want to continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  // Prompt for school year before creating any sheets
  const defaultYear = getDefaultSchoolYear();
  const yearResponse = ui.prompt(
    'School Year Configuration',
    'Enter the school year for this spreadsheet.\n\n' +
    'Format: YYYY-YYYY (e.g., "2023-2024" for a historical year)\n\n' +
    'Leave blank to use the current school year (' + defaultYear + ').',
    ui.ButtonSet.OK_CANCEL
  );

  if (yearResponse.getSelectedButton() === ui.Button.CANCEL) return;

  const yearInput = yearResponse.getResponseText().trim();
  let schoolYear = defaultYear;

  if (yearInput !== '') {
    // Validate format: YYYY-YYYY with consecutive years
    const yearMatch = yearInput.match(/^(\d{4})-(\d{4})$/);
    if (!yearMatch) {
      ui.alert('Invalid School Year',
        'School year must be in YYYY-YYYY format (e.g., "2023-2024").\n\nSetup cancelled.',
        ui.ButtonSet.OK);
      return;
    }
    const startYear = parseInt(yearMatch[1], 10);
    const endYear = parseInt(yearMatch[2], 10);
    if (endYear !== startYear + 1) {
      ui.alert('Invalid School Year',
        'School year must use consecutive years (e.g., "2023-2024", not "' + yearInput + '").\n\nSetup cancelled.',
        ui.ButtonSet.OK);
      return;
    }
    schoolYear = yearInput;
  }

  const created = [];

  // Create data sheets (required) - these delete existing sheets first
  setupInstructionsSheet(ss); created.push('Instructions');
  setupConfigSheet(ss, schoolYear); created.push('Config');
  setupTicketDataSheet(ss); created.push('TicketData');
  setupTeamsSheet(ss); created.push('Teams');
  setupCustomFieldsSheet(ss); created.push('CustomFields');
  setupDailySnapshotSheet(ss); created.push('DailySnapshot');
  setupLogsSheet(ss); created.push('Logs');

  // Create default analytics sheets
  // Additional sheets can be added via iiQ Data > Add Analytics Sheet menu
  setupMonthlyVolumeSheet(ss); created.push('MonthlyVolume');
  setupBacklogAgingSheet(ss); created.push('BacklogAging');
  setupResolutionAgingSheet(ss); created.push('ResolutionAging');
  setupTeamWorkloadSheet(ss); created.push('TeamWorkload');
  setupSLAComplianceSheet(ss); created.push('SLACompliance');
  setupPerformanceTrendsSheet(ss); created.push('PerformanceTrends');
  setupAtRiskResponseSheet(ss); created.push('AtRiskResponse');
  setupAtRiskResolutionSheet(ss); created.push('AtRiskResolution');

  // Reorder sheets for better UX
  reorderSheets(ss);

  const message = [];
  message.push('Created ' + created.length + ' sheets: ' + created.join(', '));
  message.push('\nSchool year configured: ' + schoolYear);
  message.push('\nNext steps:');
  message.push('1. Fill in Config sheet with API credentials');
  message.push('2. Run "Verify Configuration" to check settings');
  message.push('3. Run "Refresh Teams" to load team directory');
  message.push('4. (Optional) Run "Refresh Custom Fields" to list your ticket and location custom field IDs');
  message.push('5. Run "Continue Loading" to start loading ticket data');

  ui.alert('Setup Complete', message.join('\n'), ui.ButtonSet.OK);
}

/**
 * Setup Instructions sheet with setup and usage guide
 * Deletes existing sheet if present for clean slate
 */
function setupInstructionsSheet(ss) {
  deleteSheetIfExists(ss, 'Instructions');
  const sheet = ss.insertSheet('Instructions');

  // Set column width for readability
  sheet.setColumnWidth(1, 800);

  // Instructions content
  const content = [
    ['iiQ TICKET & SLA DATA - SETUP AND USAGE GUIDE'],
    [''],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    ['OVERVIEW'],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    [''],
    ['This spreadsheet extracts ticket and SLA data from Incident IQ (iiQ) for analytics'],
    ['and dashboard consumption. Data is loaded via Google Apps Script and refreshed'],
    ['automatically via time-driven triggers.'],
    [''],
    ['SCHOOL YEAR MODEL: Each spreadsheet contains ONE school year\'s data.'],
    ['Example: A "2025-2026" spreadsheet covers July 1, 2025 through June 30, 2026.'],
    ['Create a new spreadsheet for each school year.'],
    [''],
    ['Data Flow: iiQ API → Google Apps Script → This Spreadsheet → Google Data Studio / Power BI'],
    [''],
    [''],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    ['INITIAL SETUP'],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    [''],
    ['1. SETUP SPREADSHEET'],
    ['   • Menu: iiQ Data > Setup > Setup Spreadsheet'],
    ['   • When prompted, enter the school year (e.g., "2023-2024" for historical data)'],
    ['   • Leave blank for the current school year (auto-detected)'],
    ['   • Creates all sheets and configures analytics formulas'],
    [''],
    ['2. CONFIGURE API CREDENTIALS (Config sheet)'],
    ['   • API_BASE_URL: Your iiQ instance (e.g., https://district.incidentiq.com)'],
    ['   • BEARER_TOKEN: JWT token from iiQ (Admin > Developer Tools)'],
    ['   • SITE_ID: Optional - only needed for multi-site instances'],
    ['   • MODULE: Select "Ticketing" (IT), "Facilities", or "HRSD" from the dropdown (default: Ticketing)'],
    [''],
    ['3. VERIFY CONFIGURATION'],
    ['   • Menu: iiQ Data > Setup > Verify Configuration'],
    ['   • Fix any issues reported before proceeding'],
    [''],
    ['3b. (OPTIONAL) ADD CUSTOM FIELD COLUMNS'],
    ['   • Menu: iiQ Data > Setup > Refresh Custom Fields'],
    ['   • Populates the CustomFields sheet with every custom field your district'],
    ['     defines, labeled by Entity: "Ticket" or "Location"'],
    ['   • Copy the CustomFieldTypeId (column C) of a field and paste it into Config:'],
    [''],
    ['     | Config key                | Slots | Values are read from          |'],
    ['     |---------------------------|-------|-------------------------------|'],
    ['     | CUSTOM_FIELD_1/2/3        |   3   | the ticket itself             |'],
    ['     | LOCATION_CUSTOM_FIELD_1-5 |   5   | the ticket\'s location        |'],
    [''],
    ['     Location fields hold values set on a school or building — site code,'],
    ['     region, building number — and are joined onto every ticket at that'],
    ['     location. Useful for pivoting a dashboard by your own site identifier.'],
    ['   • Check the Entity column: a Location field pasted into a Ticket slot will'],
    ['     never resolve, and vice versa.'],
    ['   • A field name also works, but ids are unambiguous — two different fields'],
    ['     can share the same display name.'],
    ['   • Set these BEFORE loading data: they are locked once loading starts.'],
    ['   • Re-run this if you later change a field definition in iiQ, so the'],
    ['     sheet picks up the new type right away.'],
    [''],
    ['4. LOAD TEAM DIRECTORY'],
    ['   • Menu: iiQ Data > Refresh Teams'],
    ['   • Then fill in FunctionalArea column for each team (optional)'],
    [''],
    ['5. START LOADING TICKET DATA'],
    ['   • Menu: iiQ Data > Ticket Data > Continue Loading (Initial)'],
    ['   • Script runs for ~5.5 minutes then pauses'],
    ['   • Run again to continue (or set up triggers for automation)'],
    ['   • SLA data and device info are included automatically — no separate step needed'],
    [''],
    [''],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    ['AUTOMATED TRIGGERS (Recommended)'],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    [''],
    ['EASY SETUP: Use iiQ Data > Setup > Setup Automated Triggers'],
    ['This creates all recommended triggers automatically.'],
    [''],
    ['MANUAL SETUP: Extensions > Apps Script > Triggers (clock icon)'],
    [''],
    ['| Function                  | Schedule         | Purpose                           |'],
    ['|---------------------------|------------------|-----------------------------------|'],
    ['| triggerDataContinue       | Every 10 min     | Continue any in-progress loading  |'],
    ['| triggerOpenTicketRefresh  | Every 2 hours    | Start open ticket + SLA refresh   |'],
    ['| triggerNewTickets         | Every 30 min     | Fetch newly created tickets       |'],
    ['| triggerDailySnapshot      | Daily 7:00 PM    | Capture backlog metrics           |'],
    ['| triggerWeeklyFullRefresh  | Weekly Sun 2 AM  | Full reload (catch deletions)     |'],
    [''],
    ['About triggerDataContinue (the "keep things moving" trigger):'],
    ['• If initial load is not complete → continues loading historical data'],
    ['• If open refresh is in progress → continues until complete'],
    ['• If both are done → does nothing (safe to leave enabled permanently)'],
    [''],
    ['Historical school years: Triggers are automatically removed once all data is loaded'],
    ['and all tickets are closed (data is static — no further updates needed).'],
    [''],
    ['Data Freshness with this schedule:'],
    ['• Open ticket SLA data: max 2 hours stale'],
    ['• New tickets: appear within 30 minutes'],
    ['• Status changes: captured within 2 hours'],
    ['• Deletions/corrections: captured weekly'],
    [''],
    [''],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    ['SHEETS REFERENCE'],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    [''],
    ['DATA SHEETS (populated by scripts):'],
    [''],
    ['• TicketData (52 columns)'],
    ['  Main ticket data including SLA metrics and device info. Columns include:'],
    ['  - Ticket info: ID, Number, Subject, School Year, Status, Priority'],
    ['  - Dates: Created, Started, Modified, Closed, Due'],
    ['  - Assignment: Team, Owner, Location (with LocationType)'],
    ['  - Issue: Category, Type'],
    ['  - Requester: ID, Name'],
    ['  - SLA: Thresholds, Actuals, Breach status (Response & Resolution), IsRunning'],
    ['  - Device/Asset: AssetTag, ModelName, SerialNumber'],
    ['  - CustomField1-3: values from custom fields set on the ticket itself'],
    ['  - LocationCustomField1-5: values from custom fields set on the ticket\'s'],
    ['    LOCATION (site code, region, building number, etc.), joined on LocationId'],
    ['  - Values: IsClosed="Closed"/"Open", IsPastDue="Overdue"/"On Track", Breach=1/0'],
    [''],
    ['• Teams'],
    ['  Team directory loaded from iiQ. Add FunctionalArea values for grouping.'],
    [''],
    ['• CustomFields'],
    ['  Reference list of every custom field your district defines, refreshed from'],
    ['  iiQ. Entity column marks each as "Ticket" or "Location". Copy a'],
    ['  CustomFieldTypeId from column C into the Config sheet to add it as a'],
    ['  TicketData column. Safe to regenerate at any time — nothing reads it during'],
    ['  a data load.'],
    [''],
    ['• DailySnapshot'],
    ['  Daily backlog metrics captured at 7 PM. Used for trending analysis.'],
    ['  Cannot be calculated retroactively — daily capture is essential.'],
    [''],
    ['• Logs'],
    ['  Operation logs for troubleshooting. Auto-pruned to 500 rows.'],
    [''],
    [''],
    ['DEFAULT ANALYTICS (7 sheets, created by Setup):'],
    [''],
    ['• MonthlyVolume — Tickets created vs closed by month, net change, closure rate'],
    ['• BacklogAging — Open tickets by age bucket (0-15, 16-30, 31-60, 61-90, 90+ days)'],
    ['• TeamWorkload — Open tickets, created/closed, and aging by team'],
    ['• SLACompliance — Monthly SLA breach rates, avg response/resolution times'],
    ['• PerformanceTrends — Key metrics trending over time ("Are we getting better?")'],
    ['• AtRiskResponse — Open tickets approaching Response SLA breach (75%+ threshold)'],
    ['• AtRiskResolution — Open tickets approaching Resolution SLA breach (75%+ threshold)'],
    [''],
    ['OPTIONAL ANALYTICS (18 sheets, add via iiQ Data > Add Analytics Sheet):'],
    [''],
    ['Volume & Trends:'],
    ['  • Seasonal Comparison (YoY) — Year-over-year monthly comparison'],
    ['  • Temporal Patterns — Day-of-week and hour-of-day distributions'],
    [''],
    ['Backlog & Quality:'],
    ['  • Stale Tickets — Open tickets with no recent activity'],
    ['  • Reopen Rate — Reopened ticket detection and quality indicators'],
    [''],
    ['SLA & Response:'],
    ['  • First Contact Resolution — Same-day and 4-hour resolution rates'],
    ['  • Response Distribution — Response time percentiles and consistency'],
    ['  • Response Trends — Monthly response time trends'],
    ['  • Queue Time Analysis — Time tickets wait before pickup'],
    ['  • Queue Time by Team — Queue time comparison across teams'],
    ['  • Queue Time Trend — Monthly queue time trends'],
    [''],
    ['Team & Staff:'],
    ['  • Technician Performance — Per-technician workload and metrics'],
    ['  • Functional Area Summary — Aggregated metrics by functional area'],
    [''],
    ['Location:'],
    ['  • Location Breakdown — Tickets by location with open/created/closed counts'],
    ['  • Location Type Comparison — Metrics by school type (Elementary/Middle/High)'],
    [''],
    ['Issue & Requester:'],
    ['  • Issue Category Volume — Open/Closed by category, breach rate'],
    ['  • Issue Type Volume — Top 50 issue types, filterable by category'],
    ['  • Priority Analysis — Metrics by priority level'],
    ['  • Frequent Requesters — Top 50 requesters with category data'],
    [''],
    ['Device:'],
    ['  • Device Reliability — Ticket counts and resolution by device model'],
    [''],
    ['All analytics sheets are formula-based and auto-calculate from TicketData.'],
    ['Any sheet can be deleted and recreated via the menu. Default sheets marked with ★.'],
    [''],
    [''],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    ['MENU REFERENCE'],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    [''],
    ['iiQ Data > Setup'],
    ['  • Setup Spreadsheet — Create all sheets (safe to run multiple times)'],
    ['  • Verify Configuration — Check API settings'],
    ['  • Refresh Custom Fields — List your district\'s ticket and location custom'],
    ['    fields (with their IDs) on the CustomFields sheet'],
    ['  • Setup Automated Triggers — Create all recommended triggers'],
    ['  • View Trigger Status — Show installed triggers and schedule'],
    ['  • Remove Automated Triggers — Remove all triggers (required before destructive ops)'],
    [''],
    ['iiQ Data > Refresh Teams'],
    ['  Load/refresh team directory from iiQ'],
    [''],
    ['iiQ Data > Ticket Data'],
    ['  • Continue Loading (Initial) — Resume initial data load (runs ~5.5 min)'],
    ['  • Clear Data + Reset Progress — Clear ticket data and unlock configuration'],
    ['  • Open Ticket Refresh (Start) — Begin fresh open ticket refresh'],
    ['  • Open Ticket Refresh (Continue) — Resume if refresh timed out'],
    ['  • Open Refresh Status — Show open refresh progress'],
    ['  • Full Reload — Clear ALL data and start fresh (use sparingly)'],
    ['  • Show Status — Display loading progress'],
    [''],
    ['iiQ Data > Add Analytics Sheet'],
    ['  Add optional analytics sheets organized by category:'],
    ['  Volume & Trends, Backlog & Quality, SLA & Response, Team & Staff,'],
    ['  Location, Issue & Requester, Device'],
    ['  • Regenerate All Monthly Sheets — Rebuild date-dependent sheets'],
    [''],
    ['iiQ Data > Daily Snapshot'],
    ['  • Capture Now — Record today\'s backlog metrics'],
    ['  • Populate Historical Estimates — Create estimated past snapshots'],
    [''],
    [''],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    ['TROUBLESHOOTING'],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    [''],
    ['"API configuration missing"'],
    ['  → Check Config sheet has API_BASE_URL and BEARER_TOKEN filled in'],
    [''],
    ['"Rate limited" or 429 errors'],
    ['  → Increase THROTTLE_MS in Config (default 1000ms)'],
    ['  → Script automatically retries with exponential backoff'],
    [''],
    ['Loading seems stuck'],
    ['  → Check Logs sheet for errors'],
    ['  → Run "Show Status" to see progress'],
    ['  → Each batch processes ~1250 tickets, large districts take multiple runs'],
    [''],
    ['Custom field columns are blank'],
    ['  → Run "Refresh Custom Fields", then confirm the value in Config matches a'],
    ['    CustomFieldTypeId on the CustomFields sheet'],
    ['  → Check the Entity column: a Location field in CUSTOM_FIELD_1/2/3 (or a'],
    ['    Ticket field in LOCATION_CUSTOM_FIELD_1-5) will never resolve'],
    ['  → If the Config _ID cell reads NOT_FOUND, the value was not recognized'],
    ['  → Custom fields are read at row-write time, so existing rows stay blank'],
    ['    until reloaded. Use "Clear Data + Reset Progress" then reload.'],
    [''],
    ['Location custom field is blank for some tickets only'],
    ['  → Expected when that ticket\'s location has no value set for the field in'],
    ['    iiQ. Values are per location, not per ticket.'],
    [''],
    ['"Configuration mismatch detected"'],
    ['  → A locked value (school year, page size, batch size, or a custom field)'],
    ['    was changed after loading began'],
    ['  → Use "Clear Data + Reset Progress" to unlock, then reload'],
    [''],
    ['Missing SLA data'],
    ['  → SLA data is fetched per-batch during ticket loading'],
    ['  → Tickets without assigned SLA policies will have blank SLA columns'],
    [''],
    ['Formula errors in analytics sheets'],
    ['  → Ensure TicketData has data loaded'],
    ['  → Check column headers match expected names'],
    [''],
    ['Trigger not running'],
    ['  → Check Apps Script > Triggers for errors'],
    ['  → Verify trigger is enabled and function name is correct'],
    ['  → Use iiQ Data > Setup > View Trigger Status for diagnostics'],
    [''],
    ['"Another operation is in progress"'],
    ['  → A script is already running (menu action or trigger)'],
    ['  → Wait a few minutes for it to complete, then try again'],
    ['  → Use iiQ Data > Setup > View Trigger Status to check what\'s running'],
    [''],
    ['"Remove triggers first"'],
    ['  → Destructive operations (Full Reload, Clear Data) require triggers to be removed'],
    ['  → Use iiQ Data > Setup > Remove Automated Triggers first'],
    ['  → Re-add triggers after the operation completes'],
    [''],
    [''],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    ['DASHBOARD INTEGRATION'],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    [''],
    ['GOOGLE DATA STUDIO (Recommended):'],
    ['1. Go to datastudio.google.com > Create > Report'],
    ['2. Add Google Sheets connector > Select this spreadsheet > TicketData sheet'],
    ['3. Add DailySnapshot as a second data source if needed'],
    ['4. See the Google Data Studio build guides in the project repository for step-by-step instructions'],
    [''],
    ['POWER BI:'],
    ['1. In Power BI Desktop: Get Data > Web'],
    ['2. Use the shareable link for each sheet (File > Share > Publish to web)'],
    ['3. Or use the Google Sheets connector if available'],
    ['4. Set up scheduled refresh in Power BI Service'],
    [''],
    ['Data is refreshed every 2 hours, so dashboards can refresh on a similar schedule.'],
    ['Weekly full refresh completes by ~4 AM Sunday.'],
    [''],
    [''],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    ['iiQ TELEMETRY'],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    [''],
    ['Anonymous usage telemetry pings the iiQ team once per successful refresh, so'],
    ['we can see which districts run which version, on which iiQ instance, with'],
    ['how many tickets — to prioritize features and catch regressions early.'],
    [''],
    ['What is sent (per ping):'],
    ['  • Stable install ID (UUID — generated locally, contains no PII)'],
    ['  • Project name (iiq-tickets-to-sheets) and script version'],
    ['  • iiQ instance hostname (e.g., demo.incidentiq.com)'],
    ['  • TicketData row count'],
    ['  • Names of analytics sheets present (only the canonical set this project ships)'],
    ['  • Script time zone, install timestamp, send timestamp'],
    [''],
    ['What is NOT sent: ticket data, API tokens, custom field values or names,'],
    ['custom sheet names you add yourself, user names or emails, anything from row contents.'],
    [''],
    ['POLICY: Automated polling requires telemetry opt-in.'],
    ['  • To opt out, set TELEMETRY_ENABLED to FALSE in the Config sheet.'],
    ['  • This DISABLES automated polling: time-based triggers uninstall on next fire.'],
    ['  • Manual menu refreshes (iiQ Data > Ticket Data > ...) continue to work.'],
    [''],
    ['To re-enable: set TELEMETRY_ENABLED back to TRUE, then run'],
    ['iiQ Data > Setup > Setup Automated Triggers to reinstall the triggers.'],
    [''],
    [''],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    ['SUPPORT'],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    [''],
    ['For issues or feature requests, check the Logs sheet first for error details.'],
    [''],
    ['Last updated: ' + new Date().toISOString().split('T')[0]],
  ];

  // Write content
  sheet.getRange(1, 1, content.length, 1).setValues(content);

  // Format title
  sheet.getRange(1, 1).setFontSize(16).setFontWeight('bold').setFontColor('#1a73e8');

  // Set text wrapping for long content
  sheet.getRange(1, 1, content.length, 1).setWrap(true);

  // Locate section headers, dividers and tabular rows in a single pass, then
  // apply each style class as one batched getRangeList call. Matches the
  // Instructions formatting used across the sibling iiq-*-to-sheets projects.
  // Titles are matched explicitly rather than by position, so editing the text
  // above cannot shift the styling onto unrelated rows.
  const sectionTitles = new Set([
    'OVERVIEW',
    'INITIAL SETUP',
    'AUTOMATED TRIGGERS (Recommended)',
    'SHEETS REFERENCE',
    'MENU REFERENCE',
    'TROUBLESHOOTING',
    'DASHBOARD INTEGRATION',
    'iiQ TELEMETRY',
    'SUPPORT'
  ]);
  const sectionRanges = [];
  const dividerRanges = [];
  const monoRanges = [];

  content.forEach(function(row, index) {
    const value = row[0];
    if (sectionTitles.has(value)) sectionRanges.push('A' + (index + 1));
    if (typeof value === 'string' && /^═+$/.test(value)) dividerRanges.push('A' + (index + 1));
    // Tabular rows (any line with 2+ pipes) only line up in a monospace font
    if (typeof value === 'string' && (value.match(/\|/g) || []).length >= 2) {
      monoRanges.push('A' + (index + 1));
    }
  });

  if (sectionRanges.length) {
    sheet.getRangeList(sectionRanges).setFontWeight('bold').setFontColor('#1a73e8');
  }
  if (dividerRanges.length) {
    sheet.getRangeList(dividerRanges).setFontColor('#dadce0');
  }
  if (monoRanges.length) {
    sheet.getRangeList(monoRanges).setFontFamily('Roboto Mono');
  }

  // Freeze title row
  sheet.setFrozenRows(1);

  // Set tab color
  sheet.setTabColor('#1a73e8');

  return true;
}

/**
 * Setup Config sheet with required settings
 * Deletes existing sheet if present for clean slate
 * @param {Spreadsheet} ss - The spreadsheet
 * @param {string} [schoolYear] - School year to pre-fill (e.g., "2023-2024"). Defaults to current school year.
 */
function setupConfigSheet(ss, schoolYear) {
  deleteSheetIfExists(ss, 'Config');
  const sheet = ss.insertSheet('Config');

  // Use provided school year or calculate default
  const defaultSchoolYear = schoolYear || getDefaultSchoolYear();

  // Headers and initial config values
  const configData = [
    ['Key', 'Value'],                                                  // 1
    ['', ''],                                                          // 2
    ['# API Configuration (Required)', ''],                            // 3
    ['API_BASE_URL', 'https://YOUR-DISTRICT.incidentiq.com'],          // 4
    ['BEARER_TOKEN', ''],                                              // 5
    ['SITE_ID', ''],                                                   // 6
    ['MODULE', 'Ticketing'],                                           // 7
    ['', ''],                                                          // 8
    ['# School Year Configuration', ''],                               // 9
    ['SCHOOL_YEAR', defaultSchoolYear],                                // 10
    ['SCHOOL_YEAR_START', '07-01'],                                    // 11
    ['', ''],                                                          // 12
    ['# Performance Settings (Optional)', ''],                         // 13
    ['PAGE_SIZE', '100'],                                              // 14
    ['THROTTLE_MS', '1000'],                                           // 15
    ['TICKET_BATCH_SIZE', '1250'],                                     // 16
    ['STALE_DAYS', '7'],                                               // 17
    ['SLA_RISK_PERCENT', '75'],                                        // 18
    ['', ''],                                                          // 19
    ['# Ticket Custom Field Columns (Optional - paste CustomFieldTypeId, up to 3)', ''],               // 20
    ['CUSTOM_FIELD_1', ''],                                            // 21
    ['CUSTOM_FIELD_2', ''],                                            // 22
    ['CUSTOM_FIELD_3', ''],                                            // 23
    ['', ''],                                                          // 24
    ['# Location Custom Field Columns (Optional - paste CustomFieldTypeId, up to 5)', ''],      // 25
    ['LOCATION_CUSTOM_FIELD_1', ''],                                   // 26
    ['LOCATION_CUSTOM_FIELD_2', ''],                                   // 27
    ['LOCATION_CUSTOM_FIELD_3', ''],                                   // 28
    ['LOCATION_CUSTOM_FIELD_4', ''],                                   // 29
    ['LOCATION_CUSTOM_FIELD_5', ''],                                   // 30
    ['', ''],                                                          // 31
    ['# Progress Tracking - Managed Automatically', ''],               // 32
    ['TICKET_TOTAL_PAGES', ''],                                        // 33
    ['TICKET_LAST_PAGE', '-1'],                                        // 34
    ['TICKET_COMPLETE', 'FALSE'],                                      // 35
    ['TICKET_LAST_FETCH', ''],                                         // 36
    ['CUSTOM_FIELD_1_ID', ''],                                         // 37
    ['CUSTOM_FIELD_2_ID', ''],                                         // 38
    ['CUSTOM_FIELD_3_ID', ''],                                         // 39
    ['LOCATION_CUSTOM_FIELD_1_ID', ''],                                // 40
    ['LOCATION_CUSTOM_FIELD_2_ID', ''],                                // 41
    ['LOCATION_CUSTOM_FIELD_3_ID', ''],                                // 42
    ['LOCATION_CUSTOM_FIELD_4_ID', ''],                                // 43
    ['LOCATION_CUSTOM_FIELD_5_ID', ''],                                // 44
    ['LOCATION_CF_TYPE_CACHE', ''],                                    // 45
    ['', ''],                                                          // 46
    ['# Config Lock - Set when loading starts, cleared by "Clear Data + Reset"', ''], // 47
    ['SCHOOL_YEAR_LOADED', ''],                                        // 48
    ['PAGE_SIZE_LOADED', ''],                                          // 49
    ['BATCH_SIZE_LOADED', ''],                                         // 50
    ['MODULE_LOADED', ''],                                             // 51
    ['CUSTOM_FIELD_1_LOADED', ''],                                     // 52
    ['CUSTOM_FIELD_2_LOADED', ''],                                     // 53
    ['CUSTOM_FIELD_3_LOADED', ''],                                     // 54
    ['LOCATION_CUSTOM_FIELD_1_LOADED', ''],                            // 55
    ['LOCATION_CUSTOM_FIELD_2_LOADED', ''],                            // 56
    ['LOCATION_CUSTOM_FIELD_3_LOADED', ''],                            // 57
    ['LOCATION_CUSTOM_FIELD_4_LOADED', ''],                            // 58
    ['LOCATION_CUSTOM_FIELD_5_LOADED', ''],                            // 59
    ['', ''],                                                          // 60
    ['LAST_REFRESH', ''],                                              // 61
    ['', ''],                                                          // 62
    ['# Version Information', ''],                                     // 63
    ['SCRIPT_VERSION', SCRIPT_VERSION],                                // 64
    ['LATEST_VERSION', ''],                                            // 65
    ['VERSION_CHECK_DATE', ''],                                        // 66
    ['', ''],                                                          // 67
    ['# Dashboard Configuration', ''],                                 // 68
    ['DASHBOARD_URL', ''],                                             // 69
    ['', ''],                                                          // 70
    ['# iiQ Telemetry (set FALSE to opt out — also disables automated polling)', ''], // 71
    ['TELEMETRY_ENABLED', 'TRUE']                                              // 54
  ];

  sheet.getRange(1, 1, configData.length, 2).setValues(configData);

  // Format header
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');

  // Format section headers (rows starting with #)
  const sectionRows = [3, 9, 13, 20, 25, 34, 45, 50, 53];
  sectionRows.forEach(row => {
    sheet.getRange(row, 1, 1, 2).setFontWeight('bold').setBackground('#e8f0fe');
  });

  // Add dropdown validation for MODULE (row 7: SITE_ID is row 6, MODULE is row 7)
  const moduleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Ticketing', 'Facilities', 'HRSD'], true)
    .setHelpText('Ticketing = IT Ticketing module, Facilities = Facilities Ticketing module, HRSD = HR Service Delivery module')
    .build();
  sheet.getRange(7, 2).setDataValidation(moduleRule);

  // Column widths
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 400);

  sheet.setFrozenRows(1);

  return true;
}

/**
 * Setup TicketData sheet with TICKET_COLUMN_COUNT-column header
 * Deletes existing sheet if present for clean slate
 */
function setupTicketDataSheet(ss) {
  deleteSheetIfExists(ss, 'TicketData');
  const sheet = ss.insertSheet('TicketData');

  const headers = [
    'TicketId', 'TicketNumber', 'Subject', 'Year',
    'CreatedDate', 'StartedDate', 'ModifiedDate', 'ClosedDate', 'IsClosed',
    'Status', 'TeamId', 'TeamName', 'LocationId', 'LocationName', 'LocationType',
    'OwnerId', 'OwnerName', 'AgeDays', 'Priority', 'IsPastDue', 'DueDate',
    'SlaId', 'SlaName', 'IssueCategoryId', 'IssueCategoryName',
    'IssueTypeId', 'IssueTypeName', 'RequesterId', 'RequesterName',
    'ResponseThreshold', 'ResponseActual', 'ResponseBreach',
    'ResolutionThreshold', 'ResolutionActual', 'ResolutionBreach', 'IsRunning',
    'AssetTag', 'ModelName', 'SerialNumber',
    'AssignedToUserId', 'AssignedToUserName',
    'AssetId', 'AssetCategory',
    'CustomField1', 'CustomField2', 'CustomField3',
    'RequesterRole',
    'LocationCustomField1', 'LocationCustomField2', 'LocationCustomField3',
    'LocationCustomField4', 'LocationCustomField5'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format header
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('white')
    .setWrap(false);

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2); // Freeze TicketId and TicketNumber

  return true;
}

/**
 * Setup Teams sheet
 * Deletes existing sheet if present for clean slate
 */
function setupTeamsSheet(ss) {
  deleteSheetIfExists(ss, 'Teams');
  const sheet = ss.insertSheet('Teams');

  const headers = ['TeamId', 'TeamName', 'FunctionalArea', 'IsActive'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format header
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('white');

  // Column widths
  sheet.setColumnWidth(1, 300);  // TeamId (GUID)
  sheet.setColumnWidth(2, 200);  // TeamName
  sheet.setColumnWidth(3, 150);  // FunctionalArea (user fills in)
  sheet.setColumnWidth(4, 80);   // IsActive

  sheet.setFrozenRows(1);

  // Note: FunctionalArea column (C) is free text - district fills in their own values
  // No data validation - allows pasting values directly

  return true;
}

/**
 * Setup DailySnapshot sheet
 * Headers match GUIDE.md for PerformanceTrends lookups
 * Deletes existing sheet if present for clean slate
 */
function setupDailySnapshotSheet(ss) {
  deleteSheetIfExists(ss, 'DailySnapshot');
  const sheet = ss.insertSheet('DailySnapshot');

  // Headers matching GUIDE.md
  const headers = ['Date', 'Open Tickets', 'Aged 30+ Count', '% Aged 30+'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format header
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('white');

  // Format % Aged 30+ column as percentage
  sheet.getRange('D:D').setNumberFormat('0.0%');

  // Format Date column
  sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd');

  // Column widths
  sheet.setColumnWidth(1, 100);  // Date
  sheet.setColumnWidth(2, 100);  // Open Tickets
  sheet.setColumnWidth(3, 120);  // Aged 30+ Count
  sheet.setColumnWidth(4, 100);  // % Aged 30+

  sheet.setFrozenRows(1);

  // Add note
  sheet.getRange('A1').setNote('Populated by triggerDailySnapshot (7 PM daily). Cannot be calculated retroactively - captures point-in-time backlog state.');

  return true;
}

/**
 * Setup Logs sheet
 * Deletes existing sheet if present for clean slate
 */
function setupLogsSheet(ss) {
  deleteSheetIfExists(ss, 'Logs');
  const sheet = ss.insertSheet('Logs');

  const headers = ['Timestamp', 'Operation', 'Status', 'Details'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format header
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('white');

  // Column widths
  sheet.setColumnWidth(1, 180);  // Timestamp
  sheet.setColumnWidth(2, 150);  // Operation
  sheet.setColumnWidth(3, 100);  // Status
  sheet.setColumnWidth(4, 500);  // Details

  sheet.setFrozenRows(1);

  return true;
}

/**
 * Setup SLACompliance sheet with formulas
 * Deletes existing sheet if present for clean slate
 */
function setupSLAComplianceSheet(ss) {
  deleteSheetIfExists(ss, 'SLACompliance');
  const sheet = ss.insertSheet('SLACompliance');

  const headers = ['Month', 'Year', 'Closed', 'Breaches', 'Breach Rate', 'Avg Response (hrs)', 'Avg Resolution (hrs)'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Get month range from actual data (uses ClosedDate column G)
  const monthRange = getMonthRangeFromData(ss, 'G');

  const dataRows = [];
  for (const period of monthRange) {
    const rowNum = dataRows.length + 2;
    const y = period.year;
    const monthNum = period.monthNum;

    // Build date filter for this month
    const startDate = `TEXT(DATE(${y},${monthNum},1), "YYYY-MM-DD")`;
    const endDate = `TEXT(DATE(${y},${monthNum}+1,1), "YYYY-MM-DD")`;

    dataRows.push([
      period.monthName,
      y,
      // Closed: Count tickets closed in this month
      `=COUNTIFS(TicketData!I:I, "Closed", TicketData!H:H, ">="&${startDate}, TicketData!H:H, "<"&${endDate})`,
      // Breaches: Count where ResponseBreach=1 OR ResolutionBreach=1 (closed tickets in this month)
      `=LET(startDate, ${startDate}, endDate, ${endDate}, COUNTIFS(TicketData!I:I, "Closed", TicketData!H:H, ">="&startDate, TicketData!H:H, "<"&endDate, TicketData!AF:AF, 1) + COUNTIFS(TicketData!I:I, "Closed", TicketData!H:H, ">="&startDate, TicketData!H:H, "<"&endDate, TicketData!AI:AI, 1) - COUNTIFS(TicketData!I:I, "Closed", TicketData!H:H, ">="&startDate, TicketData!H:H, "<"&endDate, TicketData!AF:AF, 1, TicketData!AI:AI, 1))`,
      // Breach Rate
      `=IF(C${rowNum}>0, D${rowNum}/C${rowNum}, "N/A")`,
      // Avg Response (hrs): Average of ResponseActual (col AD) for closed tickets, convert minutes to hours
      `=IFERROR(AVERAGEIFS(TicketData!AE:AE, TicketData!I:I, "Closed", TicketData!H:H, ">="&${startDate}, TicketData!H:H, "<"&${endDate}, TicketData!AE:AE, ">0")/60, "N/A")`,
      // Avg Resolution (hrs): Average of ResolutionActual (col AG) for closed tickets, convert minutes to hours
      `=IFERROR(AVERAGEIFS(TicketData!AH:AH, TicketData!I:I, "Closed", TicketData!H:H, ">="&${startDate}, TicketData!H:H, "<"&${endDate}, TicketData!AH:AH, ">0")/60, "N/A")`
    ]);
  }

  if (dataRows.length > 0) {
    sheet.getRange(2, 1, dataRows.length, 7).setValues(dataRows);
  }

  // Format header
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#34a853')
    .setFontColor('white');

  // Format columns
  sheet.getRange('E:E').setNumberFormat('0.0%');   // Breach Rate
  sheet.getRange('F:G').setNumberFormat('0.0');    // Hours

  // Column widths
  sheet.setColumnWidth(1, 100);  // Month
  sheet.setColumnWidth(6, 140);  // Avg Response
  sheet.setColumnWidth(7, 150);  // Avg Resolution

  sheet.setFrozenRows(1);

  // Add note explaining the formulas
  sheet.getRange('A1').setNote(
    'SLA Compliance Metrics\n\n' +
    'Formula-based - auto-calculates from TicketData.\n' +
    'Date range based on actual ticket data.\n\n' +
    'Columns:\n' +
    '- Closed: Tickets closed in the month\n' +
    '- Breaches: Response OR Resolution SLA breached\n' +
    '- Breach Rate: Breaches / Closed\n' +
    '- Avg Response: First response time in hours\n' +
    '- Avg Resolution: Resolution time in hours\n\n' +
    'Recreate sheet to update date range if data grows.'
  );

  return true;
}

/**
 * Setup MonthlyVolume sheet with formulas
 * Deletes existing sheet if present for clean slate
 */
function setupMonthlyVolumeSheet(ss) {
  deleteSheetIfExists(ss, 'MonthlyVolume');
  const sheet = ss.insertSheet('MonthlyVolume');

  // Headers matching GUIDE.md: Month, Year, Created, Closed, Net Change, Closure Rate
  const headers = ['Month', 'Year', 'Created', 'Closed', 'Net Change', 'Closure Rate'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Get month range from actual data (uses CreatedDate column E)
  const monthRange = getMonthRangeFromData(ss, 'E');

  const dataRows = [];
  for (const period of monthRange) {
    const rowNum = dataRows.length + 2;
    const y = period.year;
    const monthNum = period.monthNum;

    // Build date filter for this month
    const startDate = `TEXT(DATE(${y},${monthNum},1), "YYYY-MM-DD")`;
    const endDate = `TEXT(DATE(${y},${monthNum}+1,1), "YYYY-MM-DD")`;

    dataRows.push([
      period.monthName,
      y,
      // Created formula
      `=COUNTIFS(TicketData!E:E, ">="&${startDate}, TicketData!E:E, "<"&${endDate})`,
      // Closed formula
      `=COUNTIFS(TicketData!H:H, ">="&${startDate}, TicketData!H:H, "<"&${endDate})`,
      // Net Change
      `=C${rowNum}-D${rowNum}`,
      // Closure Rate
      `=IF(C${rowNum}>0, D${rowNum}/C${rowNum}, "N/A")`
    ]);
  }

  if (dataRows.length > 0) {
    sheet.getRange(2, 1, dataRows.length, 6).setValues(dataRows);
  }

  // Format header
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#fbbc04')
    .setFontColor('black');

  // Format Closure Rate as percentage
  sheet.getRange('F:F').setNumberFormat('0.0%');

  // Conditional formatting for Net Change (red if positive/backlog growing, green if negative/backlog shrinking)
  if (dataRows.length > 0) {
    const netChangeRange = sheet.getRange(2, 5, dataRows.length, 1);
    const positiveRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground('#fce8e6')
      .setRanges([netChangeRange])
      .build();
    const negativeRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground('#e6f4ea')
      .setRanges([netChangeRange])
      .build();
    sheet.setConditionalFormatRules([positiveRule, negativeRule]);
  }

  sheet.setFrozenRows(1);

  return true;
}

/**
 * Setup BacklogAging sheet with formulas
 * Matches GUIDE.md structure with Sample Ticket and Last Refreshed columns
 * Deletes existing sheet if present for clean slate
 */
function setupBacklogAgingSheet(ss) {
  deleteSheetIfExists(ss, 'BacklogAging');
  const sheet = ss.insertSheet('BacklogAging');

  // Headers matching GUIDE.md
  const headers = ['Age Bucket', 'Count', '% of Total', 'Sample Ticket', 'Last Refreshed'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Age buckets matching GUIDE.md: 0-15, 16-30, 31-60, 61-90, 90+
  const bucketData = [
    // Row 2: 0-15 days
    [
      '0-15 days',
      '=COUNTIFS(TicketData!I:I, "Open", TicketData!R:R, ">=0", TicketData!R:R, "<=15")',
      '=IF($B$7>0, B2/$B$7, 0)',
      '=IFERROR(INDEX(SORT(FILTER({TicketData!B2:B,TicketData!R2:R},(TicketData!I2:I="Open")*(TicketData!R2:R>=0)*(TicketData!R2:R<=15)),2,FALSE),1,1),"")',
      '=IFERROR(VLOOKUP("LAST_REFRESH", Config!A:B, 2, FALSE), "")'
    ],
    // Row 3: 16-30 days
    [
      '16-30 days',
      '=COUNTIFS(TicketData!I:I, "Open", TicketData!R:R, ">=16", TicketData!R:R, "<=30")',
      '=IF($B$7>0, B3/$B$7, 0)',
      '=IFERROR(INDEX(SORT(FILTER({TicketData!B2:B,TicketData!R2:R},(TicketData!I2:I="Open")*(TicketData!R2:R>=16)*(TicketData!R2:R<=30)),2,FALSE),1,1),"")',
      '=IFERROR(VLOOKUP("LAST_REFRESH", Config!A:B, 2, FALSE), "")'
    ],
    // Row 4: 31-60 days
    [
      '31-60 days',
      '=COUNTIFS(TicketData!I:I, "Open", TicketData!R:R, ">=31", TicketData!R:R, "<=60")',
      '=IF($B$7>0, B4/$B$7, 0)',
      '=IFERROR(INDEX(SORT(FILTER({TicketData!B2:B,TicketData!R2:R},(TicketData!I2:I="Open")*(TicketData!R2:R>=31)*(TicketData!R2:R<=60)),2,FALSE),1,1),"")',
      '=IFERROR(VLOOKUP("LAST_REFRESH", Config!A:B, 2, FALSE), "")'
    ],
    // Row 5: 61-90 days
    [
      '61-90 days',
      '=COUNTIFS(TicketData!I:I, "Open", TicketData!R:R, ">=61", TicketData!R:R, "<=90")',
      '=IF($B$7>0, B5/$B$7, 0)',
      '=IFERROR(INDEX(SORT(FILTER({TicketData!B2:B,TicketData!R2:R},(TicketData!I2:I="Open")*(TicketData!R2:R>=61)*(TicketData!R2:R<=90)),2,FALSE),1,1),"")',
      '=IFERROR(VLOOKUP("LAST_REFRESH", Config!A:B, 2, FALSE), "")'
    ],
    // Row 6: 90+ days
    [
      '90+ days',
      '=COUNTIFS(TicketData!I:I, "Open", TicketData!R:R, ">90")',
      '=IF($B$7>0, B6/$B$7, 0)',
      '=IFERROR(INDEX(SORT(FILTER({TicketData!B2:B,TicketData!R2:R},(TicketData!I2:I="Open")*(TicketData!R2:R>90)),2,FALSE),1,1),"")',
      '=IFERROR(VLOOKUP("LAST_REFRESH", Config!A:B, 2, FALSE), "")'
    ],
    // Row 7: TOTAL
    [
      'TOTAL',
      '=SUM(B2:B6)',
      '100%',
      '',
      ''
    ]
  ];

  sheet.getRange(2, 1, bucketData.length, 5).setValues(bucketData);

  // Format header
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#ea4335')
    .setFontColor('white');

  // Format TOTAL row
  sheet.getRange(7, 1, 1, 5).setFontWeight('bold');

  // Format percentage column
  sheet.getRange('C:C').setNumberFormat('0.0%');

  // Column widths
  sheet.setColumnWidth(1, 100);  // Age Bucket
  sheet.setColumnWidth(2, 80);   // Count
  sheet.setColumnWidth(3, 90);   // % of Total
  sheet.setColumnWidth(4, 120);  // Sample Ticket
  sheet.setColumnWidth(5, 180);  // Last Refreshed

  sheet.setFrozenRows(1);

  // Add note
  sheet.getRange('D1').setNote('Shows the oldest ticket number in each age bucket (sorted by AgeDays descending)');

  return true;
}

/**
 * Setup ResolutionAging sheet — "Days to Resolve" distribution for closed tickets
 * Shows how quickly tickets are being resolved, bucketed by age at closure.
 * Uses AgeDays (column R) which stores days-from-creation-to-closure for closed tickets.
 * Deletes existing sheet if present for clean slate.
 */
function setupResolutionAgingSheet(ss) {
  deleteSheetIfExists(ss, 'ResolutionAging');
  const sheet = ss.insertSheet('ResolutionAging');

  const headers = ['Resolution Time', 'Count', '% of Total', 'Sample Ticket', 'Last Refreshed'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Buckets match BacklogAging: 0-15, 16-30, 31-60, 61-90, 90+
  const bucketData = [
    // Row 2: 0-15 days
    [
      '0-15 days',
      '=COUNTIFS(TicketData!I:I, "Closed", TicketData!R:R, ">=0", TicketData!R:R, "<=15")',
      '=IF($B$7>0, B2/$B$7, 0)',
      '=IFERROR(INDEX(SORT(FILTER({TicketData!B2:B,TicketData!R2:R},(TicketData!I2:I="Closed")*(TicketData!R2:R>=0)*(TicketData!R2:R<=15)),2,FALSE),1,1),"")',
      '=IFERROR(VLOOKUP("LAST_REFRESH", Config!A:B, 2, FALSE), "")'
    ],
    // Row 3: 16-30 days
    [
      '16-30 days',
      '=COUNTIFS(TicketData!I:I, "Closed", TicketData!R:R, ">=16", TicketData!R:R, "<=30")',
      '=IF($B$7>0, B3/$B$7, 0)',
      '=IFERROR(INDEX(SORT(FILTER({TicketData!B2:B,TicketData!R2:R},(TicketData!I2:I="Closed")*(TicketData!R2:R>=16)*(TicketData!R2:R<=30)),2,FALSE),1,1),"")',
      '=IFERROR(VLOOKUP("LAST_REFRESH", Config!A:B, 2, FALSE), "")'
    ],
    // Row 4: 31-60 days
    [
      '31-60 days',
      '=COUNTIFS(TicketData!I:I, "Closed", TicketData!R:R, ">=31", TicketData!R:R, "<=60")',
      '=IF($B$7>0, B4/$B$7, 0)',
      '=IFERROR(INDEX(SORT(FILTER({TicketData!B2:B,TicketData!R2:R},(TicketData!I2:I="Closed")*(TicketData!R2:R>=31)*(TicketData!R2:R<=60)),2,FALSE),1,1),"")',
      '=IFERROR(VLOOKUP("LAST_REFRESH", Config!A:B, 2, FALSE), "")'
    ],
    // Row 5: 61-90 days
    [
      '61-90 days',
      '=COUNTIFS(TicketData!I:I, "Closed", TicketData!R:R, ">=61", TicketData!R:R, "<=90")',
      '=IF($B$7>0, B5/$B$7, 0)',
      '=IFERROR(INDEX(SORT(FILTER({TicketData!B2:B,TicketData!R2:R},(TicketData!I2:I="Closed")*(TicketData!R2:R>=61)*(TicketData!R2:R<=90)),2,FALSE),1,1),"")',
      '=IFERROR(VLOOKUP("LAST_REFRESH", Config!A:B, 2, FALSE), "")'
    ],
    // Row 6: 90+ days
    [
      '90+ days',
      '=COUNTIFS(TicketData!I:I, "Closed", TicketData!R:R, ">90")',
      '=IF($B$7>0, B6/$B$7, 0)',
      '=IFERROR(INDEX(SORT(FILTER({TicketData!B2:B,TicketData!R2:R},(TicketData!I2:I="Closed")*(TicketData!R2:R>90)),2,FALSE),1,1),"")',
      '=IFERROR(VLOOKUP("LAST_REFRESH", Config!A:B, 2, FALSE), "")'
    ],
    // Row 7: TOTAL
    [
      'TOTAL',
      '=SUM(B2:B6)',
      '100%',
      '',
      ''
    ]
  ];

  sheet.getRange(2, 1, bucketData.length, 5).setValues(bucketData);

  // Format header
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1565c0')
    .setFontColor('white');

  // Format TOTAL row
  sheet.getRange(7, 1, 1, 5).setFontWeight('bold');

  // Format percentage column
  sheet.getRange('C:C').setNumberFormat('0.0%');

  // Column widths
  sheet.setColumnWidth(1, 120);  // Resolution Time
  sheet.setColumnWidth(2, 80);   // Count
  sheet.setColumnWidth(3, 90);   // % of Total
  sheet.setColumnWidth(4, 120);  // Sample Ticket
  sheet.setColumnWidth(5, 180);  // Last Refreshed

  sheet.setFrozenRows(1);

  // Add notes
  sheet.getRange('A1').setNote(
    'Days to Resolve Aging Report\n\n' +
    'Shows how long it takes to resolve tickets, bucketed by days from creation to closure.\n' +
    'Uses AgeDays (column R) for closed tickets only.\n\n' +
    'To recreate: iiQ Data > Add Analytics Sheet > Backlog & Quality > Resolution Aging');
  sheet.getRange('D1').setNote('Shows the slowest-resolved ticket number in each bucket (sorted by AgeDays descending)');

  return true;
}

/**
 * Setup TeamWorkload sheet with formulas
 * Uses a single array formula for sortable output
 * Deletes existing sheet if present for clean slate
 */
function setupTeamWorkloadSheet(ss) {
  deleteSheetIfExists(ss, 'TeamWorkload');
  const sheet = ss.insertSheet('TeamWorkload');

  // Headers matching GUIDE.md structure
  const headers = ['Team Name', 'Functional Area', 'Open', 'Created (MTD)', 'Closed (MTD)', 'Aged 30+', 'Last Refreshed', 'Sort Col#', 'Desc?'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Single LET formula that outputs the entire sortable table
  // This populates columns A-F dynamically based on TicketData
  // Note: Uses column L (TeamName) consistently for matching
  const mainFormula =
    '=LET(' +
    'teams, UNIQUE(FILTER(TicketData!L2:L, TicketData!L2:L<>"", TicketData!L2:L<>"TeamName")),' +
    'mtdStart, DATE(YEAR(TODAY()),MONTH(TODAY()),1),' +
    'mtdEnd, DATE(YEAR(TODAY()),MONTH(TODAY())+1,1),' +
    'col_a, teams,' +
    'col_b, BYROW(teams, LAMBDA(t, IFERROR(VLOOKUP(t, Teams!B:C, 2, FALSE), ""))),' +
    'col_c, BYROW(teams, LAMBDA(t, COUNTIFS(TicketData!L:L, t, TicketData!I:I, "Open"))),' +
    'col_d, BYROW(teams, LAMBDA(t, COUNTIFS(TicketData!L:L, t, TicketData!E:E, ">="&mtdStart, TicketData!E:E, "<"&mtdEnd))),' +
    'col_e, BYROW(teams, LAMBDA(t, COUNTIFS(TicketData!L:L, t, TicketData!H:H, ">="&mtdStart, TicketData!H:H, "<"&mtdEnd))),' +
    'col_f, BYROW(teams, LAMBDA(t, COUNTIFS(TicketData!L:L, t, TicketData!I:I, "Open", TicketData!R:R, ">=30"))),' +
    'data, HSTACK(col_a, col_b, col_c, col_d, col_e, col_f),' +
    'SORT(data, $H$2, $I$2))';

  sheet.getRange('A2').setValue(mainFormula);

  // Last Refreshed formula
  sheet.getRange('G2').setValue('=IFERROR(VLOOKUP("LAST_REFRESH", Config!A:B, 2, FALSE), "")');

  // Default sort settings (column 3 = Open, descending)
  sheet.getRange('H2').setValue(3);
  sheet.getRange('I2').setValue('FALSE');

  // Format header
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#673ab7')
    .setFontColor('white');

  // Column widths
  sheet.setColumnWidth(1, 180);  // Team Name
  sheet.setColumnWidth(2, 150);  // Functional Area
  sheet.setColumnWidth(7, 180);  // Last Refreshed
  sheet.setColumnWidth(8, 80);   // Sort Col#
  sheet.setColumnWidth(9, 60);   // Desc?

  sheet.setFrozenRows(1);

  // Add data validation for sort column
  const sortColRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['1', '2', '3', '4', '5', '6'], true)
    .setHelpText('1=Team, 2=FA, 3=Open, 4=Created, 5=Closed, 6=Aged')
    .build();
  sheet.getRange('H2').setDataValidation(sortColRule);

  // Add data validation for sort order
  const sortOrderRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['FALSE', 'TRUE'], true)
    .setHelpText('FALSE=Descending, TRUE=Ascending')
    .build();
  sheet.getRange('I2').setDataValidation(sortOrderRule);

  // Add notes
  sheet.getRange('A2').setNote('This formula auto-populates from TicketData. Functional Area is looked up from Teams sheet column C.');
  sheet.getRange('H2').setNote('Sort column: 1=Team, 2=FA, 3=Open, 4=Created, 5=Closed, 6=Aged');
  sheet.getRange('I2').setNote('FALSE=Descending (high to low), TRUE=Ascending (low to high)');

  return true;
}

/**
 * Setup LocationBreakdown sheet with formulas
 * Deletes existing sheet if present for clean slate
 */
function setupLocationBreakdownSheet(ss) {
  deleteSheetIfExists(ss, 'LocationBreakdown');
  const sheet = ss.insertSheet('LocationBreakdown');

  // Headers - includes sort controls
  const headers = ['Location Name', 'Location Type', 'Open', 'Created (MTD)', 'Closed (MTD)', 'Last Refreshed', 'Sort Col#', 'Desc?'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Single LET formula that outputs the entire sortable table
  // Note: Uses column N (LocationName) consistently for matching
  const mainFormula =
    '=LET(' +
    'locs, UNIQUE(FILTER(TicketData!N2:N, TicketData!N2:N<>"", TicketData!N2:N<>"LocationName")),' +
    'mtdStart, DATE(YEAR(TODAY()),MONTH(TODAY()),1),' +
    'mtdEnd, DATE(YEAR(TODAY()),MONTH(TODAY())+1,1),' +
    'col_a, locs,' +
    'col_b, BYROW(locs, LAMBDA(l, IFERROR(INDEX(TicketData!O:O, MATCH(l, TicketData!N:N, 0)), ""))),' +
    'col_c, BYROW(locs, LAMBDA(l, COUNTIFS(TicketData!N:N, l, TicketData!I:I, "Open"))),' +
    'col_d, BYROW(locs, LAMBDA(l, COUNTIFS(TicketData!N:N, l, TicketData!E:E, ">="&mtdStart, TicketData!E:E, "<"&mtdEnd))),' +
    'col_e, BYROW(locs, LAMBDA(l, COUNTIFS(TicketData!N:N, l, TicketData!H:H, ">="&mtdStart, TicketData!H:H, "<"&mtdEnd))),' +
    'data, HSTACK(col_a, col_b, col_c, col_d, col_e),' +
    'SORT(data, $G$2, $H$2))';

  sheet.getRange('A2').setValue(mainFormula);
  sheet.getRange('F2').setValue('=IFERROR(VLOOKUP("LAST_REFRESH", Config!A:B, 2, FALSE), "")');

  // Default sort settings (column 3 = Open, descending)
  sheet.getRange('G2').setValue(3);
  sheet.getRange('H2').setValue('FALSE');

  // Format header
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#0f9d58')
    .setFontColor('white');

  // Column widths
  sheet.setColumnWidth(1, 200);  // Location Name
  sheet.setColumnWidth(2, 120);  // Location Type
  sheet.setColumnWidth(6, 180);  // Last Refreshed
  sheet.setColumnWidth(7, 80);   // Sort Col#
  sheet.setColumnWidth(8, 60);   // Desc?

  sheet.setFrozenRows(1);

  // Add data validation for sort column
  const sortColRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['1', '2', '3', '4', '5'], true)
    .setHelpText('1=Location, 2=Type, 3=Open, 4=Created, 5=Closed')
    .build();
  sheet.getRange('G2').setDataValidation(sortColRule);

  // Add data validation for sort order
  const sortOrderRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['FALSE', 'TRUE'], true)
    .setHelpText('FALSE=Descending, TRUE=Ascending')
    .build();
  sheet.getRange('H2').setDataValidation(sortOrderRule);

  // Add notes
  sheet.getRange('A2').setNote('This formula auto-populates from TicketData. Use Sort Col# and Desc? to change sorting.');
  sheet.getRange('G2').setNote('Sort column: 1=Location, 2=Type, 3=Open, 4=Created, 5=Closed');
  sheet.getRange('H2').setNote('FALSE=Descending (high to low), TRUE=Ascending (low to high)');

  return true;
}

/**
 * Setup FunctionalAreaSummary sheet with formulas
 * Deletes existing sheet if present for clean slate
 */
function setupFunctionalAreaSummarySheet(ss) {
  deleteSheetIfExists(ss, 'FunctionalAreaSummary');
  const sheet = ss.insertSheet('FunctionalAreaSummary');

  // Headers - includes sort controls
  const headers = ['Functional Area', 'Teams', 'Open', 'Created (MTD)', 'Closed (MTD)', 'Aged 30+', '% Aged 30+', 'Last Refreshed', 'Sort Col#', 'Desc?'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Single LET formula that outputs the entire sortable table
  // Aggregates from TeamWorkload sheet
  const mainFormula =
    '=LET(' +
    'fas, UNIQUE(FILTER(TeamWorkload!B2:B, TeamWorkload!B2:B<>"", TeamWorkload!A2:A<>"")),' +
    'col_a, fas,' +
    'col_b, BYROW(fas, LAMBDA(f, COUNTIF(TeamWorkload!B:B, f))),' +
    'col_c, BYROW(fas, LAMBDA(f, SUMIF(TeamWorkload!B:B, f, TeamWorkload!C:C))),' +
    'col_d, BYROW(fas, LAMBDA(f, SUMIF(TeamWorkload!B:B, f, TeamWorkload!D:D))),' +
    'col_e, BYROW(fas, LAMBDA(f, SUMIF(TeamWorkload!B:B, f, TeamWorkload!E:E))),' +
    'col_f, BYROW(fas, LAMBDA(f, SUMIF(TeamWorkload!B:B, f, TeamWorkload!F:F))),' +
    'col_g, BYROW(fas, LAMBDA(f, LET(open, SUMIF(TeamWorkload!B:B, f, TeamWorkload!C:C), aged, SUMIF(TeamWorkload!B:B, f, TeamWorkload!F:F), IF(open>0, aged/open, 0)))),' +
    'data, HSTACK(col_a, col_b, col_c, col_d, col_e, col_f, col_g),' +
    'SORT(data, $I$2, $J$2))';

  sheet.getRange('A2').setValue(mainFormula);
  sheet.getRange('H2').setValue('=IFERROR(VLOOKUP("LAST_REFRESH", Config!A:B, 2, FALSE), "")');

  // Default sort settings (column 3 = Open, descending)
  sheet.getRange('I2').setValue(3);
  sheet.getRange('J2').setValue('FALSE');

  // Format header
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#9c27b0')
    .setFontColor('white');

  // Format percentage column
  sheet.getRange('G:G').setNumberFormat('0.0%');

  // Column widths
  sheet.setColumnWidth(1, 180);  // Functional Area
  sheet.setColumnWidth(8, 180);  // Last Refreshed
  sheet.setColumnWidth(9, 80);   // Sort Col#
  sheet.setColumnWidth(10, 60);  // Desc?

  sheet.setFrozenRows(1);

  // Add data validation for sort column
  const sortColRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['1', '2', '3', '4', '5', '6', '7'], true)
    .setHelpText('1=FA, 2=Teams, 3=Open, 4=Created, 5=Closed, 6=Aged, 7=%Aged')
    .build();
  sheet.getRange('I2').setDataValidation(sortColRule);

  // Add data validation for sort order
  const sortOrderRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['FALSE', 'TRUE'], true)
    .setHelpText('FALSE=Descending, TRUE=Ascending')
    .build();
  sheet.getRange('J2').setDataValidation(sortOrderRule);

  // Add notes
  sheet.getRange('A2').setNote('This formula auto-populates from TeamWorkload. Prerequisite: TeamWorkload must have data and Teams must have FunctionalArea filled in.');
  sheet.getRange('I2').setNote('Sort column: 1=FA, 2=Teams, 3=Open, 4=Created, 5=Closed, 6=Aged, 7=%Aged');
  sheet.getRange('J2').setNote('FALSE=Descending (high to low), TRUE=Ascending (low to high)');

  return true;
}

/**
 * Setup AtRiskResponse sheet with formulas
 * Shows tickets approaching Response SLA breach threshold
 * Deletes existing sheet if present for clean slate
 */
function setupAtRiskResponseSheet(ss) {
  deleteSheetIfExists(ss, 'AtRiskResponse');
  const sheet = ss.insertSheet('AtRiskResponse');

  // Headers
  const headers = ['Ticket Number', 'Subject', 'Team', 'Threshold (hrs)', 'Elapsed (hrs)', '% of SLA', 'Time Remaining (hrs)'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Response SLA at-risk formula
  // Shows open tickets where Response SLA usage is between risk% and 100%
  const riskLookup = 'IFERROR(VLOOKUP("SLA_RISK_PERCENT",Config!A:B,2,FALSE)/100,0.75)';
  const formula =
    '=IFERROR(SORT(FILTER(' +
    '{TicketData!B:B, LEFT(TicketData!C:C,60), TicketData!L:L, ' +
    'TicketData!AD:AD/60, TicketData!AE:AE/60, ' +
    'TicketData!AE:AE/TicketData!AD:AD, (TicketData!AD:AD-TicketData!AE:AE)/60}, ' +
    '(TicketData!I:I="Open")*' +
    '(TicketData!AD:AD>0)*' +
    '(TicketData!AF:AF<>1)*' +
    '(TicketData!AE:AE/TicketData!AD:AD>=' + riskLookup + ')*' +
    '(TicketData!AE:AE/TicketData!AD:AD<1)' +
    '), 6, FALSE), "No at-risk Response tickets")';

  sheet.getRange('A2').setValue(formula);

  // Format header
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#f44336')
    .setFontColor('white');

  // Format columns
  sheet.getRange('D:E').setNumberFormat('0.0');   // Hours
  sheet.getRange('F:F').setNumberFormat('0.0%');  // % of SLA
  sheet.getRange('G:G').setNumberFormat('0.0');   // Time Remaining

  // Column widths
  sheet.setColumnWidth(1, 120);  // Ticket Number
  sheet.setColumnWidth(2, 300);  // Subject
  sheet.setColumnWidth(3, 150);  // Team

  sheet.setFrozenRows(1);

  // Add note
  sheet.getRange('A1').setNote(
    'Response SLA At-Risk Queue\n\n' +
    'Shows open tickets where Response SLA usage is between\n' +
    'SLA_RISK_PERCENT (default 75%) and 100%.\n\n' +
    'Sorted by % of SLA descending (most urgent first).\n\n' +
    'Response SLA = time to first response/acknowledgment.'
  );

  return true;
}

/**
 * Setup AtRiskResolution sheet with formulas
 * Shows tickets approaching Resolution SLA breach threshold
 * Deletes existing sheet if present for clean slate
 */
function setupAtRiskResolutionSheet(ss) {
  deleteSheetIfExists(ss, 'AtRiskResolution');
  const sheet = ss.insertSheet('AtRiskResolution');

  // Headers
  const headers = ['Ticket Number', 'Subject', 'Team', 'Threshold (hrs)', 'Elapsed (hrs)', '% of SLA', 'Time Remaining (hrs)'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Resolution SLA at-risk formula
  // Shows open tickets where Resolution SLA usage is between risk% and 100%
  const riskLookup = 'IFERROR(VLOOKUP("SLA_RISK_PERCENT",Config!A:B,2,FALSE)/100,0.75)';
  const formula =
    '=IFERROR(SORT(FILTER(' +
    '{TicketData!B:B, LEFT(TicketData!C:C,60), TicketData!L:L, ' +
    'TicketData!AG:AG/60, TicketData!AH:AH/60, ' +
    'TicketData!AH:AH/TicketData!AG:AG, (TicketData!AG:AG-TicketData!AH:AH)/60}, ' +
    '(TicketData!I:I="Open")*' +
    '(TicketData!AG:AG>0)*' +
    '(TicketData!AI:AI<>1)*' +
    '(TicketData!AH:AH/TicketData!AG:AG>=' + riskLookup + ')*' +
    '(TicketData!AH:AH/TicketData!AG:AG<1)' +
    '), 6, FALSE), "No at-risk Resolution tickets")';

  sheet.getRange('A2').setValue(formula);

  // Format header
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#ff5722')
    .setFontColor('white');

  // Format columns
  sheet.getRange('D:E').setNumberFormat('0.0');   // Hours
  sheet.getRange('F:F').setNumberFormat('0.0%');  // % of SLA
  sheet.getRange('G:G').setNumberFormat('0.0');   // Time Remaining

  // Column widths
  sheet.setColumnWidth(1, 120);  // Ticket Number
  sheet.setColumnWidth(2, 300);  // Subject
  sheet.setColumnWidth(3, 150);  // Team

  sheet.setFrozenRows(1);

  // Add note
  sheet.getRange('A1').setNote(
    'Resolution SLA At-Risk Queue\n\n' +
    'Shows open tickets where Resolution SLA usage is between\n' +
    'SLA_RISK_PERCENT (default 75%) and 100%.\n\n' +
    'Sorted by % of SLA descending (most urgent first).\n\n' +
    'Resolution SLA = time to fully resolve/close the ticket.'
  );

  return true;
}

/**
 * @deprecated Use setupAtRiskResponseSheet and setupAtRiskResolutionSheet instead
 * Kept for backward compatibility - creates both new sheets
 */
function setupAtRiskQueueSheet(ss) {
  const createdResponse = setupAtRiskResponseSheet(ss);
  const createdResolution = setupAtRiskResolutionSheet(ss);
  return createdResponse || createdResolution;
}

/**
 * Setup PerformanceTrends sheet with formulas
 * Answers "Are we getting better?" with trending metrics
 * Deletes existing sheet if present for clean slate
 */
function setupPerformanceTrendsSheet(ss) {
  deleteSheetIfExists(ss, 'PerformanceTrends');
  const sheet = ss.insertSheet('PerformanceTrends');

  // Headers
  const headers = ['Month', 'Year', 'Closed', 'Avg Resolution (days)', 'Closure Rate', 'Breach Rate', 'Backlog EOM', '% Aged 30+ EOM'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Generate month rows based on school year configuration
  const config = getConfig();
  const monthRange = getMonthRangeFromData(ss, 'E'); // Use CreatedDate column

  const months = ['January', 'February', 'March', 'April', 'May', 'June',
                  'July', 'August', 'September', 'October', 'November', 'December'];

  const dataRows = [];
  for (const period of monthRange) {
    const rowNum = dataRows.length + 2;
    const year = period.year;
    const m = period.month;

    dataRows.push([
      months[m],
      year,
      // C: Count tickets closed in that month
      `=LET(m, MATCH(A${rowNum}, {"January";"February";"March";"April";"May";"June";"July";"August";"September";"October";"November";"December"}, 0), COUNTIFS(TicketData!H:H, ">="&TEXT(DATE(B${rowNum},m,1), "YYYY-MM-DD"), TicketData!H:H, "<"&TEXT(DATE(B${rowNum},m+1,1), "YYYY-MM-DD")))`,
      // D: Average resolution time in DAYS for tickets closed that month
      `=LET(m, MATCH(A${rowNum}, {"January";"February";"March";"April";"May";"June";"July";"August";"September";"October";"November";"December"}, 0), IFERROR(AVERAGEIFS(TicketData!R:R, TicketData!I:I, "Closed", TicketData!H:H, ">="&TEXT(DATE(B${rowNum},m,1), "YYYY-MM-DD"), TicketData!H:H, "<"&TEXT(DATE(B${rowNum},m+1,1), "YYYY-MM-DD")), "N/A"))`,
      // E: Closure rate (Closed ÷ Created)
      `=LET(m, MATCH(A${rowNum}, {"January";"February";"March";"April";"May";"June";"July";"August";"September";"October";"November";"December"}, 0), created, COUNTIFS(TicketData!E:E, ">="&TEXT(DATE(B${rowNum},m,1), "YYYY-MM-DD"), TicketData!E:E, "<"&TEXT(DATE(B${rowNum},m+1,1), "YYYY-MM-DD")), IF(created>0, C${rowNum}/created, "N/A"))`,
      // F: Breach rate - lookup from SLACompliance
      `=IFERROR(INDEX(SLACompliance!E:E, MATCH(A${rowNum}&B${rowNum}, SLACompliance!A:A&SLACompliance!B:B, 0)), "N/A")`,
      // G: Backlog at end of month - lookup from DailySnapshot
      `=IFERROR(INDEX(DailySnapshot!B:B, MATCH(EOMONTH(DATE(B${rowNum}, MATCH(A${rowNum}, {"January";"February";"March";"April";"May";"June";"July";"August";"September";"October";"November";"December"}, 0), 1), 0), DailySnapshot!A:A, 0)), "No snapshot")`,
      // H: % Aged 30+ at end of month - lookup from DailySnapshot
      `=IFERROR(INDEX(DailySnapshot!D:D, MATCH(EOMONTH(DATE(B${rowNum}, MATCH(A${rowNum}, {"January";"February";"March";"April";"May";"June";"July";"August";"September";"October";"November";"December"}, 0), 1), 0), DailySnapshot!A:A, 0)), "No snapshot")`
    ]);
  }

  if (dataRows.length > 0) {
    sheet.getRange(2, 1, dataRows.length, 8).setValues(dataRows);
  }

  // Format header
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#ff9800')
    .setFontColor('white');

  // Format columns
  sheet.getRange('D:D').setNumberFormat('0.0');   // Avg Resolution
  sheet.getRange('E:E').setNumberFormat('0.0%'); // Closure Rate
  sheet.getRange('F:F').setNumberFormat('0.0%'); // Breach Rate
  sheet.getRange('H:H').setNumberFormat('0.0%'); // % Aged 30+

  // Column widths
  sheet.setColumnWidth(1, 100);  // Month
  sheet.setColumnWidth(4, 150);  // Avg Resolution
  sheet.setColumnWidth(7, 100);  // Backlog EOM
  sheet.setColumnWidth(8, 120);  // % Aged 30+

  sheet.setFrozenRows(1);

  // Add interpretation guide as a note
  sheet.getRange('A1').setNote(
    '"Are We Getting Better?" Interpretation:\n' +
    '• Avg Resolution ↓ = Faster ticket resolution\n' +
    '• Closure Rate ↑ (≥100%) = Keeping up with demand\n' +
    '• Breach Rate ↓ = Better SLA compliance\n' +
    '• Backlog EOM ↓ = Shrinking ticket debt\n' +
    '• % Aged 30+ ↓ = Addressing older tickets\n\n' +
    'Note: Backlog EOM and % Aged 30+ require DailySnapshot data.'
  );

  return true;
}

/**
 * Setup StaleTickets sheet with formulas
 * Shows tickets with no update in X days (from STALE_DAYS config)
 * Deletes existing sheet if present for clean slate
 */
function setupStaleTicketsSheet(ss) {
  deleteSheetIfExists(ss, 'StaleTickets');
  const sheet = ss.insertSheet('StaleTickets');

  // Headers
  const headers = ['Ticket Number', 'Subject', 'Team', 'Days Since Update', 'Last Update', 'Created Date', 'Status'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Single formula that returns all stale tickets, sorted by days since update
  // Uses STALE_DAYS from Config (default 7)
  const staleFormula =
    '=IFERROR(SORT(FILTER({TicketData!B2:B, LEFT(TicketData!C2:C,80), TicketData!L2:L, ' +
    'INT(TODAY()-DATEVALUE(LEFT(TicketData!G2:G,10))), LEFT(TicketData!G2:G,10), ' +
    'LEFT(TicketData!E2:E,10), TicketData!I2:I}, ' +
    '(TicketData!I2:I="Open")*(INT(TODAY()-DATEVALUE(LEFT(TicketData!G2:G,10)))>=' +
    'IFERROR(VLOOKUP("STALE_DAYS",Config!A:B,2,FALSE),7))), 4, FALSE), "No stale tickets found")';

  sheet.getRange('A2').setValue(staleFormula);

  // Format header
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#795548')
    .setFontColor('white');

  // Column widths
  sheet.setColumnWidth(1, 120);  // Ticket Number
  sheet.setColumnWidth(2, 300);  // Subject
  sheet.setColumnWidth(3, 150);  // Team
  sheet.setColumnWidth(4, 130);  // Days Since Update
  sheet.setColumnWidth(5, 100);  // Last Update
  sheet.setColumnWidth(6, 100);  // Created Date
  sheet.setColumnWidth(7, 100);  // Status

  sheet.setFrozenRows(1);

  // Add note
  sheet.getRange('A2').setNote('Shows open tickets with no update in STALE_DAYS (default 7 days). Sorted by days since update descending. Adjust STALE_DAYS in Config sheet to change threshold.');

  return true;
}

/**
 * Reorder sheets for better user experience
 */
function reorderSheets(ss) {
  // Order for default sheets - additional analytics sheets will appear after these
  const preferredOrder = [
    'Instructions',
    'Config',
    'TicketData',
    'Teams',
    'MonthlyVolume',
    'BacklogAging',
    'TeamWorkload',
    'SLACompliance',
    'PerformanceTrends',
    'AtRiskResponse',
    'AtRiskResolution',
    'DailySnapshot',
    'Logs'
  ];

  const sheets = ss.getSheets();

  preferredOrder.forEach((name, idx) => {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(idx + 1);
    }
  });

  // Activate Config sheet
  const configSheet = ss.getSheetByName('Config');
  if (configSheet) {
    ss.setActiveSheet(configSheet);
  }
}

/**
 * Add Setup option to menu (call from onOpen in Menu.gs)
 */
function addSetupMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('iiQ Setup')
    .addItem('Setup Spreadsheet', 'setupSpreadsheet')
    .addItem('Verify Configuration', 'verifyConfiguration')
    .addToUi();
}

// =============================================================================
// CUSTOM FIELD MIGRATION + DISCOVERY
// =============================================================================

/**
 * Non-destructive migration: add custom field config rows to existing Config sheets.
 * Safe to call repeatedly — checks if CUSTOM_FIELD_1 already exists before modifying.
 * Called from: runTicketDataLoader, verifyConfiguration, listAvailableCustomFields.
 */
function migrateConfigForCustomFields() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Config');
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();

  // Ticket custom field rows are migrated separately from location ones: a sheet
  // upgraded before v1.7.0 already has the ticket rows but none of the location
  // rows, so each block checks for its own anchor key.
  let hasTicketRows = false;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === 'CUSTOM_FIELD_1') { hasTicketRows = true; break; }
  }

  if (hasTicketRows) {
    migrateConfigForLocationCustomFields();
    return;
  }

  // Find insertion points by looking for known anchor keys
  let slaRiskRow = -1;       // After Performance Settings
  let ticketLastFetchRow = -1; // After Progress Tracking
  let moduleLoadedRow = -1;  // After Config Lock

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === 'SLA_RISK_PERCENT') slaRiskRow = i + 1; // 1-indexed
    if (data[i][0] === 'TICKET_LAST_FETCH') ticketLastFetchRow = i + 1;
    if (data[i][0] === 'MODULE_LOADED') moduleLoadedRow = i + 1;
  }

  // Insert in reverse order (bottom to top) so row numbers don't shift
  // 1. Config Lock section: add _LOADED keys after MODULE_LOADED
  if (moduleLoadedRow > 0) {
    sheet.insertRowsAfter(moduleLoadedRow, 3);
    sheet.getRange(moduleLoadedRow + 1, 1, 3, 2).setValues([
      ['CUSTOM_FIELD_1_LOADED', ''],
      ['CUSTOM_FIELD_2_LOADED', ''],
      ['CUSTOM_FIELD_3_LOADED', '']
    ]);
  }

  // Re-read data since rows shifted
  const data2 = sheet.getDataRange().getValues();
  ticketLastFetchRow = -1;
  for (let i = 0; i < data2.length; i++) {
    if (data2[i][0] === 'TICKET_LAST_FETCH') ticketLastFetchRow = i + 1;
  }

  // 2. Progress Tracking section: add _ID keys after TICKET_LAST_FETCH
  if (ticketLastFetchRow > 0) {
    sheet.insertRowsAfter(ticketLastFetchRow, 3);
    sheet.getRange(ticketLastFetchRow + 1, 1, 3, 2).setValues([
      ['CUSTOM_FIELD_1_ID', ''],
      ['CUSTOM_FIELD_2_ID', ''],
      ['CUSTOM_FIELD_3_ID', '']
    ]);
  }

  // Re-read data since rows shifted again
  const data3 = sheet.getDataRange().getValues();
  slaRiskRow = -1;
  for (let i = 0; i < data3.length; i++) {
    if (data3[i][0] === 'SLA_RISK_PERCENT') slaRiskRow = i + 1;
  }

  // 3. User-facing section: add CUSTOM_FIELD names after SLA_RISK_PERCENT
  if (slaRiskRow > 0) {
    sheet.insertRowsAfter(slaRiskRow, 5); // blank + header + 3 fields
    sheet.getRange(slaRiskRow + 1, 1, 5, 2).setValues([
      ['', ''],
      ['# Ticket Custom Field Columns (Optional - paste CustomFieldTypeId, up to 3)', ''],
      ['CUSTOM_FIELD_1', ''],
      ['CUSTOM_FIELD_2', ''],
      ['CUSTOM_FIELD_3', '']
    ]);
    // Format section header
    sheet.getRange(slaRiskRow + 2, 1, 1, 2).setFontWeight('bold').setBackground('#e8f0fe');
  }

  logOperation('Config', 'MIGRATED', 'Added custom field configuration rows');

  migrateConfigForLocationCustomFields();
}

/**
 * Non-destructive migration: add LOCATION_CUSTOM_FIELD_* rows to Config sheets.
 * Split out from migrateConfigForCustomFields so sheets already carrying the
 * ticket custom field rows (anything from v1.3.0 onward) still pick these up.
 * Safe to call repeatedly — returns early once LOCATION_CUSTOM_FIELD_1 exists.
 */
function migrateConfigForLocationCustomFields() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Config');
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === 'LOCATION_CUSTOM_FIELD_1') return; // Already migrated
  }

  const N = LOCATION_CUSTOM_FIELD_COUNT;

  // Insert bottom-to-top so earlier row numbers stay valid between blocks.
  // 1. Config Lock section: _LOADED keys after the last CUSTOM_FIELD_*_LOADED
  let anchor = -1;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).indexOf('CUSTOM_FIELD_') === 0 && String(data[i][0]).indexOf('_LOADED') > 0) anchor = i + 1;
  }
  if (anchor > 0) {
    sheet.insertRowsAfter(anchor, N);
    const rows = [];
    for (let n = 1; n <= N; n++) rows.push(['LOCATION_CUSTOM_FIELD_' + n + '_LOADED', '']);
    sheet.getRange(anchor + 1, 1, N, 2).setValues(rows);
  }

  // 2. Progress Tracking section: _ID keys after the last CUSTOM_FIELD_*_ID
  const data2 = sheet.getDataRange().getValues();
  anchor = -1;
  for (let i = 0; i < data2.length; i++) {
    const key = String(data2[i][0]);
    if (key.indexOf('CUSTOM_FIELD_') === 0 && key.indexOf('_ID') === key.length - 3) anchor = i + 1;
  }
  if (anchor > 0) {
    sheet.insertRowsAfter(anchor, N + 1);
    const rows = [];
    for (let n = 1; n <= N; n++) rows.push(['LOCATION_CUSTOM_FIELD_' + n + '_ID', '']);
    rows.push(['LOCATION_CF_TYPE_CACHE', '']);
    sheet.getRange(anchor + 1, 1, N + 1, 2).setValues(rows);
  }

  // 3. User-facing section: names after CUSTOM_FIELD_3 (or the last plain slot)
  const data3 = sheet.getDataRange().getValues();
  anchor = -1;
  for (let i = 0; i < data3.length; i++) {
    if (/^CUSTOM_FIELD_\d+$/.test(String(data3[i][0]))) anchor = i + 1;
  }
  if (anchor > 0) {
    sheet.insertRowsAfter(anchor, N + 2); // blank + header + N fields
    const rows = [['', ''], ['# Location Custom Field Columns (Optional - paste CustomFieldTypeId, up to ' + N + ')', '']];
    for (let n = 1; n <= N; n++) rows.push(['LOCATION_CUSTOM_FIELD_' + n, '']);
    sheet.getRange(anchor + 1, 1, N + 2, 2).setValues(rows);
    sheet.getRange(anchor + 2, 1, 1, 2).setFontWeight('bold').setBackground('#e8f0fe');
  }

  logOperation('Config', 'MIGRATED', 'Added location custom field configuration rows');
}

/**
 * Widen a sheet's grid to at least TICKET_COLUMN_COUNT columns.
 *
 * Reading or writing a fixed TICKET_COLUMN_COUNT-wide range throws
 * "out of bounds" when the grid is narrower, which is exactly the state a sheet
 * upgraded from an earlier column count is in. Called before any full-width
 * range operation on TicketData.
 */
function ensureTicketGridWidth(sheet) {
  if (!sheet) return;
  const maxCols = sheet.getMaxColumns();
  if (maxCols < TICKET_COLUMN_COUNT) {
    sheet.insertColumnsAfter(maxCols, TICKET_COLUMN_COUNT - maxCols);
  }
}

/**
 * Update TicketData header row to include newer columns.
 * Extends older headers to TICKET_COLUMN_COUNT (52) if needed.
 * Handles upgrades from 41-, 44-, 46- and 47-column sheets.
 *
 * newHeaders lists every column appended since the 41-column baseline, in
 * order, so an upgrading sheet takes exactly the tail it is missing.
 *
 * @param {Sheet} sheet - TicketData sheet
 * @param {Object} config - Config object (unused currently, reserved for future dynamic naming)
 */
function updateCustomFieldHeaders(sheet, config) {
  if (!sheet) return;
  ensureTicketGridWidth(sheet);
  const lastCol = sheet.getLastColumn();
  if (lastCol >= TICKET_COLUMN_COUNT) return; // Already has enough columns

  const newHeaders = [
    'AssetId', 'AssetCategory', 'CustomField1', 'CustomField2', 'CustomField3', 'RequesterRole',
    'LocationCustomField1', 'LocationCustomField2', 'LocationCustomField3',
    'LocationCustomField4', 'LocationCustomField5'
  ];
  const colsToAdd = TICKET_COLUMN_COUNT - lastCol;
  if (colsToAdd > 0 && colsToAdd <= newHeaders.length) {
    const headers = newHeaders.slice(newHeaders.length - colsToAdd);
    sheet.getRange(1, lastCol + 1, 1, colsToAdd).setValues([headers]);
    sheet.getRange(1, lastCol + 1, 1, colsToAdd)
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('white')
      .setWrap(false);
  }
}

/**
 * Setup CustomFields sheet with headers
 * This sheet lists available custom field definitions; districts copy a
 * CustomFieldTypeId from column C into the Config sheet
 */
function setupCustomFieldsSheet(ss) {
  deleteSheetIfExists(ss, 'CustomFields');
  const sheet = ss.insertSheet('CustomFields');

  const headers = ['Entity', 'Name', 'CustomFieldTypeId', 'EditorType'];
  sheet.getRange(1, 1, 1, 4).setValues([headers]);

  // Format header
  sheet.getRange(1, 1, 1, 4)
    .setFontWeight('bold')
    .setBackground('#7b1fa2')
    .setFontColor('white');

  // Column widths
  sheet.setColumnWidth(1, 90);   // Entity
  sheet.setColumnWidth(2, 250);  // Name
  sheet.setColumnWidth(3, 300);  // CustomFieldTypeId
  sheet.setColumnWidth(4, 130);  // EditorType

  sheet.setFrozenRows(1);

  // Add note explaining the sheet
  sheet.getRange('A1').setNote(
    'Custom Fields Reference\n\n' +
    'This sheet lists all custom fields available in your district, for both\n' +
    'entities that feed TicketData:\n\n' +
    '  Ticket   — set on the ticket itself\n' +
    '  Location — set on the school/building; joined onto each ticket by its location\n\n' +
    'To use a custom field:\n' +
    '1. Run "Refresh Custom Fields" from the menu to populate this sheet\n' +
    '2. Copy the CustomFieldTypeId (column C) of the field you want\n' +
    '3. Paste it into Config: CUSTOM_FIELD_1/2/3 for Ticket fields, or\n' +
    '   LOCATION_CUSTOM_FIELD_1-5 for Location fields\n\n' +
    'Match the Entity column to the slot: a Location field pasted into a Ticket\n' +
    'slot will never resolve, and vice versa.\n\n' +
    'Field names are also accepted, but ids are unambiguous — a district can have\n' +
    'two different fields sharing the same display name.'
  );

  return true;
}

/**
 * Refresh the CustomFields sheet from the API
 * Populates available custom field definitions for both entities (Ticket and
 * Location) and strips any stale list validation from the Config cells
 */
function refreshCustomFields() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Ensure Config rows exist
  migrateConfigForCustomFields();

  // Fetching is separated from sheet-building so a sheet/validation failure is
  // not reported as an API credentials problem, which sent one district chasing
  // the wrong cause.
  let ticketDefs;
  try {
    ticketDefs = getTicketCustomFieldDefinitions();
  } catch (e) {
    ui.alert('Error Fetching Custom Fields',
      'Could not reach the iiQ API: ' + e.message + '\n\nCheck that API_BASE_URL and BEARER_TOKEN are configured correctly in the Config sheet.',
      ui.ButtonSet.OK);
    return;
  }

  // Location fields are optional infrastructure — a district with none should
  // still get its ticket fields listed, so this failure is not fatal.
  let locationDefs = [];
  try {
    locationDefs = getLocationCustomFieldDefinitions();
  } catch (e) {
    logOperation('LocationCustomFields', 'WARNING', 'Could not fetch location custom fields: ' + e.message);
  }

  try {

    if ((!ticketDefs || ticketDefs.length === 0) && (!locationDefs || locationDefs.length === 0)) {
      ui.alert('No Custom Fields', 'No ticket or location custom fields are defined in your district.', ui.ButtonSet.OK);
      return;
    }

    // Create or get CustomFields sheet
    let sheet = ss.getSheetByName('CustomFields');
    if (!sheet) {
      setupCustomFieldsSheet(ss);
      sheet = ss.getSheetByName('CustomFields');
    }

    // Sheets created before v1.7.0 have 3 columns and no Entity column — rebuild
    // rather than try to shift existing data under a new leading column.
    const headerRow = sheet.getRange(1, 1, 1, 4).getValues()[0];
    if (String(headerRow[0]) !== 'Entity') {
      setupCustomFieldsSheet(ss);
      sheet = ss.getSheetByName('CustomFields');
    }

    // Clear existing data (keep header)
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 4).clear();
    }

    const rows = buildCustomFieldRows_('Ticket', ticketDefs)
      .concat(buildCustomFieldRows_('Location', locationDefs));

    // Write data
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 4).setValues(rows);
    }

    // These cells take a pasted CustomFieldTypeId, so strip any list validation
    clearCustomFieldValidations(ss);

    // Refresh the editor type cache while the definitions are in hand. This is
    // the district-facing way to correct a stale type without waiting out the TTL.
    writeLocationCfTypeCache(locationDefs);

    const ticketCount = rows.filter(r => r[0] === 'Ticket').length;
    const locationCount = rows.filter(r => r[0] === 'Location').length;

    ui.alert('Custom Fields Refreshed',
      `Found ${ticketCount} ticket custom field(s) and ${locationCount} location custom field(s).\n\n` +
      'The CustomFields sheet has been updated.\n\n' +
      'To use a field, copy its CustomFieldTypeId (column C) and paste it into the Config sheet:\n' +
      '  Ticket fields  → CUSTOM_FIELD_1/2/3\n' +
      `  Location fields → LOCATION_CUSTOM_FIELD_1-${LOCATION_CUSTOM_FIELD_COUNT}\n\n` +
      'Location field values are joined onto each ticket by the ticket\'s location.\n' +
      'Check the Entity column so a Location field does not land in a Ticket slot.',
      ui.ButtonSet.OK);

  } catch (e) {
    logOperation('CustomFields', 'ERROR', 'Failed to build CustomFields sheet: ' + e.message);
    ui.alert('Error Building CustomFields Sheet',
      'The custom fields were fetched from iiQ successfully, but writing them to the sheet failed:\n\n' + e.message,
      ui.ButtonSet.OK);
  }
}

/**
 * iiQ EditorTypes enum (Spark.Shared/Enums.cs) — only the values a district is
 * likely to see on a ticket or location custom field are named.
 */
const EDITOR_TYPE_LABELS = {
  0: 'None', 1: 'Text', 2: 'MultilineText', 3: 'RichText', 4: 'Number',
  5: 'NumberRange', 6: 'Date', 7: 'DateRange', 8: 'OnOff', 9: 'Select',
  10: 'MultiSelect', 11: 'Email', 12: 'Phone', 13: 'Address', 14: 'FileUpload',
  18: 'IPAddress', 21: 'IiqUser', 22: 'IiqLocation', 23: 'IiqAsset',
  29: 'IiqModel', 33: 'IiqTeam', 35: 'IiqRoom'
};

/**
 * Turn custom field definitions into CustomFields sheet rows, sorted by name.
 *
 * Deduplicated by CustomFieldTypeId. The /custom-fields/for/* endpoints return
 * one Item per field-to-filter-set mapping, not one per field, so a single field
 * can come back dozens of times — same CustomFieldTypeId, same name, differing
 * CustomFieldId and FilterSetId. Listing every mapping made the sheet unusable
 * and pushed districts past the dropdown item cap.
 *
 * Distinct field types that happen to share a display name are kept as separate
 * rows: they are genuinely different fields, and collapsing them would hide one.
 *
 * @param {string} entity - 'Ticket' or 'Location'
 * @param {Array} definitions - CustomFieldDetail objects
 * @returns {Array} - Rows of [entity, name, uuid, editorTypeLabel]
 */
function buildCustomFieldRows_(entity, definitions) {
  const rows = [];
  const seenTypeIds = {};
  const nameCounts = {};
  let mappingCount = 0;

  for (const def of (definitions || [])) {
    const name = (def.CustomFieldType && def.CustomFieldType.Name) || '';
    if (!name) continue; // Skip unnamed fields

    const uuid = def.CustomFieldTypeId || '';
    mappingCount++;
    if (uuid && seenTypeIds[uuid]) continue; // Another mapping of a field already listed
    if (uuid) seenTypeIds[uuid] = true;

    const editorTypeId = def.EditorTypeId || (def.CustomFieldType && def.CustomFieldType.EditorType) || 0;
    const typeLabel = EDITOR_TYPE_LABELS[editorTypeId] || 'Type ' + editorTypeId;

    const key = name.trim().toLowerCase();
    nameCounts[key] = (nameCounts[key] || 0) + 1;

    rows.push([entity, name, uuid, typeLabel]);
  }

  if (mappingCount > rows.length) {
    logOperation('CustomFields', 'DEDUPED',
      `${entity}: collapsed ${mappingCount} field/filter-set mappings into ${rows.length} distinct field(s)`);
  }

  // Config stores a field NAME, so a name shared by two field types cannot be
  // resolved unambiguously — surface it rather than silently picking one.
  const ambiguous = Object.keys(nameCounts).filter(k => nameCounts[k] > 1).length;
  if (ambiguous > 0) {
    logOperation('CustomFields', 'WARNING',
      `${entity}: ${ambiguous} field name(s) are used by more than one field type. Selecting one by name may resolve to either — check the CustomFieldTypeId column.`);
  }

  rows.sort((a, b) => a[1].localeCompare(b[1]));
  return rows;
}

/**
 * Remove any list validation from the Config custom field cells.
 *
 * These cells take a CustomFieldTypeId pasted from the CustomFields sheet, not a
 * selection from a list. A dropdown was the wrong affordance for three reasons:
 * field names are not unique within a district, Sheets caps list validation at
 * 500 items and districts exceed that, and the id is what actually gets stored
 * and used. Names are still accepted for backward compatibility — see
 * looksLikeCustomFieldGuid_ in Config.gs — but the id is the documented input.
 *
 * Called on Refresh Custom Fields so sheets carrying a validation rule from an
 * earlier version get it stripped rather than keeping a stale 3-item list.
 */
function clearCustomFieldValidations(ss) {
  const configSheet = ss.getSheetByName('Config');
  if (!configSheet) return;

  const configData = configSheet.getRange(1, 1, configSheet.getLastRow(), 1).getValues();
  let cleared = 0;
  for (let i = 0; i < configData.length; i++) {
    const key = String(configData[i][0]).trim();
    if (/^CUSTOM_FIELD_[123]$/.test(key) || /^LOCATION_CUSTOM_FIELD_\d+$/.test(key)) {
      const cell = configSheet.getRange(i + 1, 2);
      cell.clearDataValidations();
      cell.setNote('Paste a CustomFieldTypeId from the CustomFields sheet (column C). A field name also works, but ids are unambiguous.');
      cleared++;
    }
  }
  if (cleared > 0) {
    logOperation('CustomFields', 'INFO', `Cleared list validation on ${cleared} custom field config cell(s)`);
  }
}

/**
 * Verify configuration is complete
 */
function verifyConfiguration() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const issues = [];
  const warnings = [];

  // Ensure custom field config rows exist for upgraded sheets
  migrateConfigForCustomFields();

  // Get all sheet names at once (faster than checking each individually)
  const sheetNames = ss.getSheets().map(s => s.getName());

  // Check required sheets first (fast - uses cached sheet names)
  const requiredSheets = ['Config', 'TicketData', 'Teams', 'Logs'];
  requiredSheets.forEach(name => {
    if (!sheetNames.includes(name)) {
      issues.push(`Missing sheet: ${name}`);
    }
  });

  // Only check config values if Config sheet exists
  if (sheetNames.includes('Config')) {
    try {
      const config = getConfig();

      if (!config.baseUrl || config.baseUrl.includes('YOUR-DISTRICT')) {
        issues.push('API_BASE_URL not configured');
      }

      if (!config.bearerToken) {
        issues.push('BEARER_TOKEN not configured');
      }

      // Validate school year configuration
      if (!config.schoolYear) {
        issues.push('SCHOOL_YEAR not configured (e.g., "2025-2026")');
      } else {
        // Validate format: YYYY-YYYY with consecutive years
        const yearMatch = config.schoolYear.match(/^(\d{4})-(\d{4})$/);
        if (!yearMatch) {
          issues.push('SCHOOL_YEAR format invalid. Expected: YYYY-YYYY (e.g., "2025-2026")');
        } else {
          const startYear = parseInt(yearMatch[1], 10);
          const endYear = parseInt(yearMatch[2], 10);
          if (endYear !== startYear + 1) {
            issues.push('SCHOOL_YEAR must be consecutive years (e.g., "2025-2026", not "2025-2027")');
          }
        }
      }

      // Validate school year start format
      if (config.schoolYearStart) {
        const startMatch = config.schoolYearStart.match(/^(\d{2})-(\d{2})$/);
        if (!startMatch) {
          issues.push('SCHOOL_YEAR_START format invalid. Expected: MM-DD (e.g., "07-01")');
        } else {
          const month = parseInt(startMatch[1], 10);
          const day = parseInt(startMatch[2], 10);
          if (month < 1 || month > 12) {
            issues.push('SCHOOL_YEAR_START month must be 01-12');
          }
          if (day < 1 || day > 31) {
            issues.push('SCHOOL_YEAR_START day must be 01-31');
          }
        }
      }

      // Check for config lock mismatches (school year, page size, batch size, custom fields)
      const lockStatus = checkConfigLock(config);
      if (lockStatus.locked && !lockStatus.matches) {
        lockStatus.mismatches.forEach(m => {
          issues.push(
            `${m.key} MISMATCH! Config has "${m.current}" but data was loaded with "${m.locked}". ` +
            `Use "Clear Data + Reset Progress" to change locked values.`
          );
        });
      }

      // Check custom field configuration
      const cfFields = [
        { name: config.customField1, id: config.customField1Id, label: '1' },
        { name: config.customField2, id: config.customField2Id, label: '2' },
        { name: config.customField3, id: config.customField3Id, label: '3' }
      ];
      for (const cf of cfFields) {
        if (cf.name && cf.id === 'NOT_FOUND') {
          warnings.push(`Custom Field ${cf.label} "${cf.name}" was not found in your district. Check spelling or use "Refresh Custom Fields" to see valid names.`);
        } else if (cf.name && !cf.id) {
          warnings.push(`Custom Field ${cf.label} "${cf.name}" will be resolved on the next data load.`);
        }
      }

      // Check location custom field configuration
      for (let i = 0; i < LOCATION_CUSTOM_FIELD_COUNT; i++) {
        const name = config.locationCustomFields[i];
        const id = config.locationCustomFieldIds[i];
        if (name && id === 'NOT_FOUND') {
          warnings.push(`Location Custom Field ${i + 1} "${name}" was not found in your district. Note that Location fields are separate from Ticket fields — use "Refresh Custom Fields" and check the Entity column.`);
        } else if (name && !id) {
          warnings.push(`Location Custom Field ${i + 1} "${name}" will be resolved on the next data load.`);
        }
      }

    } catch (e) {
      issues.push('Config sheet error: ' + e.message);
    }
  }

  if (issues.length === 0) {
    const config = getConfig();
    const dates = getSchoolYearDates(config);
    const isCurrent = isSchoolYearCurrent(config);
    const lockStatus = checkConfigLock(config);
    const dateRange = dates ?
      `${Utilities.formatDate(dates.startDate, Session.getScriptTimeZone(), 'MMM d, yyyy')} - ${Utilities.formatDate(dates.endDate, Session.getScriptTimeZone(), 'MMM d, yyyy')}` :
      'Unknown';

    let lockMessage = '';
    if (lockStatus.locked) {
      lockMessage = `\n\nConfig Lock: LOCKED (data loading started)\n` +
        `  SCHOOL_YEAR: ${config.schoolYearLoaded}\n` +
        `  PAGE_SIZE: ${config.pageSizeLoaded}\n` +
        `  BATCH_SIZE: ${config.batchSizeLoaded}\n` +
        `  MODULE: ${config.moduleLoaded}\n` +
        `To change these values, use "Clear Data + Reset Progress"`;
    } else {
      lockMessage = `\n\nConfig Lock: Unlocked (can change settings)`;
    }

    // Custom fields summary
    let cfMessage = '\n\nCustom Fields:';
    const cfFields = [
      { name: config.customField1, id: config.customField1Id, label: '1' },
      { name: config.customField2, id: config.customField2Id, label: '2' },
      { name: config.customField3, id: config.customField3Id, label: '3' }
    ];
    for (const cf of cfFields) {
      if (!cf.name) {
        cfMessage += `\n  ${cf.label}: (not configured)`;
      } else if (cf.id === 'NOT_FOUND') {
        cfMessage += `\n  ${cf.label}: ${cf.name} (not found)`;
      } else if (cf.id) {
        cfMessage += `\n  ${cf.label}: ${cf.name} (resolved)`;
      } else {
        cfMessage += `\n  ${cf.label}: ${cf.name} (pending resolution)`;
      }
    }

    // Location custom fields summary — only listed when at least one is set, so
    // districts not using them don't see five "(not configured)" lines.
    const anyLocationFields = config.locationCustomFields.some(n => n);
    if (anyLocationFields) {
      cfMessage += '\n\nLocation Custom Fields (joined by ticket location):';
      for (let i = 0; i < LOCATION_CUSTOM_FIELD_COUNT; i++) {
        const name = config.locationCustomFields[i];
        const id = config.locationCustomFieldIds[i];
        if (!name) continue;
        if (id === 'NOT_FOUND') {
          cfMessage += `\n  ${i + 1}: ${name} (not found)`;
        } else if (id) {
          cfMessage += `\n  ${i + 1}: ${name} (resolved)`;
        } else {
          cfMessage += `\n  ${i + 1}: ${name} (pending resolution)`;
        }
      }
      const cache = config.locationCfTypeCache;
      if (isLocationCfTypeCacheFresh(config)) {
        const ageDays = Math.floor((Date.now() - new Date(cache.cached).getTime()) / 86400000);
        cfMessage += `\n  Editor type cache: valid (${ageDays}d old, refreshes after ${LOCATION_CF_TYPE_CACHE_DAYS}d)`;
      } else {
        cfMessage += `\n  Editor type cache: will refresh on next data load`;
      }
    }

    let warningMessage = '';
    if (warnings.length > 0) {
      warningMessage = `\n\nWarnings:\n- ${warnings.join('\n- ')}`;
    }

    ui.alert('Configuration Valid',
      `All required settings are configured.\n\n` +
      `School Year: ${config.schoolYear}\n` +
      `Module: ${config.module}\n` +
      `Date Range: ${dateRange}\n` +
      `Status: ${isCurrent ? 'Current (incremental updates enabled)' : 'Historical'}` +
      lockMessage + cfMessage + warningMessage + `\n\n` +
      `You can now run "Continue Loading" to start loading data.`,
      ui.ButtonSet.OK);
  } else {
    let warningMessage = '';
    if (warnings.length > 0) {
      warningMessage = '\n\nWarnings:\n- ' + warnings.join('\n- ');
    }
    ui.alert('Configuration Issues', 'Please fix the following issues:\n\n- ' + issues.join('\n- ') + warningMessage, ui.ButtonSet.OK);
  }
}
