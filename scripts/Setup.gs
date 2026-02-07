/**
 * Setup Script - Initial Spreadsheet Configuration
 *
 * Creates required data sheets and default analytics sheets.
 * Running setup provides a CLEAN SLATE - existing sheets are overwritten.
 *
 * Data Sheets (always created):
 * - Instructions: Setup and usage guide
 * - Config: API settings and progress tracking
 * - TicketData: Main data (39 columns including SLA metrics and device/asset)
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
    '- TicketData (39 columns) - ALL DATA WILL BE LOST!\n' +
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
  setupDailySnapshotSheet(ss); created.push('DailySnapshot');
  setupLogsSheet(ss); created.push('Logs');

  // Create default analytics sheets
  // Additional sheets can be added via iiQ Data > Add Analytics Sheet menu
  setupMonthlyVolumeSheet(ss); created.push('MonthlyVolume');
  setupBacklogAgingSheet(ss); created.push('BacklogAging');
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
  message.push('4. Run "Continue Loading" to start loading ticket data');

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
    ['and Power BI consumption. Data is loaded via Google Apps Script and refreshed'],
    ['automatically via time-driven triggers.'],
    [''],
    ['SCHOOL YEAR MODEL: Each spreadsheet contains ONE school year\'s data.'],
    ['Example: A "2025-2026" spreadsheet covers July 1, 2025 through June 30, 2026.'],
    ['Create a new spreadsheet for each school year.'],
    [''],
    ['Data Flow: iiQ API → Google Apps Script → This Spreadsheet → Power BI'],
    [''],
    [''],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    ['INITIAL SETUP'],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    [''],
    ['1. CONFIGURE API CREDENTIALS (Config sheet)'],
    ['   • API_BASE_URL: Your iiQ instance (e.g., https://district.incidentiq.com)'],
    ['   • BEARER_TOKEN: JWT token from iiQ (Admin > Integrations > API)'],
    ['   • SITE_ID: Optional - only needed for multi-site instances'],
    [''],
    ['2. CONFIGURE SCHOOL YEAR (Config sheet)'],
    ['   • SCHOOL_YEAR: The school year for this spreadsheet (e.g., "2025-2026")'],
    ['   • SCHOOL_YEAR_START: First day of school year in MM-DD format (default "07-01")'],
    [''],
    ['   Examples:'],
    ['   • Traditional: SCHOOL_YEAR=2025-2026, SCHOOL_YEAR_START=07-01 (July 1 - June 30)'],
    ['   • Fall start:  SCHOOL_YEAR=2025-2026, SCHOOL_YEAR_START=08-15 (Aug 15 - Aug 14)'],
    [''],
    ['3. VERIFY CONFIGURATION'],
    ['   • Menu: iiQ Data > Setup > Verify Configuration'],
    ['   • Fixes any issues reported before proceeding'],
    [''],
    ['4. LOAD TEAM DIRECTORY'],
    ['   • Menu: iiQ Data > Refresh Teams'],
    ['   • Then fill in FunctionalArea column for each team (optional)'],
    [''],
    ['5. START LOADING TICKET DATA'],
    ['   • Menu: iiQ Data > Ticket Data > Continue Loading'],
    ['   • Script runs for ~5.5 minutes then pauses'],
    ['   • Run again to continue (or set up triggers for automation)'],
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
    ['• TicketData (39 columns)'],
    ['  Main ticket data including SLA metrics. Columns include:'],
    ['  - Ticket info: ID, Number, Subject, Status, Priority'],
    ['  - Dates: Created, Modified, Closed'],
    ['  - Assignment: Team, Owner, Location'],
    ['  - Issue: Category, Type'],
    ['  - SLA: Thresholds, Actuals, Breach status (Response & Resolution)'],
    [''],
    ['• Teams'],
    ['  Team directory loaded from iiQ. Add FunctionalArea values for grouping.'],
    [''],
    ['• DailySnapshot'],
    ['  Daily backlog metrics captured at 7 PM. Used for trending analysis.'],
    ['  Cannot be calculated retroactively - daily capture is essential.'],
    [''],
    ['• Logs'],
    ['  Operation logs for troubleshooting. Auto-pruned to 500 rows.'],
    [''],
    [''],
    ['ANALYTICS SHEETS (formula-based, auto-calculate):'],
    [''],
    ['• SLACompliance'],
    ['  Monthly SLA breach rates and average response/resolution times.'],
    ['  Refresh via: iiQ Data > Analytics > Refresh SLA Compliance'],
    [''],
    ['• MonthlyVolume'],
    ['  Tickets created vs closed by month. Shows net backlog change.'],
    ['  Green = backlog decreased, Red = backlog increased.'],
    [''],
    ['• BacklogAging'],
    ['  Distribution of current open tickets by age (0-7 days, 8-14, etc.)'],
    [''],
    ['• TeamWorkload'],
    ['  Open tickets and aging metrics by team.'],
    ['  Note: Copy formulas down after Teams sheet is populated.'],
    [''],
    [''],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    ['MENU REFERENCE'],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    [''],
    ['iiQ Data > Setup'],
    ['  • Setup Spreadsheet - Create all sheets (safe to run multiple times)'],
    ['  • Verify Configuration - Check API settings'],
    [''],
    ['iiQ Data > Refresh Teams'],
    ['  Load/refresh team directory from iiQ'],
    [''],
    ['iiQ Data > Ticket Data'],
    ['  • Continue Loading (Initial) - Resume initial data load (runs ~5.5 min)'],
    ['  • Open Ticket Refresh (Start) - Begin fresh open ticket refresh'],
    ['  • Open Ticket Refresh (Continue) - Resume if refresh timed out'],
    ['  • Open Refresh Status - Show open refresh progress'],
    ['  • Full Reload - Clear ALL data and start fresh (use sparingly)'],
    ['  • Show Status - Display loading progress by year'],
    [''],
    ['iiQ Data > Analytics'],
    ['  • Refresh SLA Compliance - Recalculate SLA metrics'],
    [''],
    ['iiQ Data > Daily Snapshot'],
    ['  • Capture Now - Record today\'s backlog metrics'],
    ['  • Populate Historical Estimates - Create estimated past snapshots'],
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
    ['  → Each batch processes ~2000 tickets, large years take multiple runs'],
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
    [''],
    [''],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    ['POWER BI INTEGRATION'],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    [''],
    ['To connect Power BI to this spreadsheet:'],
    [''],
    ['1. In Power BI Desktop: Get Data > Web'],
    ['2. Use the shareable link for each sheet (File > Share > Publish to web)'],
    ['3. Or use the Google Sheets connector if available'],
    ['4. Set up scheduled refresh in Power BI Service'],
    [''],
    ['Recommended refresh schedule: Data is refreshed every 2 hours, so Power BI can'],
    ['refresh more frequently if needed. Weekly full refresh completes by ~4 AM Sunday.'],
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

  // Format section headers
  const sectionRows = [4, 14, 49, 63, 100, 126, 143, 155];
  sectionRows.forEach(row => {
    if (row <= content.length) {
      sheet.getRange(row, 1).setFontWeight('bold').setFontColor('#1a73e8');
    }
  });

  // Format divider lines
  const dividerRows = [3, 13, 48, 62, 99, 125, 142, 154];
  dividerRows.forEach(row => {
    if (row <= content.length) {
      sheet.getRange(row, 1).setFontColor('#dadce0');
    }
  });

  // Set text wrapping for long content
  sheet.getRange(1, 1, content.length, 1).setWrap(true);

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
    ['Key', 'Value'],
    ['', ''],
    ['# API Configuration (Required)', ''],
    ['API_BASE_URL', 'https://YOUR-DISTRICT.incidentiq.com'],
    ['BEARER_TOKEN', ''],
    ['SITE_ID', ''],
    ['', ''],
    ['# School Year Configuration', ''],
    ['SCHOOL_YEAR', defaultSchoolYear],
    ['SCHOOL_YEAR_START', '07-01'],
    ['', ''],
    ['# Performance Settings (Optional)', ''],
    ['PAGE_SIZE', '100'],
    ['THROTTLE_MS', '1000'],
    ['TICKET_BATCH_SIZE', '2000'],
    ['STALE_DAYS', '7'],
    ['SLA_RISK_PERCENT', '75'],
    ['', ''],
    ['# Progress Tracking - Managed Automatically', ''],
    ['TICKET_TOTAL_PAGES', ''],
    ['TICKET_LAST_PAGE', '-1'],
    ['TICKET_COMPLETE', 'FALSE'],
    ['TICKET_LAST_FETCH', ''],
    ['', ''],
    ['# Config Lock - Set when loading starts, cleared by "Clear Data + Reset"', ''],
    ['SCHOOL_YEAR_LOADED', ''],
    ['PAGE_SIZE_LOADED', ''],
    ['BATCH_SIZE_LOADED', ''],
    ['', ''],
    ['LAST_REFRESH', '']
  ];

  sheet.getRange(1, 1, configData.length, 2).setValues(configData);

  // Format header
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');

  // Format section headers (rows starting with #)
  const sectionRows = [3, 8, 12, 19, 25];
  sectionRows.forEach(row => {
    sheet.getRange(row, 1, 1, 2).setFontWeight('bold').setBackground('#e8f0fe');
  });

  // Column widths
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 400);

  sheet.setFrozenRows(1);

  return true;
}

/**
 * Setup TicketData sheet with 39-column header
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
    'AssetTag', 'ModelName', 'SerialNumber'
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
      '=IFERROR(INDEX(SORT(FILTER({TicketData!B2:B,TicketData!Q2:Q},(TicketData!I2:I="Open")*(TicketData!Q2:Q>=0)*(TicketData!Q2:Q<=15)),2,FALSE),1,1),"")',
      '=IFERROR(VLOOKUP("LAST_REFRESH", Config!A:B, 2, FALSE), "")'
    ],
    // Row 3: 16-30 days
    [
      '16-30 days',
      '=COUNTIFS(TicketData!I:I, "Open", TicketData!R:R, ">=16", TicketData!R:R, "<=30")',
      '=IF($B$7>0, B3/$B$7, 0)',
      '=IFERROR(INDEX(SORT(FILTER({TicketData!B2:B,TicketData!Q2:Q},(TicketData!I2:I="Open")*(TicketData!Q2:Q>=16)*(TicketData!Q2:Q<=30)),2,FALSE),1,1),"")',
      '=IFERROR(VLOOKUP("LAST_REFRESH", Config!A:B, 2, FALSE), "")'
    ],
    // Row 4: 31-60 days
    [
      '31-60 days',
      '=COUNTIFS(TicketData!I:I, "Open", TicketData!R:R, ">=31", TicketData!R:R, "<=60")',
      '=IF($B$7>0, B4/$B$7, 0)',
      '=IFERROR(INDEX(SORT(FILTER({TicketData!B2:B,TicketData!Q2:Q},(TicketData!I2:I="Open")*(TicketData!Q2:Q>=31)*(TicketData!Q2:Q<=60)),2,FALSE),1,1),"")',
      '=IFERROR(VLOOKUP("LAST_REFRESH", Config!A:B, 2, FALSE), "")'
    ],
    // Row 5: 61-90 days
    [
      '61-90 days',
      '=COUNTIFS(TicketData!I:I, "Open", TicketData!R:R, ">=61", TicketData!R:R, "<=90")',
      '=IF($B$7>0, B5/$B$7, 0)',
      '=IFERROR(INDEX(SORT(FILTER({TicketData!B2:B,TicketData!Q2:Q},(TicketData!I2:I="Open")*(TicketData!Q2:Q>=61)*(TicketData!Q2:Q<=90)),2,FALSE),1,1),"")',
      '=IFERROR(VLOOKUP("LAST_REFRESH", Config!A:B, 2, FALSE), "")'
    ],
    // Row 6: 90+ days
    [
      '90+ days',
      '=COUNTIFS(TicketData!I:I, "Open", TicketData!R:R, ">90")',
      '=IF($B$7>0, B6/$B$7, 0)',
      '=IFERROR(INDEX(SORT(FILTER({TicketData!B2:B,TicketData!Q2:Q},(TicketData!I2:I="Open")*(TicketData!Q2:Q>90)),2,FALSE),1,1),"")',
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
    'mtdStart, TEXT(DATE(YEAR(TODAY()),MONTH(TODAY()),1), "YYYY-MM-DD"),' +
    'mtdEnd, TEXT(DATE(YEAR(TODAY()),MONTH(TODAY())+1,1), "YYYY-MM-DD"),' +
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
    'mtdStart, TEXT(DATE(YEAR(TODAY()),MONTH(TODAY()),1), "YYYY-MM-DD"),' +
    'mtdEnd, TEXT(DATE(YEAR(TODAY()),MONTH(TODAY())+1,1), "YYYY-MM-DD"),' +
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
    'INT(TODAY()-DATEVALUE(LEFT(TicketData!F2:F,10))), LEFT(TicketData!F2:F,10), ' +
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

/**
 * Verify configuration is complete
 */
function verifyConfiguration() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const issues = [];

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

      // Check for config lock mismatches (school year, page size, batch size)
      const lockStatus = checkConfigLock(config);
      if (lockStatus.locked && !lockStatus.matches) {
        lockStatus.mismatches.forEach(m => {
          issues.push(
            `${m.key} MISMATCH! Config has "${m.current}" but data was loaded with "${m.locked}". ` +
            `Use "Clear Data + Reset Progress" to change locked values.`
          );
        });
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
        `To change these values, use "Clear Data + Reset Progress"`;
    } else {
      lockMessage = `\n\nConfig Lock: Unlocked (can change settings)`;
    }

    ui.alert('Configuration Valid',
      `All required settings are configured.\n\n` +
      `School Year: ${config.schoolYear}\n` +
      `Date Range: ${dateRange}\n` +
      `Status: ${isCurrent ? 'Current (incremental updates enabled)' : 'Historical'}` +
      lockMessage + `\n\n` +
      `You can now run "Continue Loading" to start loading data.`,
      ui.ButtonSet.OK);
  } else {
    ui.alert('Configuration Issues', 'Please fix the following issues:\n\n- ' + issues.join('\n- '), ui.ButtonSet.OK);
  }
}
