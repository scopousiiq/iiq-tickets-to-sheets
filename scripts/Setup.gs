/**
 * Setup Script - Initial Spreadsheet Configuration
 *
 * Creates required data sheets and default analytics sheets.
 * Safe to run multiple times - only creates missing sheets/headers.
 *
 * Data Sheets (always created):
 * - Instructions: Setup and usage guide
 * - Config: API settings and progress tracking
 * - TicketData: Main data (35 columns including SLA metrics)
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
 * Main setup function - creates all sheets and configurations
 */
function setupSpreadsheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const response = ui.alert(
    'Setup iiQ Data Spreadsheet',
    'This will create/configure the following sheets:\n\n' +
    'DATA SHEETS:\n' +
    '- Instructions (setup guide)\n' +
    '- Config (API settings)\n' +
    '- TicketData (35 columns)\n' +
    '- Teams (directory)\n' +
    '- DailySnapshot (trending)\n' +
    '- Logs (operations)\n\n' +
    'DEFAULT ANALYTICS:\n' +
    '- MonthlyVolume\n' +
    '- BacklogAging\n' +
    '- TeamWorkload\n' +
    '- SLACompliance\n' +
    '- PerformanceTrends\n' +
    '- AtRiskResponse\n' +
    '- AtRiskResolution\n\n' +
    'Additional analytics sheets can be added later via:\n' +
    'iiQ Data > Add Analytics Sheet\n\n' +
    'Existing sheets will not be overwritten.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  let created = [];
  let skipped = [];

  // Create data sheets (required)
  if (setupInstructionsSheet(ss)) created.push('Instructions'); else skipped.push('Instructions');
  if (setupConfigSheet(ss)) created.push('Config'); else skipped.push('Config');
  if (setupTicketDataSheet(ss)) created.push('TicketData'); else skipped.push('TicketData');
  if (setupTeamsSheet(ss)) created.push('Teams'); else skipped.push('Teams');
  if (setupDailySnapshotSheet(ss)) created.push('DailySnapshot'); else skipped.push('DailySnapshot');
  if (setupLogsSheet(ss)) created.push('Logs'); else skipped.push('Logs');

  // Create default analytics sheets
  // Additional sheets can be added via iiQ Data > Add Analytics Sheet menu
  if (setupMonthlyVolumeSheet(ss)) created.push('MonthlyVolume'); else skipped.push('MonthlyVolume');
  if (setupBacklogAgingSheet(ss)) created.push('BacklogAging'); else skipped.push('BacklogAging');
  if (setupTeamWorkloadSheet(ss)) created.push('TeamWorkload'); else skipped.push('TeamWorkload');
  if (setupSLAComplianceSheet(ss)) created.push('SLACompliance'); else skipped.push('SLACompliance');
  if (setupPerformanceTrendsSheet(ss)) created.push('PerformanceTrends'); else skipped.push('PerformanceTrends');
  if (setupAtRiskResponseSheet(ss)) created.push('AtRiskResponse'); else skipped.push('AtRiskResponse');
  if (setupAtRiskResolutionSheet(ss)) created.push('AtRiskResolution'); else skipped.push('AtRiskResolution');

  // Reorder sheets for better UX
  reorderSheets(ss);

  const message = [];
  if (created.length > 0) message.push('Created: ' + created.join(', '));
  if (skipped.length > 0) message.push('Already existed: ' + skipped.join(', '));
  message.push('\nNext steps:');
  message.push('1. Fill in Config sheet with API credentials');
  message.push('2. Run "Refresh Teams" to load team directory');
  message.push('3. Run "Continue Loading" to start loading ticket data');

  ui.alert('Setup Complete', message.join('\n'), ui.ButtonSet.OK);
}

/**
 * Setup Instructions sheet with setup and usage guide
 * @returns {boolean} true if created, false if already exists
 */
function setupInstructionsSheet(ss) {
  if (ss.getSheetByName('Instructions')) return false;

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
    ['2. CONFIGURE YEAR TRACKING (Config sheet)'],
    ['   For historical years (full reload via pagination):'],
    ['   • Add row: TICKET_2024_LAST_PAGE = -1'],
    ['   • Add row: TICKET_2024_COMPLETE = FALSE'],
    [''],
    ['   For current year (incremental updates via date windowing):'],
    ['   • Add row: TICKET_2025_LAST_FETCH = (leave empty)'],
    [''],
    ['   Example for loading 2023-2025 data with 2025 as current year:'],
    ['   • TICKET_2023_LAST_PAGE, TICKET_2023_COMPLETE'],
    ['   • TICKET_2024_LAST_PAGE, TICKET_2024_COMPLETE'],
    ['   • TICKET_2025_LAST_FETCH'],
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
    ['• TicketData (35 columns)'],
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
 * @returns {boolean} true if created, false if already exists
 */
function setupConfigSheet(ss) {
  if (ss.getSheetByName('Config')) return false;

  const sheet = ss.insertSheet('Config');

  // Headers and initial config values
  const configData = [
    ['Key', 'Value'],
    ['', ''],
    ['# API Configuration (Required)', ''],
    ['API_BASE_URL', 'https://YOUR-DISTRICT.incidentiq.com'],
    ['BEARER_TOKEN', ''],
    ['SITE_ID', ''],
    ['', ''],
    ['# Performance Settings (Optional)', ''],
    ['PAGE_SIZE', '100'],
    ['THROTTLE_MS', '1000'],
    ['TICKET_BATCH_SIZE', '2000'],
    ['STALE_DAYS', '7'],
    ['SLA_RISK_PERCENT', '75'],
    ['', ''],
    ['# Progress Tracking - Managed Automatically', ''],
    ['# Add TICKET_{YEAR}_LAST_PAGE rows for historical years', ''],
    ['# Add TICKET_{YEAR}_LAST_FETCH row for current year', ''],
    ['', ''],
    ['# Example for 2024 historical + 2025 current:', ''],
    ['TICKET_2024_LAST_PAGE', '-1'],
    ['TICKET_2024_COMPLETE', 'FALSE'],
    ['TICKET_2025_LAST_PAGE', '-1'],
    ['TICKET_2025_COMPLETE', 'FALSE'],
    ['TICKET_2026_LAST_FETCH', ''],
    ['', ''],
    ['LAST_REFRESH', '']
  ];

  sheet.getRange(1, 1, configData.length, 2).setValues(configData);

  // Format header
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');

  // Format section headers (rows starting with #)
  const sectionRows = [3, 8, 15, 19];
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
 * Setup TicketData sheet with 35-column header
 * @returns {boolean} true if created, false if already exists
 */
function setupTicketDataSheet(ss) {
  if (ss.getSheetByName('TicketData')) return false;

  const sheet = ss.insertSheet('TicketData');

  const headers = [
    'TicketId', 'TicketNumber', 'Subject', 'Year',
    'CreatedDate', 'ModifiedDate', 'ClosedDate', 'IsClosed',
    'Status', 'TeamId', 'TeamName', 'LocationId', 'LocationName', 'LocationType',
    'OwnerId', 'OwnerName', 'AgeDays', 'Priority', 'IsPastDue', 'DueDate',
    'SlaId', 'SlaName', 'IssueCategoryId', 'IssueCategoryName',
    'IssueTypeId', 'IssueTypeName', 'RequesterId', 'RequesterName',
    'ResponseThreshold', 'ResponseActual', 'ResponseBreach',
    'ResolutionThreshold', 'ResolutionActual', 'ResolutionBreach', 'IsRunning'
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
 * @returns {boolean} true if created, false if already exists
 */
function setupTeamsSheet(ss) {
  if (ss.getSheetByName('Teams')) return false;

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
 * @returns {boolean} true if created, false if already exists
 */
function setupDailySnapshotSheet(ss) {
  if (ss.getSheetByName('DailySnapshot')) return false;

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
 * @returns {boolean} true if created, false if already exists
 */
function setupLogsSheet(ss) {
  if (ss.getSheetByName('Logs')) return false;

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
 * @returns {boolean} true if created, false if already exists
 */
function setupSLAComplianceSheet(ss) {
  if (ss.getSheetByName('SLACompliance')) return false;

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
      `=COUNTIFS(TicketData!H:H, "Yes", TicketData!G:G, ">="&${startDate}, TicketData!G:G, "<"&${endDate})`,
      // Breaches: Count where ResponseBreach=TRUE OR ResolutionBreach=TRUE (closed tickets in this month)
      `=LET(startDate, ${startDate}, endDate, ${endDate}, COUNTIFS(TicketData!H:H, "Yes", TicketData!G:G, ">="&startDate, TicketData!G:G, "<"&endDate, TicketData!AE:AE, TRUE) + COUNTIFS(TicketData!H:H, "Yes", TicketData!G:G, ">="&startDate, TicketData!G:G, "<"&endDate, TicketData!AH:AH, TRUE) - COUNTIFS(TicketData!H:H, "Yes", TicketData!G:G, ">="&startDate, TicketData!G:G, "<"&endDate, TicketData!AE:AE, TRUE, TicketData!AH:AH, TRUE))`,
      // Breach Rate
      `=IF(C${rowNum}>0, D${rowNum}/C${rowNum}, "N/A")`,
      // Avg Response (hrs): Average of ResponseActual (col AD) for closed tickets, convert minutes to hours
      `=IFERROR(AVERAGEIFS(TicketData!AD:AD, TicketData!H:H, "Yes", TicketData!G:G, ">="&${startDate}, TicketData!G:G, "<"&${endDate}, TicketData!AD:AD, ">0")/60, "N/A")`,
      // Avg Resolution (hrs): Average of ResolutionActual (col AG) for closed tickets, convert minutes to hours
      `=IFERROR(AVERAGEIFS(TicketData!AG:AG, TicketData!H:H, "Yes", TicketData!G:G, ">="&${startDate}, TicketData!G:G, "<"&${endDate}, TicketData!AG:AG, ">0")/60, "N/A")`
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
 * @returns {boolean} true if created, false if already exists
 */
function setupMonthlyVolumeSheet(ss) {
  if (ss.getSheetByName('MonthlyVolume')) return false;

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
      `=COUNTIFS(TicketData!G:G, ">="&${startDate}, TicketData!G:G, "<"&${endDate})`,
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

  sheet.setFrozenRows(1);

  return true;
}

/**
 * Setup BacklogAging sheet with formulas
 * Matches GUIDE.md structure with Sample Ticket and Last Refreshed columns
 * @returns {boolean} true if created, false if already exists
 */
function setupBacklogAgingSheet(ss) {
  if (ss.getSheetByName('BacklogAging')) return false;

  const sheet = ss.insertSheet('BacklogAging');

  // Headers matching GUIDE.md
  const headers = ['Age Bucket', 'Count', '% of Total', 'Sample Ticket', 'Last Refreshed'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Age buckets matching GUIDE.md: 0-15, 16-30, 31-60, 61-90, 90+
  const bucketData = [
    // Row 2: 0-15 days
    [
      '0-15 days',
      '=COUNTIFS(TicketData!H:H, "No", TicketData!Q:Q, ">=0", TicketData!Q:Q, "<=15")',
      '=IF($B$7>0, B2/$B$7, 0)',
      '=IFERROR(INDEX(SORT(FILTER({TicketData!B2:B,TicketData!Q2:Q},(TicketData!H2:H="No")*(TicketData!Q2:Q>=0)*(TicketData!Q2:Q<=15)),2,FALSE),1,1),"")',
      '=IFERROR(VLOOKUP("LAST_REFRESH", Config!A:B, 2, FALSE), "")'
    ],
    // Row 3: 16-30 days
    [
      '16-30 days',
      '=COUNTIFS(TicketData!H:H, "No", TicketData!Q:Q, ">=16", TicketData!Q:Q, "<=30")',
      '=IF($B$7>0, B3/$B$7, 0)',
      '=IFERROR(INDEX(SORT(FILTER({TicketData!B2:B,TicketData!Q2:Q},(TicketData!H2:H="No")*(TicketData!Q2:Q>=16)*(TicketData!Q2:Q<=30)),2,FALSE),1,1),"")',
      '=IFERROR(VLOOKUP("LAST_REFRESH", Config!A:B, 2, FALSE), "")'
    ],
    // Row 4: 31-60 days
    [
      '31-60 days',
      '=COUNTIFS(TicketData!H:H, "No", TicketData!Q:Q, ">=31", TicketData!Q:Q, "<=60")',
      '=IF($B$7>0, B4/$B$7, 0)',
      '=IFERROR(INDEX(SORT(FILTER({TicketData!B2:B,TicketData!Q2:Q},(TicketData!H2:H="No")*(TicketData!Q2:Q>=31)*(TicketData!Q2:Q<=60)),2,FALSE),1,1),"")',
      '=IFERROR(VLOOKUP("LAST_REFRESH", Config!A:B, 2, FALSE), "")'
    ],
    // Row 5: 61-90 days
    [
      '61-90 days',
      '=COUNTIFS(TicketData!H:H, "No", TicketData!Q:Q, ">=61", TicketData!Q:Q, "<=90")',
      '=IF($B$7>0, B5/$B$7, 0)',
      '=IFERROR(INDEX(SORT(FILTER({TicketData!B2:B,TicketData!Q2:Q},(TicketData!H2:H="No")*(TicketData!Q2:Q>=61)*(TicketData!Q2:Q<=90)),2,FALSE),1,1),"")',
      '=IFERROR(VLOOKUP("LAST_REFRESH", Config!A:B, 2, FALSE), "")'
    ],
    // Row 6: 90+ days
    [
      '90+ days',
      '=COUNTIFS(TicketData!H:H, "No", TicketData!Q:Q, ">90")',
      '=IF($B$7>0, B6/$B$7, 0)',
      '=IFERROR(INDEX(SORT(FILTER({TicketData!B2:B,TicketData!Q2:Q},(TicketData!H2:H="No")*(TicketData!Q2:Q>90)),2,FALSE),1,1),"")',
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
 * @returns {boolean} true if created, false if already exists
 */
function setupTeamWorkloadSheet(ss) {
  if (ss.getSheetByName('TeamWorkload')) return false;

  const sheet = ss.insertSheet('TeamWorkload');

  // Headers matching GUIDE.md structure
  const headers = ['Team Name', 'Functional Area', 'Open', 'Created (MTD)', 'Closed (MTD)', 'Aged 30+', 'Last Refreshed', 'Sort Col#', 'Desc?'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Single LET formula that outputs the entire sortable table
  // This populates columns A-F dynamically based on TicketData
  const mainFormula =
    '=LET(' +
    'teams, UNIQUE(FILTER(TicketData!K2:K, TicketData!K2:K<>"", TicketData!K2:K<>"TeamName")),' +
    'mtdStart, TEXT(DATE(YEAR(TODAY()),MONTH(TODAY()),1), "YYYY-MM-DD"),' +
    'mtdEnd, TEXT(DATE(YEAR(TODAY()),MONTH(TODAY())+1,1), "YYYY-MM-DD"),' +
    'col_a, teams,' +
    'col_b, BYROW(teams, LAMBDA(t, IFERROR(VLOOKUP(t, Teams!B:C, 2, FALSE), ""))),' +
    'col_c, BYROW(teams, LAMBDA(t, COUNTIFS(TicketData!K:K, t, TicketData!H:H, "No"))),' +
    'col_d, BYROW(teams, LAMBDA(t, COUNTIFS(TicketData!K:K, t, TicketData!E:E, ">="&mtdStart, TicketData!E:E, "<"&mtdEnd))),' +
    'col_e, BYROW(teams, LAMBDA(t, COUNTIFS(TicketData!K:K, t, TicketData!G:G, ">="&mtdStart, TicketData!G:G, "<"&mtdEnd))),' +
    'col_f, BYROW(teams, LAMBDA(t, COUNTIFS(TicketData!K:K, t, TicketData!H:H, "No", TicketData!Q:Q, ">=30"))),' +
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
 * @returns {boolean} true if created, false if already exists
 */
function setupLocationBreakdownSheet(ss) {
  if (ss.getSheetByName('LocationBreakdown')) return false;

  const sheet = ss.insertSheet('LocationBreakdown');

  // Headers - includes sort controls
  const headers = ['Location Name', 'Location Type', 'Open', 'Created (MTD)', 'Closed (MTD)', 'Last Refreshed', 'Sort Col#', 'Desc?'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Single LET formula that outputs the entire sortable table
  const mainFormula =
    '=LET(' +
    'locs, UNIQUE(FILTER(TicketData!M2:M, TicketData!M2:M<>"", TicketData!M2:M<>"LocationName")),' +
    'mtdStart, TEXT(DATE(YEAR(TODAY()),MONTH(TODAY()),1), "YYYY-MM-DD"),' +
    'mtdEnd, TEXT(DATE(YEAR(TODAY()),MONTH(TODAY())+1,1), "YYYY-MM-DD"),' +
    'col_a, locs,' +
    'col_b, BYROW(locs, LAMBDA(l, IFERROR(INDEX(TicketData!N:N, MATCH(l, TicketData!M:M, 0)), ""))),' +
    'col_c, BYROW(locs, LAMBDA(l, COUNTIFS(TicketData!M:M, l, TicketData!H:H, "No"))),' +
    'col_d, BYROW(locs, LAMBDA(l, COUNTIFS(TicketData!M:M, l, TicketData!E:E, ">="&mtdStart, TicketData!E:E, "<"&mtdEnd))),' +
    'col_e, BYROW(locs, LAMBDA(l, COUNTIFS(TicketData!M:M, l, TicketData!G:G, ">="&mtdStart, TicketData!G:G, "<"&mtdEnd))),' +
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
 * @returns {boolean} true if created, false if already exists
 */
function setupFunctionalAreaSummarySheet(ss) {
  if (ss.getSheetByName('FunctionalAreaSummary')) return false;

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
 * @returns {boolean} true if created, false if already exists
 */
function setupAtRiskResponseSheet(ss) {
  if (ss.getSheetByName('AtRiskResponse')) return false;

  const sheet = ss.insertSheet('AtRiskResponse');

  // Headers
  const headers = ['Ticket Number', 'Subject', 'Team', 'Threshold (hrs)', 'Elapsed (hrs)', '% of SLA', 'Time Remaining (hrs)'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Response SLA at-risk formula
  // Shows open tickets where Response SLA usage is between risk% and 100%
  const riskLookup = 'IFERROR(VLOOKUP("SLA_RISK_PERCENT",Config!A:B,2,FALSE)/100,0.75)';
  const formula =
    '=IFERROR(SORT(FILTER(' +
    '{TicketData!B:B, LEFT(TicketData!C:C,60), TicketData!K:K, ' +
    'TicketData!AC:AC/60, TicketData!AD:AD/60, ' +
    'TicketData!AD:AD/TicketData!AC:AC, (TicketData!AC:AC-TicketData!AD:AD)/60}, ' +
    '(TicketData!H:H="No")*' +
    '(TicketData!AC:AC>0)*' +
    '(TicketData!AE:AE<>TRUE)*' +
    '(TicketData!AD:AD/TicketData!AC:AC>=' + riskLookup + ')*' +
    '(TicketData!AD:AD/TicketData!AC:AC<1)' +
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
 * @returns {boolean} true if created, false if already exists
 */
function setupAtRiskResolutionSheet(ss) {
  if (ss.getSheetByName('AtRiskResolution')) return false;

  const sheet = ss.insertSheet('AtRiskResolution');

  // Headers
  const headers = ['Ticket Number', 'Subject', 'Team', 'Threshold (hrs)', 'Elapsed (hrs)', '% of SLA', 'Time Remaining (hrs)'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Resolution SLA at-risk formula
  // Shows open tickets where Resolution SLA usage is between risk% and 100%
  const riskLookup = 'IFERROR(VLOOKUP("SLA_RISK_PERCENT",Config!A:B,2,FALSE)/100,0.75)';
  const formula =
    '=IFERROR(SORT(FILTER(' +
    '{TicketData!B:B, LEFT(TicketData!C:C,60), TicketData!K:K, ' +
    'TicketData!AF:AF/60, TicketData!AG:AG/60, ' +
    'TicketData!AG:AG/TicketData!AF:AF, (TicketData!AF:AF-TicketData!AG:AG)/60}, ' +
    '(TicketData!H:H="No")*' +
    '(TicketData!AF:AF>0)*' +
    '(TicketData!AH:AH<>TRUE)*' +
    '(TicketData!AG:AG/TicketData!AF:AF>=' + riskLookup + ')*' +
    '(TicketData!AG:AG/TicketData!AF:AF<1)' +
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
 * @returns {boolean} true if created, false if already exists
 */
function setupPerformanceTrendsSheet(ss) {
  if (ss.getSheetByName('PerformanceTrends')) return false;

  const sheet = ss.insertSheet('PerformanceTrends');

  // Headers
  const headers = ['Month', 'Year', 'Closed', 'Avg Resolution (days)', 'Closure Rate', 'Breach Rate', 'Backlog EOM', '% Aged 30+ EOM'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Generate month rows for last 2 years
  const currentYear = new Date().getFullYear();
  const months = ['January', 'February', 'March', 'April', 'May', 'June',
                  'July', 'August', 'September', 'October', 'November', 'December'];

  const dataRows = [];
  for (let year = currentYear - 1; year <= currentYear; year++) {
    for (let m = 0; m < 12; m++) {
      const rowNum = dataRows.length + 2;
      dataRows.push([
        months[m],
        year,
        // C: Count tickets closed in that month
        `=LET(m, MATCH(A${rowNum}, {"January";"February";"March";"April";"May";"June";"July";"August";"September";"October";"November";"December"}, 0), COUNTIFS(TicketData!G:G, ">="&TEXT(DATE(B${rowNum},m,1), "YYYY-MM-DD"), TicketData!G:G, "<"&TEXT(DATE(B${rowNum},m+1,1), "YYYY-MM-DD")))`,
        // D: Average resolution time in DAYS for tickets closed that month
        `=LET(m, MATCH(A${rowNum}, {"January";"February";"March";"April";"May";"June";"July";"August";"September";"October";"November";"December"}, 0), IFERROR(AVERAGEIFS(TicketData!Q:Q, TicketData!H:H, "Yes", TicketData!G:G, ">="&TEXT(DATE(B${rowNum},m,1), "YYYY-MM-DD"), TicketData!G:G, "<"&TEXT(DATE(B${rowNum},m+1,1), "YYYY-MM-DD")), "N/A"))`,
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
 * @returns {boolean} true if created, false if already exists
 */
function setupStaleTicketsSheet(ss) {
  if (ss.getSheetByName('StaleTickets')) return false;

  const sheet = ss.insertSheet('StaleTickets');

  // Headers
  const headers = ['Ticket Number', 'Subject', 'Team', 'Days Since Update', 'Last Update', 'Created Date', 'Status'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Single formula that returns all stale tickets, sorted by days since update
  // Uses STALE_DAYS from Config (default 7)
  const staleFormula =
    '=IFERROR(SORT(FILTER({TicketData!B2:B, LEFT(TicketData!C2:C,80), TicketData!K2:K, ' +
    'INT(TODAY()-DATEVALUE(LEFT(TicketData!F2:F,10))), LEFT(TicketData!F2:F,10), ' +
    'LEFT(TicketData!E2:E,10), TicketData!I2:I}, ' +
    '(TicketData!H2:H="No")*(INT(TODAY()-DATEVALUE(LEFT(TicketData!F2:F,10)))>=' +
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

      if (config.historicalYears.length === 0 && !config.currentYear) {
        issues.push('No TICKET_{YEAR}_LAST_PAGE or TICKET_{YEAR}_LAST_FETCH rows found');
      }

    } catch (e) {
      issues.push('Config sheet error: ' + e.message);
    }
  }

  if (issues.length === 0) {
    ui.alert('Configuration Valid', 'All required settings are configured.\n\nYou can now run "Continue Loading" to start loading data.', ui.ButtonSet.OK);
  } else {
    ui.alert('Configuration Issues', 'Please fix the following issues:\n\n- ' + issues.join('\n- '), ui.ButtonSet.OK);
  }
}
