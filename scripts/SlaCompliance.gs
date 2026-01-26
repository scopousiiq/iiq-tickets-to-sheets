/**
 * SLA Compliance Analytics
 *
 * Calculates monthly SLA metrics from consolidated TicketData.
 * Script-based for reliability and performance with large datasets.
 *
 * Metrics calculated:
 * - Closed: Count of tickets closed in the month
 * - Breaches: Count of tickets with response or resolution SLA breach
 * - Breach Rate: Breaches / Closed
 * - Avg Response (hrs): Average first response time in hours
 * - Avg Resolution (hrs): Average resolution time in hours
 *
 * Data Source:
 * - TicketData sheet (35 columns including SLA metrics in columns 29-35)
 *
 * Menu: IIQ Data > Analytics > Refresh SLA Compliance
 */

/**
 * Refresh SLA Compliance data (menu entry point)
 */
function refreshSlaCompliance() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const response = ui.alert(
    'Refresh SLA Compliance',
    'This will recalculate SLA metrics from TicketData.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const result = runSlaComplianceRefresh(ss);

  ui.alert(
    'SLA Compliance Updated',
    `Calculated metrics for ${result.monthCount} months.\n` +
    `Processed ${result.ticketCount} closed tickets.\n` +
    `Runtime: ${result.runtime} seconds`,
    ui.ButtonSet.OK
  );
}

/**
 * Run SLA Compliance refresh (can be called from menu or trigger)
 */
function runSlaComplianceRefresh(ss) {
  const startTime = Date.now();

  logOperation('SLA Compliance', 'START', 'Calculating SLA metrics from consolidated TicketData');

  // Get closed tickets with SLA data (now from single source)
  const closedTickets = getClosedTicketsWithSla(ss);

  logOperation('SLA Compliance', 'DATA_LOADED', `${closedTickets.length} closed tickets with SLA data`);

  // Calculate metrics by month
  const metrics = calculateMonthlyMetricsFromTickets(closedTickets);

  // Write to sheet
  writeSlaComplianceSheet(ss, metrics);

  const runtime = ((Date.now() - startTime) / 1000).toFixed(1);
  logOperation('SLA Compliance', 'COMPLETE',
    `${metrics.length} months calculated in ${runtime}s`);

  return {
    monthCount: metrics.length,
    ticketCount: closedTickets.length,
    runtime: runtime
  };
}

/**
 * Get closed tickets with their SLA metrics from consolidated TicketData
 * @param {Spreadsheet} ss - Active spreadsheet
 * @returns {Array} - Array of ticket objects with SLA data
 *
 * TicketData columns used:
 * - A (1): TicketId
 * - G (7): ClosedDate
 * - H (8): IsClosed
 * - AD (30): ResponseActual
 * - AE (31): ResponseBreach
 * - AG (33): ResolutionActual
 * - AH (34): ResolutionBreach
 */
function getClosedTicketsWithSla(ss) {
  const sheet = ss.getSheetByName('TicketData');
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const lastRow = sheet.getLastRow();
  // Read all 35 columns for efficiency
  const data = sheet.getRange(2, 1, lastRow - 1, 35).getValues();

  const tickets = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const ticketId = row[0];         // Column A
    const closedDate = row[6];       // Column G (0-indexed = 6)
    const isClosed = row[7];         // Column H (0-indexed = 7)
    const responseActual = row[29];  // Column AD (0-indexed = 29)
    const responseBreach = row[30];  // Column AE (0-indexed = 30)
    const resolutionActual = row[32]; // Column AG (0-indexed = 32)
    const resolutionBreach = row[33]; // Column AH (0-indexed = 33)

    // Only process closed tickets
    if (ticketId && (isClosed === 'Yes' || isClosed === true) && closedDate) {
      const date = new Date(closedDate);
      if (!isNaN(date.getTime())) {
        tickets.push({
          ticketId: ticketId,
          closedDate: date,
          year: date.getFullYear(),
          month: date.getMonth() + 1, // 1-indexed
          responseActual: responseActual,
          responseBreach: responseBreach,
          resolutionActual: resolutionActual,
          resolutionBreach: resolutionBreach
        });
      }
    }
  }

  return tickets;
}

/**
 * Calculate metrics for each month from closed tickets with SLA data
 * @param {Array} tickets - Array of closed ticket objects with SLA data
 * @returns {Array} - Array of monthly metrics objects
 */
function calculateMonthlyMetricsFromTickets(tickets) {
  // Group closed tickets by year-month
  const monthlyData = new Map();

  for (const ticket of tickets) {
    const key = `${ticket.year}-${String(ticket.month).padStart(2, '0')}`;

    if (!monthlyData.has(key)) {
      monthlyData.set(key, {
        year: ticket.year,
        month: ticket.month,
        closed: 0,
        breaches: 0,
        responseTimes: [],
        resolutionTimes: []
      });
    }

    const monthData = monthlyData.get(key);
    monthData.closed++;

    // Count breaches (either response or resolution)
    if (ticket.responseBreach === true || ticket.resolutionBreach === true) {
      monthData.breaches++;
    }

    // Collect times for averaging (only positive values)
    if (ticket.responseActual && ticket.responseActual > 0) {
      monthData.responseTimes.push(ticket.responseActual);
    }
    if (ticket.resolutionActual && ticket.resolutionActual > 0) {
      monthData.resolutionTimes.push(ticket.resolutionActual);
    }
  }

  // Convert to array and calculate averages
  const results = [];
  const monthNames = ['', 'January', 'February', 'March', 'April', 'May', 'June',
                      'July', 'August', 'September', 'October', 'November', 'December'];

  for (const [key, data] of monthlyData) {
    // Calculate averages, converting minutes to hours
    const avgResponse = data.responseTimes.length > 0
      ? data.responseTimes.reduce((a, b) => a + b, 0) / data.responseTimes.length / 60
      : null;

    const avgResolution = data.resolutionTimes.length > 0
      ? data.resolutionTimes.reduce((a, b) => a + b, 0) / data.resolutionTimes.length / 60
      : null;

    results.push({
      month: monthNames[data.month],
      year: data.year,
      monthNum: data.month,
      closed: data.closed,
      breaches: data.breaches,
      breachRate: data.closed > 0 ? data.breaches / data.closed : 0,
      avgResponseHrs: avgResponse,
      avgResolutionHrs: avgResolution
    });
  }

  // Sort by year desc, month desc (most recent first)
  results.sort((a, b) => {
    if (a.year !== b.year) return b.year - a.year;
    return b.monthNum - a.monthNum;
  });

  return results;
}

/**
 * Write metrics to SLACompliance sheet
 * @param {Spreadsheet} ss - Active spreadsheet
 * @param {Array} metrics - Array of monthly metrics
 */
function writeSlaComplianceSheet(ss, metrics) {
  let sheet = ss.getSheetByName('SLACompliance');

  if (!sheet) {
    sheet = ss.insertSheet('SLACompliance');
  }

  // Clear existing data
  sheet.clear();

  // Write headers
  const headers = ['Month', 'Year', 'Closed', 'Breaches', 'Breach Rate',
                   'Avg Response (hrs)', 'Avg Resolution (hrs)'];
  sheet.getRange(1, 1, 1, 7).setValues([headers]);
  sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
  sheet.setFrozenRows(1);

  // Write data
  if (metrics.length > 0) {
    const rows = metrics.map(m => [
      m.month,
      m.year,
      m.closed,
      m.breaches,
      m.breachRate,
      m.avgResponseHrs !== null ? m.avgResponseHrs : 'N/A',
      m.avgResolutionHrs !== null ? m.avgResolutionHrs : 'N/A'
    ]);

    sheet.getRange(2, 1, rows.length, 7).setValues(rows);

    // Format breach rate as percentage
    sheet.getRange(2, 5, rows.length, 1).setNumberFormat('0.00%');

    // Format hours with 1 decimal place
    sheet.getRange(2, 6, rows.length, 2).setNumberFormat('0.0');
  }

  // Auto-resize columns
  for (let i = 1; i <= 7; i++) {
    sheet.autoResizeColumn(i);
  }
}

/**
 * Trigger-safe version for automated refresh
 * Can be called from a time-driven trigger
 */
function triggerSlaComplianceRefresh() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    const result = runSlaComplianceRefresh(ss);
    logOperation('Trigger', 'SLA_COMPLIANCE',
      `Refreshed: ${result.monthCount} months, ${result.ticketCount} tickets, ${result.runtime}s`);
  } catch (error) {
    logOperation('Trigger', 'ERROR', `SLA Compliance refresh failed: ${error.message}`);
  }
}
