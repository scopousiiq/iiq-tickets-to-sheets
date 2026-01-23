/**
 * Daily Snapshot - Captures backlog metrics for trending analysis
 *
 * These metrics cannot be calculated retroactively, so daily capture is essential.
 * Used by PerformanceTrends sheet to show end-of-month backlog and aging.
 *
 * Recommended Trigger: Daily at 7:00 PM (after business hours)
 */

/**
 * Append today's snapshot to the DailySnapshot sheet
 * Safe to run multiple times per day - only adds one row per date
 */
function appendDailySnapshot() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const snapshotSheet = ss.getSheetByName('DailySnapshot');
  const ticketSheet = ss.getSheetByName('TicketData');

  if (!snapshotSheet) {
    logOperation('DailySnapshot', 'ERROR', 'DailySnapshot sheet not found');
    return;
  }

  if (!ticketSheet) {
    logOperation('DailySnapshot', 'ERROR', 'TicketData sheet not found');
    return;
  }

  const today = new Date();
  const todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  // Check if we already have a snapshot for today
  const existingData = snapshotSheet.getDataRange().getValues();
  for (let i = 1; i < existingData.length; i++) {
    if (existingData[i][0]) {
      const rowDate = Utilities.formatDate(new Date(existingData[i][0]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (rowDate === todayStr) {
        logOperation('DailySnapshot', 'SKIP', 'Snapshot already exists for ' + todayStr);
        return;
      }
    }
  }

  // Calculate metrics from TicketData
  const ticketData = ticketSheet.getDataRange().getValues();
  const headers = ticketData[0];

  // Find column indexes
  const isClosedCol = headers.indexOf('IsClosed');
  const ageDaysCol = headers.indexOf('AgeDays');

  if (isClosedCol === -1 || ageDaysCol === -1) {
    // Try by position if headers not found (H=7, Q=16)
    const isClosedIdx = 7;  // Column H (0-indexed)
    const ageDaysIdx = 16;  // Column Q (0-indexed)

    let openCount = 0;
    let aged30Count = 0;

    for (let i = 1; i < ticketData.length; i++) {
      const row = ticketData[i];
      if (row[isClosedIdx] === 'No') {
        openCount++;
        if (row[ageDaysIdx] >= 30) {
          aged30Count++;
        }
      }
    }

    const percentAged = openCount > 0 ? (aged30Count / openCount) : 0;

    // Append the snapshot
    snapshotSheet.appendRow([
      today,
      openCount,
      aged30Count,
      percentAged
    ]);

    logOperation('DailySnapshot', 'SUCCESS',
      `Captured: ${openCount} open, ${aged30Count} aged 30+ (${(percentAged * 100).toFixed(1)}%)`);
    return;
  }

  // Use found column indexes
  let openCount = 0;
  let aged30Count = 0;

  for (let i = 1; i < ticketData.length; i++) {
    const row = ticketData[i];
    if (row[isClosedCol] === 'No') {
      openCount++;
      if (row[ageDaysCol] >= 30) {
        aged30Count++;
      }
    }
  }

  const percentAged = openCount > 0 ? (aged30Count / openCount) : 0;

  // Append the snapshot
  snapshotSheet.appendRow([
    today,
    openCount,
    aged30Count,
    percentAged
  ]);

  logOperation('DailySnapshot', 'SUCCESS',
    `Captured: ${openCount} open, ${aged30Count} aged 30+ (${(percentAged * 100).toFixed(1)}%)`);
}

/**
 * Trigger-safe function for automated daily snapshots
 * Use this with a time-driven trigger (daily at 7 PM recommended)
 */
function triggerDailySnapshot() {
  logOperation('Trigger', 'START', 'Daily snapshot triggered');

  try {
    appendDailySnapshot();
  } catch (error) {
    logOperation('Trigger', 'ERROR', 'Daily snapshot failed: ' + error.message);
  }
}

/**
 * Manually populate historical snapshots (approximate)
 * This creates estimated snapshots for past months based on current data
 * NOTE: These are estimates only - actual historical values may differ
 */
function populateHistoricalSnapshots() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Populate Historical Estimates',
    'This will create ESTIMATED snapshots for the last day of each past month.\n\n' +
    'These are approximations based on current data and may not reflect actual historical values.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const snapshotSheet = ss.getSheetByName('DailySnapshot');
  const ticketSheet = ss.getSheetByName('TicketData');

  if (!snapshotSheet || !ticketSheet) {
    ui.alert('Error', 'Required sheets not found.', ui.ButtonSet.OK);
    return;
  }

  const ticketData = ticketSheet.getDataRange().getValues();
  const today = new Date();

  // Get min/max dates from ticket data
  let minDate = today;
  let maxDate = new Date(2020, 0, 1);

  for (let i = 1; i < ticketData.length; i++) {
    const createdDate = ticketData[i][4]; // Column E - CreatedDate
    if (createdDate) {
      const d = new Date(createdDate);
      if (d < minDate) minDate = d;
      if (d > maxDate) maxDate = d;
    }
  }

  // Generate end-of-month dates from minDate to last month
  const snapshots = [];
  let current = new Date(minDate.getFullYear(), minDate.getMonth() + 1, 0); // End of min month
  const lastMonth = new Date(today.getFullYear(), today.getMonth(), 0); // End of last month

  while (current <= lastMonth) {
    // For each end-of-month date, estimate the backlog
    // This is approximate: count tickets created before this date that weren't closed before this date
    let openCount = 0;
    let aged30Count = 0;
    const thirtyDaysAgo = new Date(current.getTime() - 30 * 24 * 60 * 60 * 1000);

    for (let i = 1; i < ticketData.length; i++) {
      const createdDate = new Date(ticketData[i][4]); // Column E
      const closedDate = ticketData[i][6] ? new Date(ticketData[i][6]) : null; // Column G
      const isClosed = ticketData[i][7]; // Column H

      // Was this ticket open on this date?
      if (createdDate <= current) {
        if (isClosed === 'No' || (closedDate && closedDate > current)) {
          openCount++;
          if (createdDate <= thirtyDaysAgo) {
            aged30Count++;
          }
        }
      }
    }

    const percentAged = openCount > 0 ? (aged30Count / openCount) : 0;
    snapshots.push([current, openCount, aged30Count, percentAged]);

    // Move to next month
    current = new Date(current.getFullYear(), current.getMonth() + 2, 0);
  }

  // Append all snapshots
  if (snapshots.length > 0) {
    const lastRow = snapshotSheet.getLastRow();
    snapshotSheet.getRange(lastRow + 1, 1, snapshots.length, 4).setValues(snapshots);
  }

  logOperation('DailySnapshot', 'HISTORICAL', `Created ${snapshots.length} estimated historical snapshots`);
  ui.alert('Complete', `Created ${snapshots.length} estimated historical snapshots.\n\nNote: These are estimates based on current data.`, ui.ButtonSet.OK);
}
