/**
 * Teams Directory
 * Loads all teams for Functional Area mapping
 */

function refreshTeams() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Teams');
  const ui = SpreadsheetApp.getUi();

  try {
    const teams = getAllTeams();

    if (!teams || teams.length === 0) {
      ui.alert('Warning', 'No teams returned from API.', ui.ButtonSet.OK);
      return;
    }

    // Preserve existing Functional Area mappings
    const existingData = {};
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const existing = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
      existing.forEach(row => {
        if (row[0]) {
          existingData[row[0]] = row[2]; // TeamId -> Functional Area
        }
      });
    }

    // Clear data rows (keep header)
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 4).clear();
    }

    // Populate with fresh team data
    const teamRows = teams.map(team => {
      const teamId = team.TeamId || team.Id;
      return [
        teamId,
        team.TeamName || team.Name || '',
        existingData[teamId] || '', // Preserve FA mapping
        team.IsDeleted === false ? 'Yes' : 'No'
      ];
    });

    if (teamRows.length > 0) {
      sheet.getRange(2, 1, teamRows.length, 4).setValues(teamRows);
    }

    updateLastRefresh();
    logOperation('Teams Refresh', 'SUCCESS', `Loaded ${teamRows.length} teams`);

    ui.alert('Complete', `Loaded ${teamRows.length} teams. Please update Functional Area mappings in column C.`, ui.ButtonSet.OK);

  } catch (e) {
    logOperation('Teams Refresh', 'ERROR', e.message);
    ui.alert('Error', 'Failed to refresh teams: ' + e.message, ui.ButtonSet.OK);
  }
}
