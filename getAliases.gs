/**
 * Lists users and groups with email aliases and adds them to the "Aliases" sheet
 * in the current Google Sheet. 
 *
 * The aliases are printed in a one-per-row format with columns for Alias,
 * Target (primary email/address), and TargetType (User/Group).
 * This version includes optimizations for faster execution and error prevention.
 *
 * © 2021 xFanatical, Inc.
 * @license MIT
 * This script has been redesigned for the DoiT AdminPulse for Workspace marketplace app.
 */
function listAliases() {
  const functionName = 'listAliases';
  const startTime = new Date();
  Logger.log(`-- Starting ${functionName} at: ${startTime.toLocaleString()}`);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Delete existing "Aliases" sheet if it exists
    const existingSheet = ss.getSheetByName('Aliases');
    if (existingSheet) {
      ss.deleteSheet(existingSheet);
    }

    // Create a new "Aliases" sheet
    const aliasSheet = ss.insertSheet('Aliases');

    // Append header row
    aliasSheet.appendRow(['Alias', 'Target', 'TargetType']);

    // Format header row
    const headerRange = aliasSheet.getRange(1, 1, 1, 3);
    headerRange.setFontFamily('Montserrat')
      .setFontWeight('bold')
      .setBackground('#fc3156') // Note: Original was fc3156, you might want #fc3165 to match other sheets
      .setFontColor("#ffffff");
    aliasSheet.setFrozenRows(1);


    // Get user and group aliases
    // Note: These calls fetch all users/groups. For very large domains, pagination may be needed.
    const users = AdminDirectory.Users.list({
      customer: 'my_customer',
      fields: 'users(primaryEmail,aliases)'
    }).users;
    const groups = AdminDirectory.Groups.list({
      customer: 'my_customer',
      fields: 'groups(email,aliases)'
    }).groups;

    // Prepare data for batch update
    const data = [];
    if (users) {
      users.forEach(user => {
        if (user.aliases) {
          user.aliases.forEach(alias => data.push([alias, user.primaryEmail, 'User']));
        }
      });
    } else {
      Logger.log(`In ${functionName}: No users found.`);
    }

    if (groups) {
      groups.forEach(group => {
        if (group.aliases) {
          group.aliases.forEach(alias => data.push([alias, group.email, 'Group']));
        }
      });
    } else {
      Logger.log(`In ${functionName}: No groups found.`);
    }

    // Batch update the sheet with the data
    if (data.length > 0) {
      aliasSheet.getRange(aliasSheet.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
    }

    // Format columns
    aliasSheet.autoResizeColumns(1, 3);
    aliasSheet.setColumnWidth(2, 300);
    aliasSheet.setColumnWidth(1, 300);

    // Delete empty rows (avoiding the header row)
    const maxRows = aliasSheet.getMaxRows();
    const lastRow = aliasSheet.getLastRow();
    if (maxRows > lastRow) {
      aliasSheet.deleteRows(lastRow + 1, maxRows - lastRow);
    }

    // Delete columns D-Z if they exist
    if (aliasSheet.getMaxColumns() > 3) {
      aliasSheet.deleteColumns(4, aliasSheet.getMaxColumns() - 3);
    }


  } catch (e) {
    Logger.log(`!! ERROR in ${functionName}: ${e.toString()}`);
    // Re-throw the error to potentially stop the main script if something goes wrong
    throw e;
  } finally {
    const endTime = new Date();
    const duration = (endTime.getTime() - startTime.getTime()) / 1000; // Duration in seconds
    Logger.log(`-- Finished ${functionName} at: ${endTime.toLocaleString()} (Duration: ${duration.toFixed(2)}s)`);
  }
}