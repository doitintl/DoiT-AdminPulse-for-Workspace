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
    .setBackground('#fc3156')
    .setFontColor("#ffffff");
  aliasSheet.setFrozenRows(1);


  // Get user and group aliases
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
    Logger.log('No users found.');
  }

  if (groups) {
    groups.forEach(group => {
      if (group.aliases) {
        group.aliases.forEach(alias => data.push([alias, group.email, 'Group']));
      }
    });
  } else {
    Logger.log('No groups found.');
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

  // Delete columns D-Z if they exist - MOVED TO THE END
  aliasSheet.deleteColumns(4, 23); // Delete from column 4 onwards

}