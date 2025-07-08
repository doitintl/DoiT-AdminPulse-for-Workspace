/**
 * This script will inventory names of all configured App Passwords from all users in an organization.
 */
function getAppPasswords() {
  const functionName = 'getAppPasswords';
  const startTime = new Date();
  Logger.log(`-- Starting ${functionName} at: ${startTime.toLocaleString()}`);

  try {
    const userEmail = Session.getActiveUser().getEmail();
    const domain = userEmail.split("@").pop();

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let appPasswordsSheet = spreadsheet.getSheetByName("App Passwords");
    if (appPasswordsSheet) {
      spreadsheet.deleteSheet(appPasswordsSheet);
    }

    appPasswordsSheet = spreadsheet.insertSheet("App Passwords", spreadsheet.getNumSheets());

    const headerRange = appPasswordsSheet.getRange("A1:E1");
    headerRange.setFontFamily("Montserrat")
      .setBackground("#fc3165")
      .setFontColor("white")
      .setFontWeight("bold")
      .setValues([["CodeID", "Name", "Creation Time", "Last Time Used", "User"]]);
    appPasswordsSheet.setFrozenRows(1);

    let pageToken = null;
    do {
      const response = AdminDirectory.Users.list({
        domain: domain,
        customer: "my_customer",
        maxResults: 100,
        projection: "FULL",
        viewType: "admin_view",
        orderBy: "email",
        pageToken: pageToken,
      });

      if (response.users && response.users.length > 0) {
        processUsers(response.users, appPasswordsSheet);
      }

      pageToken = response.nextPageToken;
    } while (pageToken);

    // --- Post-processing and cleanup ---
    const lastRow = appPasswordsSheet.getLastRow();

    // ======================== MODIFIED CODE BLOCK ========================
    if (lastRow > 1) { // Data was found, so format it.
      // Auto-resize columns
      appPasswordsSheet.autoResizeColumns(1, 5);
      
      // Add filter
      appPasswordsSheet.getRange('A1:E' + lastRow).createFilter();

      // Add conditional formatting for "Never Used"
      const neverUsedRange = appPasswordsSheet.getRange("D2:D" + lastRow);
      const neverUsedRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("Never Used")
        .setBackground("#f4cccc") // Light red
        .setRanges([neverUsedRange])
        .build();
      const rules = appPasswordsSheet.getConditionalFormatRules();
      rules.push(neverUsedRule);
      appPasswordsSheet.setConditionalFormatRules(rules);

    } else { // No data was found, so write a message.
      const messageRange = appPasswordsSheet.getRange("A2:E2");
      messageRange.merge()
        .setValue("No App Passwords found.")
        .setHorizontalAlignment("center")
        .setFontStyle("italic")
        .setFontColor("#666666");
    }
    // ===================================================================

    // Delete extra columns from F onwards
    const maxCols = appPasswordsSheet.getMaxColumns();
    if (maxCols > 5) {
      appPasswordsSheet.deleteColumns(6, maxCols - 5);
    }
    
    // Delete empty rows from the bottom of the sheet, ensuring we don't cause an error.
    const maxRows = appPasswordsSheet.getMaxRows();
    // Re-fetch lastRow in case we added the "No App Passwords" message.
    const finalLastRow = appPasswordsSheet.getLastRow();
    
    // We must always keep at least one non-frozen row. The header is frozen.
    // So if finalLastRow is 1 (only header) or 2 (header + message), this logic works.
    const rowsToKeep = Math.max(finalLastRow, 2); 

    if (maxRows > rowsToKeep) {
      appPasswordsSheet.deleteRows(rowsToKeep + 1, maxRows - rowsToKeep);
    }

  } catch (e) {
    Logger.log(`!! ERROR in ${functionName}: ${e.toString()}`);
    if (e.stack) {
      Logger.log(`Stack trace: ${e.stack}`);
    }
    if (e.message.indexOf("User does not have credentials to perform this operation") > -1) {
      const ui = SpreadsheetApp.getUi();
      ui.alert(
        'Insufficient Permissions',
        'You need Super Admin privileges to use this feature',
        ui.ButtonSet.OK
      );
    } else {
      throw e;
    }
  } finally {
    const endTime = new Date();
    const duration = (endTime.getTime() - startTime.getTime()) / 1000;
    Logger.log(`-- Finished ${functionName} at: ${endTime.toLocaleString()} (Duration: ${duration.toFixed(2)}s)`);
  }
}

// NOTE: No changes were needed in the functions below.
function processUsers(users, sheet) {
  const data = [];

  users.forEach(function (user) {
    Utilities.sleep(200); 
    try {
      const asps = AdminDirectory.Asps.list(user.primaryEmail);

      if (asps && asps.items) {
        asps.items.forEach(function (asp) {
          data.push([
            asp.codeId,
            asp.name,
            formatTimestamp(asp.creationTime),
            asp.lastTimeUsed ? formatTimestamp(asp.lastTimeUsed) : "Never Used",
            user.primaryEmail,
          ]);
        });
      }
    } catch (err) {
      Logger.log(`Could not process user ${user.primaryEmail}. Error: ${err.message}`);
    }
  });
  
  if (data.length > 0) {
     sheet.getRange(sheet.getLastRow() + 1, 1, data.length, data[0].length)
       .setValues(data);
  }
}

function formatTimestamp(timestampString) {
  if (timestampString === "0" || timestampString === 0) {
    return "Never Used";
  } else if (typeof timestampString === "string" && timestampString.length >= 13) {
    const timestamp = parseInt(timestampString, 10);
    const date = new Date(timestamp);
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  } else {
    Logger.log("Unknown timestamp format: " + timestampString);
    return "Invalid Timestamp";
  }
}