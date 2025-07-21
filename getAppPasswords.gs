/**
 * Inventories all configured App Passwords from all users across ALL DOMAINS in the organization.
 * Displays user email on the sheet but logs only the user ID for privacy.
 * Provides UI feedback during execution.
 */
function getAppPasswords() {
  const functionName = 'getAppPasswords';
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi(); 
  const startTime = new Date();
  Logger.log(`-- Starting ${functionName} at: ${startTime.toLocaleString()}`);

  spreadsheet.toast(
    'Processing... This may take several minutes.',
    'Starting App Password Audit',
    -1 // Indefinite duration
  );

  try {
    const sheetName = "App Passwords";
    let appPasswordsSheet = spreadsheet.getSheetByName(sheetName);

    if (!appPasswordsSheet) {
      appPasswordsSheet = spreadsheet.insertSheet(sheetName, 0);
    } else {
      const oldFilter = appPasswordsSheet.getFilter();
      if (oldFilter) {
        oldFilter.remove();
      }
      appPasswordsSheet.clear();
    }

    const headers = ["CodeID", "Name", "Creation Time", "Last Time Used", "User"];
    const headerRange = appPasswordsSheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers])
      .setFontFamily("Montserrat")
      .setBackground("#fc3165")
      .setFontColor("white")
      .setFontWeight("bold");
    appPasswordsSheet.setFrozenRows(1);

    let pageToken = null;
    let allPasswordsData = [];

    let pageNumber = 1;
    do {
      spreadsheet.toast(`Fetching user page ${pageNumber}...`, 'Processing...do not close or edit the sheet App Passwords page.', -1);
      
      // MODIFICATION: The 'domain' parameter has been removed from this API call.
      // Using only 'customer: "my_customer"' fetches users from all domains.
      const response = AdminDirectory.Users.list({
        customer: "my_customer",
        maxResults: 100,
        projection: "basic",
        viewType: "admin_view",
        orderBy: "email",
        pageToken: pageToken,
      });

      if (response.users && response.users.length > 0) {
        response.users.forEach(function(user) {
          Utilities.sleep(250);
          try {
            const asps = AdminDirectory.Asps.list(user.id);
            if (asps && asps.items) {
              asps.items.forEach(function(asp) {
                allPasswordsData.push([
                  asp.codeId,
                  asp.name,
                  formatTimestamp(asp.creationTime),
                  asp.lastTimeUsed ? formatTimestamp(asp.lastTimeUsed) : "Never Used",
                  user.primaryEmail, 
                ]);
              });
            }
          } catch (err) {
            Logger.log(`Could not process App Passwords for user ID ${user.id}. Error: ${err.message}`);
          }
        });
      }
      pageToken = response.nextPageToken;
      pageNumber++;
    } while (pageToken);

    if (allPasswordsData.length > 0) {
      spreadsheet.toast('Writing data to sheet...', 'Processing...', -1);
      appPasswordsSheet.getRange(2, 1, allPasswordsData.length, headers.length).setValues(allPasswordsData);

      const lastRow = appPasswordsSheet.getLastRow();

      for (let i = 1; i <= headers.length; i++) {
        appPasswordsSheet.autoResizeColumn(i);
      }
      
      appPasswordsSheet.getRange(1, 1, lastRow, headers.length).createFilter();

      const neverUsedRange = appPasswordsSheet.getRange("D2:D" + lastRow);
      const neverUsedRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("Never Used")
        .setBackground("#f4cccc")
        .setRanges([neverUsedRange])
        .build();
      const rules = appPasswordsSheet.getConditionalFormatRules();
      rules.push(neverUsedRule);
      appPasswordsSheet.setConditionalFormatRules(rules);

    } else {
      appPasswordsSheet.getRange("A2").setValue("No App Passwords found in the domain.");
    }

    const maxCols = appPasswordsSheet.getMaxColumns();
    if (maxCols > headers.length) {
      appPasswordsSheet.deleteColumns(headers.length + 1, maxCols - headers.length);
    }
    const maxRows = appPasswordsSheet.getMaxRows();
    const finalLastRow = appPasswordsSheet.getLastRow();
    if (maxRows > finalLastRow) {
      appPasswordsSheet.deleteRows(finalLastRow + 1, maxRows - finalLastRow);
    }
    // Log the successful completion of the function.
    const endTime = new Date();
    const duration = (endTime.getTime() - startTime.getTime()) / 1000; // duration in seconds
    Logger.log(`-- Successfully completed ${functionName} at: ${endTime.toLocaleString()}. Total duration: ${duration.toFixed(2)} seconds.`);
    
    spreadsheet.toast('Audit complete!', 'Success!', 10);
    
    const alertMessage = "The App Password inventory is complete.\n\n" +
      "IMPORTANT: App Passwords are 16-digit passcodes that grant access to an account. They are a security risk because they bypass 2-Step Verification (2SV).\n\n" +
      "Review any unfamiliar or old entries. Rows highlighted for 'Never Used' indicates the password was never used by an app to authenticate to Google services.";

    ui.alert('Audit Summary & Security Warning', alertMessage, ui.ButtonSet.OK);

  } catch (e) {
    Logger.log(`!! FATAL ERROR in ${functionName}: ${e.toString()}\n${e.stack}`);
    spreadsheet.toast('An error occurred. Check logs.', 'Error!', 15);

    if (e.message.includes("User does not have credentials to perform this operation")) {
      ui.alert(
        'Insufficient Permissions',
        'This script must be run by a Super Administrator to view App Passwords for all users.',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert('An unexpected error occurred. Please check the script execution logs for details.');
    }
  }
}

/**
 * Formats a Unix timestamp string into a human-readable date.
 * @param {string} timestampString A string representing milliseconds since epoch.
 * @returns {string} The formatted date string or a status message.
 */
function formatTimestamp(timestampString) {
  if (!timestampString || timestampString === "0") {
    return "Never Used";
  }
  const timestamp = parseInt(timestampString, 10);
  if (isNaN(timestamp)) {
    return "Invalid Timestamp";
  }
  const date = new Date(timestamp);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
}