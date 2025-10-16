/**
 * @fileoverview Inventories all App Passwords from all users.
 * This script is designed to handle very large Google Workspace environments by using a highly
 * scalable batch processing pattern. It processes users page by page, using time-based
 * triggers to avoid exceeding script execution time limits, without ever needing to store the
 * full user list in memory or properties.
 */

// --- Configuration ---
const APP_PASSWORDS_SHEET_NAME = "App Passwords";
const APP_PASSWORDS_USER_PAGE_SIZE = 200; // Number of users to fetch in each page/batch.
const APP_PASSWORDS_TRIGGER_FUNCTION = "processAppPasswordBatch";

/**
 * Main function to be run from the menu.
 * Kicks off the App Password inventory process.
 */
function getAppPasswords() {
  const functionName = 'getAppPasswords';
  const startTime = new Date();
  Logger.log(`-- Starting ${functionName} at: ${startTime.toLocaleString()}`);
  
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheet.toast('Starting App Password inventory...', 'Setup', 10);

  try {
    // 1. Clean up from any previous runs
    _deleteTriggersByName(APP_PASSWORDS_TRIGGER_FUNCTION);
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.deleteProperty('appPasswords_userPageToken');
    scriptProperties.setProperty('appPasswords_startTime', startTime.getTime());

    // 2. Set up the spreadsheet
    _setupAppPasswordsSheet();
    
    // 3. Start the first batch
    spreadsheet.toast('Starting first batch of users...', 'Processing', 10);
    processAppPasswordBatch();

  } catch (e) {
    Logger.log(`!! FATAL ERROR in ${functionName}: ${e.toString()}\n${e.stack}`);
    SpreadsheetApp.getUi().alert(`A critical error occurred during setup: ${e.message}. Please check the logs.`);
  }
}

/**
 * Processes a batch of users to fetch their App Passwords.
 * This function fetches a single page of users from the Admin SDK, processes them,
 * and then triggers itself for the next page.
 */
function processAppPasswordBatch() {
  const functionName = APP_PASSWORDS_TRIGGER_FUNCTION;
  const scriptProperties = PropertiesService.getScriptProperties();
  
  try {
    const pageToken = scriptProperties.getProperty('appPasswords_userPageToken');
    Logger.log(`Processing user page with token: ${pageToken || ' (first page)'}`);

    // 1. Fetch a single page of users
    const userPage = AdminDirectory.Users.list({
      customer: "my_customer",
      maxResults: APP_PASSWORDS_USER_PAGE_SIZE,
      projection: "basic",
      viewType: "admin_view",
      orderBy: "email",
      fields: "nextPageToken,users(id,primaryEmail)",
      pageToken: pageToken,
    });

    let allPasswordsData = [];
    if (userPage.users && userPage.users.length > 0) {
      Logger.log(`Processing batch of ${userPage.users.length} users.`);
      // 2. Process app passwords for the fetched users
      userPage.users.forEach((user) => {
        Utilities.sleep(250); // Prevent hitting API rate limits
        try {
          const asps = AdminDirectory.Asps.list(user.id);
          if (asps && asps.items) {
            asps.items.forEach((asp) => {
              allPasswordsData.push([
                asp.codeId,
                asp.name,
                _formatTimestamp(asp.creationTime),
                asp.lastTimeUsed ? _formatTimestamp(asp.lastTimeUsed) : "Never Used",
                user.primaryEmail,
              ]);
            });
          }
        } catch (err) {
          Logger.log(`Could not process App Passwords for user ID ${user.id}. Error: ${err.message}`);
        }
      });
    }

    // 3. Write the collected data for this batch to the sheet
    if (allPasswordsData.length > 0) {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(APP_PASSWORDS_SHEET_NAME);
      sheet.getRange(sheet.getLastRow() + 1, 1, allPasswordsData.length, allPasswordsData[0].length).setValues(allPasswordsData);
    }

    // 4. Check if there is a next page and trigger the next run
    const nextPageToken = userPage.nextPageToken;
    if (nextPageToken) {
      scriptProperties.setProperty('appPasswords_userPageToken', nextPageToken);
      _createTrigger(APP_PASSWORDS_TRIGGER_FUNCTION, 5);
      Logger.log(`Batch complete. Trigger created for next user page.`);
    } else {
      // 5. No more pages, finalize the process
      Logger.log("All user pages have been processed. Finalizing sheet.");
      _finalizeAppPasswordsSheet();
      
      // Clean up properties
      scriptProperties.deleteProperty('appPasswords_userPageToken');
      
      const totalStartTime = new Date(parseInt(scriptProperties.getProperty('appPasswords_startTime'), 10));
      const totalEndTime = new Date();
      const totalDuration = (totalEndTime.getTime() - totalStartTime.getTime()) / 1000;
      Logger.log(`-- Successfully completed App Password inventory at: ${totalEndTime.toLocaleString()}. Total duration: ${totalDuration.toFixed(2)} seconds.`);
    }
  } catch (e) {
    Logger.log(`!! FATAL ERROR in ${functionName}: ${e.toString()}\n${e.stack}`);
    _deleteTriggersByName(APP_PASSWORDS_TRIGGER_FUNCTION);
  }
}


// --- Helper Functions (Setup, Finalization, Formatting) ---

/**
 * Sets up the initial state of the "App Passwords" sheet.
 * @private
 */
function _setupAppPasswordsSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(APP_PASSWORDS_SHEET_NAME);

  if (sheet) {
    sheet.clear();
    spreadsheet.deleteSheet(sheet);
  }
  
  sheet = spreadsheet.insertSheet(APP_PASSWORDS_SHEET_NAME, 0);

  const headers = ["CodeID", "Name", "Creation Time", "Last Time Used", "User"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontFamily("Montserrat")
    .setBackground("#fc3165")
    .setFontColor("white")
    .setFontWeight("bold");
  sheet.setFrozenRows(1);
}

/**
 * Applies final formatting and cleanup to the sheet.
 * @private
 */
function _finalizeAppPasswordsSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(APP_PASSWORDS_SHEET_NAME);
  if (sheet.getLastRow() <= 1) {
    sheet.getRange("A2").setValue("No App Passwords found in the domain.");
    return;
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  // Create filter
  sheet.getRange(1, 1, lastRow, lastCol).createFilter();

  // Add conditional formatting for "Never Used"
  const neverUsedRange = sheet.getRange("D2:D" + lastRow);
  const neverUsedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Never Used")
    .setBackground("#f4cccc")
    .setRanges([neverUsedRange])
    .build();
  const rules = sheet.getConditionalFormatRules();
  rules.push(neverUsedRule);
  sheet.setConditionalFormatRules(rules);

  // Auto-resize columns
  for (let i = 1; i <= lastCol; i++) {
    sheet.autoResizeColumn(i);
  }
  
  // Clean up extra rows/columns
  const maxCols = sheet.getMaxColumns();
  if (maxCols > lastCol) {
    sheet.deleteColumns(lastCol + 1, maxCols - lastCol);
  }
  const maxRows = sheet.getMaxRows();
  if (maxRows > lastRow) {
    sheet.deleteRows(lastRow + 1, maxRows - lastRow);
  }
}

/**
 * Formats a Unix timestamp string into a human-readable date.
 * @param {string} timestampString A string representing milliseconds since epoch.
 * @returns {string} The formatted date string or a status message.
 * @private
 */
function _formatTimestamp(timestampString) {
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

// --- Generic Trigger Management ---

/**
 * Deletes all script triggers with a specific handler function name.
 * @param {string} functionName The name of the handler function for the triggers to delete.
 * @private
 */
function _deleteTriggersByName(functionName) {
  try {
    ScriptApp.getProjectTriggers().forEach(trigger => {
      if (trigger.getHandlerFunction() === functionName) {
        ScriptApp.deleteTrigger(trigger);
      }
    });
  } catch (e) {
    Logger.log(`Error deleting triggers: ${e.message}`);
  }
}

/**
 * Creates a time-based trigger to run a function after a short delay.
 * @param {string} functionName The name of the function to trigger.
 * @param {number} delayInSeconds The delay in seconds before the trigger runs.
 * @private
 */
function _createTrigger(functionName, delayInSeconds) {
  ScriptApp.newTrigger(functionName)
    .timeBased()
    .after(delayInSeconds * 1000)
    .create();
}