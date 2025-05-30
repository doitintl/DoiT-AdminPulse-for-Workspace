/**
 * This script inventories App Passwords for all users in an organization.
 * Designed for Google Sheets editor add-ons, it runs in batches to avoid timeouts,
 * stores state in DocumentProperties, and anonymizes user identifiers in logs.
 */

// --- CONFIGURATION ---
const SPREADSHEET_NAME = "App Passwords";
const MAX_EXECUTION_TIME_MINUTES = 5;
const MAX_USERS_PER_API_CALL = 100;
const MAX_ROWS_PER_SHEET_WRITE = 500;
const TRIGGER_HANDLER_FUNCTION = 'processUserBatchAndContinue_AppPasswords';

// --- DOCUMENT PROPERTIES SERVICE KEYS ---
const PROP_USER_PAGE_TOKEN_PREFIX = 'APP_PW_USER_PAGE_TOKEN_';
const PROP_USERS_TO_PROCESS_PREFIX = 'APP_PW_USERS_TO_PROCESS_JSON_';
const PROP_LAST_ROW_WRITTEN_PREFIX = 'APP_PW_LAST_ROW_WRITTEN_';
const PROP_IS_FIRST_RUN_PREFIX = 'APP_PW_IS_FIRST_RUN_';
const PROP_TRIGGER_ID_PREFIX = 'APP_PW_TRIGGER_ID_';

/**
 * Main function to be called by the add-on menu item.
 * Initiates or restarts the App Password inventory for the active spreadsheet.
 */
function getAppPasswords() {
  const doc = SpreadsheetApp.getActiveSpreadsheet();
  const docId = doc.getId();
  const docProperties = PropertiesService.getDocumentProperties();

  Logger.log("getAppPasswords called. Initiating App Password Inventory.");
  Logger.log(`INFO: This script will delete and recreate the "${SPREADSHEET_NAME}" sheet if it exists.`);

  const existingTriggerId = docProperties.getProperty(PROP_TRIGGER_ID_PREFIX + docId);
  if (existingTriggerId) {
    deleteTriggerById_(existingTriggerId);
    docProperties.deleteProperty(PROP_TRIGGER_ID_PREFIX + docId);
  }
  docProperties.deleteProperty(PROP_USER_PAGE_TOKEN_PREFIX + docId);
  docProperties.deleteProperty(PROP_USERS_TO_PROCESS_PREFIX + docId);
  docProperties.deleteProperty(PROP_LAST_ROW_WRITTEN_PREFIX + docId);

  docProperties.setProperty(PROP_IS_FIRST_RUN_PREFIX + docId, 'true');

  Logger.log("Previous state and trigger (if any) for this document cleared. Attempting initial batch processing directly.");
  processUserBatchAndContinue_AppPasswords();
  Logger.log("getAppPasswords initial invocation complete. If more processing is needed, a trigger has been set.");
}

/**
 * Processes batches of users and their App Passwords.
 */
function processUserBatchAndContinue_AppPasswords() {
  const doc = SpreadsheetApp.getActiveSpreadsheet();
  const docId = doc.getId();
  const docProperties = PropertiesService.getDocumentProperties();
  const scriptLock = LockService.getScriptLock();

  if (!scriptLock.tryLock(30000)) {
    Logger.log(`Could not acquire script lock for ${TRIGGER_HANDLER_FUNCTION}. Another instance for this user might be running.`);
    return;
  }

  const startTime = new Date().getTime();
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let appPasswordsSheet;
  let isCompletingSuccessfully = false;

  try {
    Logger.log(`${TRIGGER_HANDLER_FUNCTION} started.`);

    if (docProperties.getProperty(PROP_IS_FIRST_RUN_PREFIX + docId) === 'true') {
      appPasswordsSheet = spreadsheet.getSheetByName(SPREADSHEET_NAME);
      if (appPasswordsSheet) {
        Logger.log(`Deleting existing sheet: ${SPREADSHEET_NAME}`);
        spreadsheet.deleteSheet(appPasswordsSheet);
      }
      appPasswordsSheet = spreadsheet.insertSheet(SPREADSHEET_NAME, spreadsheet.getNumSheets());
      const headerRange = appPasswordsSheet.getRange("A1:E1");
      headerRange.setFontFamily("Montserrat")
        .setBackground("#fc3165")
        .setFontColor("white")
        .setFontWeight("bold")
        .setValues([["CodeID", "Name", "Creation Time", "Last Time Used", "User (Email)"]]);
      appPasswordsSheet.setFrozenRows(1);
      docProperties.setProperty(PROP_LAST_ROW_WRITTEN_PREFIX + docId, '1');
      docProperties.deleteProperty(PROP_IS_FIRST_RUN_PREFIX + docId);
      Logger.log("Sheet setup complete.");
    } else {
      appPasswordsSheet = spreadsheet.getSheetByName(SPREADSHEET_NAME);
      if (!appPasswordsSheet) {
        Logger.log(`Sheet "${SPREADSHEET_NAME}" not found. Stopping. Please run getAppPasswords() again.`);
        const triggerId = docProperties.getProperty(PROP_TRIGGER_ID_PREFIX + docId);
        if (triggerId) deleteTriggerById_(triggerId);
        cleanupDocumentProperties_(docId, docProperties);
        isCompletingSuccessfully = true;
        return;
      }
      Logger.log("Continuing process on existing sheet.");
    }

    let userPageToken = docProperties.getProperty(PROP_USER_PAGE_TOKEN_PREFIX + docId) || null;
    let usersToProcessJson = docProperties.getProperty(PROP_USERS_TO_PROCESS_PREFIX + docId);
    let usersToProcess = usersToProcessJson ? JSON.parse(usersToProcessJson) : [];
    
    const allAspData = [];

    while (true) {
      const currentTimeForLoopCheck = new Date().getTime();
      const executionTimeSecondsForLoopCheck = (currentTimeForLoopCheck - startTime) / 1000;
      if (executionTimeSecondsForLoopCheck >= (MAX_EXECUTION_TIME_MINUTES * 60 - 30)) {
          Logger.log("Time limit approaching at start of loop. Saving state and scheduling continuation.");
          docProperties.setProperty(PROP_USERS_TO_PROCESS_PREFIX + docId, JSON.stringify(usersToProcess));
          docProperties.setProperty(PROP_USER_PAGE_TOKEN_PREFIX + docId, userPageToken);
          setupContinuationTrigger_(docId, docProperties);
          return;
      }

      if (usersToProcess.length === 0) {
        if (userPageToken === 'NO_MORE_USERS') {
          Logger.log("All users have been fetched and processed.");
          isCompletingSuccessfully = true;
          break; 
        }
        
        Logger.log(`Fetching users with pageToken: ${userPageToken}`);
        const userListResponse = AdminDirectory.Users.list({
          customer: "my_customer",
          maxResults: MAX_USERS_PER_API_CALL,
          projection: "BASIC",
          viewType: "admin_view",
          orderBy: "email",
          pageToken: userPageToken,
        });

        if (userListResponse.users && userListResponse.users.length > 0) {
          usersToProcess = userListResponse.users.map(user => ({ id: user.id, primaryEmail: user.primaryEmail }));
          Logger.log(`Fetched ${usersToProcess.length} users.`);
        } else {
          Logger.log("No more users found from API (or no users in the domain).");
          usersToProcess = []; 
        }
        
        userPageToken = userListResponse.nextPageToken || 'NO_MORE_USERS';
        
        if (usersToProcess.length === 0 && userPageToken === 'NO_MORE_USERS') {
            Logger.log("All users fetched. No users in current batch and no next page. Process likely complete.");
            isCompletingSuccessfully = true;
            break;
        }
      }

      while (usersToProcess.length > 0) {
        const currentUserObject = usersToProcess.shift(); 
        const userIdForLog = currentUserObject.id;
        const userKeyForApi = currentUserObject.id;
        const userEmailForSheet = currentUserObject.primaryEmail;

        Logger.log(`Processing ASPs for user ID: ${userIdForLog}`);
        try {
          const aspsResponse = AdminDirectory.Asps.list(userKeyForApi);
          if (aspsResponse && aspsResponse.items) {
            aspsResponse.items.forEach(asp => {
              allAspData.push([
                asp.codeId,
                asp.name,
                formatTimestamp(asp.creationTime),
                asp.lastTimeUsed ? formatTimestamp(asp.lastTimeUsed) : "",
                userEmailForSheet,
              ]);
            });
          }
        } catch (e) {
          Logger.log(`Error fetching ASPs for user ID ${userIdForLog}: ${e.toString()}. Skipping user.`);
        }

        const currentTime = new Date().getTime();
        const executionTimeSeconds = (currentTime - startTime) / 1000;

        if (allAspData.length >= MAX_ROWS_PER_SHEET_WRITE || 
            (executionTimeSeconds >= (MAX_EXECUTION_TIME_MINUTES * 60 - 45))) {
            
          if (allAspData.length > 0) {
            writeDataToSheet_(appPasswordsSheet, allAspData, docProperties, docId);
            allAspData.length = 0; 
          }

          if (executionTimeSeconds >= (MAX_EXECUTION_TIME_MINUTES * 60 - 45)) {
            docProperties.setProperty(PROP_USERS_TO_PROCESS_PREFIX + docId, JSON.stringify(usersToProcess));
            docProperties.setProperty(PROP_USER_PAGE_TOKEN_PREFIX + docId, userPageToken);
            Logger.log(`Time limit approaching (${executionTimeSeconds}s). Saving state. ${usersToProcess.length} users pending in sub-batch.`);
            setupContinuationTrigger_(docId, docProperties);
            return;
          }
        }
      } 
      
      if (usersToProcess.length === 0) {
        docProperties.deleteProperty(PROP_USERS_TO_PROCESS_PREFIX + docId);
        docProperties.setProperty(PROP_USER_PAGE_TOKEN_PREFIX + docId, userPageToken);
        Logger.log("Finished current sub-batch of users.");
      }
    }

    if (isCompletingSuccessfully) {
        if (allAspData.length > 0) {
            writeDataToSheet_(appPasswordsSheet, allAspData, docProperties, docId);
        }
        Logger.log("All users processed. Finalizing sheet and cleaning up.");
        if (appPasswordsSheet.getLastRow() > 0) {
            appPasswordsSheet.autoResizeColumns(1, 5);
            const lastRowData = parseInt(docProperties.getProperty(PROP_LAST_ROW_WRITTEN_PREFIX + docId) || appPasswordsSheet.getLastRow());
            if (lastRowData > 1) {
                try {
                    appPasswordsSheet.getRange(1, 1, lastRowData, 5).createFilter();
                } catch (filterError) {
                    Logger.log(`Could not create filter: ${filterError.toString()}`);
                }
            }
        } else {
            Logger.log("Sheet is empty, skipping autoResize/filter.");
        }

        const triggerId = docProperties.getProperty(PROP_TRIGGER_ID_PREFIX + docId);
        if (triggerId) deleteTriggerById_(triggerId);
        cleanupDocumentProperties_(docId, docProperties);
        Logger.log("App Password Inventory process COMPLETE. Trigger and properties for this document deleted.");
    } else {
        Logger.log("Process loop exited unexpectedly. Saving state and setting trigger as precaution.");
        docProperties.setProperty(PROP_USERS_TO_PROCESS_PREFIX + docId, JSON.stringify(usersToProcess));
        docProperties.setProperty(PROP_USER_PAGE_TOKEN_PREFIX + docId, userPageToken);
        setupContinuationTrigger_(docId, docProperties);
    }

  } catch (e) {
    Logger.log(`CRITICAL Error in ${TRIGGER_HANDLER_FUNCTION}: ${e.toString()}\nStack: ${e.stack}`);
    try {
        Logger.log("Attempting to save state before exiting due to critical error...");
        docProperties.setProperty(PROP_USERS_TO_PROCESS_PREFIX + docId, JSON.stringify(usersToProcess));
        docProperties.setProperty(PROP_USER_PAGE_TOKEN_PREFIX + docId, userPageToken);
        setupContinuationTrigger_(docId, docProperties);
        Logger.log("State saved and trigger set for retry after critical error.");
    } catch (saveError) {
        Logger.log(`Failed to save state or set trigger after critical error: ${saveError.toString()}`);
    }
  } finally {
    scriptLock.releaseLock();
    Logger.log(`${TRIGGER_HANDLER_FUNCTION} finished its current execution slice.`);
  }
}

/**
 * Sets up a time-driven trigger for continuation.
 */
function setupContinuationTrigger_(docId, docProperties) {
  const existingTriggerId = docProperties.getProperty(PROP_TRIGGER_ID_PREFIX + docId);
  if (existingTriggerId) {
    deleteTriggerById_(existingTriggerId);
  }
  
  const newTrigger = ScriptApp.newTrigger(TRIGGER_HANDLER_FUNCTION)
    .timeBased()
    .after(20 * 1000)
    .create();
  docProperties.setProperty(PROP_TRIGGER_ID_PREFIX + docId, newTrigger.getUniqueId());
  Logger.log(`Continuation trigger ID ${newTrigger.getUniqueId()} created for ${TRIGGER_HANDLER_FUNCTION}.`);
}

/**
 * Deletes a trigger by its unique ID.
 */
function deleteTriggerById_(triggerId) {
  const allTriggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < allTriggers.length; i++) {
    if (allTriggers[i].getUniqueId() === triggerId) {
      try {
        ScriptApp.deleteTrigger(allTriggers[i]);
        Logger.log(`Deleted trigger ID: ${triggerId}`);
        return true;
      } catch (e) {
        Logger.log(`Error deleting trigger ID ${triggerId}: ${e.message}`);
        return false;
      }
    }
  }
  Logger.log(`Trigger ID ${triggerId} not found for deletion.`);
  return false;
}

/**
 * Cleans up all relevant document properties for this script.
 */
function cleanupDocumentProperties_(docId, docProperties) {
  docProperties.deleteProperty(PROP_USER_PAGE_TOKEN_PREFIX + docId);
  docProperties.deleteProperty(PROP_USERS_TO_PROCESS_PREFIX + docId);
  docProperties.deleteProperty(PROP_LAST_ROW_WRITTEN_PREFIX + docId);
  docProperties.deleteProperty(PROP_IS_FIRST_RUN_PREFIX + docId);
  docProperties.deleteProperty(PROP_TRIGGER_ID_PREFIX + docId);
  Logger.log(`All document-specific properties for App Password script cleared for docId: ${docId}.`);
}

/**
 * Writes data to the sheet and updates the last written row property.
 */
function writeDataToSheet_(sheet, data, docProperties, docId) {
  if (!data || data.length === 0) return;

  let lastRow = parseInt(docProperties.getProperty(PROP_LAST_ROW_WRITTEN_PREFIX + docId) || sheet.getLastRow());
  if (isNaN(lastRow) || lastRow < 1) lastRow = 0; 
  
  const startRow = lastRow + 1;
  sheet.getRange(startRow, 1, data.length, data[0].length).setValues(data);
  const newLastRow = startRow + data.length - 1;
  docProperties.setProperty(PROP_LAST_ROW_WRITTEN_PREFIX + docId, newLastRow.toString());
  Logger.log(`Wrote ${data.length} rows. Last written row is now ${newLastRow}.`);
}

/**
 * Formats a timestamp string into a human-readable date string.
 */
function formatTimestamp(timestampString) {
  if (!timestampString || timestampString === "0" || timestampString === 0) {
    return "Never Used";
  }
  const timestamp = parseInt(timestampString);
  if (isNaN(timestamp)) {
    Logger.log("Unknown timestamp format: " + timestampString);
    return "Invalid Timestamp";
  }
  const date = new Date(timestamp < 3000000000 ? timestamp * 1000 : timestamp);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
}