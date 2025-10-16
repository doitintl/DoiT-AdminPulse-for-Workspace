/**
 * @fileoverview Lists all non-native OAuth tokens for all users.
 * This script is designed to handle very large Google Workspace environments by using a highly
 * scalable batch processing pattern. It processes users page by page, using time-based
 * triggers to avoid exceeding script execution time limits.
 */

// --- Configuration ---
const OAUTH_TOKENS_SHEET_NAME = "OAuth Tokens";
const OAUTH_TOKENS_USER_PAGE_SIZE = 200; // Number of users to fetch in each page/batch.
const OAUTH_TOKENS_TRIGGER_FUNCTION = "processOAuthTokenBatch";

// Using a Set for highly efficient lookups of high-risk scopes.
const HIGH_RISK_SCOPES_SET = new Set([
  "https://mail.google.com/", "https://www.googleapis.com/auth/gmail.compose",
  "https://www.googleapis.com/auth/gmail.insert", "https://www.googleapis.com/auth/gmail.metadata",
  "https://www.googleapis.com/auth/gmail.modify", "https://www.googleapis.com/auth/gmail.readonly",
  "https://www.googleapis.com/auth/gmail.send", "https://www.googleapis.com/auth/gmail.settings.basic",
  "https://www.googleapis.com/auth/gmail.settings.sharing", "https://www.googleapis.com/auth/documents",
  "https://www.googleapis.com/auth/documents.readonly", "https://www.googleapis.com/auth/drive",
  "https://www.googleapis.com/auth/drive.activity", "https://www.googleapis.com/auth/drive.activity.readonly",
  "https://www.googleapis.com/auth/drive.admin", "https://www.googleapis.com/auth/drive.admin.labels",
  "https://www.googleapis.com/auth/drive.admin.labels.readonly", "https://www.googleapis.com/auth/drive.admin.readonly",
  "https://www.googleapis.com/auth/drive.admin.shareddrive", "https://www.googleapis.com/auth/drive.admin.shareddrive.readonly",
  "https://www.googleapis.com/auth/drive.apps", "https://www.googleapis.com/auth/drive.apps.readonly",
  "https://www.googleapis.com/auth/drive.categories.readonly", "https://www.googleapis.com/auth/drive.labels.readonly",
  "https://www.googleapis.com/auth/drive.meet.readonly", "https://www.googleapis.com/auth/drive.metadata",
  "https://www.googleapis.com/auth/drive.metadata.readonly", "https://www.googleapis.com/auth/drive.photos.readonly",
  "https://www.googleapis.com/auth/drive.readonly", "https://www.googleapis.com/auth/drive.scripts",
  "https://www.googleapis.com/auth/drive.teams", "https://www.googleapis.com/auth/forms.body",
  "https://www.googleapis.com/auth/forms.body.readonly", "https://www.googleapis.com/auth/forms.currentonly",
  "https://www.googleapis.com/auth/forms.responses.readonly", "https://www.googleapis.com/auth/presentations",
  "https://www.googleapis.com/auth/presentations.readonly", "https://www.googleapis.com/auth/script.addons.curation",
  "https://www.googleapis.com/auth/script.projects", "https://www.googleapis.com/auth/sites",
  "https://www.googleapis.com/auth/sites.readonly", "https://www.googleapis.com/auth/spreadsheets",
  "https://www.googleapis.com/auth/spreadsheets.readonly", "https://www.googleapis.com/auth/chat.delete",
  "https://www.googleapis.com/auth/chat.import", "https://www.googleapis.com/auth/chat.messages",
  "https://www.googleapis.com/auth/chat.messages.readonly"
]);

/**
 * Main function to be run from the menu.
 * Kicks off the OAuth Token inventory process.
 */
function getTokens() {
  const functionName = 'getOAuthTokens';
  const startTime = new Date();
  Logger.log(`-- Starting ${functionName} at: ${startTime.toLocaleString()}`);
  
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheet.toast('Starting OAuth Token inventory...', 'Setup', 10);

  try {
    // 1. Clean up from any previous runs
    _deleteTriggersByName(OAUTH_TOKENS_TRIGGER_FUNCTION);
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.deleteProperty('oauthTokens_userPageToken');
    scriptProperties.setProperty('oauthTokens_startTime', startTime.getTime());

    // 2. Set up the spreadsheet
    _setupOAuthTokensSheet();
    
    // 3. Start the first batch
    spreadsheet.toast('Starting first batch of users...', 'Processing', 10);
    processOAuthTokenBatch();

  } catch (e) {
    Logger.log(`!! FATAL ERROR in ${functionName}: ${e.toString()}\n${e.stack}`);
    SpreadsheetApp.getUi().alert(`A critical error occurred during setup: ${e.message}. Please check the logs.`);
  }
}

/**
 * Processes a batch of users to fetch their OAuth tokens.
 * This function fetches a single page of users, processes their tokens,
 * and then triggers itself for the next page.
 */
function processOAuthTokenBatch() {
  const functionName = OAUTH_TOKENS_TRIGGER_FUNCTION;
  const scriptProperties = PropertiesService.getScriptProperties();
  
  try {
    const pageToken = scriptProperties.getProperty('oauthTokens_userPageToken');
    Logger.log(`Processing user page for tokens with token: ${pageToken || ' (first page)'}`);

    // 1. Fetch a single page of users
    const userPage = AdminDirectory.Users.list({
      customer: "my_customer",
      maxResults: OAUTH_TOKENS_USER_PAGE_SIZE,
      projection: "basic",
      viewType: "admin_view",
      orderBy: "email",
      fields: "nextPageToken,users(id,primaryEmail,suspended)",
      pageToken: pageToken,
    });

    let collectedTokens = []; // {rowData: [], isHighRisk: boolean}
    if (userPage.users && userPage.users.length > 0) {
      Logger.log(`Processing tokens for ${userPage.users.length} users.`);
      // 2. Process tokens for the fetched users
      userPage.users.forEach((user) => {
        if (user.suspended) return;
        Utilities.sleep(100);
        try {
          const currentTokens = AdminDirectory.Tokens.list(user.primaryEmail);
          if (currentTokens && currentTokens.items && currentTokens.items.length > 0) {
            currentTokens.items.forEach((tok) => {
              if (!tok.nativeApp) {
                const isHighRisk = _tokenHasHighRiskScope(tok.scopes);
                collectedTokens.push({
                  rowData: [
                    user.primaryEmail, tok.displayText, tok.clientId,
                    tok.anonymous, tok.nativeApp, tok.scopes.join(" "),
                  ],
                  isHighRisk: isHighRisk
                });
              }
            });
          }
        } catch (err) {
          Logger.log(`Could not process tokens for user ID: ${user.id}. Error: ${err.message}`);
        }
      });
    }

    // 3. Write the collected data for this batch to the sheet
    if (collectedTokens.length > 0) {
      _writeTokensToSpreadsheet(collectedTokens);
    }

    // 4. Check for next page and trigger next run
    const nextPageToken = userPage.nextPageToken;
    if (nextPageToken) {
      scriptProperties.setProperty('oauthTokens_userPageToken', nextPageToken);
      _createTrigger(OAUTH_TOKENS_TRIGGER_FUNCTION, 5);
      Logger.log(`Token batch complete. Trigger created for next user page.`);
    } else {
      // 5. No more pages, finalize the process
      Logger.log("All user pages have been processed for tokens. Finalizing sheet.");
      _finalizeOAuthTokensSheet();
      
      // Clean up properties
      scriptProperties.deleteProperty('oauthTokens_userPageToken');
      
      const totalStartTime = new Date(parseInt(scriptProperties.getProperty('oauthTokens_startTime'), 10));
      const totalEndTime = new Date();
      const totalDuration = (totalEndTime.getTime() - totalStartTime.getTime()) / 1000;
      Logger.log(`-- Successfully completed OAuth Token inventory at: ${totalEndTime.toLocaleString()}. Total duration: ${totalDuration.toFixed(2)} seconds.`);
    }
  } catch (e) {
    Logger.log(`!! FATAL ERROR in ${functionName}: ${e.toString()}\n${e.stack}`);
    _deleteTriggersByName(OAUTH_TOKENS_TRIGGER_FUNCTION);
  }
}


// --- Helper Functions (Setup, Writing, Finalization) ---

/**
 * Sets up the initial state of the "OAuth Tokens" sheet.
 * @private
 */
function _setupOAuthTokensSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(OAUTH_TOKENS_SHEET_NAME);

  if (sheet) {
    sheet.clear();
    spreadsheet.deleteSheet(sheet);
  }
  
  sheet = spreadsheet.insertSheet(OAUTH_TOKENS_SHEET_NAME, 0);
  const headers = [['User Email', 'Application Name', 'Client ID', 'Is Anonymous', 'Is Native App', 'Granted Scopes']];

  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers)
    .setFontFamily("Montserrat").setBackground('#fc3165').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.getRange("F1").setNote("A light red highlighted row indicates the app uses a high-risk scope.");
}

/**
 * Writes a batch of token data to the spreadsheet, including conditional formatting.
 * @param {Array<Object>} collectedTokens An array of token objects to write.
 * @private
 */
function _writeTokensToSpreadsheet(collectedTokens) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OAUTH_TOKENS_SHEET_NAME);
  if (!sheet || collectedTokens.length === 0) return;

  const tokenDataRows = collectedTokens.map(token => token.rowData);
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, tokenDataRows.length, tokenDataRows[0].length).setValues(tokenDataRows);

  let backgroundColors = [];
  collectedTokens.forEach(token => {
    backgroundColors.push(Array(tokenDataRows[0].length).fill(token.isHighRisk ? "#f4cccc" : null));
  });
  sheet.getRange(startRow, 1, backgroundColors.length, backgroundColors[0].length).setBackgrounds(backgroundColors);
}

/**
 * Applies final formatting and cleanup to the sheet.
 * @private
 */
function _finalizeOAuthTokensSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OAUTH_TOKENS_SHEET_NAME);
  if (sheet.getLastRow() <= 1) {
    sheet.getRange("A2").setValue("No non-native OAuth tokens found for any users.");
    return;
  }

  sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).createFilter();
  
  sheet.autoResizeColumns(1, 1);
  sheet.autoResizeColumn(4);
  sheet.autoResizeColumn(5);
  sheet.setColumnWidth(2, 320); // Application Name
  sheet.setColumnWidth(3, 300); // Client ID
  sheet.setColumnWidth(6, 350); // Scopes

  const maxCols = sheet.getMaxColumns();
  const headersLength = 6;
  if (maxCols > headersLength) {
    sheet.deleteColumns(headersLength + 1, maxCols - headersLength);
  }
}

/**
 * Checks if a token's scopes contain any high-risk permissions.
 * @param {string[]} tokenScopes The scopes of a single token.
 * @returns {boolean} True if a high-risk scope is found.
 * @private
 */
function _tokenHasHighRiskScope(tokenScopes) {
  if (!tokenScopes) return false;
  for (const scope of tokenScopes) {
    if (HIGH_RISK_SCOPES_SET.has(scope)) {
      return true;
    }
  }
  return false;
}

// Note: The trigger management functions (_deleteTriggersByName, _createTrigger)
// are assumed to be present in the project, for example, from the refactored getAppPasswords.gs script.