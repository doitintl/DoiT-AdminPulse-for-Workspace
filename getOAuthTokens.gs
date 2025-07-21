/**
 * @fileoverview Lists all non-native OAuth tokens granted to users in a Google Workspace domain.
 * This script is optimized for performance and reliability in large environments.
 */

// --- Configuration ---
const CUSTOMER_ID = "my_customer";
const OAUTH_TOKEN_SHEET_NAME = "OAuth Tokens";
const SHEET_HEADERS = [['User Email', 'Application Name', 'Client ID', 'Is Anonymous', 'Is Native App', 'Granted Scopes']];

// MODIFICATION: Use a Set for highly efficient lookups.
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

// --- Main Function ---
function getTokens() {
  const functionName = 'getTokens';
  const startTime = new Date();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  spreadsheet.toast('Starting OAuth Token Audit...', 'Processing... Do not close or edit the page.', -1);
  Logger.log(`-- Starting ${functionName} at: ${startTime.toLocaleString()}`);

  try {
    let collectedTokens = []; // Will store objects: {rowData: [], isHighRisk: boolean}

    // 1. Collect all token data and determine risk in-script
    _listAllUsersPaged(function(user) {
      if (user.suspended) return;

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
    });

    // 2. Write all collected data to the spreadsheet
    spreadsheet.toast('Writing data to sheet...', 'Finalizing...', -1);
    _writeTokensToSpreadsheet(collectedTokens);

    // 3. Add final summary alert
    spreadsheet.toast('Audit complete!', 'Success!', 10);
    const alertMessage = "The OAuth Token audit is complete.\n\n" +
      "Rows highlighted in light red represent third-party applications that have been granted high-risk permissions to access Google Workspace data (e.g., read/write to Gmail, Drive or Chat).\n\n" +
      "Please review these entries carefully to ensure they are legitimate and necessary.";
    ui.alert('Audit Summary & Security Notice', alertMessage, ui.ButtonSet.OK);

  } catch (e) {
    Logger.log(`!! ERROR in ${functionName}: ${e.toString()}\n${e.stack}`);
    spreadsheet.toast('A critical error occurred. Check logs.', 'Error!', 15);
    ui.alert(`A critical error occurred in ${functionName}. Check the logs for details.`);
  } finally {
    const endTime = new Date();
    const duration = (endTime.getTime() - startTime.getTime()) / 1000;
    Logger.log(`-- Finished ${functionName} at: ${endTime.toLocaleString()} (Duration: ${duration.toFixed(2)}s)`);
  }
}

// --- Helper Functions ---

function _listAllUsersPaged(callback) {
  let pageToken;
  let pageNumber = 1;
  do {
    let page;
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast(`Fetching user page ${pageNumber}...`, 'Processing...', -1);
      page = AdminDirectory.Users.list({
        customer: CUSTOMER_ID,
        orderBy: "email",
        maxResults: 500,
        pageToken: pageToken,
        fields: "nextPageToken,users(id,primaryEmail,suspended)"
      });
    } catch (e) {
      throw new Error(`Critical error listing users from AdminDirectory: ${e.message}. Halting script.`);
    }

    if (page.users) {
      page.users.forEach((user) => {
        Utilities.sleep(100);
        try {
          callback(user);
        } catch (errCallback) {
          Logger.log(`Could not process tokens for user ID: ${user.id}. Error: ${errCallback.message}`);
        }
      });
    }
    pageToken = page.nextPageToken;
    pageNumber++;
  } while (pageToken);
}

function _writeTokensToSpreadsheet(collectedTokens) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(OAUTH_TOKEN_SHEET_NAME);

  if (sheet) {
    const oldFilter = sheet.getFilter();
    if (oldFilter) oldFilter.remove();
    sheet.clear();
  } else {
    sheet = spreadsheet.insertSheet(OAUTH_TOKEN_SHEET_NAME, 0);
  }

  sheet.getRange(1, 1, 1, SHEET_HEADERS[0].length).setValues(SHEET_HEADERS)
    .setFontFamily("Montserrat").setBackground('#fc3165').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.getRange("F1").setNote("A light red highlighted row indicates the app uses a high-risk scope.");

  if (collectedTokens.length > 0) {
    const tokenDataRows = collectedTokens.map(token => token.rowData);
    sheet.getRange(2, 1, tokenDataRows.length, tokenDataRows[0].length).setValues(tokenDataRows);

    // MODIFICATION: Apply background colors directly, which is safer than a formula.
    let backgroundColors = [];
    collectedTokens.forEach(token => {
      backgroundColors.push(Array(SHEET_HEADERS[0].length).fill(token.isHighRisk ? "#f4cccc" : null));
    });
    sheet.getRange(2, 1, backgroundColors.length, backgroundColors[0].length).setBackgrounds(backgroundColors);
    
    sheet.getRange(1, 1, sheet.getLastRow(), SHEET_HEADERS[0].length).createFilter();
  } else {
    sheet.getRange("A2").setValue("No non-native OAuth tokens found for any users.");
  }
  
  _formatSheet(sheet);
  SpreadsheetApp.flush();
}

/**
 * A new helper to check for high risk scopes using the efficient Set.
 * @param {string[]} tokenScopes The scopes of a single token.
 * @returns {boolean} True if a high-risk scope is found.
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

/**
 * A new helper for final sheet formatting and cleanup.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to format.
 */
function _formatSheet(sheet) {
  sheet.autoResizeColumns(1, 1);
  sheet.autoResizeColumn(4);
  sheet.autoResizeColumn(5);
  sheet.setColumnWidth(2, 320); // Application Name
  sheet.setColumnWidth(3, 300); // Client ID
  sheet.setColumnWidth(6, 350); // Scopes

  const maxCols = sheet.getMaxColumns();
  if (maxCols > SHEET_HEADERS[0].length) {
    sheet.deleteColumns(SHEET_HEADERS[0].length + 1, maxCols - SHEET_HEADERS[0].length);
  }
}