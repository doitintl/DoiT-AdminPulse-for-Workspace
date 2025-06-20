/**
 * @fileoverview Lists all non-native OAuth tokens granted to users in a Google Workspace domain.
 * This script fetches user and token information via the Admin SDK and outputs it to a Google Sheet.
 * It provides a detailed, per-user, per-app view of token grants, highlighting those with high-risk scopes.
 */

// --- Configuration ---
const CUSTOMER_ID = "my_customer";
const OAUTH_TOKEN_SHEET_NAME = "OAuth Tokens";
const SHEET_HEADERS = [['primaryEmail', 'displayText', 'clientId', 'anonymous', 'nativeApp', 'User Key', 'Scopes']];
const HIGH_RISK_SCOPES = [
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
];

function _buildSensitiveScopesFormula(sensitiveScopesArray, scopeCellReference = "$G2") {
  if (!sensitiveScopesArray || sensitiveScopesArray.length === 0) {
    return "=FALSE";
  }
  const regexMatchClauses = sensitiveScopesArray.map(scope => {
    const escapedScope = scope.replace(/"/g, '""');
    return `REGEXMATCH(TO_TEXT(${scopeCellReference}), "${escapedScope}")`;
  });
  return `=OR(${regexMatchClauses.join(", ")})`;
}

const SENSITIVE_SCOPES_FORMULA = _buildSensitiveScopesFormula(HIGH_RISK_SCOPES, "$G2");

// --- Main Function ---
function getTokens() {
  const functionName = 'getTokens';
  const startTime = new Date();
  Logger.log(`-- Starting ${functionName} at: ${startTime.toLocaleString()}`);

  try {
    const collectedTokensData = [];

    // 1. Collect all token data
    _listAllUsersPaged(function(user) {
      try {
        if (user.suspended) {
          return;
        }
        const currentTokens = AdminDirectory.Tokens.list(user.primaryEmail);
        if (currentTokens && currentTokens.items && currentTokens.items.length > 0) {
          currentTokens.items.forEach((tok) => {
            if (!tok.nativeApp) {
              collectedTokensData.push([
                user.primaryEmail, tok.displayText, tok.clientId,
                tok.anonymous, tok.nativeApp, tok.userKey,
                tok.scopes.join(" "),
              ]);
            }
          });
        }
      } catch (e) {
        // Silently catch errors for individual users to allow the script to continue
      }
    });

    // 2. Write all collected data to the spreadsheet
    _writeTokensToSpreadsheet(collectedTokensData);

  } catch (e) {
    Logger.log(`!! ERROR in ${functionName}: ${e.toString()}`);
    SpreadsheetApp.getUi().alert(`A critical error occurred in ${functionName}. Check the logs for details.`);
  } finally {
    const endTime = new Date();
    const duration = (endTime.getTime() - startTime.getTime()) / 1000;
    Logger.log(`-- Finished ${functionName} at: ${endTime.toLocaleString()} (Duration: ${duration.toFixed(2)}s)`);
  }
}

// --- Helper Functions ---

function _listAllUsersPaged(callback) {
  let pageToken;
  do {
    let page;
    try {
      page = AdminDirectory.Users.list({
        customer: CUSTOMER_ID,
        orderBy: "givenName",
        maxResults: 500,
        pageToken: pageToken,
        fields: "nextPageToken,users(id,primaryEmail,suspended)"
      });
    } catch (e) {
      throw new Error(`Critical error listing users from AdminDirectory: ${e.message}. Halting script.`);
    }

    const users = page.users;
    if (users && users.length > 0) {
      users.forEach((user) => {
        if (callback) {
          try {
            callback(user);
          } catch (errCallback) {
            // Silently ignore errors in the callback for a single user to let the main loop continue
          }
        }
      });
    }
    pageToken = page.nextPageToken;
  } while (pageToken);
}

function _writeTokensToSpreadsheet(tokenDataRows) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let oAuthTokenSheet = spreadsheet.getSheetByName(OAUTH_TOKEN_SHEET_NAME);

  if (oAuthTokenSheet) {
    spreadsheet.deleteSheet(oAuthTokenSheet);
  }
  
  const sheets = spreadsheet.getSheets();
  const lastSheetIndex = sheets.length;
  oAuthTokenSheet = spreadsheet.insertSheet(OAUTH_TOKEN_SHEET_NAME, lastSheetIndex);

  // Apply headers and base formatting
  oAuthTokenSheet.getRange(1, 1, 1, SHEET_HEADERS[0].length).setValues(SHEET_HEADERS)
    .setFontFamily("Montserrat")
    .setBackground('#fc3165')
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  oAuthTokenSheet.setFrozenRows(1);
  oAuthTokenSheet.hideColumns(6);
  oAuthTokenSheet.getRange("G1").setNote("A light red highlighted row indicates the app uses a restricted or sensitive scope.");

  // Apply conditional formatting
  if (SENSITIVE_SCOPES_FORMULA !== "=FALSE") {
    const conditionalFormatRange = oAuthTokenSheet.getRange(2, 1, oAuthTokenSheet.getMaxRows() - 1, SHEET_HEADERS[0].length);
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(SENSITIVE_SCOPES_FORMULA)
      .setBackground("#f4cccc")
      .setRanges([conditionalFormatRange])
      .build();
    oAuthTokenSheet.setConditionalFormatRules([rule]);
  }

  // Write token data to the sheet
  if (tokenDataRows.length > 0) {
    const numRows = tokenDataRows.length;
    const numCols = tokenDataRows[0].length;
    oAuthTokenSheet.getRange(2, 1, numRows, numCols).setValues(tokenDataRows);
  } else {
    oAuthTokenSheet.getRange("A2").setValue("No non-native OAuth tokens found for any users.");
  }

  // Add Filter
  const lastRowForFilter = Math.max(1, oAuthTokenSheet.getLastRow());
  const filterRange = oAuthTokenSheet.getRange(1, 1, lastRowForFilter, SHEET_HEADERS[0].length);
  if (filterRange.getFilter()) {
    filterRange.getFilter().remove();
  }
  filterRange.createFilter();

  // Adjust column widths and clean up
  oAuthTokenSheet.autoResizeColumns(1, 1);
  oAuthTokenSheet.autoResizeColumn(4);
  oAuthTokenSheet.autoResizeColumn(5);
  oAuthTokenSheet.setColumnWidth(2, 320);
  oAuthTokenSheet.setColumnWidth(3, 300);
  oAuthTokenSheet.setColumnWidth(7, 350);

  const maxCols = oAuthTokenSheet.getMaxColumns();
  if (maxCols > SHEET_HEADERS[0].length) {
    oAuthTokenSheet.deleteColumns(SHEET_HEADERS[0].length + 1, maxCols - SHEET_HEADERS[0].length);
  }

  SpreadsheetApp.flush();
}