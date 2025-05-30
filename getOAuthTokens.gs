/**
 * @fileoverview Lists all non-native OAuth tokens granted to users in a Google Workspace domain.
 * This script fetches user and token information via the Admin SDK and outputs it to a Google Sheet.
 * It provides a detailed, per-user, per-app view of token grants, highlighting those with high-risk scopes.
 */

// --- Configuration ---
/**
 * The Google Workspace Customer ID (e.g., "C0xxxxxxx") or "my_customer" for the current admin's domain.
 * @type {string}
 */
const CUSTOMER_ID = "my_customer";

/**
 * Name of the sheet where OAuth token data will be written.
 * @type {string}
 */
const OAUTH_TOKEN_SHEET_NAME = "OAuth Tokens";

/**
 * Header row for the OAuth token sheet.
 * @type {Array<string>}
 */
const SHEET_HEADERS = [['primaryEmail', 'displayText', 'clientId', 'anonymous', 'nativeApp', 'User Key', 'Scopes']];

/**
 * Array of OAuth scopes considered high-risk. These will be used to conditionally format rows in the sheet.
 * @type {string[]}
 */
const HIGH_RISK_SCOPES = [
  "https://mail.google.com/",
  "https://www.googleapis.com/auth/gmail.compose",
  "https://www.googleapis.com/auth/gmail.insert",
  "https://www.googleapis.com/auth/gmail.metadata",
  "https://www.googleapis.com/auth/gmail.modify",
  "https://www.googleapis.com/auth/gmail.readonly",
  "https://www.googleapis.com/auth/gmail.send",
  "https://www.googleapis.com/auth/gmail.settings.basic",
  "https://www.googleapis.com/auth/gmail.settings.sharing",
  "https://www.googleapis.com/auth/documents",
  "https://www.googleapis.com/auth/documents.readonly",
  "https://www.googleapis.com/auth/drive",
  "https://www.googleapis.com/auth/drive.activity",
  "https://www.googleapis.com/auth/drive.activity.readonly",
  "https://www.googleapis.com/auth/drive.admin",
  "https://www.googleapis.com/auth/drive.admin.labels",
  "https://www.googleapis.com/auth/drive.admin.labels.readonly",
  "https://www.googleapis.com/auth/drive.admin.readonly",
  "https://www.googleapis.com/auth/drive.admin.shareddrive",
  "https://www.googleapis.com/auth/drive.admin.shareddrive.readonly",
  "https://www.googleapis.com/auth/drive.apps",
  "https://www.googleapis.com/auth/drive.apps.readonly",
  "https://www.googleapis.com/auth/drive.categories.readonly",
  "https://www.googleapis.com/auth/drive.labels.readonly",
  "https://www.googleapis.com/auth/drive.meet.readonly",
  "https://www.googleapis.com/auth/drive.metadata",
  "https://www.googleapis.com/auth/drive.metadata.readonly",
  "https://www.googleapis.com/auth/drive.photos.readonly",
  "https://www.googleapis.com/auth/drive.readonly",
  "https://www.googleapis.com/auth/drive.scripts",
  "https://www.googleapis.com/auth/drive.teams",
  "https://www.googleapis.com/auth/forms.body",
  "https://www.googleapis.com/auth/forms.body.readonly",
  "https://www.googleapis.com/auth/forms.currentonly",
  "https://www.googleapis.com/auth/forms.responses.readonly",
  "https://www.googleapis.com/auth/presentations",
  "https://www.googleapis.com/auth/presentations.readonly",
  "https://www.googleapis.com/auth/script.addons.curation",
  "https://www.googleapis.com/auth/script.projects",
  "https://www.googleapis.com/auth/sites",
  "https://www.googleapis.com/auth/sites.readonly",
  "https://www.googleapis.com/auth/spreadsheets",
  "https://www.googleapis.com/auth/spreadsheets.readonly",
  "https://www.googleapis.com/auth/chat.delete",
  "https://www.googleapis.com/auth/chat.import",
  "https://www.googleapis.com/auth/chat.messages",
  "https://www.googleapis.com/auth/chat.messages.readonly"
];

/**
 * Builds the Google Sheets formula string for conditional formatting based on a list of sensitive scopes.
 * The formula checks if the target cell (containing space-separated scopes) contains any of the specified sensitive scopes.
 * @param {string[]} sensitiveScopesArray - An array of scope strings.
 * @param {string} scopeCellReference - The cell reference for the scopes in the first data row (e.g., "$G2").
 * @return {string} The OR-concatenated REGEXMATCH formula. Returns "=FALSE" if no scopes are provided.
 * @private
 */
function _buildSensitiveScopesFormula(sensitiveScopesArray, scopeCellReference = "$G2") {
  if (!sensitiveScopesArray || sensitiveScopesArray.length === 0) {
    console.warn("No sensitive scopes provided for conditional formatting. Rule will evaluate to FALSE.");
    return "=FALSE"; // No rule will be applied if the list is empty.
  }
  const regexMatchClauses = sensitiveScopesArray.map(scope => {
    // Escape double quotes within the scope string itself if they could ever occur (highly unlikely for OAuth scope URLs)
    const escapedScope = scope.replace(/"/g, '""');
    return `REGEXMATCH(TO_TEXT(${scopeCellReference}), "${escapedScope}")`;
  });
  return `=OR(${regexMatchClauses.join(", ")})`;
}

/**
 * Dynamically generated formula for conditional formatting to highlight rows with sensitive/restricted scopes.
 * It checks the 'Scopes' column (assumed to be G, starting at row 2) for any of the scopes listed in HIGH_RISK_SCOPES.
 * @type {string}
 */
const SENSITIVE_SCOPES_FORMULA = _buildSensitiveScopesFormula(HIGH_RISK_SCOPES, "$G2");


// --- Main Function ---

/**
 * Fetches all users and their OAuth tokens, then writes the data to a new sheet.
 * This is the primary function to be run.
 */
function getTokens() {
  const collectedTokensData = [];

  console.log("Starting token collection process for customer: %s", CUSTOMER_ID);

  // 1. Collect all token data
  _listAllUsersPaged(function (user) {
    try {
      if (user.suspended) {
        console.log("[INFO] Skipping suspended User ID: %s", user.id);
        return;
      }

      const currentTokens = AdminDirectory.Tokens.list(user.primaryEmail);

      if (currentTokens && currentTokens.items && currentTokens.items.length > 0) {
        currentTokens.items.forEach((tok) => {
          // Filter out native apps, as these are often less of a security concern for this type of audit.
          if (!tok.nativeApp) {
            collectedTokensData.push([
              user.primaryEmail,
              tok.displayText,
              tok.clientId,
              tok.anonymous,
              tok.nativeApp, // Will be false here due to the 'if' condition
              tok.userKey,
              tok.scopes.join(" "), // Consolidate scopes into a single string
            ]);
          }
        });
      }
    } catch (e) {
      const errorMessage = e.message + (e.stack ? `\nStack: ${e.stack}` : '');
      console.log("[ERROR] User ID: %s, error fetching tokens: %s", user.id, errorMessage);
    }
  });

  console.log("Token data collection complete. Total non-native tokens found: %s", collectedTokensData.length);
  _writeTokensToSpreadsheet(collectedTokensData);
  console.log("Script finished successfully.");
}

// --- Helper Functions ---

/**
 * Iterates through all users in the domain, invoking a callback for each user.
 * Handles pagination automatically.
 * @param {function(AdminDirectory.Schema.User):void} callback - Function to call for each user object.
 * @private
 */
function _listAllUsersPaged(callback) {
  let pageToken;

  do {
    let page;
    try {
      page = AdminDirectory.Users.list({
        customer: CUSTOMER_ID,
        orderBy: "givenName",
        maxResults: 500, // Max allowed by API, good for performance
        pageToken: pageToken,
        fields: "nextPageToken,users(id,primaryEmail,suspended)" // Request only necessary fields
      });
    } catch (e) {
      const errorMessage = e.message + (e.stack ? `\nStack: ${e.stack}` : '');
      console.error("Error listing users from AdminDirectory: %s. Halting user processing.", errorMessage);
      break; // Stop processing if user listing fails
    }

    const users = page.users;

    if (users && users.length > 0) {
      users.forEach((user) => {
        if (callback) {
          try {
            callback(user);
          } catch (errCallback) {
            // Catch errors from the callback to ensure the main loop continues for other users
            const userIdForError = user && user.id ? user.id : (user && user.primaryEmail ? user.primaryEmail : "unknown user (error in callback)");
            const cbErrorMessage = errCallback.message + (errCallback.stack ? `\nStack: ${errCallback.stack}` : '');
            console.error("Error in callback for User ID %s: %s", userIdForError, cbErrorMessage);
          }
        }
      });
    } else {
      if (!pageToken) { // Only log "No users found" on the very first attempt
        console.log("No users found for the specified customer/domain.");
      }
    }
    pageToken = page.nextPageToken;
  } while (pageToken);
}

/**
 * Handles all spreadsheet operations: creating/clearing sheet, formatting, and writing data.
 * @param {Array<Array<any>>} tokenDataRows - An array of arrays, where each inner array is a row of token data.
 * @private
 */
function _writeTokensToSpreadsheet(tokenDataRows) {
  console.log("Starting spreadsheet operations...");
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let oAuthTokenSheet = spreadsheet.getSheetByName(OAUTH_TOKEN_SHEET_NAME);

  // Delete or clear the old sheet if it exists
  if (oAuthTokenSheet) {
    try {
      spreadsheet.deleteSheet(oAuthTokenSheet);
      console.log("Deleted existing sheet: '%s'", OAUTH_TOKEN_SHEET_NAME);
      oAuthTokenSheet = null; // Clear reference after deletion
    } catch (e) {
      console.warn("Could not delete sheet '%s': %s. Will attempt to clear it instead.", OAUTH_TOKEN_SHEET_NAME, e.message);
      try {
        oAuthTokenSheet.clearContents();
        oAuthTokenSheet.clearFormats();
        // Remove all conditional formatting rules one by one
        let rules = oAuthTokenSheet.getConditionalFormatRules();
        while (rules.length > 0) {
          // Removing the first rule repeatedly. Get a fresh copy of rules in each iteration.
          oAuthTokenSheet.removeConditionalFormatRule(rules[0]);
          rules = oAuthTokenSheet.getConditionalFormatRules();
        }
        oAuthTokenSheet.setFrozenRows(0);
        // Consider unhiding columns if a full reset is needed
        console.log("Cleared contents and formats from existing sheet: '%s'", OAUTH_TOKEN_SHEET_NAME);
      } catch (eClear) {
        const clearErrorMessage = eClear.message + (eClear.stack ? `\nStack: ${eClear.stack}` : '');
        console.error("Failed to clear existing sheet '%s': %s. Halting script as sheet setup is compromised.", OAUTH_TOKEN_SHEET_NAME, clearErrorMessage);
        SpreadsheetApp.getUi().alert(`Error: Failed to prepare sheet '${OAUTH_TOKEN_SHEET_NAME}'. Please check logs. Script halted.`);
        return; // Critical error
      }
    }
  }

  // Create a new sheet if it was successfully deleted or never existed
  if (!oAuthTokenSheet) {
    const sheets = spreadsheet.getSheets();
    const lastSheetIndex = sheets.length;
    oAuthTokenSheet = spreadsheet.insertSheet(OAUTH_TOKEN_SHEET_NAME, lastSheetIndex);
    console.log("Created new sheet: '%s'", OAUTH_TOKEN_SHEET_NAME);
  } else {
    console.log("Re-using existing sheet: '%s' after clearing.", OAUTH_TOKEN_SHEET_NAME);
  }

  // Apply headers and base formatting
  oAuthTokenSheet.getRange(1, 1, 1, SHEET_HEADERS[0].length).setValues(SHEET_HEADERS)
    .setFontFamily("Montserrat")
    .setBackground('#fc3165')
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  oAuthTokenSheet.setFrozenRows(1);
  oAuthTokenSheet.hideColumns(6); // Hide Column F ('User Key') by default
  oAuthTokenSheet.getRange("G1").setNote("A light red highlighted row indicates the app uses a restricted or sensitive scope.");

  // Apply conditional formatting
  // The SENSITIVE_SCOPES_FORMULA is now generated dynamically at the top of the script
  if (SENSITIVE_SCOPES_FORMULA !== "=FALSE") { // Only apply if there are scopes to check
    const conditionalFormatRange = oAuthTokenSheet.getRange(2, 1, oAuthTokenSheet.getMaxRows() - 1, SHEET_HEADERS[0].length); // Apply from row 2 to sheet max rows
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(SENSITIVE_SCOPES_FORMULA)
      .setBackground("#f4cccc")
      .setRanges([conditionalFormatRange])
      .build();
    const rules = oAuthTokenSheet.getConditionalFormatRules(); // Get current rules (might be empty if sheet was just created)
    rules.push(rule);
    oAuthTokenSheet.setConditionalFormatRules(rules);
    console.log("Conditional formatting rule for sensitive scopes applied.");
  } else {
    console.log("No sensitive scopes defined; skipping conditional formatting rule application.");
  }


  // Write token data to the sheet
  if (tokenDataRows.length > 0) {
    const numRows = tokenDataRows.length;
    const numCols = tokenDataRows[0].length;
    oAuthTokenSheet.getRange(2, 1, numRows, numCols).setValues(tokenDataRows);
    console.log("%s token records written to sheet '%s'.", tokenDataRows.length, OAUTH_TOKEN_SHEET_NAME);
  } else {
    console.log("No token data to write to the sheet.");
  }

  // Add Filter
  const lastRowForFilter = Math.max(1, oAuthTokenSheet.getLastRow());
  const filterRange = oAuthTokenSheet.getRange(1, 1, lastRowForFilter, SHEET_HEADERS[0].length);
  const existingFilter = filterRange.getFilter();
  if (existingFilter) {
    existingFilter.remove();
  }
  filterRange.createFilter();
  console.log("Filter applied to range A1:G%s.", lastRowForFilter);

  // Adjust column widths
  oAuthTokenSheet.autoResizeColumns(1, 1); // Column A (primaryEmail)
  oAuthTokenSheet.autoResizeColumn(4);     // Column D (anonymous)
  oAuthTokenSheet.autoResizeColumn(5);     // Column E (nativeApp)

  oAuthTokenSheet.setColumnWidth(2, 320); // Column B (displayText)
  oAuthTokenSheet.setColumnWidth(3, 300); // Column C (clientId)
  oAuthTokenSheet.setColumnWidth(7, 350); // Column G (Scopes)

  // Delete unused columns (H onwards)
  const maxCols = oAuthTokenSheet.getMaxColumns();
  if (maxCols > SHEET_HEADERS[0].length) {
    oAuthTokenSheet.deleteColumns(SHEET_HEADERS[0].length + 1, maxCols - SHEET_HEADERS[0].length);
  }
  console.log("Column sizes adjusted and extra columns deleted.");

  SpreadsheetApp.flush(); // Ensure all spreadsheet changes are committed
}