/**
 * This script will list all Connected Applications to user accounts, one per row.
 * This information can be more helpful compared to Google's admin console
 * reporting of the number of users per app.
 * 
 */

// Get all users. Specify 'domain' to filter search to one domain
function listAllUsers(cb) {
  const tokens = [];
  let pageToken;

  do {
    const page = AdminDirectory.Users.list({
      customer: "my_customer",
      orderBy: "givenName",
      maxResults: 500,
      pageToken: pageToken,
    });

    const users = page.users;

    if (users) {
      users.forEach((user) => {
        if (cb) {
          cb(user);
        }
      });
    } else {
      console.log("No users found.");
    }

    pageToken = page.nextPageToken || "";
  } while (pageToken);
}

// Gets all users and tokens
function getTokens() {
  const oAuthTokenSheetName = "OAuth Tokens";
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const lastSheetIndex = sheets.length;

  // Check if "OAuth Tokens" sheet exists and delete if it does
  let oAuthTokenSheet = spreadsheet.getSheetByName(oAuthTokenSheetName);
  if (oAuthTokenSheet) {
    spreadsheet.deleteSheet(oAuthTokenSheet);
  }

  // Create new sheet at the last index
  oAuthTokenSheet = spreadsheet.insertSheet(oAuthTokenSheetName, lastSheetIndex);

  
  // Apply font
  const headerRange = oAuthTokenSheet.getRange('A1:G1');
  const font = headerRange.getFontFamily();
  if (font != "Montserrat") {
    headerRange.setFontFamily("Montserrat");
  }
  
  // Add headers
  oAuthTokenSheet.getRange('A1:G1').setValues([['primaryEmail', 'displayText', 'clientId', 'anonymous', 'nativeApp', 'User Key', 'Scopes']]);
  
  // Format header row
  oAuthTokenSheet.getRange('A1:G1').setBackground('#fc3165').setFontColor('#ffffff').setFontWeight('bold');
  
  // Freeze header row
  oAuthTokenSheet.setFrozenRows(1);
  
  // Hide column F
  oAuthTokenSheet.hideColumns(6);
  
  // Apply conditional formatting
  const range = oAuthTokenSheet.getRange('A2:G999');
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=OR(REGEXMATCH($G2, "https://mail.google.com/"), REGEXMATCH($G2, "https://www.googleapis.com/auth/gmail.compose"), REGEXMATCH($G2, "https://www.googleapis.com/auth/gmail.insert"), REGEXMATCH($G2, "https://www.googleapis.com/auth/gmail.metadata"), REGEXMATCH($G2, "https://www.googleapis.com/auth/gmail.modify"), REGEXMATCH($G2, "https://www.googleapis.com/auth/gmail.readonly"), REGEXMATCH($G2, "https://www.googleapis.com/auth/gmail.send"), REGEXMATCH($G2, "https://www.googleapis.com/auth/gmail.settings.basic"), REGEXMATCH($G2, "https://www.googleapis.com/auth/gmail.settings.sharing"), REGEXMATCH($G2, "https://www.googleapis.com/auth/drive"), REGEXMATCH($G2, "https://www.googleapis.com/auth/drive.apps.readonly"), REGEXMATCH($G2, "https://www.googleapis.com/auth/drive.metadata"), REGEXMATCH($G2, "https://www.googleapis.com/auth/drive.metadata.readonly"), REGEXMATCH($G2, "https://www.googleapis.com/auth/drive.readonly"), REGEXMATCH($G2, "https://www.googleapis.com/auth/drive.scripts"), REGEXMATCH($G2, "https://www.googleapis.com/auth/documents"))')
    .setBackground("#f4cccc")
    .setRanges([range])
    .build();
  const rules = oAuthTokenSheet.getConditionalFormatRules();
  rules.push(rule);
  oAuthTokenSheet.setConditionalFormatRules(rules);
  
  const tokens = [];

  listAllUsers(function (user) {
    try {
      if (user.suspended) {
        console.log(
          "[suspended] %s (%s)",
          user.name.fullName,
          user.primaryEmail
        );
        return;
      }

      const currentTokens = AdminDirectory.Tokens.list(user.primaryEmail);

      if (
        currentTokens &&
        currentTokens.items &&
        currentTokens.items.length
      ) {
        currentTokens.items.forEach((tok) => {
          if (!tok.nativeApp) {
            tokens.push([
              user.primaryEmail,
              tok.displayText,
              tok.clientId,
              tok.anonymous,
              tok.nativeApp,
              tok.userKey,
              tok.scopes.join(" "),
            ]);
          }
        });
      }
    } catch (e) {
      console.log("[error] %s: %s", user.primaryEmail, e);
    }
  });

  console.log("Tokens written to Sheet Users: %s", tokens.length);
  const dataRange = oAuthTokenSheet.getRange(2, 1, tokens.length, tokens[0].length);
  dataRange.setValues(tokens);
  
  // Auto resize columns A, D, E
  oAuthTokenSheet.autoResizeColumns(1, 1);
  oAuthTokenSheet.autoResizeColumns(4, 2);

  // Resize column G
  oAuthTokenSheet.setColumnWidth(7, 320);
  oAuthTokenSheet.setColumnWidth(2, 315);
  
  // Delete columns H-Z
  oAuthTokenSheet.deleteColumns(8, 18);
  oAuthTokenSheet.deleteColumns(8);
}
