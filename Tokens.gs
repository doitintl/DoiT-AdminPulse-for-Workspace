/**
 * This script will list all Connected Applications to user accounts, one per row.
 * This information can be more helpful compared to Google's admin console
 * reporting of the number of users per app.
 * @OnlyCurrentDoc
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
  const tokens = [];

  // Check if there is data in row 2 and clear the sheet contents accordingly
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("OAuth Tokens");
  const dataRange = sheet.getRange(2, 1, 1, sheet.getLastColumn());
  const isDataInRow2 = dataRange.getValues().flat().some(Boolean);

  if (isDataInRow2) {
    sheet
      .getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
      .clearContent();
  }

  listAllUsers(function (user) {
    try {
      if (user.suspended) {
        console.log(
          "[suspended] %s (%s)",
          user.name.fullName,
          user.primaryEmail,
        );
        return;
      }

      const currentTokens = AdminDirectory.Tokens.list(user.primaryEmail);

      if (currentTokens && currentTokens.items && currentTokens.items.length) {
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
  const dataRange2 = sheet.getRange(2, 1, tokens.length, tokens[0].length);
  dataRange2.setValues(tokens);
  sheet.hideColumns(6);
}
