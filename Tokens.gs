/** 
 * This script will list all Connected Applications to user accounts, one per row. 
 * @OnlyCurrentDoc
 */

// Get all users. Specify 'domain' to filter search to one domain  
function listAllUsers(cb) {
    var pageToken, page;
    do {
        page = AdminDirectory.Users.list({
            customer: 'my_customer',
            orderBy: 'givenName',
            maxResults: 500,
            pageToken: pageToken
        });

        var users = page.users;
        if (users) {
            for (var i = 0; i < users.length; i++) {
                var user = users[i];
                if (cb) {
                    cb(user)
                }
            }
        } else {
            Logger.log('No users found.');
        }
        pageToken = page.nextPageToken;
    } while (pageToken);
}

// Gets all users and tokens
function getTokens() {
    var tokens = [];

    // Check if there is data in row 2 and clear the sheet contents accordingly
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("OAuth Tokens");
    var dataRange = sheet.getRange(2, 1, 1, sheet.getLastColumn());
    var isDataInRow2 = dataRange.getValues().flat().some(Boolean);

    if (isDataInRow2) {
        sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }

    listAllUsers(function(user) {
        try {
            if (user.suspended) {
                Logger.log('[suspended] %s (%s)', user.name.fullName, user.primaryEmail);
                return;
            }

            var currentTokens = AdminDirectory.Tokens.list(user.primaryEmail);
            if (currentTokens && currentTokens.items && currentTokens.items.length) {
                for (var i = 0; i < currentTokens.items.length; i++) {
                    var tok = currentTokens.items[i];
                    if (tok.nativeApp == false) {
                        tokens.push([
                            user.primaryEmail,
                            tok.displayText,
                            tok.clientId,
                            tok.anonymous,
                            tok.nativeApp,
                            tok.userKey,
                            tok.scopes.join(' '),
                        ]);
                    }
                }
            }
        } catch (e) {
            Logger.log("[error] %s: %s", user.primaryEmail, e);
        }
    });

    Logger.log('Tokens written to Sheet Users: %s', tokens.length);
    var dataRange = sheet.getRange(2, 1, tokens.length, tokens[0].length);
    dataRange.setValues(tokens);
    sheet.hideColumns(6);
}
