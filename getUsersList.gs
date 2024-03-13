/**
 * This script will use OAUTH with an admin account to print Google Workspace user data to a Google Sheet.
 * @OnlyCurrentDoc
 */

function getUsersList() {
  const users = [];
  const userEmail = Session.getActiveUser().getEmail();
  const domain = userEmail.split("@").pop();
  const options = {
    domain: domain, 
    customer: "my_customer",
    maxResults: 100, 
    projection: "FULL", 
    viewType: "admin_view", 
    orderBy: "email", 
  };

  var userSheet = SpreadsheetApp.getActiveSpreadsheet();
  var existingSheet = userSheet.getSheetByName("Users");

  if (existingSheet) {
    userSheet.deleteSheet(existingSheet);
  }

  var usersSheet = userSheet.insertSheet("Users");
  var headers = ["Name", "Email", "Super Admin", "Delegated Admin", "Suspended",
                  "Archived", "Last Login Time", "Enrolled in 2SV", "Enforced in 2SV",
                  "Org Path"];
  var headerRange = usersSheet.getRange("A1:J1");
  headerRange.setValues([headers]);
  headerRange.setFontColor("#ffffff");
  headerRange.setFontSize(10);
  headerRange.setFontFamily("Montserrat");
  headerRange.setBackground("#fc3165");
  headerRange.setFontWeight("bold");

  // Delete cells K to Z
  usersSheet.deleteColumns(11, 16); // K to Z

  let response;
  do {
    response = AdminDirectory.Users.list(options); //Types of info pulled from API https://developers.google.com/admin-sdk/directory/reference/rest/v1/users
    users.push(
      ...response.users.map((user) => [
        user.name.fullName,
        user.primaryEmail,
        user.isAdmin,
        user.isDelegatedAdmin,
        user.suspended,
        user.archived,
        user.lastLoginTime,
        user.isEnrolledIn2Sv,
        user.isEnforcedIn2Sv,
        user.orgUnitPath,
      ]),
    );

    // For domains with many users, the results are paged
    if (response.nextPageToken) {
      options.pageToken = response.nextPageToken;
    }
  } while (response.nextPageToken);

  // Insert data in a spreadsheet
  usersSheet.setFrozenRows(1); // Sets headers in sheet and freezes row
  usersSheet.getRange(2, 1, users.length, users[0].length).setValues(users); // Adds in user info starting from row 1

  // Auto resize all columns
  usersSheet.autoResizeColumns(1, usersSheet.getLastColumn());

  // Apply conditional formatting
  var rangeC = usersSheet.getRange("C2:C1000");
  var ruleC = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("TRUE")
    .setBackground("#ffb6c1")
    .setRanges([rangeC])
    .build();
  var ruleCFalse = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("FALSE")
    .setBackground("#b7e1cd")
    .setRanges([rangeC])
    .build();

  var rangeH = usersSheet.getRange("H2:H1000");
  var ruleH = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("FALSE")
    .setBackground("#ffb6c1")
    .setRanges([rangeH])
    .build();
  var ruleHFalse = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("TRUE")
    .setBackground("#b7e1cd")
    .setRanges([rangeH])
    .build();

  var rangeI = usersSheet.getRange("I2:I1000");
  var ruleI = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("FALSE")
    .setBackground("#ffb6c1")
    .setRanges([rangeI])
    .build();
  var ruleIFalse = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("TRUE")
    .setBackground("#b7e1cd")
    .setRanges([rangeI])
    .build();

  var rangeD = usersSheet.getRange("D2:D1000");
  var ruleD = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("TRUE")
    .setBackground("#ffb6c1")
    .setRanges([rangeD])
    .build();
  var ruleDFalse = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("FALSE")
    .setBackground("#b7e1cd")
    .setRanges([rangeD])
    .build();

  var rules = [ruleC, ruleCFalse, ruleH, ruleHFalse, ruleI, ruleIFalse, ruleD, ruleDFalse];
  usersSheet.setConditionalFormatRules(rules);

  // Create named range for columns B, C, D, and E
  var namedRange = userSheet.setNamedRange('UserStatus', usersSheet.getRange('B2:E1000'));
}
