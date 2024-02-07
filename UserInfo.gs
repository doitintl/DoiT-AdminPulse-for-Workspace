/**
 * This script will use OAUTH with an admin account to print Google Workspace user data to a Google Sheet.
 * @OnlyCurrentDoc
 */

function getUsersList() {
  const users = [];
  const userEmail = Session.getActiveUser().getEmail();
  const domain = userEmail.split("@").pop();
  const options = {
    domain: domain, // Google Workspace domain name - To Do: Pull domain from google sheet cell
    customer: "my_customer",
    maxResults: 100, //To do: figure out if this limits results to first 100 users
    projection: "FULL", // Fetch basic details of users
    viewType: "admin_view", //Admin view of users instead of domain public view
    orderBy: "email", // Sort results by users
  };

  // Check if there is data in row 2
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  const dataRange = sheet.getRange(2, 1, 1, sheet.getLastColumn());
  const isDataInRow2 = dataRange.getValues().flat().some(Boolean);

  // Clear existing data starting from row 2 if there is data
  if (isDataInRow2) {
    sheet
      .getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
      .clearContent();
  }

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
  sheet.setFrozenRows(1); // Sets headers in sheet and freezes row
  sheet.getRange(2, 1, users.length, users[0].length).setValues(users); // Adds in user info starting from row 1
}
