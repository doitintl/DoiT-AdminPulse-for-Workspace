/** This script will inventory names of all configured App Passwords from all users in an organization.
 * @OnlyCurrentDoc
 */

function getAppPasswords() {
  const userEmail = Session.getActiveUser().getEmail();
  const domain = userEmail.split("@").pop();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("App Passwords");

  // Check if there is data in row 2 and clear the sheet contents accordingly
  const dataRange = sheet.getRange(2, 1, 1, sheet.getLastColumn());
  const isDataInRow2 = dataRange.getValues().flat().some(Boolean);

  if (isDataInRow2) {
    sheet
      .getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
      .clearContent();
  }

  // Set pagination parameters.
  let pageToken = null;

  do {
    // Make an API call to retrieve users.
    const options = {
      domain: domain,
      customer: "my_customer",
      maxResults: 100,
      projection: "FULL",
      viewType: "admin_view",
      orderBy: "email",
      pageToken: pageToken,
    };

    const response = AdminDirectory.Users.list(options);

    // Process the retrieved users.
    processUsers(response.users, sheet);

    // Update the page token for the next iteration.
    pageToken = response.nextPageToken;
  } while (pageToken);
}

function processUsers(users, sheet) {
  // Iterate over the retrieved users.
  const data = [];

  users.forEach(function (user) {
    // Retrieve app passwords for the user.
    const asps = AdminDirectory.Asps.list(user.primaryEmail);

    if (asps && asps.items) {
      asps.items.forEach(function (asp) {
        data.push([
          asp.codeId,
          asp.name,
          asp.creationTime,
          asp.lastTimeUsed,
          user.primaryEmail,
        ]);
      });
    }
  });

  // Write the data to the sheet in one go
  if (data.length > 0) {
    sheet
      .getRange(sheet.getLastRow() + 1, 1, data.length, data[0].length)
      .setValues(data);
  }
}
