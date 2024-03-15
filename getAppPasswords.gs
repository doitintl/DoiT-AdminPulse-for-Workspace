/**
 * This script will inventory names of all configured App Passwords from all users in an organization.
 * @OnlyCurrentDoc
 */

function getAppPasswords() {
  const userEmail = Session.getActiveUser().getEmail();
  const domain = userEmail.split("@").pop();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create or clear the "App Passwords" sheet
  let appPasswordsSheet = spreadsheet.getSheetByName("App Passwords");
  if (appPasswordsSheet) {
    appPasswordsSheet.clear(); // Clear existing data
  } else {
    appPasswordsSheet = spreadsheet.insertSheet("App Passwords"); // Create new sheet
  }
  
  // Set font to Montserrat
  appPasswordsSheet.getRange("A1:Z1").setFontFamily("Montserrat");

  // Add headers
  const headers = ["CodeID", "Name", "Creation Time", "Last Time Used", "User"];
  appPasswordsSheet.appendRow(headers);

  // Apply formatting to the header row
  const headerRange = appPasswordsSheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground("#fc3165").setFontColor("white").setFontWeight("bold");

  // Freeze the header row
  appPasswordsSheet.setFrozenRows(1);

  // Delete columns F-Z. deleteColumns is called twice because (6, 20) was not deleting the last column. 
  appPasswordsSheet.deleteColumns(7, 20);
  appPasswordsSheet.deleteColumns(6);

  // Set pagination parameters
  let pageToken = null;

  do {
    // Make an API call to retrieve users
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

    // Process the retrieved users
    processUsers(response.users, appPasswordsSheet);

    // Update the page token for the next iteration
    pageToken = response.nextPageToken;
  } while (pageToken);

  // Auto resize the columns
  appPasswordsSheet.autoResizeColumns(1, 5);
}

function processUsers(users, sheet) {
  // Prepare data array
  const data = [];

  // Iterate over the retrieved users
  users.forEach(function (user) {
    // Retrieve app passwords for the user
    const asps = AdminDirectory.Asps.list(user.primaryEmail);

    // Process app passwords
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

  // Write the data to the sheet in chunks to reduce API calls
  const batchSize = 1000;
  for (let i = 0; i < data.length; i += batchSize) {
    const chunk = data.slice(i, i + batchSize);
    if (chunk.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, chunk.length, chunk[0].length).setValues(chunk);
    }
  }
}
