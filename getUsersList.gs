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

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const lastSheetIndex = sheets.length;

  // Check if "Users" sheet exists, delete it if it does
  let usersSheet = spreadsheet.getSheetByName("Users");
  if (usersSheet) {
    spreadsheet.deleteSheet(usersSheet);
  }

  // Create a new sheet at the last index
  usersSheet = spreadsheet.insertSheet("Users", lastSheetIndex);

  var headers = ["Name", "Email", "Super Admin", "Delegated Admin", "Suspended",
    "Archived", "Last Login Time", "Creation Time", "Enrolled in 2SV", "Enforced in 2SV",
    "Org Path"];
  var headerRange = usersSheet.getRange("A1:K1"); // Changed to K1 to accommodate new column
  headerRange.setValues([headers]);
  headerRange.setFontColor("#ffffff");
  headerRange.setFontSize(10);
  headerRange.setFontFamily("Montserrat");
  headerRange.setBackground("#fc3165");
  headerRange.setFontWeight("bold");

  // Delete cells L to Z
  usersSheet.deleteColumns(12, 15); // L to Z. Adjusted to account for added column.

  let response;
  do {
    response = AdminDirectory.Users.list(options); //Types of info pulled from API https://developers.google.com/admin-sdk/directory/reference/rest/v1/users
    users.push(
      ...response.users.map((user) => {
        // Format the dates.  If null or undefined, return null.
        let lastLoginTime = user.lastLoginTime ? user.lastLoginTime : null;

        // Check for the epoch time (1970-01-01T00:00:00.000Z) and replace with "Never logged in"
        if (lastLoginTime === "1970-01-01T00:00:00.000Z") {
          lastLoginTime = "Never logged in";
        } else {
          lastLoginTime = formatDate(lastLoginTime);
        }

        const creationTime = user.creationTime ? formatDate(user.creationTime) : null;

        return [
          user.name.fullName,
          user.primaryEmail,
          user.isAdmin,
          user.isDelegatedAdmin,
          user.suspended,
          user.archived,
          lastLoginTime, // Formatted Last Login Time
          creationTime,  // Formatted Creation Time
          user.isEnrolledIn2Sv,
          user.isEnforcedIn2Sv,
          user.orgUnitPath,
        ];
      }),
    );

    // For domains with many users, the results are paged
    if (response.nextPageToken) {
      options.pageToken = response.nextPageToken;
    }
  } while (response.nextPageToken);

  // Insert data in a spreadsheet
  usersSheet.setFrozenRows(1); // Sets headers in sheet and freezes row
  usersSheet.getRange(2, 1, users.length, users[0].length).setValues(users); // Adds in user info starting from row 1

  // Auto resize specific columns
  usersSheet.autoResizeColumns(2, 10);  // Auto-resize columns B through K

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

  var rangeI = usersSheet.getRange("I2:I1000"); // Changed from H to I due to the addition of creationTime.
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

  var rangeJ = usersSheet.getRange("J2:J1000"); // Changed from I to J due to the addition of creationTime.
  var ruleJ = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("FALSE")
    .setBackground("#ffb6c1")
    .setRanges([rangeJ])
    .build();
  var ruleJFalse = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("TRUE")
    .setBackground("#b7e1cd")
    .setRanges([rangeJ])
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

  var rules = [ruleC, ruleCFalse, ruleI, ruleIFalse, ruleJ, ruleJFalse, ruleD, ruleDFalse]; //Changed H to I and I to J
  usersSheet.setConditionalFormatRules(rules);

  // Create named range for columns B, C, D, and E
  var namedRange = spreadsheet.setNamedRange('UserStatus', usersSheet.getRange('B2:E1000'));

  // Add a filter to columns C-K
  try {
    const rangeForFilter = usersSheet.getRange(1, 3, usersSheet.getLastRow(), 9); // C to K (9 columns)
    let filter = usersSheet.getFilter();

    if (filter) {
      filter.remove(); //remove existing filters.

    }
    usersSheet.getRange(1, 1, usersSheet.getLastRow(), usersSheet.getLastColumn()).createFilter();

  } catch (e) {
    Logger.log(`Error adding filter: ${e}`);
  }

  // Apply yellow background to "Never logged in" cells
  const lastColumn = usersSheet.getLastColumn();
  const lastRow = usersSheet.getLastRow();
  const lastLoginColumnIndex = 7; // Column G - "Last Login Time"

  const lastLoginRange = usersSheet.getRange(2, lastLoginColumnIndex, lastRow - 1, 1); // Get range of last login values

  const lastLoginValues = lastLoginRange.getValues(); // Get all values

  for (let i = 0; i < lastLoginValues.length; i++) {
    if (lastLoginValues[i][0] === "Never logged in") { // If value is "Never logged in"
      usersSheet.getRange(i + 2, lastLoginColumnIndex).setBackground("yellow"); // Set the yellow background. Adding 2 because of index and header row.
    }
  }
}

/**
 * Formats a date string into a more readable format.
 * @param {string} dateString The date string to format (e.g., "2024-08-08T12:41:09.000Z").
 * @returns {string} The formatted date string (e.g., "August 8, 2024, 12:41 PM"), or null if the input is null/undefined.
 */
function formatDate(dateString) {
  if (!dateString) {
    return null; // Handle null or undefined input
  }

  try {
    const date = new Date(dateString);
    const options = {
      year: 'numeric',
      month: 'long',
      day: 'numeric',
      hour: 'numeric',
      minute: 'numeric',
      hour12: true, // Use 12-hour time format
    };
    return date.toLocaleString('en-US', options); // Format to US English
  } catch (e) {
    Logger.log(`Error formatting date: ${dateString} - ${e}`);
    return "Invalid Date"; // Return a user-friendly error message
  }
}