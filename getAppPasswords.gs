/**
 * This script will inventory names of all configured App Passwords from all users in an organization.
 * 
 */

function getAppPasswords() {
  try { 
    const userEmail = Session.getActiveUser().getEmail();
    const domain = userEmail.split("@").pop();

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    // Check for existing sheet and handle duplicates
    let appPasswordsSheet = spreadsheet.getSheetByName("App Passwords");
    if (appPasswordsSheet) {
      // If the sheet exists, delete it first
      spreadsheet.deleteSheet(appPasswordsSheet);
    }

    // Create a new sheet with the desired name
    appPasswordsSheet = spreadsheet.insertSheet("App Passwords", spreadsheet.getNumSheets());

    // Sheet styling (concise)
    const headerRange = appPasswordsSheet.getRange("A1:E1");
    headerRange.setFontFamily("Montserrat")
      .setBackground("#fc3165")
      .setFontColor("white")
      .setFontWeight("bold")
      .setValues([["CodeID", "Name", "Creation Time", "Last Time Used", "User"]]);

    appPasswordsSheet.setFrozenRows(1);

    let pageToken = null;
    do {
      const response = AdminDirectory.Users.list({
        domain: domain,
        customer: "my_customer",
        maxResults: 100,
        projection: "FULL",
        viewType: "admin_view",
        orderBy: "email",
        pageToken: pageToken,
      });

      if (response.users && response.users.length > 0) {
        processUsers(response.users, appPasswordsSheet);
      }

      pageToken = response.nextPageToken;
    } while (pageToken);

    // Auto-resize and add filter after data is populated
    appPasswordsSheet.autoResizeColumns(1, 5);
    const lastRow = appPasswordsSheet.getLastRow();
    appPasswordsSheet.getRange('B1:E' + lastRow).createFilter();

  } catch (e) {
    if (e.message.indexOf("User does not have credentials to perform this operation") > -1) {
      // Display a modal dialog box for the error message
      const ui = SpreadsheetApp.getUi();
      ui.alert(
        'Insufficient Permissions',
        'You need Super Admin privileges to use this feature',
        ui.ButtonSet.OK
      );
      // Log the detailed error for debugging
      Logger.log(e);
    } else {
      // For other errors, re-throw the exception
      throw e;
    }
  }
}

function processUsers(users, sheet) {
  const data = [];

  users.forEach(function (user) {
    const asps = AdminDirectory.Asps.list(user.primaryEmail);

    if (asps && asps.items) {
      asps.items.forEach(function (asp) {
        data.push([
          asp.codeId,
          asp.name,
          formatTimestamp(asp.creationTime),
          asp.lastTimeUsed ? formatTimestamp(asp.lastTimeUsed) : "",
          user.primaryEmail,
        ]);
      });
    }
  });

  // Dynamic batch write:
  const batchSize = 500; // Maximum batch size

  for (let i = 0; i < data.length; i += batchSize) {
    const chunk = data.slice(i, i + batchSize); // Get the current chunk of data
    const numRows = chunk.length; // Number of rows in the current chunk

    // Write only the necessary number of rows
    sheet.getRange(sheet.getLastRow() + 1, 1, numRows, chunk[0].length)
      .setValues(chunk);
  }
}

// Corrected timestamp formatting:
function formatTimestamp(timestampString) {
  if (timestampString === "0" || timestampString === 0) {
    return "Never Used"; // Handle 0 timestamps
  } else if (typeof timestampString === "string" && timestampString.length >= 13) {
    const timestamp = parseInt(timestampString.slice(0, 13));
    const date = new Date(timestamp);
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  } else {
    Logger.log("Unknown timestamp format: " + timestampString);
    return "Invalid Timestamp";
  }
}