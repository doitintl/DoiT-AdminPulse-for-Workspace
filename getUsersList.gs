const HEADER_BG_COLOR = "#fc3165";
const HEADER_FONT_COLOR = "#ffffff";
const FONT_FAMILY = "Montserrat";
const FONT_SIZE = 10;
const SHEET_NAME = "Users"; // Name the sheet to consolidate code.

function getUsersList() {
  const startTime = new Date();
  Logger.log(`Function getUsersList started at: ${startTime.toLocaleString()}`);

  const users = [];
  const userEmail = Session.getActiveUser().getEmail();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const options = {
    customer: "my_customer", // Keep this!
    maxResults: 100,
    projection: "FULL",
    viewType: "admin_view",
    orderBy: "email",
  };

  try {
    // Get or create the "Users" sheet.
    let usersSheet = getOrCreateSheet(spreadsheet, SHEET_NAME);

    // Headers
    const headers = ["Name", "Email", "Super Admin", "Delegated Admin", "Suspended",
      "Archived", "Last Login Time", "Creation Date", "Enrolled in 2SV", "Enforced in 2SV",
      "Org Path"];
    const headerRange = usersSheet.getRange("A1:K1");
    headerRange.setValues([headers]);
    headerRange.setFontColor(HEADER_FONT_COLOR);
    headerRange.setFontSize(FONT_SIZE);
    headerRange.setFontFamily(FONT_FAMILY);
    headerRange.setBackground(HEADER_BG_COLOR);
    headerRange.setFontWeight("bold");

    // Insert data in the spreadsheet
    do {
      try {
        Utilities.sleep(200);  // Add some delay to avoid hitting rate limits
        let response = AdminDirectory.Users.list(options);

        if (response && response.users) {
          // Process users and push into the 'users' array
          users.push(
            ...response.users.map((user) => {
              const lastLoginTime = user.lastLoginTime === "1970-01-01T00:00:00.000Z" ? "Never logged in" : (user.lastLoginTime ? formatDate(user.lastLoginTime, true) : null);
              const creationTime = user.creationTime ? formatDate(user.creationTime, false) : null;

              return [
                user.name.fullName,
                user.primaryEmail,
                user.isAdmin || false,
                user.isDelegatedAdmin || false,
                user.suspended || false,
                user.archived || false,
                lastLoginTime,
                creationTime,
                user.isEnrolledIn2Sv || false,
                user.isEnforcedIn2Sv || false,
                user.orgUnitPath || ""
              ];
            })
          );
        } else {
          Logger.log("No users found in this page or invalid response.");
        }


        if (response && response.nextPageToken) {
          options.pageToken = response.nextPageToken;
        } else {
          options.pageToken = null;
        }
      } catch (apiError) {
        Logger.log(`API Error: ${apiError} - PageToken: ${options.pageToken || 'First Page'}`);
        break;
      }
    } while (options.pageToken);

    // Set the values after the loop
    if (users.length > 0) {
      usersSheet.getRange(2, 1, users.length, headers.length).setValues(users);
      usersSheet.setFrozenRows(1);
    } else {
      Logger.log("No users found to insert into the sheet.");
      usersSheet.getRange("A2").setValue("No users found."); //Display "No users found." message below header
    }

    // Set Column Widths (Batch Update)
    const columnWidths = [null, null, 112, 145, 105, 90, 135, 120, 129, 130, null]; // null for auto
    for (let i = 0; i < columnWidths.length; i++) {
      if (columnWidths[i] !== null) {
        usersSheet.setColumnWidth(i + 1, columnWidths[i]);
      } else {
        usersSheet.autoResizeColumn(i + 1);
      }
    }

    // Delete extra columns (L onwards)
    usersSheet.deleteColumns(12, usersSheet.getMaxColumns() - 11); //Deletes from L to the end of sheet.

    // Conditional Formatting (refactored for clarity)
    applyConditionalFormatting(usersSheet);

    // Re-calculate lastRow after data is written to the sheet
    const lastRow = usersSheet.getLastRow();
    // Dynamic Named Range Calculation
    const rangeForNamedRange = usersSheet.getRange(2, 2, Math.max(1, lastRow - 1), 4); // B2:E[lastRow] Use Math.max to ensure height is never 0

    // Named Range
    // First remove the existing named range
    let namedRange = spreadsheet.getNamedRanges();
    for (let i = 0; i < namedRange.length; i++) {
      if (namedRange[i].getName() == "UserStatus") {
        try {
          namedRange[i].remove();
        } catch (remove_error) {
          Logger.log("Unable to remove named range");
        }
      }
    }
    // Now create a new one.
    spreadsheet.setNamedRange('UserStatus', rangeForNamedRange);

    // Filter
    addFilter(usersSheet);

    // Highlight "Never logged in"
    highlightNeverLoggedIn(usersSheet);

    Logger.log("User list generation complete.");

  } catch (e) {
    Logger.log(`Error during user list generation: ${e}`);
    SpreadsheetApp.getUi().alert(`An error occurred: ${e}. Check the logs.`);
  } finally {
    const endTime = new Date();
    const duration = (endTime.getTime() - startTime.getTime()) / 1000; // Duration in seconds
    Logger.log(`Function getUsersList completed at: ${endTime.toLocaleString()}`);
    Logger.log(`Total execution time: ${duration.toFixed(2)} seconds.`);
  }
}

function applyConditionalFormatting(usersSheet) {
  // Clear any existing conditional formatting
  usersSheet.clearConditionalFormatRules();

  // Find the last row of *data* in column A (Name) - most reliable indicator
  const lastRow = usersSheet.getLastRow(); //Used to capture all the user records.
  const ranges = {
    "C": usersSheet.getRange("C2:C" + lastRow), // Super Admin
    "D": usersSheet.getRange("D2:D" + lastRow), // Delegated Admin
    "I": usersSheet.getRange("I2:I" + lastRow), // Enrolled in 2SV
    "J": usersSheet.getRange("J2:J" + lastRow)  // Enforced in 2SV
  };

  let rules = [];  //Create rules list

  for (const col in ranges) {
    const range = ranges[col];
    const trueRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("TRUE")
      .setBackground("#b7e1cd")
      .setRanges([range])
      .build();
    const falseRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("FALSE")
      .setBackground("#ffb6c1")
      .setRanges([range])
      .build();
    rules = rules.concat([trueRule, falseRule]); //Add the rules to the list.
  }

  usersSheet.setConditionalFormatRules(rules); //Set all the rules at once.
}

function addFilter(usersSheet) {
  try {
    const lastRow = usersSheet.getLastRow();
    const lastColumn = usersSheet.getLastColumn();

    //Data range should only include the content.  Start at first column, first row, last row of data.
    const dataRange = usersSheet.getRange(1, 1, lastRow, lastColumn);
    let filter = usersSheet.getFilter();

    if (filter) {
      filter.remove();
    }
    dataRange.createFilter();  //Create a new filter.

  } catch (e) {
    Logger.log(`Error adding filter: ${e}`);
  }
}

function highlightNeverLoggedIn(usersSheet) {
  const lastLoginColumnIndex = 7; // Column G
  const lastRow = usersSheet.getLastRow();

  const lastLoginRange = usersSheet.getRange(2, lastLoginColumnIndex, lastRow - 1, 1);
  const lastLoginValues = lastLoginRange.getValues();

  for (let i = 0; i < lastLoginValues.length; i++) {
    if (lastLoginValues[i][0] === "Never logged in") {
      usersSheet.getRange(i + 2, lastLoginColumnIndex).setBackground("yellow");
    }
  }
}

/**
 * Formats a date string into a more readable format (without timezone).
 * @param {string} dateString The date string to format (e.g., "2024-08-08T12:41:09.000Z").
 * @param {boolean} includeTime Whether to include the time in the formatted output.
 * @returns {string} The formatted date string (e.g., "1/28/24 12:41 PM"), or null if the input is null/undefined.
 */
function formatDate(dateString, includeTime) {
  if (!dateString) {
    return null;
  }

  try {
    const date = new Date(dateString);
    const month = date.getMonth() + 1;
    const day = date.getDate();
    const year = date.getFullYear().toString().slice(-2);

    let formattedDate = `${month}/${day}/${year}`;

    if (includeTime) {
      let hours = date.getHours();
      const minutes = date.getMinutes();
      const ampm = hours >= 12 ? 'PM' : 'AM';
      hours = hours % 12;
      hours = hours ? hours : 12;
      const formattedTime = `${hours}:${minutes.toString().padStart(2, '0')} ${ampm}`;
      formattedDate += ` ${formattedTime}`;
    }

    return formattedDate;
  } catch (e) {
    Logger.log(`Error formatting date: ${dateString} - ${e}`);
    return "Invalid Date";
  }
}

function getOrCreateSheet(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (sheet) {
    try {
      spreadsheet.deleteSheet(sheet);  //Delete existing sheet
    } catch (err) {
      Logger.log("Unable to delete sheet with name " + sheetName + ".")
    }

  }
  sheet = spreadsheet.insertSheet(sheetName, spreadsheet.getSheets().length); //Insert new sheet at end
  return sheet;  //Return the sheet
}