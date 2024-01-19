/**
 * This script will list group memberships and the user's group role, one per row.
 * @OnlyCurrentDoc
 */

function getGroupMembers() {
  const userEmail = Session.getActiveUser().getEmail();
  const domain = userEmail.split("@").pop();

  const groupEmails = []; // More descriptive variable name
  let nextPageToken;

  do {
    const page = AdminDirectory.Groups.list({ domain, maxResults: 100, pageToken: nextPageToken });
    const groups = page.groups;

    if (groups) {
      for (let i = 0; i < groups.length; i++) {
        const group = groups[i];
        groupEmails.push(group.email);
      }
    }

    nextPageToken = page.nextPageToken;
  } while (nextPageToken);

  const groupMembers = []; // More descriptive variable name
  let page2Token;

  for (let j = 0; j < groupEmails.length; j++) {
    do {
      try {
        page2 = AdminDirectory.Members.list(groupEmails[j], {
          domainName: domain,
          maxResults: 500,
          pageToken: page2Token,
        });
        const members = page2.members;

        if (members) {
          for (let i = 0; i < members.length; i++) {
            const member = members[i];
            const row = [
              groupEmails[j],
              member.email,
              member.role,
              member.status,
              member.type,
              member.id,
            ];
            groupMembers.push(row);
          }
        }
      } catch (error) {
        console.error(`Error retrieving members for group ${groupEmails[j]}: ${error.message}`);
      }

      page2Token = page2.nextPageToken;
    } while (page2Token);
  }

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName('Group Members') || spreadsheet.insertSheet('Group Members'); // Check for existing sheet and create if not found
    
    // Check if there is data in row 2 and clear the sheet contents accordingly
    const dataRange = sheet.getRange(2, 1, 1, sheet.getLastColumn());
    const isDataInRow2 = dataRange.getValues().flat().some(Boolean);

    if (isDataInRow2) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }

    sheet.setFrozenRows(1); // Freeze headers

    const lastRow = sheet.getLastRow(); // Append data to the end of the sheet
    sheet.getRange(lastRow + 1, 1, groupMembers.length, groupMembers[0].length).setValues(groupMembers);
  } catch (error) {
    console.error(`Error writing group members to spreadsheet: ${error.message}`);
  }
}
