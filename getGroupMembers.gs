 function getGroupMembers() {
  const groupEmails = [];
  let nextPageToken = "";

  try {
    do {
      const page = AdminDirectory.Groups.list({
        customer: 'my_customer',
        maxResults: 100,
        pageToken: nextPageToken,
      });
      const groups = page.groups;

      if (groups) {
        groups.forEach((group) => {
          groupEmails.push(group.email);
        });
      }

      nextPageToken = page.nextPageToken || "";
    } while (nextPageToken);
  } catch (error) {
    console.error(`Error listing groups: ${error.message}`);
    console.error(`Error details: ${JSON.stringify(error)}`);
    throw new Error(`Failed to retrieve groups. Check permissions and API availability. ${error.message}`);
  }

  const groupMembers = [];

  for (let j = 0; j < groupEmails.length; j++) {
    let page2;
    let page2Token = "";

    do {
      try {
        page2 = AdminDirectory.Members.list(groupEmails[j], {
          maxResults: 500,
          pageToken: page2Token,
        });
        const members = page2.members;

        if (members) {
          members.forEach((member) => {
            let memberEmail = member.email;

            if (!memberEmail) {
              memberEmail = "All members in the organization";
            } else {
              const spaceRegex = /^space\//; // Regular expression to match "space/"

              if (spaceRegex.test(memberEmail)) {
                // It's a chat space!
                const spaceId = memberEmail.substring(6); // Extract the space ID
                memberEmail = "Chat Space (ID: " + spaceId + ")"; // Display a placeholder name. Could extract real name via API if available.
              }
            }

            const row = [
              groupEmails[j],
              memberEmail,
              member.role,
              member.status,
              member.type,
              member.id,
            ];
            groupMembers.push(row);
          });
        }
      } catch (error) {
        console.error(
          `Error retrieving members for group ${groupEmails[j]}: ${error.message}`,
        );
        console.error(`Error details: ${JSON.stringify(error)}`);
      }

      page2Token = page2.nextPageToken || "";
    } while (page2Token);
  }

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    const lastSheetIndex = sheets.length;

    let groupMembersSheet = spreadsheet.getSheetByName("Group Members");

    // Check if the sheet exists and delete it if it does
    if (groupMembersSheet) {
      spreadsheet.deleteSheet(groupMembersSheet);
    }

    // Create the "Group Members" sheet at the last index
    groupMembersSheet = spreadsheet.insertSheet("Group Members", lastSheetIndex);

    // Set up the sheet with headers and formatting
    const headers = [["Group Email", "Member Email", "Member Role", "Member Status", "Member Type", "Member ID"]];
    groupMembersSheet.getRange("A1:F1").setValues(headers);
    groupMembersSheet.getRange("A1:F1").setFontColor("#ffffff").setFontSize(10).setFontFamily("Montserrat").setBackground("#fc3165").setFontWeight("bold");
    groupMembersSheet.setFrozenRows(1);
    // Delete columns G-Z
    groupMembersSheet.deleteColumns(7, 20);

    // Append data to the end of the sheet
    const lastRow = groupMembersSheet.getLastRow();
    groupMembersSheet.getRange(groupMembersSheet.getLastRow() + 1, 1, groupMembers.length, groupMembers[0].length).setValues(groupMembers);
    groupMembersSheet.autoResizeColumns(1, 6);

    // --- Optimization: Batch operations for notes ---
    const numRows = groupMembers.length;
    const memberEmailValues = groupMembersSheet.getRange(2, 2, numRows, 1).getValues();

    const notesData = []; // Array to store note values
    const boldRanges = []; // Array to store ranges for bolding "All members..."

    for (let i = 0; i < numRows; i++) {
      const memberType = groupMembers[i][4];
      const memberStatus = groupMembers[i][3];
      const memberEmail = memberEmailValues[i][0];
      let note = null;

      if (memberEmail === "All members in the organization") {
        boldRanges.push(groupMembersSheet.getRange(i + 2, 2));
      } else if (memberEmail.startsWith("Chat Space")) {
        note = "Chat Space Membership";
      } else if (memberType !== "CUSTOMER" && !memberStatus) {
        note = "External Membership";
      }

      notesData.push([note]);
    }

    // Set notes in batch
    groupMembersSheet.getRange(2, 4, numRows, 1).setNotes(notesData); // Set all notes at once

    // Apply bold formatting in batch
    boldRanges.forEach(range => range.setFontWeight("bold"));

    // Apply conditional formatting
    const range = groupMembersSheet.getRange("D2:D" + (lastRow + groupMembers.length));
    const rules = [
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains("suspended")
        .setBackground("#ffc9c9")
        .setRanges([range])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains("archived")
        .setBackground("#ffc9c9")
        .setRanges([range])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains("active")
        .setBackground("#b7e1cd")
        .setRanges([range])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('OR($B2="All members in the organization", LEFT($B2, 10)="Chat Space")')
        .setBackground("#b7e1cd") // Light green for All members AND Chat Space
        .setRanges([range])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND(ISBLANK(D2), NOT(OR($B2="All members in the organization", LEFT($B2, 10)="Chat Space")))')
        .setBackground("#fff2cc") // Yellow for external, non All members and Chat Space
        .setRanges([range])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=OR(ISBLANK(C2), ISBLANK(D2), ISBLANK(A2))')
        .setBackground("#b7e1cd")
        .setRanges([range])
        .build()
    ];
    groupMembersSheet.setConditionalFormatRules(rules);

  } catch (error) {
    console.error(
      `Error writing group members to spreadsheet: ${error.message}`,
    );
  }
}