function getGroupMembers() {
  const userEmail = Session.getActiveUser().getEmail();
  const domain = userEmail.split("@").pop();

  const groupEmails = [];
  let nextPageToken = "";

  do {
    const page = AdminDirectory.Groups.list({
      domain: domain,
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

  const groupMembers = [];

  for (let j = 0; j < groupEmails.length; j++) {
    let page2;
    let page2Token = "";

    do {
      try {
        page2 = AdminDirectory.Members.list(groupEmails[j], {
          domainName: domain,
          maxResults: 500,
          pageToken: page2Token,
        });
        const members = page2.members;

        if (members) {
          members.forEach((member) => {
            const row = [
              groupEmails[j],
              member.email,
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
    groupMembersSheet.getRange("A1:F1").setValues([["Group Email", "Member Email", "Member Role", "Member Status", "Member Type", "Member ID"]]);
    groupMembersSheet.getRange("A1:F1").setFontColor("#ffffff").setFontSize(10).setFontFamily("Montserrat").setBackground("#fc3165").setFontWeight    ("bold");
    groupMembersSheet.setFrozenRows(1); 

    // --- Add Note to Cell D1 ---
    groupMembersSheet.getRange("D1").setNote("A yellow highlighted row indicates a group member from an external organization.");

    // --- Add Filter View ---
    const lastRow = groupMembersSheet.getLastRow();
    const filterRange = groupMembersSheet.getRange('A1:F' + lastRow);  // Filter columns A through F, starting from row 1
    filterRange.createFilter(); 

    // Append data to the end of the sheet (starting from row 3)
    groupMembersSheet.getRange(lastRow + 1, 1, groupMembers.length, groupMembers[0].length).setValues(groupMembers);  
    groupMembersSheet.autoResizeColumns(1, 6);

    // Delete columns G-Z (starting from row 3)
    groupMembersSheet.deleteColumns(7, 20); // G to Z

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
        .whenFormulaSatisfied('=ISBLANK(D2)')
        .setBackground("#fff2cc")
        .setRanges([range])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=OR(ISBLANK(C2), ISBLANK(D2), ISBLANK(A2))')
        .setBackground("#fff2cc")
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