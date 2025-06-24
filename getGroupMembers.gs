function getGroupMembers() {
  const functionName = 'getGroupMembers';
  const startTime = new Date();
  Logger.log(`-- Starting ${functionName} at: ${startTime.toLocaleString()}`);

  try {
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
      Logger.log(`!! ERROR in ${functionName} (Listing Groups): ${error.message}`);
      Logger.log(`!! Error details: ${JSON.stringify(error)}`);
      throw new Error(`Failed to retrieve groups. Check permissions and API availability. ${error.message}`);
    }

    const groupMembers = [];

    for (let j = 0; j < groupEmails.length; j++) {
      let page2;
      let page2Token = "";
      
      // Add a small delay to avoid hitting rate limits on the Members.list API call
      Utilities.sleep(250);

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
              let memberStatus = member.status;
              let memberType = member.type;

              if (!memberEmail) {
                memberEmail = "All members in the organization";
                memberStatus = "ORGANIZATION";
              } else {
                const spaceRegex = /^space\//;
                if (spaceRegex.test(memberEmail)) {
                  const spaceId = memberEmail.substring(6);
                  memberEmail = "Chat Space (ID: " + spaceId + ")";
                }
                if (memberEmail.startsWith("Chat Space")) {
                  memberStatus = "ACTIVE";
                } else if (member.type !== "CUSTOMER" && !member.status) {
                  memberStatus = "EXTERNAL";
                }
              }
              const row = [groupEmails[j], memberEmail, member.role, memberStatus, memberType, member.id];
              groupMembers.push(row);
            });
          }
        } catch (error) {
          Logger.log(`!! ERROR in ${functionName} (Retrieving members for group ${groupEmails[j]}): ${error.message}`);
          Logger.log(`!! Error details: ${JSON.stringify(error)}`);
        }
        page2Token = page2 ? (page2.nextPageToken || "") : "";
      } while (page2Token);
    }

    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const sheets = spreadsheet.getSheets();
      const lastSheetIndex = sheets.length;
      let groupMembersSheet = spreadsheet.getSheetByName("Group Members");

      if (groupMembersSheet) {
        spreadsheet.deleteSheet(groupMembersSheet);
      }
      groupMembersSheet = spreadsheet.insertSheet("Group Members", lastSheetIndex);

      const headers = [["Group Email", "Member Email", "Member Role", "Member Status", "Member Type", "Member ID"]];
      groupMembersSheet.getRange("A1:F1").setValues(headers);
      groupMembersSheet.getRange("A1:F1").setFontColor("#ffffff").setFontSize(10).setFontFamily("Montserrat").setBackground("#fc3165").setFontWeight("bold");
      groupMembersSheet.setFrozenRows(1);
      
      if(groupMembersSheet.getMaxColumns() > 6) {
        groupMembersSheet.deleteColumns(7, groupMembersSheet.getMaxColumns() - 6);
      }
      
      if (groupMembers.length === 0) {
        groupMembersSheet.getRange("A2").setValue("No group members found.");
        return;
      }
      
      const lastRow = groupMembersSheet.getLastRow();
      groupMembersSheet.getRange(groupMembersSheet.getLastRow() + 1, 1, groupMembers.length, groupMembers[0].length).setValues(groupMembers);

      groupMembersSheet.setColumnWidth(3, 100);
      groupMembersSheet.setColumnWidth(4, 114);
      groupMembersSheet.setColumnWidth(5, 107);
      groupMembersSheet.autoResizeColumns(1, 2);
      groupMembersSheet.hideColumn(groupMembersSheet.getRange("F:F"));
      const noteText = "Green = Internal Member\nYellow = External Member\nRed = Inactive Member";
      groupMembersSheet.getRange("D1").setNote(noteText);

      const numRows = groupMembers.length;
      const memberEmailValues = groupMembersSheet.getRange(2, 2, numRows, 1).getValues();
      const memberStatusValues = groupMembersSheet.getRange(2, 4, numRows, 1).getValues();
      const memberTypeValues = groupMembersSheet.getRange(2, 5, numRows, 1).getValues();

      const notesData = [];
      const boldRanges = [];

      for (let i = 0; i < numRows; i++) {
        const memberStatus = memberStatusValues[i][0];
        const memberEmail = memberEmailValues[i][0];
        let note = null;

        if (memberEmail === "All members in the organization") {
          boldRanges.push(groupMembersSheet.getRange(i + 2, 2));
          note = "Organization Membership";
        } else if (memberEmail.startsWith("Chat Space")) {
          note = "Chat Space Membership";
        } else if (memberStatus === "EXTERNAL") {
          note = "External Membership";
        } else if (memberStatus === "ORGANIZATION") {
          note = "Organization Membership";
        }
        notesData.push([note]);
      }
      groupMembersSheet.getRange(2, 4, numRows, 1).setNotes(notesData);
      boldRanges.forEach(range => range.setFontWeight("bold"));

      const range = groupMembersSheet.getRange("D2:D" + (lastRow + groupMembers.length));
      const rules = [
        SpreadsheetApp.newConditionalFormatRule().whenTextContains("suspended").setBackground("#ffc9c9").setRanges([range]).build(),
        SpreadsheetApp.newConditionalFormatRule().whenTextContains("archived").setBackground("#ffc9c9").setRanges([range]).build(),
        SpreadsheetApp.newConditionalFormatRule().whenTextContains("active").setBackground("#b7e1cd").setRanges([range]).build(),
        SpreadsheetApp.newConditionalFormatRule().whenTextContains("ORGANIZATION").setBackground("#b7e1cd").setRanges([range]).build(),
        SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('OR($B2="All members in the organization", LEFT($B2, 10)="Chat Space")').setBackground("#b7e1cd").setRanges([range]).build(),
        SpreadsheetApp.newConditionalFormatRule().whenTextContains("EXTERNAL").setBackground("#fff2cc").setRanges([range]).build(),
        SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=OR(ISBLANK(C2), ISBLANK(D2), ISBLANK(A2))').setBackground("#b7e1cd").setRanges([range]).build()
      ];
      groupMembersSheet.setConditionalFormatRules(rules);

    } catch (error) {
      Logger.log(`!! ERROR in ${functionName} (Writing to spreadsheet): ${error.message}`);
    }

  } catch (error) {
    // This outer catch will now only catch errors from the API/Data Fetching part.
    Logger.log(`!! FATAL ERROR in ${functionName}: ${error.toString()}`);
    SpreadsheetApp.getUi().alert(`A critical error occurred while fetching data in ${functionName}. Check the logs for details.`);
  } finally {
    // This will always run, regardless of where an error occurred.
    const endTime = new Date();
    const duration = (endTime.getTime() - startTime.getTime()) / 1000;
    Logger.log(`-- Finished ${functionName} at: ${endTime.toLocaleString()} (Duration: ${duration.toFixed(2)}s)`);
  }
}