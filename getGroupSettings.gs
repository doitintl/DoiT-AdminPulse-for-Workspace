/**
 * This script lists all Google Groups and their associated Group Settings to a Google Sheet, including the Group ID.
 * This version includes performance logging, rate-limiting, and robust error handling.
 */
function getGroupsSettings() {
  const functionName = 'getGroupsSettings';
  const startTime = new Date();
  Logger.log(`-- Starting ${functionName} at: ${startTime.toLocaleString()}`);

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let groupSettingsSheet = spreadsheet.getSheetByName("Group Settings");

    if (groupSettingsSheet) {
      spreadsheet.deleteSheet(groupSettingsSheet);
    }
    groupSettingsSheet = spreadsheet.insertSheet("Group Settings", spreadsheet.getNumSheets());

    const headers = [
      "ID", "name", "email", "description", "whoCanJoin", "whoCanPostMessage", "whoCanViewMembership",
      "whoCanViewGroup", "whoCanDiscoverGroup", "allowExternalMembers", "allowWebPosting", "primaryLanguage",
      "isArchived", "archiveOnly", "messageModerationLevel", "spamModerationLevel", "replyTo",
      "customReplyTo", "includeCustomFooter", "customFooterText", "sendMessageDenyNotification",
      "defaultMessageDenyNotificationText", "membersCanPostAsTheGroup", "includeInGlobalAddressList",
      "whoCanLeaveGroup", "whoCanContactOwner", "favoriteRepliesOnTop", "whoCanBanUsers", "whoCanModerateMembers",
      "whoCanModerateContent", "whoCanAssistContent", "customRolesEnabledForSettingsToBeMerged",
      "enableCollaborativeInbox", "defaultSender",
    ];
    groupSettingsSheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontFamily("Montserrat").setBackground("#fc3165").setFontWeight("bold").setFontColor("#ffffff");
    groupSettingsSheet.setFrozenRows(1);
    groupSettingsSheet.setFrozenColumns(2);
    groupSettingsSheet.hideColumns(1); // Hide Column A (ID)

    // Notes are set once during sheet setup
    _addGroupSettingsHeaderNotes(groupSettingsSheet);

    // --- DATA FETCHING ---
    const allRows = [];
    let pageToken;
    do {
      const page = AdminDirectory.Groups.list({
        pageToken: pageToken,
        customer: "my_customer",
        orderBy: "email",
        sortOrder: "ASCENDING",
      });
      if (page.groups && page.groups.length > 0) {
        for (const group of page.groups) {
          Utilities.sleep(250); // IMPORTANT: Prevents hitting API rate limits.
          try {
            const settings = AdminGroupSettings.Groups.get(group.email);
            allRows.push([
              group.id, group.name, group.email, group.description, settings.whoCanJoin, settings.whoCanPostMessage,
              settings.whoCanViewMembership, settings.whoCanViewGroup, settings.whoCanDiscoverGroup, settings.allowExternalMembers,
              settings.allowWebPosting, settings.primaryLanguage, settings.isArchived, settings.archiveOnly,
              settings.messageModerationLevel, settings.spamModerationLevel, settings.replyTo, settings.customReplyTo,
              settings.includeCustomFooter, settings.customFooterText, settings.sendMessageDenyNotification,
              settings.defaultMessageDenyNotificationText, settings.membersCanPostAsTheGroup, settings.includeInGlobalAddressList,
              settings.whoCanLeaveGroup, settings.whoCanContactOwner, settings.favoriteRepliesOnTop, settings.whoCanBanUsers,
              settings.whoCanModerateMembers, settings.whoCanModerateContent, settings.whoCanAssistContent,
              settings.customRolesEnabledForSettingsToBeMerged, settings.enableCollaborativeInbox, settings.defaultSender,
            ]);
          } catch (e) {
            Logger.log(`Could not fetch settings for group ${group.email}. Error: ${e.message}`);
          }
        }
      }
      pageToken = page.nextPageToken;
    } while (pageToken);

    // --- DATA WRITING AND FORMATTING ---
    if (allRows.length > 0) {
      groupSettingsSheet.getRange(2, 1, allRows.length, headers.length).setValues(allRows);
      
      const lastRow = groupSettingsSheet.getLastRow();

      // Apply Conditional Formatting
      _applyGroupSettingsConditionalFormatting(groupSettingsSheet);

      // Create Named Range
      if (spreadsheet.getRangeByName('GroupID')) {
        spreadsheet.removeNamedRange('GroupID');
      }
      const dataRowCount = Math.max(1, lastRow - 1);
      spreadsheet.setNamedRange("GroupID", groupSettingsSheet.getRange("A2:C" + (dataRowCount + 1)));

      // Apply Filter
      const dataRange = groupSettingsSheet.getDataRange();
      if (dataRange.getFilter()) dataRange.getFilter().remove();
      dataRange.createFilter();

      // Auto-resize columns
      groupSettingsSheet.autoResizeColumn(2);
      groupSettingsSheet.autoResizeColumn(3);

    } else {
      groupSettingsSheet.getRange("A2").setValue("No groups found.");
      // Ensure the named range is still created even if there are no groups
      if (spreadsheet.getRangeByName('GroupID')) {
        spreadsheet.removeNamedRange('GroupID');
      }
      spreadsheet.setNamedRange("GroupID", groupSettingsSheet.getRange("A2:C2"));
    }
    
    if (groupSettingsSheet.getMaxColumns() > headers.length) {
      groupSettingsSheet.deleteColumns(headers.length + 1, groupSettingsSheet.getMaxColumns() - headers.length);
    }

  } catch (e) {
    Logger.log(`!! ERROR in ${functionName}: ${e.toString()}`);
    SpreadsheetApp.getUi().alert(`An error occurred in ${functionName}: ${e.message}`);
  } finally {
    const endTime = new Date();
    const duration = (endTime.getTime() - startTime.getTime()) / 1000;
    Logger.log(`-- Finished ${functionName} at: ${endTime.toLocaleString()} (Duration: ${duration.toFixed(2)}s)`);
  }
}

/**
 * Helper function to apply all conditional formatting rules to the Group Settings sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The target sheet.
 * @param {number} lastRow The last row with data.
 * @private
 */
function _applyGroupSettingsConditionalFormatting(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const green = "#c6efce";
  const red = "#ffc7ce";
  const rules = [];

  const addRule = (range, text, color) => {
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains(text).setBackground(color).setRanges([range]).build());
  };

  // Column E: whoCanJoin
  const rangeE = sheet.getRange("E2:E" + lastRow);
  addRule(rangeE, "INVITED_CAN_JOIN", green);
  addRule(rangeE, "CAN_REQUEST_TO_JOIN", red);
  addRule(rangeE, "ALL_IN_DOMAIN_CAN_JOIN", red);

  // Column F: whoCanPostMessage
  const rangeF = sheet.getRange("F2:F" + lastRow);
  addRule(rangeF, "ANYONE_CAN_POST", red);
  addRule(rangeF, "ALL_IN_DOMAIN_CAN_POST", green);  
  addRule(rangeF, "ALL_MANAGERS_CAN_POST", green);
  addRule(rangeF, "ALL_MEMBERS_CAN_POST", green);
  addRule(rangeF, "ALL_OWNERS_CAN_POST", green);
  addRule(rangeF, "NONE_CAN_POST", green);
  
  // Column G: whoCanViewMembership
  const rangeG = sheet.getRange("G2:G" + lastRow);
  addRule(rangeG, "ALL_IN_DOMAIN_CAN_VIEW", red);
  addRule(rangeG, "ALL_MEMBERS_CAN_VIEW", green);
  addRule(rangeG, "ALL_MANAGERS_CAN_VIEW", green);
  addRule(rangeG, "ALL_OWNERS_CAN_VIEW", green);  

  // Column H: whoCanViewGroup
  const rangeH = sheet.getRange("H2:H" + lastRow);
  addRule(rangeH, "ANYONE_CAN_VIEW", red);
  addRule(rangeH, "ALL_IN_DOMAIN_CAN_VIEW", red);
  addRule(rangeH, "ALL_MEMBERS_CAN_VIEW", green);
  addRule(rangeH, "ALL_MANAGERS_CAN_VIEW", green);
  addRule(rangeH, "ALL_OWNERS_CAN_VIEW", green);  

  // Column I: whoCanDiscoverGroup
  const rangeI = sheet.getRange("I2:I" + lastRow);
  addRule(rangeI, "ANYONE_CAN_DISCOVER", red);
  addRule(rangeI, "ALL_IN_DOMAIN_CAN_DISCOVER", red);
  addRule(rangeI, "ALL_MEMBERS_CAN_DISCOVER", green);
  
  // Column M: isArchived
  const rangeM = sheet.getRange("M2:M" + lastRow);
  addRule(rangeM, "false", red); 
  addRule(rangeM, "true", green);

  // Column N: archiveOnly
  const rangeN = sheet.getRange("N2:N" + lastRow);
  addRule(rangeN, "false", green);
  addRule(rangeN, "true", red);

  // Column O: messageModerationLevel
  const rangeO = sheet.getRange("O2:O" + lastRow);
  addRule(rangeO, "MODERATE_NONE", red);
  addRule(rangeO, "MODERATE_ALL_MESSAGES", green);
  addRule(rangeO, "MODERATE_NON_MEMBERS", green);
  addRule(rangeO, "MODERATE_NEW_MEMBERS", green);
  
  // Column P: spamModerationLevel
  const rangeP = sheet.getRange("P2:P" + lastRow);
  addRule(rangeP, "ALLOW", red);
  addRule(rangeP, "MODERATE", green);
  addRule(rangeP, "SILENTLY_MODERATE", green);
  addRule(rangeP, "REJECT", green);

  sheet.setConditionalFormatRules(rules);
}

/**
 * Helper function to add all header notes to the Group Settings sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The target sheet.
 * @private
 */
function _addGroupSettingsHeaderNotes(sheet) {
  const notes = {
    'E1': "Permission to join group. Possible values are:\n   ANYONE_CAN_JOIN: Any internet user, both inside and outside your domain, can join the group.\n   ALL_IN_DOMAIN_CAN_JOIN: Anyone in the account domain can join. This includes accounts with multiple domains.\n   INVITED_CAN_JOIN: Candidates for membership can be invited to join by group owners or managers.\n   CAN_REQUEST_TO_JOIN: Non-members can request an invitation to join. Group owners or managers can then approve or reject these requests.",
    'F1': "Permissions to post messages. Possible values are:\n\nNONE_CAN_POST: The group is disabled and archived. No one can post a message to this group.\n\nWhen archiveOnly is false, updating whoCanPostMessage to NONE_CAN_POST, results in an error.\n\nIf archiveOnly is reverted from true to false, whoCanPostMessages is set to ALL_MANAGERS_CAN_POST.\n\nALL_MANAGERS_CAN_POST: Managers, including group owners, can post messages.\n\nALL_MEMBERS_CAN_POST: Any group member can post a message.\n\nALL_OWNERS_CAN_POST: Only group owners can post a message.\n\nALL_IN_DOMAIN_CAN_POST: Anyone in the account can post a message.\n\nANYONE_CAN_POST: Any internet user who outside your account can access your Google Groups service and post a message.\n\nNote: When whoCanPostMessage is set to ANYONE_CAN_POST, we recommend the messageModerationLevel be set to MODERATE_NON_MEMBERS to protect the group from possible spam.",
    'G1': "Permissions to view membership. Possible values are:\n\nALL_IN_DOMAIN_CAN_VIEW: Anyone in the account can view the group members list.\n\nIf a group already has external members, those members can still send email to this group.\n\nALL_MEMBERS_CAN_VIEW: The group members can view the group members list.\n\nALL_MANAGERS_CAN_VIEW: The group managers can view group members list.",
    'H1': "Permissions to view group messages. Possible values are:\n\nANYONE_CAN_VIEW: Any internet user can view the group's messages.\n\nALL_IN_DOMAIN_CAN_VIEW: Anyone in your account can view this group's messages.\n\nALL_MEMBERS_CAN_VIEW: All group members can view the group's messages.\n\nALL_MANAGERS_CAN_VIEW: Any group manager can view this group's messages.",
    'I1': "Specifies the set of users for whom this group is discoverable.\n\n   ANYONE_CAN_DISCOVER: The group is discoverable by anyone searching for groups.\n\n   ALL_IN_DOMAIN_CAN_DISCOVER: The group is only discoverable by users within the same domain as the group.\n\n   ALL_MEMBERS_CAN_DISCOVER: The group is only discoverable by existing members of the group.",
    'J1': "Identifies whether members external to your organization can join the group.\n\n   true: Google Workspace users external to your organization can become members of this group.\n\n   false: Users not belonging to the organization are not allowed to become members of this group.",
    'K1': "Allows posting from web.\n\n   true: Allows any member to post to the group forum.\n\n   false: Members only use Gmail to communicate with the group.",
    'L1': "The primary language for the group.",
    'M1': "Allows the Group contents to be archived.\n\n   true: Archive messages sent to the group.\n\n   false: Do not keep an archive of messages sent to this group.",
    'N1': "Allows the group to be archived only.\n\n   true: Group is archived and the group is inactive.\n\n   false: The group is active and can receive messages.",
    'O1': "Moderation level of incoming messages.\n\n   MODERATE_ALL_MESSAGES: All messages are sent to the group owner's email address for approval.\n\n   MODERATE_NON_MEMBERS: All messages from non group members are sent to the group owner's email address for approval.\n\n   MODERATE_NEW_MEMBERS: All messages from new members are sent to the group owner's email address for approval.\n\n   MODERATE_NONE: No moderator approval is required.",
    'P1': "Specifies moderation levels for messages detected as spam.\n\n   ALLOW: Post the message to the group.\n\n   MODERATE: Send the message to the moderation queue.\n\n   SILENTLY_MODERATE: Send the message to the moderation queue.\n\n   REJECT: Immediately reject the message.",
    'Q1': "Specifies who receives the default reply.\n\n   REPLY_TO_CUSTOM: For replies to messages, use the group's custom email address.\n\n   REPLY_TO_SENDER: The reply sent to author of message.\n\n   REPLY_TO_LIST: This reply message is sent to the group.\n\n   REPLY_TO_OWNER: The reply is sent to the owner(s) of the group.\n\n   REPLY_TO_IGNORE: Group users individually decide where the message reply is sent.",
    'R1': "An email address used when replying to a message if the replyTo property is set to REPLY_TO_CUSTOM.",
    'S1': "Whether to include a custom footer.\n\n   true: Include the custom footer text.\n\n   false: Don't include a custom footer.",
    'T1': "Sets the content of the custom footer text.",
    'U1': "Allows a member to be notified if their message to the group is denied by the group owner.",
    'V1': "Denied Message Notification Text.",
    'W1': "Enables members to post messages as the group.",
    'X1': "Enables the group to be included in the Global Address List.",
    'Y1': "Permission to leave the group.",
    'Z1': "Permission to contact owner of the group via web UI.",
    'AA1': "Indicates if favorite replies should be displayed before other replies.",
    'AB1': "Specifies who can deny membership to users.",
    'AC1': "Specifies who can manage members.",
    'AD1': "Specifies who can moderate content.",
    'AE1': "Specifies who can moderate metadata.",
    'AF1': "Specifies whether the group has a custom role that's included in one of the settings being merged.",
    'AG1': "Specifies whether a collaborative inbox will remain turned on for the group.",
    'AH1': "Default sender for members who can post messages as the group."
  };
  for (const cell in notes) {
    sheet.getRange(cell).setNote(notes[cell]);
  }
}