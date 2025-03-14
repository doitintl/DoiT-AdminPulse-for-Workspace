/**
 * This script lists all Google Groups and their associated Group Settings to a Google Sheet, including the Group ID.
 */

function getGroupsSettings() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let groupSettingsSheet = spreadsheet.getSheetByName("Group Settings");

  // Check if "Group Settings" sheet exists, delete it if it does
  if (groupSettingsSheet) {
    spreadsheet.deleteSheet(groupSettingsSheet);
  }

  // Get the index for the new sheet
  const newSheetIndex = spreadsheet.getNumSheets();

  // Add the "Group Settings" sheet at the end of the workbook
  groupSettingsSheet = spreadsheet.insertSheet("Group Settings", newSheetIndex);

  // Declare the rows array outside the loop
  let rows = [];

  // Add headers with Montserrat font, fill color, and freeze header row
  const headers = [
    "ID", // Added Group ID header
    "name",
    "email",
    "description",
    "whoCanJoin",
    "whoCanPostMessage",
    "whoCanViewMembership",
    "whoCanViewGroup",
    "whoCanDiscoverGroup",
    "allowExternalMembers",
    "allowWebPosting",
    "primaryLanguage",
    "isArchived",
    "archiveOnly",
    "messageModerationLevel",
    "spamModerationLevel",
    "replyTo",
    "customReplyTo",
    "includeCustomFooter",
    "customFooterText",
    "sendMessageDenyNotification",
    "defaultMessageDenyNotificationText",
    "membersCanPostAsTheGroup",
    "includeInGlobalAddressList",
    "whoCanLeaveGroup",
    "whoCanContactOwner",
    "favoriteRepliesOnTop",
    "whoCanBanUsers",
    "whoCanModerateMembers",
    "whoCanModerateContent",
    "whoCanAssistContent",
    "customRolesEnabledForSettingsToBeMerged",
    "enableCollaborativeInbox",
    "defaultSender",
  ];
  const headersRange = groupSettingsSheet.getRange(1, 1, 1, headers.length);
  headersRange
    .setFontFamily("Montserrat")
    .setBackground("#fc3165")
    .setFontWeight("bold")
    .setFontColor("#ffffff")
    .setValues([headers]);
  groupSettingsSheet.setFrozenRows(1);
  groupSettingsSheet.setColumnWidth(1, 180); // Adjust column width for ID
  groupSettingsSheet.setColumnWidth(2, 180);
  groupSettingsSheet.setColumnWidth(3, 275);
  groupSettingsSheet.setFrozenColumns(2);

  // Hide column A
  groupSettingsSheet.hideColumn(groupSettingsSheet.getRange("A:A"));

  // Add notes to specified cells
  groupSettingsSheet
    .getRange("E1")
    .setNote(
      "Permission to join group. Possible values are:\n" +
        "   ANYONE_CAN_JOIN: Any internet user, both inside and outside your domain, can join the group.\n" +
        "   ALL_IN_DOMAIN_CAN_JOIN: Anyone in the account domain can join. This includes accounts with multiple domains.\n" +
        "   INVITED_CAN_JOIN: Candidates for membership can be invited to join by group owners or managers.\n" +
        "   CAN_REQUEST_TO_JOIN: Non-members can request an invitation to join. Group owners or managers can then approve or reject these requests.",
    );

  groupSettingsSheet
    .getRange("F1")
    .setNote(
      "Permissions to post messages. Possible values are:\n\n" +
        "NONE_CAN_POST: The group is disabled and archived. No one can post a message to this group.\n\n" +
        "When archiveOnly is false, updating whoCanPostMessage to NONE_CAN_POST, results in an error.\n\n" +
        "If archiveOnly is reverted from true to false, whoCanPostMessages is set to ALL_MANAGERS_CAN_POST.\n\n" +
        "ALL_MANAGERS_CAN_POST: Managers, including group owners, can post messages.\n\n" +
        "ALL_MEMBERS_CAN_POST: Any group member can post a message.\n\n" +
        "ALL_OWNERS_CAN_POST: Only group owners can post a message.\n\n" +
        "ALL_IN_DOMAIN_CAN_POST: Anyone in the account can post a message.\n\n" +
        "ANYONE_CAN_POST: Any internet user who outside your account can access your Google Groups service and post a message.\n\n" +
        "Note: When whoCanPostMessage is set to ANYONE_CAN_POST, we recommend the messageModerationLevel be set to MODERATE_NON_MEMBERS to protect the group from possible spam.",
    );

  groupSettingsSheet
    .getRange("G1")
    .setNote(
      "Permissions to view membership. Possible values are:\n\n" +
        "ALL_IN_DOMAIN_CAN_VIEW: Anyone in the account can view the group members list.\n\n" +
        "If a group already has external members, those members can still send email to this group.\n\n" +
        "ALL_MEMBERS_CAN_VIEW: The group members can view the group members list.\n\n" +
        "ALL_MANAGERS_CAN_VIEW: The group managers can view group members list.",
    );

  groupSettingsSheet
    .getRange("I1")
    .setNote(
      "Specifies the set of users for whom this group is discoverable.\n\n" +
        "   ANYONE_CAN_DISCOVER: The group is discoverable by anyone searching for groups.\n\n" +
        "   ALL_IN_DOMAIN_CAN_DISCOVER: The group is only discoverable by users within the same domain as the group.\n\n" +
        "   ALL_MEMBERS_CAN_DISCOVER: The group is only discoverable by existing members of the group.",
    );

  groupSettingsSheet
    .getRange("H1")
    .setNote(
      "Permissions to view group messages. Possible values are:\n\n" +
        "ANYONE_CAN_VIEW: Any internet user can view the group's messages.\n\n" +
        "ALL_IN_DOMAIN_CAN_VIEW: Anyone in your account can view this group's messages.\n\n" +
        "ALL_MEMBERS_CAN_VIEW: All group members can view the group's messages.\n\n" +
        "ALL_MANAGERS_CAN_VIEW: Any group manager can view this group's messages.",
    );

  groupSettingsSheet
    .getRange("J1")
    .setNote(
      "Identifies whether members external to your organization can join the group.\n\n" +
        "   true: Google Workspace users external to your organization can become members of this group.\n\n" +
        "   false: Users not belonging to the organization are not allowed to become members of this group.",
    );

  groupSettingsSheet
    .getRange("K1")
    .setNote(
      "Allows posting from web.\n\n" +
        "   true: Allows any member to post to the group forum.\n\n" +
        "   false: Members only use Gmail to communicate with the group.",
    );

  groupSettingsSheet
    .getRange("L1")
    .setNote(
      "The primary language for the group.",
    );

  groupSettingsSheet
    .getRange("M1")
    .setNote(
      "Allows the Group contents to be archived.\n\n" +
        "   true: Archive messages sent to the group.\n\n" +
        "   false: Do not keep an archive of messages sent to this group.",
    );

  groupSettingsSheet
    .getRange("N1")
    .setNote(
      "Allows the group to be archived only.\n\n" +
        "   true: Group is archived and the group is inactive.\n\n" +
        "   false: The group is active and can receive messages.",
    );

  groupSettingsSheet
    .getRange("O1")
    .setNote(
      "Moderation level of incoming messages.\n\n" +
        "   MODERATE_ALL_MESSAGES: All messages are sent to the group owner's email address for approval.\n\n" +
        "   MODERATE_NON_MEMBERS: All messages from non group members are sent to the group owner's email address for approval.\n\n" +
        "   MODERATE_NEW_MEMBERS: All messages from new members are sent to the group owner's email address for approval.\n\n" +
        "   MODERATE_NONE: No moderator approval is required.",
    );

  groupSettingsSheet
    .getRange("P1")
    .setNote(
      "Specifies moderation levels for messages detected as spam.\n\n" +
        "   ALLOW: Post the message to the group.\n\n" +
        "   MODERATE: Send the message to the moderation queue.\n\n" +
        "   SILENTLY_MODERATE: Send the message to the moderation queue.\n\n" +
        "   REJECT: Immediately reject the message.",
    );

  groupSettingsSheet
    .getRange("Q1")
    .setNote(
      "Specifies who receives the default reply.\n\n" +
        "   REPLY_TO_CUSTOM: For replies to messages, use the group's custom email address.\n\n" +
        "   REPLY_TO_SENDER: The reply sent to author of message.\n\n" +
        "   REPLY_TO_LIST: This reply message is sent to the group.\n\n" +
        "   REPLY_TO_OWNER: The reply is sent to the owner(s) of the group.\n\n" +
        "   REPLY_TO_IGNORE: Group users individually decide where the message reply is sent.",
    );

  groupSettingsSheet
    .getRange("R1")
    .setNote(
      "An email address used when replying to a message if the replyTo property is set to REPLY_TO_CUSTOM.",
    );

  groupSettingsSheet
    .getRange("S1")
    .setNote(
      "Whether to include a custom footer.\n\n" +
        "   true: Include the custom footer text.\n\n" +
        "   false: Don't include a custom footer.",
    );

  groupSettingsSheet
    .getRange("T1")
    .setNote(
      "Sets the content of the custom footer text.",
    );

  groupSettingsSheet
    .getRange("U1")
    .setNote(
      "Allows a member to be notified if their message to the group is denied by the group owner.",
    );

  groupSettingsSheet
    .getRange("V1")
    .setNote(
      "Denied Message Notification Text.",
    );

  groupSettingsSheet
    .getRange("W1")
    .setNote(
      "Enables members to post messages as the group.",
    );

  groupSettingsSheet
    .getRange("X1")
    .setNote(
      "Enables the group to be included in the Global Address List.",
    );

  groupSettingsSheet
    .getRange("Y1")
    .setNote(
      "Permission to leave the group.",
    );

  groupSettingsSheet
    .getRange("Z1")
    .setNote(
      "Permission to contact owner of the group via web UI.",
    );

  groupSettingsSheet
    .getRange("AA1")
    .setNote(
      "Indicates if favorite replies should be displayed before other replies.",
    );

  groupSettingsSheet
    .getRange("AB1")
    .setNote(
      "Specifies who can deny membership to users.",
    );

  groupSettingsSheet
    .getRange("AC1")
    .setNote(
      "Specifies who can manage members.",
    );

  groupSettingsSheet
    .getRange("AD1")
    .setNote(
      "Specifies who can moderate content.",
    );

  groupSettingsSheet
    .getRange("AE1")
    .setNote(
      "Specifies who can moderate metadata.",
    );

  groupSettingsSheet
    .getRange("AF1")
    .setNote(
      "Specifies whether the group has a custom role that's included in one of the settings being merged.",
    );

  groupSettingsSheet
    .getRange("AG1")
    .setNote(
      "Specifies whether a collaborative inbox will remain turned on for the group.",
    );

  groupSettingsSheet
    .getRange("AH1")
    .setNote(
      "Default sender for members who can post messages as the group.",
    );

const rangeE = groupSettingsSheet.getRange("E2:E");
const rule1 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("INVITED_CAN_JOIN")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeE])
  .build();
const rule2 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("CAN_REQUEST_TO_JOIN")
  .setBackground("#ffc7ce") // Light red
  .setRanges([rangeE])
  .build();
const rule3 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ALL_IN_DOMAIN_CAN_JOIN")
  .setBackground("#ffc7ce") // Light red
  .setRanges([rangeE])
  .build();

const rangeF = groupSettingsSheet.getRange("F2:F");
const rule4 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ANYONE_CAN_POST")
  .setBackground("#ffc7ce") // Light red
  .setRanges([rangeF])
  .build();
const rule5 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ALL_MANAGERS_CAN_POST")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeF])
  .build();
const rule6 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ALL_IN_DOMAIN_CAN_POST")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeF])
  .build();
const rule7 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ALL_MEMBERS_CAN_POST")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeF])
  .build();
const rule8 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ALL_OWNERS_CAN_POST")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeF])
  .build();  
const rule9 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("NONE_CAN_POST")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeF])
  .build();

const rangeG = groupSettingsSheet.getRange("G2:G");
const rule10 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ALL_IN_DOMAIN_CAN_VIEW")
  .setBackground("#ffc7ce") // Light red
  .setRanges([rangeG])
  .build();
const rule11 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ALL_MEMBERS_CAN_VIEW")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeG])
  .build();
const rule12 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ALL_MANAGERS_CAN_VIEW")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeG])
  .build();
const rule13 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ALL_OWNERS_CAN_VIEW")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeG])
  .build();

const rangeH = groupSettingsSheet.getRange("H2:H");
const rule14 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ANYONE_CAN_VIEW")
  .setBackground("#ffc7ce") // Light red
  .setRanges([rangeH])
  .build();
const rule15 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ALL_IN_DOMAIN_CAN_VIEW")
  .setBackground("#ffc7ce") // Light red
  .setRanges([rangeH])
  .build();
const rule16 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ALL_MEMBERS_CAN_VIEW")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeH])
  .build();
const rule17 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ALL_MANAGERS_CAN_VIEW")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeH])
  .build();
const rule18 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ALL_OWNERS_CAN_VIEW")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeH])
  .build();
const rangeI = groupSettingsSheet.getRange("L2:L");
const rule19 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("True")
  .setBackground("#ffc7ce") // Light red
  .setRanges([rangeI])
  .build();
const rule20 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("False")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeI])
  .build();
const rangeJ = groupSettingsSheet.getRange("N2:N");
const rule21 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("False")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeJ])
  .build();
const rule22 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("True")
  .setBackground("#ffc7ce") // Light red
  .setRanges([rangeJ])
  .build();  
const rangeK = groupSettingsSheet.getRange("O2:O");
const rule23 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("MODERATE_NONE")
  .setBackground("#ffc7ce") // Light red
  .setRanges([rangeK])
  .build();
const rule24 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("MODERATE_ALL_MESSAGES")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeK])
  .build();
const rule25 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("MODERATE_NON_MEMBERS")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeK])
  .build();
const rule26 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("MODERATE_NEW_MEMBERS")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeK])
  .build();
const rangeL = groupSettingsSheet.getRange("P2:P");
const rule27 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ALLOW")
  .setBackground("#ffc7ce") // Light red
  .setRanges([rangeL])
  .build();
const rule28 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("MODERATE")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeL])
  .build();
const rule29 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("SILENTLY_MODERATE")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeL])
  .build();
const rule30 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("REJECT")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeL])
  .build();
const rangeAM = groupSettingsSheet.getRange("I2:I");
const rule31 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ANYONE_CAN_DISCOVER")
  .setBackground("#ffc7ce") // Light red
  .setRanges([rangeAM])
  .build();
const rule32 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ALL_IN_DOMAIN_CAN_DISCOVER")
  .setBackground("#ffc7ce") // Light red
  .setRanges([rangeAM])
  .build();
const rule33 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ALL_MEMBERS_CAN_DISCOVER")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeAM])
  .build();

const rangeN = groupSettingsSheet.getRange("M2:M");
const rule34 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("False")
  .setBackground("#ffc7ce") // Light red
  .setRanges([rangeN])
  .build();
const rule35 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("True")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeN])
  .build();


const rules = [
  rule1, rule2, rule3, rule4, rule5, rule6, rule7, rule8,
  rule9, rule10, rule11, rule12, rule13, rule14, rule15, rule16, rule17, rule18, rule19, rule20,
  rule21, rule22, rule23, rule24, rule25, rule26, rule27, rule28, rule29, rule30,
  rule31, rule32, rule33,rule34, rule35
];

groupSettingsSheet.setConditionalFormatRules(rules);

  // Add filters to columns E-AH
  const lastColumn = groupSettingsSheet.getLastColumn();
  const filterRange = groupSettingsSheet.getRange(1, 5, 1, Math.min(31, lastColumn - 4)); // Filters E to AH
  let filter = groupSettingsSheet.getFilter();
  if (filter) {
    filter.remove();
  }
  let criteria = SpreadsheetApp.newFilterCriteria().build();
  let dataRange = groupSettingsSheet.getDataRange();
  let filterCreated = dataRange.createFilter();
  for (let i = 5; i <= Math.min(34, lastColumn); i++) {
    filterCreated.setColumnFilterCriteria(i, criteria);
  }

  // Auto resize specified columns
  const columnsToResize = [5, 6, 7, 8, 9, 15, 17, 25, 26, 28, 29, 30, 31, 34]; // E, F, G, H, I, O, Q, Y, Z, AB, AC, AD, AE, AH
  columnsToResize.forEach(column => {
    groupSettingsSheet.autoResizeColumn(column);
  });


  let pageToken;

  do {
    const page = AdminDirectory.Groups.list({
      pageToken: pageToken,
      customer: "my_customer",
      orderBy: "email",
      sortOrder: "ASCENDING",
    });
    if (page.groups && page.groups.length > 0) {
      for (let i = 0; i < page.groups.length; i++) {
        const group = page.groups[i];
        const settings = getSettingsGroup(group.email);
        rows.push([
          group.id, // Added group ID here
          group.name,
          group.email,
          group.description,
          settings.whoCanJoin,
          settings.whoCanPostMessage,
          settings.whoCanViewMembership,
          settings.whoCanViewGroup,
          settings.whoCanDiscoverGroup,
          settings.allowExternalMembers,
          settings.allowWebPosting,
          settings.primaryLanguage,
          settings.isArchived,
          settings.archiveOnly,
          settings.messageModerationLevel,
          settings.spamModerationLevel,
          settings.replyTo,
          settings.customReplyTo,
          settings.includeCustomFooter,
          settings.customFooterText,
          settings.sendMessageDenyNotification,
          settings.defaultMessageDenyNotificationText,
          settings.membersCanPostAsTheGroup,
          settings.includeInGlobalAddressList,
          settings.whoCanLeaveGroup,
          settings.whoCanContactOwner,
          settings.favoriteRepliesOnTop,
          settings.whoCanBanUsers,
          settings.whoCanModerateMembers,
          settings.whoCanModerateContent,
          settings.whoCanAssistContent,
          settings.customRolesEnabledForSettingsToBeMerged,
          settings.enableCollaborativeInbox,
          settings.defaultSender,
        ]);
      }
      const numRows = rows.length;
      if (numRows > 0) {
        groupSettingsSheet
          .getRange(
            groupSettingsSheet.getLastRow() + 1,
            1,
            numRows,
            headers.length,
          )
          .setValues(rows);

          // Create the named range after data is added
          const lastRow = groupSettingsSheet.getLastRow();
          const range = groupSettingsSheet.getRange("A2:C" + lastRow);
          spreadsheet.setNamedRange("GroupID", range);
      }
       rows = []; // Clear the rows array
    }
    pageToken = page.nextPageToken;
  } while (pageToken);
}

function getSettingsGroup(email) {
  return AdminGroupSettings.Groups.get(email);
}