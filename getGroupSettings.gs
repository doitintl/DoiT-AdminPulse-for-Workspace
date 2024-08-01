/**
 * This script lists all Google Groups and their associated Group Settings to a Google Sheet.
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
  groupSettingsSheet.setColumnWidth(1, 180);
  groupSettingsSheet.setColumnWidth(2, 275);
  groupSettingsSheet.setFrozenColumns(2);

  // Add notes to specified cells
  groupSettingsSheet
    .getRange("D1")
    .setNote(
      "Permission to join group. Possible values are:\n" +
        "   ANYONE_CAN_JOIN: Any internet user, both inside and outside your domain, can join the group.\n" +
        "   ALL_IN_DOMAIN_CAN_JOIN: Anyone in the account domain can join. This includes accounts with multiple domains.\n" +
        "   INVITED_CAN_JOIN: Candidates for membership can be invited to join by group owners or managers.\n" +
        "   CAN_REQUEST_TO_JOIN: Non-members can request an invitation to join. Group owners or managers can then approve or reject these requests.",
    );

  groupSettingsSheet
    .getRange("E1")
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
    .getRange("F1")
    .setNote(
      "Permissions to view membership. Possible values are:\n\n" +
        "ALL_IN_DOMAIN_CAN_VIEW: Anyone in the account can view the group members list.\n\n" +
        "If a group already has external members, those members can still send email to this group.\n\n" +
        "ALL_MEMBERS_CAN_VIEW: The group members can view the group members list.\n\n" +
        "ALL_MANAGERS_CAN_VIEW: The group managers can view group members list.",
    );

  groupSettingsSheet
    .getRange("H1")
    .setNote(
      "Specifies the set of users for whom this group is discoverable. Possible values are:\n" +
        "   ANYONE_CAN_DISCOVER: The group is discoverable by anyone searching for groups.\n" +
        "   ALL_IN_DOMAIN_CAN_DISCOVER: The group is only discoverable by users within the same domain as the group.\n" +
        "   ALL_MEMBERS_CAN_DISCOVER: The group is only discoverable by existing members of the group.",
    );  

  groupSettingsSheet
    .getRange("G1")
    .setNote(
      "Permissions to view group messages. Possible values are:\n\n" +
        "ANYONE_CAN_VIEW: Any internet user can view the group's messages.\n\n" +
        "ALL_IN_DOMAIN_CAN_VIEW: Anyone in your account can view this group's messages.\n\n" +
        "ALL_MEMBERS_CAN_VIEW: All group members can view the group's messages.\n\n" +
        "ALL_MANAGERS_CAN_VIEW: Any group manager can view this group's messages.",
    );

  groupSettingsSheet
    .getRange("I1")
    .setNote(
      "Identifies whether members external to your organization can join the group. Possible values are:\n\n" +
        "true: Google Workspace users external to your organization can become members of this group.\n\n" +
        "false: Users not belonging to the organization are not allowed to become members of this group.",
    );

  groupSettingsSheet
    .getRange("J1")
    .setNote(
      "Allows posting from web. Possible values are:\n\n" +
        "true: Allows any member to post to the group forum.\n\n" +
        "false: Members only use Gmail to communicate with the group.",
    );

  groupSettingsSheet
    .getRange("K1")
    .setNote(
      "The primary language for the group. Use the language tags in the Supported languages table.",
    );

  groupSettingsSheet
    .getRange("L1")
    .setNote(
      "Allows the Group contents to be archived. Possible values are:\n\n" +
        "true: Archive messages sent to the group.\n\n" +
        "false: Do not keep an archive of messages sent to this group. If false, previously archived messages remain in the archive.",
    );

  groupSettingsSheet
    .getRange("M1")
    .setNote(
      "Allows the group to be archived only. Possible values are:\n\n" +
        "true: Group is archived and the group is inactive. New messages to this group are rejected. The older archived messages are browseable and searchable.\n\n" +
        "If true, the whoCanPostMessage property is set to NONE_CAN_POST.\n\n" +
        "If reverted from true to false, whoCanPostMessages is set to ALL_MANAGERS_CAN_POST.\n\n" +
        "false: The group is active and can receive messages.\n\n" +
        "When false, updating whoCanPostMessage to NONE_CAN_POST, results in an error.",
    );

  groupSettingsSheet
    .getRange("N1")
    .setNote(
      "Moderation level of incoming messages. Possible values are:\n\n" +
        "MODERATE_ALL_MESSAGES: All messages are sent to the group owner's email address for approval. If approved, the message is sent to the group.\n\n" +
        "MODERATE_NON_MEMBERS: All messages from non group members are sent to the group owner's email address for approval. If approved, the message is sent to the group.\n\n" +
        "MODERATE_NEW_MEMBERS: All messages from new members are sent to the group owner's email address for approval. If approved, the message is sent to the group.\n\n" +
        "MODERATE_NONE: No moderator approval is required. Messages are delivered directly to the group.",
    );

  groupSettingsSheet
    .getRange("O1")
    .setNote(
      "Specifies moderation levels for messages detected as spam. Possible values are:\n\n" +
        "ALLOW: Post the message to the group.\n\n" +
        "MODERATE: Send the message to the moderation queue. This is the default.\n\n" +
        "SILENTLY_MODERATE: Send the message to the moderation queue, but do not send notification to moderators.\n\n" +
        "REJECT: Immediately reject the message.",
    );

  groupSettingsSheet
    .getRange("P1")
    .setNote(
      "Specifies who receives the default reply. Possible values are:\n\n" +
        "REPLY_TO_CUSTOM: For replies to messages, use the group's custom email address.\n" +
        "  - When ReplyTo is set to REPLY_TO_CUSTOM, customReplyTo must have a value, otherwise an error is returned.\n\n" +
        "REPLY_TO_SENDER: The reply sent to author of message.\n\n" +
        "REPLY_TO_LIST: This reply message is sent to the group.\n\n" +
        "REPLY_TO_OWNER: The reply is sent to the owner(s) of the group. This does not include the group's managers.\n\n" +
        "REPLY_TO_IGNORE: Group users individually decide where the message reply is sent.\n\n" +
        "REPLY_TO_MANAGERS: This reply message is sent to the group's managers, which includes all managers and the group owner.",
    );

  groupSettingsSheet
    .getRange("Q1")
    .setNote(
      "An email address used when replying to a message if the replyTo property is set to REPLY_TO_CUSTOM. This address is defined by an account administrator.\n\n" +
        "When the group's ReplyTo property is set to REPLY_TO_CUSTOM, the customReplyTo property holds the custom email address used when replying to a message.\n\n" +
        "If the group's ReplyTo property is set to REPLY_TO_CUSTOM, the customReplyTo property must have a text value, otherwise an error is returned.",
    );

  groupSettingsSheet
    .getRange("R1")
    .setNote(
      "Whether to include a custom footer. Possible values are:\n\n" +
        "true: Include the custom footer text set in the `customFooterText` property.\n\n" +
        "false: Don't include a custom footer.",
    );

  groupSettingsSheet
    .getRange("S1")
    .setNote(
      "Sets the content of the custom footer text. Maximum characters: 1,000.\n\n" +
        "Note: Custom footers only appear in emails sent from the group, not when viewing messages within Google Groups.",
    );

  groupSettingsSheet
    .getRange("T1")
    .setNote(
      "Allows a member to be notified if their message to the group is denied by the group owner. Possible values are:\n\n" +
        "true: Send a notification to the message author when their message is rejected. The content of the notification is set in the `defaultMessageDenyNotificationText` property.\n\n" +
        "  - Note: The `defaultMessageDenyNotificationText` property only applies when `sendMessageDenyNotification` is set to `true`.\n\n" +
        "false: No notification is sent to the message author when their message is rejected.",
    );

  groupSettingsSheet
    .getRange("U1")
    .setNote(
      "Enables the group to be included in the Global Address List. For more information, see the help center. Possible values are:\n" +
        "   true: Group is included in the Global Address List.\n" +
        "   false: Group is not included in the Global Address List.",
    );

  groupSettingsSheet
    .getRange("V1")
    .setNote(
      "Enables members to post messages as the group. Possible values are:\n" +
        "   true: Group member can post messages using the group's email address instead of their own email address. Messages appear to originate from the group itself.\n" +
        "        Note: When true, any message moderation settings on individual users or new members do not apply to posts made on behalf of the group.\n" +
        "   false: Members cannot post in behalf of the group's email address.",
    );

  groupSettingsSheet
    .getRange("W1")
    .setNote(
      "Enables the group to be included in the Global Address List. For more information, see the help center. Possible values are:\n" +
        "   true: Group is included in the Global Address List.\n" +
        "   false: Group is not included in the Global Address List.",
    );

  groupSettingsSheet
    .getRange("X1")
    .setNote(
      "Permission to leave the group. Possible values are:\n" +
        "   ALL_MANAGERS_CAN_LEAVE: Group managers can leave the group.\n" +
        "   ALL_MEMBERS_CAN_LEAVE: All group members can leave the group.\n" +
        "   NONE_CAN_LEAVE: No one can leave the group. Group ownership can only be transferred.",
    );

  groupSettingsSheet
    .getRange("Y1")
    .setNote(
      "Permission to contact owner of the group via web UI. Possible values are:\n" +
        "   ALL_IN_DOMAIN_CAN_CONTACT: Anyone within the same domain as the group can contact the owner.\n" +
        "   ALL_MANAGERS_CAN_CONTACT: Only group managers can contact the owner.\n" +
        "   ALL_MEMBERS_CAN_CONTACT: All group members can contact the owner.\n" +
        "   ANYONE_CAN_CONTACT: Anyone can contact the owner via the web UI.",
    );

  groupSettingsSheet
    .getRange("Z1")
    .setNote(
      "Indicates if favorite replies should be displayed before other replies.\n" +
        "   true: Favorite replies are displayed at the top, above other replies.\n" +
        "   false: Favorite replies are displayed alongside other replies in the conversation order.",
    );

  groupSettingsSheet
    .getRange("AA1")
    .setNote(
      "Specifies who can deny membership to users. This permission will be deprecated once it is merged into the whoCanModerateMembers setting.\n" +
        "   ALL_MEMBERS: All group members can deny membership requests.\n" +
        "   OWNERS_AND_MANAGERS: Only group owners and managers can deny membership requests.\n" +
        "   OWNERS_ONLY: Only group owners can deny membership requests.\n" +
        "   NONE: No one can deny membership requests (automatic approval).",
    );

  groupSettingsSheet
    .getRange("AB1")
    .setNote(
      "Specifies who can manage members (approve/deny membership requests, remove members). Possible values are:\n" +
        "   ALL_MEMBERS: All group members can manage members.\n" +
        "   OWNERS_AND_MANAGERS: Only group owners and managers can manage members.\n" +
        "   OWNERS_ONLY: Only group owners can manage members.\n" +
        "   NONE: No one can manage members except the group owner (automatic approval for membership requests).",
    );

  groupSettingsSheet
    .getRange("AC1")
    .setNote(
      "Specifies who can moderate content (approve/reject/remove messages). Possible values are:\n" +
        "   ALL_MEMBERS: All group members can moderate content.\n" +
        "   OWNERS_AND_MANAGERS: Only group owners and managers can moderate content.\n" +
        "   OWNERS_ONLY: Only group owners can moderate content.\n" +
        "   NONE: No one can moderate content except the group owner (all messages are automatically posted).",
    );

  groupSettingsSheet
    .getRange("AD1")
    .setNote(
      "Specifies who can moderate metadata (tags, topics). Possible values are:\n" +
        "   ALL_MEMBERS: All group members can moderate metadata.\n" +
        "   OWNERS_AND_MANAGERS: Only group owners and managers can moderate metadata.\n" +
        "   MANAGERS_ONLY: Only group managers can moderate metadata (owners cannot).\n" +
        "   OWNERS_ONLY: Only group owners can moderate metadata.\n" +
        "   NONE: No one can moderate metadata except the group owner (metadata edits are automatic).",
    );

  groupSettingsSheet
    .getRange("AE1")
    .setNote(
      "Specifies whether the group has a custom role that's included in one of the settings being merged. This field is read-only and updates to it are ignored.\n" +
        "   true: The group has a custom role included in the settings being merged.\n" +
        "   false: The group does not have a custom role included in the settings being merged.",
    );

  groupSettingsSheet
    .getRange("AF1")
    .setNote(
      "Specifies whether a collaborative inbox will remain turned on for the group. Possible values are:\n" +
        "   true: The group will continue to use a collaborative inbox where members can see and manage emails sent to the group address.\n" +
        "   false: The group will not use a collaborative inbox. Emails sent to the group address will only be delivered to group owners.",
    );

  groupSettingsSheet
    .getRange("AG1")
    .setNote(
      "Default sender for members who can post messages as the group. Possible values are:\n" +
        "   DEFAULT_SELF: When a member with 'post as group' permission sends a message, it will appear to be sent from their own email address.\n" +
        "   GROUP: When a member with 'post as group' permission sends a message, it will appear to be sent from the group's email address.",
    );

const rangeD = groupSettingsSheet.getRange("D2:D");
const rule1 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("INVITED_CAN_JOIN")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeD])
  .build();
const rule2 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("CAN_REQUEST_TO_JOIN")
  .setBackground("#ffc7ce") // Light red
  .setRanges([rangeD])
  .build();
const rule3 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ALL_IN_DOMAIN_CAN_JOIN")
  .setBackground("#ffc7ce") // Light red
  .setRanges([rangeD])
  .build();

const rangeE = groupSettingsSheet.getRange("E2:E");
const rule4 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ANYONE_CAN_POST")
  .setBackground("#ffc7ce") // Light red
  .setRanges([rangeE])
  .build();
const rule5 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ALL_MANAGERS_CAN_POST")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeE])
  .build();
const rule6 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ALL_IN_DOMAIN_CAN_POST")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeE])
  .build();
const rule7 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ALL_MEMBERS_CAN_POST")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeE])
  .build();
const rule8 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ALL_OWNERS_CAN_POST")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeE])
  .build();  
const rule9 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("NONE_CAN_POST")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeE])
  .build();

const rangeF = groupSettingsSheet.getRange("F2:F");
const rule10 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ALL_IN_DOMAIN_CAN_VIEW")
  .setBackground("#ffc7ce") // Light red
  .setRanges([rangeF])
  .build();
const rule11 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ALL_MEMBERS_CAN_VIEW")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeF])
  .build();
const rule12 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ALL_MANAGERS_CAN_VIEW")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeF])
  .build();
const rule13 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ALL_OWNERS_CAN_VIEW")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeF])
  .build();

const rangeG = groupSettingsSheet.getRange("G2:G");
const rule14 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ANYONE_CAN_VIEW")
  .setBackground("#ffc7ce") // Light red
  .setRanges([rangeG])
  .build();
const rule15 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ALL_IN_DOMAIN_CAN_VIEW")
  .setBackground("#ffc7ce") // Light red
  .setRanges([rangeG])
  .build();
const rule16 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ALL_MEMBERS_CAN_VIEW")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeG])
  .build();
const rule17 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ALL_MANAGERS_CAN_VIEW")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeG])
  .build();
const rule18 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ALL_OWNERS_CAN_VIEW")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeG])
  .build();

const rangeH = groupSettingsSheet.getRange("K2:K");
const rule19 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("True")
  .setBackground("#ffc7ce") // Light red
  .setRanges([rangeH])
  .build();
const rule20 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("False")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeH])
  .build();
const rangeK = groupSettingsSheet.getRange("M2:M");
const rule21 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("False")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeK])
  .build();
const rule22 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("True")
  .setBackground("#ffc7ce") // Light red
  .setRanges([rangeK])
  .build();  

const rangeM = groupSettingsSheet.getRange("N2:N");
const rule23 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("MODERATE_NONE")
  .setBackground("#ffc7ce") // Light red
  .setRanges([rangeM])
  .build();
const rule24 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("MODERATE_ALL_MESSAGES")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeM])
  .build();
const rule25 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("MODERATE_NON_MEMBERS")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeM])
  .build();
const rule26 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("MODERATE_NEW_MEMBERS")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeM])
  .build();

const rangeN = groupSettingsSheet.getRange("O2:O");
const rule27 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ALLOW")
  .setBackground("#ffc7ce") // Light red
  .setRanges([rangeN])
  .build();
const rule28 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("MODERATE")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeN])
  .build();
const rule29 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("SILENTLY_MODERATE")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeN])
  .build();
const rule30 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("REJECT")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeN])
  .build();

const rangeAF = groupSettingsSheet.getRange("H2:H");
const rule31 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ANYONE_CAN_DISCOVER")
  .setBackground("#ffc7ce") // Light red
  .setRanges([rangeAF])
  .build();
const rule32 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ALL_IN_DOMAIN_CAN_DISCOVER")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeAF])
  .build();
const rule33 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("ALL_MEMBERS_CAN_DISCOVER")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeAF])
  .build();

const rangeL = groupSettingsSheet.getRange("L2:L");
const rule34 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("False")
  .setBackground("#ffc7ce") // Light red
  .setRanges([rangeL])
  .build();
const rule35 = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("True")
  .setBackground("#c6efce") // Light green
  .setRanges([rangeL])
  .build();


const rules = [
  rule1, rule2, rule3, rule4, rule5, rule6, rule7, rule8,
  rule9, rule10, rule11, rule12, rule13, rule14, rule15, rule16, rule17, rule18, rule19, rule20,
  rule21, rule22, rule23, rule24, rule25, rule26, rule27, rule28, rule29, rule30,
  rule31, rule32, rule33,rule34, rule35
];

groupSettingsSheet.setConditionalFormatRules(rules);


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
      }
    }
    pageToken = page.nextPageToken;
  } while (pageToken);

  // --- Add Filter View ---
  const lastRow = groupSettingsSheet.getLastRow(); 
  // Filter columns C through AG (3rd column to 33rd column)
  const filterRange = groupSettingsSheet.getRange('C1:AG' + lastRow); 
  filterRange.createFilter(); 

}

function getSettingsGroup(email) {
  return AdminGroupSettings.Groups.get(email);
}