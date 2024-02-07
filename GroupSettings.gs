/**
 * This script will use OAUTH list all Google Groups and the associated Group Settings to a Google Sheet.
 * @OnlyCurrentDoc
 */

function getGroupsSettings() {
  const rep = [
    [
      "name",
      "email",
      "description",
      "whoCanJoin",
      "whoCanPostMessage",
      "whoCanViewMembership",
      "whoCanViewGroup",
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
      "whoCanDiscoverGroup",
      "defaultSender",
    ],
  ];
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
        rep.push([
          group.name,
          group.email,
          group.description,
          settings.whoCanJoin,
          settings.whoCanPostMessage,
          settings.whoCanViewMembership,
          settings.whoCanViewGroup,
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
          settings.whoCanDiscoverGroup,
          settings.defaultSender,
        ]);
      }
    }
    pageToken = page.nextPageToken;
  } while (pageToken);

  // Insert data in a spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Group Settings");

  // Check if there is data in row 2 and clear the sheet contents accordingly
  const dataRange = sheet.getRange(2, 1, 1, sheet.getLastColumn());
  const isDataInRow2 = dataRange.getValues().flat().some(Boolean);

  if (isDataInRow2) {
    sheet
      .getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
      .clearContent();
  }

  sheet.getRange(1, 1, rep.length, rep[0].length).setValues(rep);
}

function getSettingsGroup(email) {
  return AdminGroupSettings.Groups.get(email);
}
