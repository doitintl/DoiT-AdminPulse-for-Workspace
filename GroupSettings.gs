/**
 * This script will use OAUTH list all Google Groups and the associated Group Settings to a Google Sheet.
 * @OnlyCurrentDoc
 */

function getGroupsSettings() {
  var page;
  var pageToken;
  var rep = [];
  rep.push(['name', 'email', 'description', 'whoCanJoin','whoCanPostMessage','whoCanViewMembership', 'whoCanViewGroup',
            'allowExternalMembers','allowWebPosting','primaryLanguage','isArchived','archiveOnly','messageModerationLevel','spamModerationLevel','replyTo','customReplyTo','includeCustomFooter','customFooterText',
            'sendMessageDenyNotification','defaultMessageDenyNotificationText','membersCanPostAsTheGroup','includeInGlobalAddressList','whoCanLeaveGroup','whoCanContactOwner','favoriteRepliesOnTop','whoCanBanUsers','whoCanModerateMembers','whoCanModerateContent','whoCanAssistContent','customRolesEnabledForSettingsToBeMerged','enableCollaborativeInbox','whoCanDiscoverGroup','defaultSender']);

  do {
    page = AdminDirectory.Groups.list({pageToken:pageToken, customer:'my_customer', orderBy:'email', sortOrder:'ASCENDING'});
    if (page.groups && page.groups.length > 0) {
      for (var i = 0; i < page.groups.length; i++) {
        var group = page.groups[i];
        var settings = getSettingsGroup(group.email);
        rep.push([group.name,group.email,group.description,settings.whoCanJoin,settings.whoCanPostMessage,settings.whoCanViewMembership,settings.whoCanViewGroup,settings.allowExternalMembers, settings.allowWebPosting,settings.primaryLanguage,settings.isArchived,settings.archiveOnly,settings.messageModerationLevel,settings.spamModerationLevel,settings.replyTo,settings.customReplyTo,settings.includeCustomFooter,settings.customFooterText,settings.sendMessageDenyNotification,settings.defaultMessageDenyNotificationText,settings.membersCanPostAsTheGroup,settings.includeInGlobalAddressList,settings.whoCanLeaveGroup,settings.whoCanContactOwner,settings.favoriteRepliesOnTop,settings.whoCanBanUsers,settings.whoCanModerateMembers,settings.whoCanModerateContent,settings.whoCanAssistContent,settings.customRolesEnabledForSettingsToBeMerged,settings.enableCollaborativeInbox,settings.whoCanDiscoverGroup,settings.defaultSender]);
      }
    } 
    pageToken = page.nextPageToken;
  } while (pageToken);

  // Insert data in a spreadsheet  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Group Settings');
  
  // Check if there is data in row 2 and clear the sheet contents accordingly
  var dataRange = sheet.getRange(2, 1, 1, sheet.getLastColumn());
  var isDataInRow2 = dataRange.getValues().flat().some(Boolean);

  if (isDataInRow2) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }

  sheet.getRange(1,1,rep.length,rep[0].length).setValues(rep);
}

function getSettingsGroup(email){
  return AdminGroupSettings.Groups.get(email);
}
