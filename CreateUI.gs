/**This code creates the UI button under the Extensions menu and runs all scripts or individual scripts based on admin selection. 
 * @OnlyCurrentDoc
 */

// Triggered on install
function onInstall(e) {
  onOpen(e);
}

// Creates the menu and sub-menu items under "Add-ons"
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Setup Sheet', 'setupSheet')
    .addSeparator()
    .addItem('Run all scripts', 'fetchInfo')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Run Individual Scripts')
      .addItem('List Users', 'getUsersList')
      .addItem('List Domains', 'getDomainList')
      .addItem('List Group Settings', 'getGroupsSettings')
      .addItem('List Mobile Devices', 'getMobileDevices')
      .addItem('List License Assignments', 'getLicenseAssignments')
      .addItem('List OAuth Tokens', 'getTokens')
      .addItem('List Customer Contact Info', 'getDomainInfo')
      .addItem('List App Passwords', 'getAppPasswords')
      .addItem('List Organizational Units', 'getOrgUnits')
      .addItem('List Shared Drives', 'getSharedDrives')
      .addItem('List Group Members', 'getGroupMembers'))
    .addToUi()
}


// Function to run all scripts
function fetchInfo() {
  getUsersList();
  getDomainList();
  getGroupsSettings();
  getMobileDevices();
  getLicenseAssignments();
  getTokens();
  getDomainInfo();
  getAppPasswords();
  getOrgUnits();
  getSharedDrives();
  getGroupMembers();
}