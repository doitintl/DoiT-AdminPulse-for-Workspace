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
    .addItem('Run all scripts', 'fetchInfo')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Run Individual Scripts')
      .addItem('getUsersList', 'getUsersList')
      .addItem('getDomainList', 'getDomainList')
      .addItem('getGroupsSettings', 'getGroupsSettings')
      .addItem('getMobileDevices', 'getMobileDevices')
      .addItem('getLicenseAssignments', 'getLicenseAssignments')
      .addItem('getTokens', 'getTokens')
      .addItem('getDomainInfo', 'getDomainInfo')
      .addItem('getAppPasswords', 'getAppPasswords')
      .addItem('getOrgUnits', 'getOrgUnits')
      .addItem('getSharedDrives', 'getSharedDrives')
      .addItem('getGroupMembers', 'getGroupMembers'))
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