// Triggered on install
function onInstall(e) {
  onOpen(e);
}

// Creates the menu and sub-menu items under "Add-ons"
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Fetch Info')
    .addItem('Run all scripts', 'fetchInfo')
    .addSubMenu(ui.createMenu('Run Individual Scripts')
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
    .addToUi();
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
