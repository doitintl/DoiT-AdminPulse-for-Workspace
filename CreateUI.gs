/**
 * This code creates the UI button under the Extensions menu and runs all scripts or individual scripts based on admin selection.
 * 
 */

// Triggered on install
function onInstall(e) {
  onOpen(e);
}

// Creates the menu and sub-menu items under "Add-ons"
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Set up or Refresh Sheet', 'setupSheet')
    .addSeparator()
    .addItem('Run all scripts', 'promptRunAllScripts')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Run Individual Scripts')
      .addItem('List Customer Contact Info', 'getDomainInfo')
      .addItem('List Domains', 'promptGetDomainList')
      .addItem('List Users', 'getUsersList')
      .addItem('List Mobile Devices', 'getMobileDevices')
      .addItem('List License Assignments', 'getLicenseAssignments')
      .addItem('List OAuth Tokens', 'getTokens')
      .addItem('List App Passwords', 'getAppPasswords')
      .addItem('List Organizational Units', 'getOrgUnits')
      .addItem('List Shared Drives', 'getSharedDrives')
      .addItem('List Group Members', 'getGroupMembers'))
      .addItem('List Group Settings', 'getGroupsSettings')
    .addSeparator()
    .addItem('Get Support', 'contactPartner')
    .addToUi();
}

// Function to run all scripts with a confirmation prompt
function promptRunAllScripts() {
  var response = Browser.msgBox(
    'Run All Scripts Confirmation',
    'This will execute all scripts. Click OK to proceed. An external service (Cloudflare) will be used to return DNS results.',
    Browser.Buttons.OK_CANCEL
  );

  if (response == 'ok') {
    fetchInfo();
  }
}

// Function to run 'List Domains' script with a confirmation prompt
function promptGetDomainList() {
  var response = Browser.msgBox(
    'List Domains Confirmation',
    'This will execute the script for listing domains and DNS records. Click OK to proceed. An external service (Cloudflare) will be used to return DNS results.',
    Browser.Buttons.OK_CANCEL
  );

  if (response == 'ok') {
    getDomainList();
  }
}

// Function to run all scripts
function fetchInfo() {
  getDomainInfo();  
  getDomainList();
  getUsersList();
  getMobileDevices();
  getLicenseAssignments();
  getTokens();
  getAppPasswords();
  getOrgUnits();
  getSharedDrives();
  getGroupMembers();
  getGroupsSettings();  
}
