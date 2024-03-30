/**
 * This code creates the UI button under the Extensions menu and runs all scripts or individual scripts based on admin selection.
 * 
 */

// Triggered on install
function onInstall(e) {
  onOpen(e);
}

// Creates the menu and sub-menu items under "Extensions"
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
      .addItem('List Group Settings', 'getGroupsSettings')
      .addItem('List Group Members', 'getGroupMembers'))
    .addSeparator()
    .addItem('Get Support', 'contactPartner')
    .addToUi();

  // Show alert message
  SpreadsheetApp.getUi().alert('Welcome to the Security Checklist for Workspace Admins!\n\n' +
  'This tool provides a comprehensive checklist of security controls for Business and Enterprise organizations.\n\n' +
  'To use this tool and all its functions, you must have a Super Admin account.\n\n' +
  'Many settings do not have an API, so we have included links to Google\'s documentation, best practice recommendations, and the relevant section of the admin console.\n\n' +
  'To begin, run the read-only API reports using the Run All Scripts button under Extensions > Security Checklist for Workspace Admins. These reports will help you answer questions on the Security Checklist.\n\n' +
  'After running the API reports, complete the checklist of security controls and take notes on areas where your organization can improve its security posture.\n\n' +
  'For developer support or assistance with reviewing your environment and understanding the findings, use the Get Support button in the Extensions menu.');
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
    // Display alert notification when all scripts have completed
  SpreadsheetApp.getUi().alert('All API scripts have successfully completed.');
}