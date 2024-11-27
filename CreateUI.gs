/**
 * This code creates the UI button under the Extensions menu and runs all 
 * scripts or individual scripts based on admin selection.
 */

// Triggered on install 
function onInstall(e) {
  // No need to call onOpen here, it's automatically triggered on opening the sheet
}

// Creates the menu and sub-menu items under "Extensions"
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Activate Application', 'activateApplication')
    .addSeparator()     
    .addItem('Set up or Refresh Sheet', 'setupSheet')
    .addSeparator()
    .addItem('Run all scripts', 'promptRunAllScripts')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Run Individual Scripts')
      .addItem('List Customer Contact Info', 'getDomainInfo')
      .addItem('List Domains', 'promptGetDomainList')
      .addItem('List Users', 'getUsersList')
      .addItem('List Aliases', 'listAliases')      
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
  listAliases();
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

function activateApplication() {
  // Perform activation steps here (if any)

  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Activate Application',
    'Application activated and connected to your Google Account. Next, use the "Set up or Refresh Sheet" button from the extensions menu.',
    ui.ButtonSet.OK
  );
}

function setupSheet() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Setup Sheet',
    'Sheet setup complete!',
    ui.ButtonSet.OK
  ); 
}