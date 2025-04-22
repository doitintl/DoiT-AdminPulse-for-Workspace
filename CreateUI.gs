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
    .addItem('Activate Application', 'activateApplication')
    .addSeparator()
    .addItem('Set up or Refresh Sheet', 'setupSheet')  // Make sure 'setupSheet' function exists
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Inventory Workspace Settings')
      .addItem('Check all policies', 'promptFetchAndListPolicies'))
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Run Reports')
      .addItem('Run all reports', 'promptRunAllScripts')
      .addSeparator()
      .addItem('List General Account Settings', 'getDomainInfo')
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


// Function to run 'Check all policies' script with a confirmation prompt
function promptFetchAndListPolicies() {
  var response = Browser.msgBox(
    'Check All Policies Confirmation',
    'This will perform an automated inventory of Workspace settings and dependencies including OU and Group inventory. This can take several minutes to complete.',
    Browser.Buttons.OK_CANCEL
  );

  if (response == 'ok') {
    fetchAndListPolicies();
  }
}


// Function to run 'List Domains' script with a confirmation prompt
function promptGetDomainList() {
  var response = Browser.msgBox(
    'List Domains Confirmation',
    'This will execute the script for listing domains and DNS records. Click OK to proceed.  Calls will be made to Google DNS to return DNS records.',
    Browser.Buttons.OK_CANCEL
  );

  if (response == 'ok') {
    getDomainList();
  }
}

function fetchInfo() {
  Utilities.sleep(1000);
}

//Function to run all scripts
function promptRunAllScripts() {
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
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
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