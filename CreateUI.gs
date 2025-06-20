/**
 * This code creates the UI button under the Extensions menu. 
 * It allows an administrator to activate the application, set up the sheet, 
 * run a full policy check, or run individual inventory reports.
 */

/**
 * Runs when the add-on is installed.
 * @param {Object} e The event parameter for a simple trigger.
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Creates the add-on menu in the spreadsheet UI when the document is opened.
 * @param {Object} e The event parameter for a simple trigger.
 */
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  
  // Create the main menu for reports, which will be populated with categories.
  const reportsMenu = ui.createMenu('Run Reports');

  // --- Category 1: User & Security Reports ---
  const userSecuritySubMenu = ui.createMenu('User & Security')
    .addItem('List Users', 'getUsersList')
    .addItem('List Aliases', 'listAliases')
    .addItem('List Mobile Devices', 'getMobileDevices')
    .addItem('List License Assignments', 'getLicenseAssignments')
    .addSeparator()
    .addItem('List OAuth Tokens', 'getTokens')
    .addItem('List App Passwords', 'getAppPasswords');
    
  // --- Category 2: Infrastructure & Groups Reports ---
  const infraGroupsSubMenu = ui.createMenu('Infrastructure & Groups')
    .addItem('List Organizational Units', 'getOrgUnits')
    .addItem('List Shared Drives', 'getSharedDrives')
    .addSeparator()
    .addItem('List Group Settings', 'getGroupsSettings')
    .addItem('List Group Members', 'getGroupMembers');
    
  // --- Category 3: Domain & General Reports ---
  const domainGeneralSubMenu = ui.createMenu('Domain & General')
    .addItem('List General Account Settings', 'getDomainInfo')
    .addItem('List Domains', 'getDomainList');
    
  // Add the categorized sub-menus to the main 'Run Reports' menu.
  reportsMenu
    .addSubMenu(userSecuritySubMenu)
    .addSubMenu(infraGroupsSubMenu)
    .addSubMenu(domainGeneralSubMenu);

  // Build the final Add-on menu.
  ui.createAddonMenu()
    .addItem('Activate Application', 'activateApplication')
    .addItem('Set up or Refresh Sheet', 'setupSheet')
    .addSeparator()
    .addSubMenu(ui.createMenu('Inventory Workspace Settings')
      .addItem('Check all policies', 'runFullPolicyCheck')) // This now correctly calls runFullPolicyCheck
    .addSeparator()
    .addSubMenu(reportsMenu) // Add the fully constructed reports menu
    .addSeparator()
    .addItem('Get Support', 'contactPartner')
    .addToUi();
}


/**
 * Placeholder function for the 'Activate Application' menu item.
 */
function activateApplication() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Activate Application',
    'Application activated and connected to your Google Account. Next, use the "Set up or Refresh Sheet" button from the extensions menu.',
    ui.ButtonSet.OK
  );
}

/**
 * Placeholder function for the 'Set up or Refresh Sheet' menu item.
 */
function setupSheet() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Setup Sheet',
    'Sheet setup complete!',
    ui.ButtonSet.OK
  );
}