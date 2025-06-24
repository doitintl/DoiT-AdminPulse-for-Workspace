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
      // --- THIS IS THE MODIFIED LINE ---
      .addItem('Check all policies', 'promptAndRunFullPolicyCheck')) // Calls the new wrapper function
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

/**
 * Shows a confirmation dialog before running the full policy check.
 * This is called by the menu item.
 */
function promptAndRunFullPolicyCheck() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Confirmation Required',
    'This function requires Super Administrator privileges and may take several minutes to complete.\n\nDo you want to proceed?',
    ui.ButtonSet.YES_NO
  );

  // Check the user's response
  if (response == ui.Button.YES) {
    // If they click "Yes", run the main function
    Logger.log("User confirmed. Starting full policy check.");
    SpreadsheetApp.getActiveSpreadsheet().toast('Starting policy check...', 'Please wait', 10);
    runFullPolicyCheck();
  } else {
    // If they click "No", do nothing and log the cancellation
    Logger.log("User cancelled the policy check.");
    SpreadsheetApp.getActiveSpreadsheet().toast('Policy check cancelled.', 'Info', 5);
  }
}