function setupSheet() {
  const functionName = 'setupSheet';
  const startTime = new Date();
  Logger.log(`-- Starting ${functionName} at: ${startTime.toLocaleString()}`);

  try {
    const ui = SpreadsheetApp.getUi(); // Get the UI service.

    // --- Confirmation Alert ---
    const response = ui.alert(
      'Warning',
      'This will delete any existing data in the sheet, including any notes or completed checklists. Do you want to continue?',
      ui.ButtonSet.YES_NO
    );

    if (response === ui.Button.NO) {
      Logger.log("User cancelled the setup process.");
      return; // Exit the function if the user clicks "No".
    }

    // --- Get Domain ---
    const domain = Session.getActiveUser().getEmail().split('@')[1];

    // --- Copy Template ---
    const templateId = '1rbgKhzDYDmPDKuyx9_qR3CWpTX_ouacEKViuPwAUAf8';
    let copiedFileId;
    try {
      const copiedFile = DriveApp.getFileById(templateId).makeCopy();
      copiedFileId = copiedFile.getId();
    } catch (e) {
      ui.alert('Error', 'Failed to copy the template file. Please ensure you have access to the template file and try again.\n\nError Details: ' + e, ui.ButtonSet.OK);
      Logger.log(`!! ERROR in ${functionName}: Failed to copy template. ${e.toString()}`);
      return; // Exit if the template can't be copied.
    }

    const templateSpreadsheet = SpreadsheetApp.openById(copiedFileId);
    const templateSheets = templateSpreadsheet.getSheets();

    // --- Copy Sheets to Current Spreadsheet ---
    const currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    for (let i = 0; i < templateSheets.length; i++) {
      const sheetName = templateSheets[i].getName();
      const existingSheet = currentSpreadsheet.getSheetByName(sheetName);

      if (existingSheet) {
        currentSpreadsheet.deleteSheet(existingSheet);
      }
      templateSheets[i].copyTo(currentSpreadsheet).setName(sheetName);
    }

    // --- Clean Up "Sheet1" ---
    const sheet1 = currentSpreadsheet.getSheetByName('Sheet1');
    if (sheet1) {
      currentSpreadsheet.deleteSheet(sheet1);
    }

    // --- Rename Spreadsheet ---
    currentSpreadsheet.rename('[' + domain + '] DoiT AdminPulse for Workspace');

    // --- Trash Copied Template File ---
    const copiedFile = DriveApp.getFileById(copiedFileId);
    copiedFile.setTrashed(true); // Clean up temporary file.

    // --- Setup Complete Alert ---
    ui.alert(
      'Sheet Setup Complete',
      'Welcome to the DoiT AdminPulse for Workspace!\n\n' +
        'This tool provides a comprehensive checklist of security controls for Business and Enterprise organizations.\n\n' +
        'To use this tool and all its functions, you must have a Super Admin account.\n\n' +
        'Many settings do not have an API, so we have included links to Google\'s documentation, best practice recommendations, and the relevant section of the admin console.\n\n' +
        'To Begin, Run the read-only API reports using the Check all policies button under Extensions > DoiT AdminPulse for Workspace > Inventory Workspace Settings menu. This will inventory all Google Workspace policies to the sheet and present an alert when it completes.\n\n' +
        'Run the read-only API reports using the Run Reports menu under Extensions > DoiT AdminPulse for Workspace. These reports will help you answer questions on the Security Checklist.\n\n' +
        'After running the API reports, complete the checklist of security controls and take notes on areas where your organization can improve its security posture.\n\n' +
        'For developer support or assistance with reviewing your environment and understanding the findings, use the Get Support button in the Extensions menu.',
      ui.ButtonSet.OK
    );

  } catch (e) {
    Logger.log(`!! ERROR in ${functionName}: An unexpected error occurred. ${e.toString()}`);
    SpreadsheetApp.getUi().alert("An unexpected error occurred during setup. Please check the logs for more details.");
  } finally {
    const endTime = new Date();
    const duration = (endTime.getTime() - startTime.getTime()) / 1000;
    Logger.log(`-- Finished ${functionName} at: ${endTime.toLocaleString()} (Duration: ${duration.toFixed(2)}s)`);
  }
}