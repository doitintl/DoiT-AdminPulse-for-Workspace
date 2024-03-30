function setupSheet() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Warning', 'This will delete any existing data in the sheet, including any notes or completed checklists. Do you want to continue?', ui.ButtonSet.YES_NO);

  if (response === ui.Button.NO) {
    return;
  }

  const domain = Session.getActiveUser().getEmail().split('@')[1];

  const templateId = '1rbgKhzDYDmPDKuyx9_qR3CWpTX_ouacEKViuPwAUAf8';
  const copiedFile = DriveApp.getFileById(templateId).makeCopy();
  const copiedFileId = copiedFile.getId();

  const templateSpreadsheet = SpreadsheetApp.openById(copiedFileId);
  const templateSheets = templateSpreadsheet.getSheets();

  const currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  for (let i = 0; i < templateSheets.length; i++) {
    const sheetName = templateSheets[i].getName();

    const existingSheet = currentSpreadsheet.getSheetByName(sheetName);

    if (existingSheet) {
      currentSpreadsheet.deleteSheet(existingSheet);
    }

    templateSheets[i].copyTo(currentSpreadsheet).setName(sheetName);
  }

  const sheet1 = currentSpreadsheet.getSheetByName('Sheet1');
  if (sheet1) {
    currentSpreadsheet.deleteSheet(sheet1);
  }

  currentSpreadsheet.rename('[' + domain + '] Security Checklist for Workspace Admins');

  copiedFile.setTrashed(true);

  const setupCompleteAlert = ui.alert('Sheet Setup Complete',
    'Welcome to the Security Checklist for Workspace Admins!\n\n' +
    'This tool provides a comprehensive checklist of security controls for Business and Enterprise organizations.\n\n' +
    'To use this tool and all its functions, you must have a Super Admin account.\n\n' +
    'Many settings do not have an API, so we have included links to Google\'s documentation, best practice recommendations, and the relevant section of the admin console.\n\n' +
    'To begin, run the read-only API reports using the Run All Scripts button under Extensions > Security Checklist for Workspace Admins. These reports will help you answer questions on the Security Checklist.\n\n' +
    'After running the API reports, complete the checklist of security controls and take notes on areas where your organization can improve its security posture.\n\n' +
    'For developer support or assistance with reviewing your environment and understanding the findings, use the Get Support button in the Extensions menu.',
    ui.ButtonSet.OK);
}
