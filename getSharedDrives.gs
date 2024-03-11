/**
 * This script will inventory all Shared Drives and add them to a Google Sheet including the Inventory Date,
 * Shared Drive ID, Shared Drive name and all Shared Drive restriction settings.
 * @OnlyCurrentDoc
 */

function getSharedDrives() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var shareddrivessheet = spreadsheet.getSheetByName('Shared Drives');

  // Check if the sheet exists, delete it if it does
  if (shareddrivessheet) {
    spreadsheet.deleteSheet(shareddrivessheet);
  }

  // Create a new 'Shared Drives' sheet
  shareddrivessheet = spreadsheet.insertSheet('Shared Drives');

  // Add headers to the sheet
  var headers = [
    'Audit Date',
    'ID',
    'Name',
    'Copy Requires Writer Permission',
    'Domain Users Only',
    'Drive Members Only',
    'Admin Managed Restrictions',
    'Sharing Folders Requires Organizer Permission',
    'orgUnitId',
    'Organization Unit'
  ];
  shareddrivessheet.appendRow(headers);

  // Format the headers
  var headerRange = shareddrivessheet.getRange('A1:J1');
  headerRange.setFontWeight('bold').setFontColor('#ffffff').setFontFamily('Montserrat');
  headerRange.setBackground('#fc3165');

  const startTime = new Date().getTime();
  const audit_timestamp = Utilities.formatDate(
    new Date(),
    "UTC",
    "yyyy-MM-dd'T'HH:mm:ss'Z'",
  );
  const rowsToWrite = [];

  let sharedDrives = Drive.Drives.list({
    maxResults: 100,
    useDomainAdminAccess: true,
    fields: "drives(id,name,restrictions,orgUnitId)",
  });

  let sharedDrivesItems = sharedDrives.drives || [];

  // If a next page token exists then iterate through again.
  while (sharedDrives.nextPageToken) {
    sharedDrives = Drive.Drives.list({
      pageToken: sharedDrives.nextPageToken,
      maxResults: 100,
      useDomainAdminAccess: true,
      hidden: false,
      fields: "drives(id,name,restrictions,orgUnitId)",
    });
    sharedDrivesItems = sharedDrivesItems.concat(sharedDrives.drives || []);
  }

  sharedDrivesItems.forEach(function (value, index) {
    const newRow = [
      audit_timestamp,
      value.id,
      value.name,
      value.restrictions ? value.restrictions.copyRequiresWriterPermission : null,
      value.restrictions ? value.restrictions.domainUsersOnly : null,
      value.restrictions ? value.restrictions.driveMembersOnly : null,
      value.restrictions ? value.restrictions.adminManagedRestrictions : null,
      value.restrictions ? value.restrictions.sharingFoldersRequiresOrganizerPermission : null,
      value.orgUnitId,
    ];
    // add to row array instead of append because append is slow
    rowsToWrite.push(newRow);

    // Set the formula for each row dynamically
    const formula = `=IFERROR(VLOOKUP(I${index + 2}, OrgID2Path, 2, FALSE), VLOOKUP(I${index + 2}, Org2ParentPath, 2, FALSE))`;
    shareddrivessheet.getRange(index + 2, 10).setFormula(formula);
  });

  // Write data to the sheet only if there is data to write
  if (rowsToWrite.length > 0) {
    shareddrivessheet
      .getRange(2, 1, rowsToWrite.length, rowsToWrite[0].length)
      .setValues(rowsToWrite);
  }

  shareddrivessheet.hideColumns(9);

  // Delete columns K-Z
  shareddrivessheet.deleteColumns(11, 16);

  // Auto-resize columns A, B, C, and J
  shareddrivessheet.autoResizeColumns(1, 3);
  shareddrivessheet.autoResizeColumns(10, 1);

  // Add conditional formatting rules
  var endRow = shareddrivessheet.getLastRow();
  var range = shareddrivessheet.getRange('D2:H' + endRow);
  var rule1 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('FALSE')
    .setBackground('#ffcccb')
    .setRanges([range])
    .build();
  var rule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('TRUE')
    .setBackground('#b7e1cd')
    .setRanges([range])
    .build();
  var rules = shareddrivessheet.getConditionalFormatRules();
  rules.push(rule1);
  rules.push(rule2);
  shareddrivessheet.setConditionalFormatRules(rules);

  const endTime = new Date().getTime();
  const elapsed = (endTime - startTime) / 1000;
  Logger.log("Elapsed Seconds: " + elapsed);
}
