/**
 * This script will inventory all Shared Drives and add them to a Google Sheet including the Inventory Date,
 * Shared Drive ID, Shared Drive name and all Shared Drive restriction settings.
 * @OnlyCurrentDoc
 */

function getSharedDrives() {
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
    fields: "items(id,name,restrictions,orgUnitId)",
  });

  let sharedDrivesItems = sharedDrives.items || [];

  // If a next page token exists then iterate through again.
  while (sharedDrives.nextPageToken) {
    sharedDrives = Drive.Drives.list({
      pageToken: sharedDrives.nextPageToken,
      maxResults: 100,
      useDomainAdminAccess: true,
      hidden: false,
    });
    sharedDrivesItems = sharedDrivesItems.concat(sharedDrives.items || []);
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("Shared Drives");

  // Clear existing data if any
  if (sheet.getLastRow() > 1) {
    sheet
      .getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
      .clearContent();
  }

  sharedDrivesItems.forEach(function (value, index) {
    const newRow = [
      audit_timestamp,
      value.id,
      value.name,
      value.restrictions.copyRequiresWriterPermission,
      value.restrictions.domainUsersOnly,
      value.restrictions.driveMembersOnly,
      value.restrictions.adminManagedRestrictions,
      value.restrictions.sharingFoldersRequiresOrganizerPermission,
      value.orgUnitId,
    ];
    // add to row array instead of append because append is slow
    rowsToWrite.push(newRow);

    // Set the formula for each row dynamically
    const formula = `=IFERROR(VLOOKUP(I${index + 2}, 'Org Units'!OrgID2Path, 2, FALSE), VLOOKUP(I${index + 2}, 'Org Units'!Org2ParentPath, 2, FALSE))`;
    sheet.getRange(index + 2, 10).setFormula(formula);
  });

  // Write data to the sheet only if there is data to write
  if (rowsToWrite.length > 0) {
    sheet
      .getRange(2, 1, rowsToWrite.length, rowsToWrite[0].length)
      .setValues(rowsToWrite);
  }

  sheet.hideColumns(9);

  const endTime = new Date().getTime();
  const elapsed = (endTime - startTime) / 1000;
  Logger.log("Elapsed Seconds: " + elapsed);
}
