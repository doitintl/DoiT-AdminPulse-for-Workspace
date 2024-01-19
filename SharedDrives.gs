/** 
 * This script will inventory all Shared Drives and add them to a Google Sheet including the Inventory Date, 
 * Shared Drive ID, Shared Drive name and all Shared Drive restriction settings. 
 * @OnlyCurrentDoc
 */

function getSharedDrives() {
  var startTime = new Date().getTime();
  var audit_timestamp = Utilities.formatDate(new Date(), "UTC", "yyyy-MM-dd'T'HH:mm:ss'Z'");
  var rowsToWrite = [];

  let sharedDrives = Drive.Drives.list({
    maxResults: 100,
    useDomainAdminAccess: true,
    fields: "items(id,name,restrictions,orgUnitId)",
  });

  let sharedDrivesItems = sharedDrives.items;

  // If a next page token exists then iterate through again.
  while(sharedDrives.nextPageToken){
    sharedDrives = Drive.Drives.list({
      pageToken: sharedDrives.nextPageToken,
      maxResults: 100,
      useDomainAdminAccess: true,
      hidden: false,
    });
    sharedDrivesItems = sharedDrivesItems.concat(sharedDrives.items);
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Shared Drives");

  // Check if there is data in row 2 and clear the sheet contents accordingly
  var dataRange = ss.getRange(2, 1, 1, ss.getLastColumn());
  var isDataInRow2 = dataRange.getValues().flat().some(Boolean);

  if (isDataInRow2) {
    ss.getRange(2, 1, ss.getLastRow() - 1, ss.getLastColumn()).clearContent();
  }

  sharedDrivesItems.forEach(function(value) {
    var newRow = [audit_timestamp, value.id, value.name, value.restrictions.copyRequiresWriterPermission, value.restrictions.domainUsersOnly, value.restrictions.driveMembersOnly, value.restrictions.adminManagedRestrictions, value.restrictions.sharingFoldersRequiresOrganizerPermission, value.orgUnitId];
    // add to row array instead of append because append is slow
    rowsToWrite.push(newRow);
  });

  ss.getRange(ss.getLastRow() + 1, 1, rowsToWrite.length, rowsToWrite[0].length).setValues(rowsToWrite);
  ss.hideColumns(9);

  // Iterate down each row of Column I that has data and populate the Vlookup formula
  for (var i = 2; i <= ss.getLastRow(); i++) {
    if (ss.getRange(i, 9).getValue() !== "") {
      ss.getRange(i, 10).setValue("=IFERROR(VLOOKUP(I2, 'Org Units'!OrgID2Path, 2, FALSE), VLOOKUP(I2, 'Org Units'!Org2ParentPath, 2, FALSE))");
    }
  }

  var endTime = new Date().getTime();
  var elapsed = (endTime - startTime) / 1000;
  console.log('Elapsed Seconds: ' + elapsed);
}
