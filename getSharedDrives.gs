function getSharedDrives() {
  const functionName = 'getSharedDrives';
  const startTime = new Date();
  Logger.log(`-- Starting ${functionName} at: ${startTime.toLocaleString()}`);

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // --- DEPENDENCY CHECK ---
    // Check if the 'OrgID2Path' named range exists, which is created by getOrgUnits.
    const orgDataRange = spreadsheet.getRangeByName('OrgID2Path');
    if (!orgDataRange) {
      // If the range doesn't exist, inform the user and run the prerequisite script.
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "Required Org Unit data not found. Running the update first. This may take a moment...",
        "Dependency Update", -1 // A negative duration means the toast stays until dismissed or replaced.
      );
      getOrgUnits(); // Run the function that creates the named ranges.
      SpreadsheetApp.getActiveSpreadsheet().toast("Org Unit data updated. Continuing with Shared Drives report.", "Update Complete", 5);
    }
    // --- END DEPENDENCY CHECK ---

    let sharedDrivesSheet = spreadsheet.getSheetByName('Shared Drives');
    if (sharedDrivesSheet) {
      spreadsheet.deleteSheet(sharedDrivesSheet);
    }
    sharedDrivesSheet = spreadsheet.insertSheet('Shared Drives', spreadsheet.getNumSheets());

    const headers = [
      'Audit Date', 'ID', 'Name', 'Copy Requires Writer Permission', 'Domain Users Only',
      'Drive Members Only', 'Admin Managed Restrictions', 'Sharing Folders Requires Organizer Permission',
      'orgUnitId', 'Organization Unit'
    ];
    sharedDrivesSheet.getRange('A1:J1').setValues([headers])
      .setFontWeight('bold')
      .setFontColor('#ffffff')
      .setFontFamily('Montserrat')
      .setBackground('#fc3165');

    const audit_timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-M-dd HH:mm:ss");
    const rowsToWrite = [];
    const formulasToWrite = [];

    let pageToken = null;
    let sharedDrivesItems = [];
    do {
      const response = Drive.Drives.list({
        pageToken: pageToken,
        maxResults: 100,
        useDomainAdminAccess: true,
        fields: "nextPageToken,drives(id,name,restrictions,orgUnitId)",
      });
      if (response.drives) {
        sharedDrivesItems = sharedDrivesItems.concat(response.drives);
      }
      pageToken = response.nextPageToken;
    } while (pageToken);


    if (sharedDrivesItems.length > 0) {
      sharedDrivesItems.forEach(function(value, index) {
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
        rowsToWrite.push(newRow);
        
        const formula = `=IFERROR(VLOOKUP(I${index + 2} & "", OrgID2Path, 3, FALSE), IFERROR(VLOOKUP(I${index + 2} & "", Org2ParentPath, 2, FALSE), "Not Found"))`;
        formulasToWrite.push([formula]);
      });

      sharedDrivesSheet.getRange(2, 1, rowsToWrite.length, rowsToWrite[0].length).setValues(rowsToWrite);
      SpreadsheetApp.flush();
      sharedDrivesSheet.getRange(2, 10, formulasToWrite.length, 1).setFormulas(formulasToWrite);

    } else {
        sharedDrivesSheet.getRange("A2").setValue("No Shared Drives found.");
    }

    sharedDrivesSheet.hideColumns(9);

    if (sharedDrivesSheet.getMaxColumns() > 10) {
        sharedDrivesSheet.deleteColumns(11, sharedDrivesSheet.getMaxColumns() - 10);
    }
    
    sharedDrivesSheet.autoResizeColumns(1, 3);
    sharedDrivesSheet.autoResizeColumn(10);

    const lastRow = sharedDrivesSheet.getLastRow();
    if (lastRow > 1) {
      const range = sharedDrivesSheet.getRange('D2:H' + lastRow);
      const rule1 = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('FALSE').setBackground('#ffcccb').setRanges([range]).build();
      const rule2 = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('TRUE').setBackground('#b7e1cd').setRanges([range]).build();
      sharedDrivesSheet.setConditionalFormatRules([rule1, rule2]);

      const filterRange = sharedDrivesSheet.getRange('A1:J' + lastRow);
      filterRange.createFilter();
    }

  } catch (e) {
    Logger.log(`!! ERROR in ${functionName}: ${e.toString()}`);
    SpreadsheetApp.getUi().alert(`An error occurred in ${functionName}: ${e.message}`);
  } finally {
    const endTime = new Date();
    const duration = (endTime.getTime() - startTime.getTime()) / 1000;
    Logger.log(`-- Finished ${functionName} at: ${endTime.toLocaleString()} (Duration: ${duration.toFixed(2)}s)`);
  }
}