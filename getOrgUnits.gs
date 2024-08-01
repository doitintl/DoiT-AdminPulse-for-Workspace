function getOrgUnits(){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const lastSheetIndex = sheets.length;

  let orgUnitsSheet = spreadsheet.getSheetByName("Org Units");

  // Check if the sheet exists, delete it if it does
  if (orgUnitsSheet) {
    spreadsheet.deleteSheet(orgUnitsSheet);
  }

  // Create a new 'Org Units' sheet at the last index
  orgUnitsSheet = spreadsheet.insertSheet("Org Units", lastSheetIndex);

  // Add headers to the sheet
  const headers = [
    "Org Name ID",
    "Org Unit Name",
    "OrgUnit Path",
    "Description",
    "Parent Org Unit ID",
    "Parent Org Unit Path",
  ];
  orgUnitsSheet.appendRow(headers);

  // Format the headers
  const headerRange = orgUnitsSheet.getRange("A1:F1");
  headerRange.setFontWeight("bold").setFontColor("#ffffff").setFontFamily("Montserrat");
  headerRange.setBackground("#fc3165");


  // Fetch and sort org units
  const orgUnits = AdminDirectory.Orgunits.list("my_customer", {
    type: "ALL",
  }).organizationUnits;

  // Sort the orgUnits array based on the orgUnitPath
  orgUnits.sort((a, b) => {
    // Split the paths into components
    const pathA = a.orgUnitPath.split("/");
    const pathB = b.orgUnitPath.split("/");

    // Compare paths component by component
    for (let i = 0; i < Math.min(pathA.length, pathB.length); i++) {
      if (pathA[i] < pathB[i]) return -1; // a comes before b
      if (pathA[i] > pathB[i]) return 1; // a comes after b
    }

    // If all components match, shorter path comes first
    return pathA.length - pathB.length;
  });

  // Prepare data for the sheet (including headers)
  const fileArray = [headers]; 
  orgUnits.forEach((orgUnit) => {
    fileArray.push([
      orgUnit.orgUnitId.slice(3),
      orgUnit.name,
      orgUnit.orgUnitPath,
      orgUnit.description,
      orgUnit.parentOrgUnitId ? orgUnit.parentOrgUnitId.replace(/^id:/, "") : "",
      orgUnit.parentOrgUnitPath,
    ]);
  });

  // Write data back to the sheet
  orgUnitsSheet.getRange(1, 1, fileArray.length, fileArray[0].length).setValues(fileArray);

  // Delete columns G-Z
  orgUnitsSheet.deleteColumns(7, 20);

  // Auto-resize columns A, B, C, E, and F
  orgUnitsSheet.autoResizeColumns(1, 3);
  orgUnitsSheet.autoResizeColumns(5, 2);

  // Define ranges
  spreadsheet.setNamedRange('Org2ParentPath', orgUnitsSheet.getRange('E:F'));
  spreadsheet.setNamedRange('OrgID2Path', orgUnitsSheet.getRange('A:C'));

  // --- Add Filter View ---
  const lastRow = orgUnitsSheet.getLastRow();
  const filterRange = orgUnitsSheet.getRange('A1:F' + lastRow);  // Filter columns A through F (including header)
  filterRange.createFilter(); 
}
