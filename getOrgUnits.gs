/** This code will inventory all OUs in Google Workspace. The org unit IDs are used in other sections of the workbook.
 * @OnlyCurrentDoc
 */

function getOrgUnits() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var orgUnitsSheet = spreadsheet.getSheetByName("Org Units");

  // Check if the sheet exists, delete it if it does
  if (orgUnitsSheet) {
    spreadsheet.deleteSheet(orgUnitsSheet);
  }

  // Create a new 'Org Units' sheet
  orgUnitsSheet = spreadsheet.insertSheet("Org Units");

  // Add headers to the sheet
  var headers = [
    "Org Name ID",
    "Org Unit Name",
    "OrgUnit Path",
    "Description",
    "Parent Org Unit ID",
    "Parent Org Unit Path",
  ];
  orgUnitsSheet.appendRow(headers);

  // Format the headers
  var headerRange = orgUnitsSheet.getRange("A1:F1");
  headerRange.setFontWeight("bold").setFontColor("#ffffff").setFontFamily("Montserrat");
  headerRange.setBackground("#fc3165");

  // This code will inventory all OUs in Google Workspace. The org unit IDs are used in other sections of the workbook.
  const fileArray = [
    [
      "Org Name ID",
      "Org Unit Name",
      "OrgUnit Path",
      "Description",
      "Parent Org Unit ID",
      "Parent Org Unit Path",
    ],
  ];

  const orgUnits = AdminDirectory.Orgunits.list("my_customer", {
    type: "ALL",
  }).organizationUnits;

  orgUnits.forEach((orgUnit) => {
    fileArray.push([
      orgUnit.orgUnitId.slice(3),
      orgUnit.name,
      orgUnit.orgUnitPath,
      orgUnit.description,
      orgUnit.parentOrgUnitId
        ? orgUnit.parentOrgUnitId.replace(/^id:/, "")
        : "",
      orgUnit.parentOrgUnitPath,
    ]);
  });

  // Write data back to our sheets
  orgUnitsSheet
    .getRange(1, 1, fileArray.length, fileArray[0].length)
    .setValues(fileArray);

  // Delete columns G-Z
  orgUnitsSheet.deleteColumns(7, 20);

  // Auto-resize columns A, B, C, E, and F
  orgUnitsSheet.autoResizeColumns(1, 3);
  orgUnitsSheet.autoResizeColumns(5, 2);

  // Define ranges
  spreadsheet.setNamedRange('Org2ParentPath', orgUnitsSheet.getRange('E:F'));
  spreadsheet.setNamedRange('OrgID2Path', orgUnitsSheet.getRange('A:C'));
}