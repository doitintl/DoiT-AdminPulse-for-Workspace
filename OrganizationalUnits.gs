/** This code will inventory all OUs in Google Workspace. The org unit IDs are used in other sections of the workbook.
 * @OnlyCurrentDoc
 */

function getOrgUnits() {
  // Get Google Sheet as we'll write data to it later, change sheet name to match yours
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Org Units");

  // You can add more fields based on your requirements
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
  sheet
    .getRange(1, 1, fileArray.length, fileArray[0].length)
    .setValues(fileArray);
}
