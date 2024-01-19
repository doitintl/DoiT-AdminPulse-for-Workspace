/**This code will inventory all OUs in Google Workspace. The org unit IDs are used in other sections of the workbook. 
 * @OnlyCurrentDoc
 */

function getOrgUnits() {
    // get google sheet as we'll write data to it later, change sheet name to match yours
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Org Units");

    // Check if there is data in row 2 and clear the sheet contents accordingly
    var dataRange = sheet.getRange(2, 1, 1, sheet.getLastColumn());
    var isDataInRow2 = dataRange.getValues().flat().some(Boolean);

    if (isDataInRow2) {
        sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }

    // you can add more fields based on your requirements
    const fileArray = [];

    const orgUnits = AdminDirectory.Orgunits.list("my_customer", {
        type: "ALL"
    }).organizationUnits.forEach(orgUnit => {
        fileArray.push([orgUnit.orgUnitId.slice(3), orgUnit.name, orgUnit.orgUnitPath, orgUnit.description, orgUnit.parentOrgUnitId ? orgUnit.parentOrgUnitId.replace(/^id:/, '') : '', orgUnit.parentOrgUnitPath])
    });

    // write data back to our sheets
    sheet.getRange(sheet.getLastRow() + 1, 1, fileArray.length, fileArray[0].length).setValues(fileArray);
    
}
