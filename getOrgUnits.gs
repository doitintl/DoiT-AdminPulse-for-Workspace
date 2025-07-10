function getOrgUnits() {
  const functionName = 'getOrgUnits';
  const startTime = new Date();
  Logger.log(`-- Starting ${functionName} at: ${startTime.toLocaleString()}`);

  try {
    // Reset global variables
    const orgUnitMap = new Map();
    orgUnitMap.clear();
    customerRootOuId = null;
    actualCustomerId = null;

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let orgUnitsSheet = spreadsheet.getSheetByName("Org Units");

    if (orgUnitsSheet) {
      spreadsheet.deleteSheet(orgUnitsSheet);
    }
    orgUnitsSheet = spreadsheet.insertSheet("Org Units", spreadsheet.getSheets().length);

    const headers = ["Org Unit ID (Raw)", "Org Unit Name", "OrgUnit Path", "Description", "Parent Org Unit ID (Raw)", "Parent Org Unit Path"];
    orgUnitsSheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight("bold").setFontColor("#ffffff").setFontFamily("Montserrat").setBackground("#fc3165");

    let allOUsFromApiList = [];
    let pageToken;
    const customerApiAlias = "my_customer";

    // Step 1: Fetch Customer and Root OU details
    try {
      const customerDetails = AdminDirectory.Customers.get(customerApiAlias);
      actualCustomerId = customerDetails.id;
      if (customerDetails.orgUnitId) {
        customerRootOuId = customerDetails.orgUnitId;
      }
    } catch (e) {
      // It's okay if this fails, the next step can still work.
    }

    // Step 2: Fetch all OUs using Orgunits.list
    const identifierForListCall = actualCustomerId || customerApiAlias;
    do {
      const orgUnitsResponsePage = AdminDirectory.Orgunits.list(identifierForListCall, {
        type: "ALL",
        maxResults: 500,
        pageToken: pageToken,
      });
      if (orgUnitsResponsePage && orgUnitsResponsePage.organizationUnits) {
        allOUsFromApiList = allOUsFromApiList.concat(orgUnitsResponsePage.organizationUnits);
      }
      pageToken = orgUnitsResponsePage.nextPageToken;
    } while (pageToken);

    // Step 3: Populate the global map and find root if needed
    allOUsFromApiList.forEach((ou) => {
      if (ou.orgUnitId && ou.hasOwnProperty('orgUnitPath')) {
        orgUnitMap.set(ou.orgUnitId, ou.orgUnitPath);
        if (!customerRootOuId && ou.orgUnitPath === "/") {
          customerRootOuId = ou.orgUnitId;
        }
      }
    });
    // Ensure the identified root has the correct path in the map
    if (customerRootOuId) {
      orgUnitMap.set(customerRootOuId, "/");
      PropertiesService.getScriptProperties().setProperty('customerRootOuId', customerRootOuId);
    }

    // ---- PREPARE DATA FOR WRITING TO SHEET ----
    const fileArray = [];
    allOUsFromApiList.forEach((orgUnit) => {
      // *** Strip the "id:" prefix from both OU IDs before writing to the sheet. ***
      const cleanOrgUnitId = (orgUnit.orgUnitId || "").replace('id:', '');
      const cleanParentOrgUnitId = (orgUnit.parentOrgUnitId || "").replace('id:', '');
      
      let orgPath = orgUnit.orgUnitPath;
      if(customerRootOuId && orgUnit.orgUnitId === customerRootOuId) {
          orgPath = "/"; // Enforce root path is "/"
      }

      fileArray.push([
        cleanOrgUnitId,
        orgUnit.name || "",
        orgPath,
        orgUnit.description || "",
        cleanParentOrgUnitId,
        orgUnit.parentOrgUnitPath || "",
      ]);
    });
    
    // Sort for presentation
    fileArray.sort((a, b) => {
        const pathA = (a[2] || "").toLowerCase();
        const pathB = (b[2] || "").toLowerCase();
        if (pathA === "/") return -1;
        if (pathB === "/") return 1;
        return pathA.localeCompare(pathB);
    });

    if (fileArray.length > 0) {
      orgUnitsSheet.getRange(2, 1, fileArray.length, headers.length).setValues(fileArray);
    } else {
      orgUnitsSheet.getRange("A2").setValue("No Organizational Units found.");
    }

    // Sheet finalization
    const lastRowWithData = orgUnitsSheet.getLastRow();
    if (orgUnitsSheet.getMaxColumns() > headers.length) {
      orgUnitsSheet.deleteColumns(headers.length + 1, orgUnitsSheet.getMaxColumns() - headers.length);
    }
    if (lastRowWithData > 1) {
      orgUnitsSheet.autoResizeColumns(1, headers.length);
      const dataRowCount = lastRowWithData - 1;
      spreadsheet.setNamedRange('Org2ParentPath', orgUnitsSheet.getRange(2, 5, dataRowCount, 2)); // E2:F...
      spreadsheet.setNamedRange('OrgID2Path', orgUnitsSheet.getRange(2, 1, dataRowCount, 3));     // A2:C...

      const filterRange = orgUnitsSheet.getRange(1, 1, lastRowWithData, headers.length);
      if (filterRange.getFilter()) filterRange.getFilter().remove();
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
