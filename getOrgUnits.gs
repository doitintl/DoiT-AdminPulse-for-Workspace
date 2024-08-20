function getOrgUnits() {
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


  // Fetch org units, explicitly including the root organizational unit
  const orgUnitsResponse = AdminDirectory.Orgunits.list("my_customer", {
    type: "ALL",
    customer: 'my_customer' 
  });

  let orgUnits = [];
  if (orgUnitsResponse && orgUnitsResponse.organizationUnits) {
    orgUnits = orgUnitsResponse.organizationUnits;

    // Sort the orgUnits array based on the orgUnitPath, then by name (case-insensitive) if paths are equal
    orgUnits.sort((a, b) => {
      // Split the paths into components
      const pathA = a.orgUnitPath.split("/");
      const pathB = b.orgUnitPath.split("/");

      // Compare paths component by component
      for (let i = 0; i < Math.min(pathA.length, pathB.length); i++) {
        if (pathA[i] < pathB[i]) return -1; 
        if (pathA[i] > pathB[i]) return 1; 
      }

      // If all components match, sort by name (case-insensitive)
      if (pathA.length === pathB.length) {
        return a.name.toLowerCase().localeCompare(b.name.toLowerCase());
      }

      // If all components match and names are equal, shorter path comes first (unlikely)
      return pathA.length - pathB.length;
    });
  } 
  // If no OUs were fetched (meaning only the root OU exists), 
  // extract its details from the orgUnitsResponse itself or fetch it directly
  if (orgUnits.length === 0) {
    let rootOrgUnit = {
      orgUnitId: orgUnitsResponse.orgUnitId || "",
      name: orgUnitsResponse.name || "",
      orgUnitPath: orgUnitsResponse.orgUnitPath || "/",
      description: orgUnitsResponse.description || "",
      parentOrgUnitId: "", 
      parentOrgUnitPath: "",
    };

    // If name is still empty, fetch the domain info and use the organization name
    if (rootOrgUnit.name === "") {
      const domainInfo = AdminDirectory.Customers.get('my_customer');
      rootOrgUnit.name = domainInfo.postalAddress.organizationName || ""; 
    }

    // Set description to "Root level OU with no sub-OUs" if there are no other OUs
    rootOrgUnit.description = "Root level OU with no sub-OUs"; 

    orgUnits.push(rootOrgUnit); 
  }

  // Prepare data for the sheet (including headers)
  const fileArray = [headers];
  orgUnits.forEach((orgUnit) => {
    fileArray.push([
      orgUnit.orgUnitId ? orgUnit.orgUnitId.slice(3) : "", 
      orgUnit.name,
      orgUnit.orgUnitPath,
      orgUnit.description, // Moved description to the 4th column
      orgUnit.parentOrgUnitId ? orgUnit.parentOrgUnitId.replace(/^id:/, "") : "",
      orgUnit.parentOrgUnitPath,
    ]);
  });

  // Write data back to the sheet
  orgUnitsSheet.getRange(1, 1, fileArray.length, fileArray[0].length).setValues(fileArray);

  // Delete columns G-Z
  orgUnitsSheet.deleteColumns(7, 20);

  // Auto-resize all columns
  orgUnitsSheet.autoResizeColumns(1, orgUnitsSheet.getLastColumn()); 

  // Define ranges 
  // Only set named ranges if there's data in the sheet (orgUnits is not empty)
  if (orgUnits.length > 0) {
    spreadsheet.setNamedRange('Org2ParentPath', orgUnitsSheet.getRange('E:F')); // Updated range for Org2ParentPath
    spreadsheet.setNamedRange('OrgID2Path', orgUnitsSheet.getRange('A:C'));

    // --- Add Filter View ---
    const lastRow = orgUnitsSheet.getLastRow();
    const filterRange = orgUnitsSheet.getRange('A1:F' + lastRow);  // Adjust filter range if needed
    filterRange.createFilter();
  }
}