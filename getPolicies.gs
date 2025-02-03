let additionalServicesData = [];

function setupSheetHeader(sheet) {
  const header = ["Category", "Policy Name", "Org Unit ID", "Setting Value", "Policy Query"];
  const headerRange = sheet.getRange(1, 1, 1, header.length);

  headerRange.setValues([header]);
  headerRange.setFontWeight("bold")
    .setFontColor("#ffffff")
    .setFontFamily("Montserrat")
    .setBackground("#fc3165");

  sheet.setFrozenRows(1);  // Freeze the header row

  // Check if a filter exists before trying to create a new one
  if (!sheet.getFilter()) {
    sheet.getRange(1, 1, 1, header.length).createFilter(); //Add the filter to the header.
  }

}
function finalizeCloudIdentitySheet(sheet) {
  sheet.autoResizeColumns(1, 4);

  // Delete columns E to Z if they exist
  const lastColumn = sheet.getLastColumn();
  if (lastColumn > 5) {
    sheet.deleteColumns(5, lastColumn - 5);
  }

  const lastRow = sheet.getLastRow();
    // Determine the number of existing rows in the sheet
    const maxRows = sheet.getMaxRows(); //Get the max number of rows in the sheet
    
    //Check to make sure there is data in the sheet before deleting rows
  if (lastRow < maxRows) {
      const rowsToDelete = maxRows - lastRow; //calculate the number of rows to be deleted
      if (rowsToDelete > 0) {
      sheet.deleteRows(lastRow + 1, rowsToDelete); //This will always not delete the frozen rows.
      }
  }

  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).sort({ column: 1, ascending: true });
  }
}

function fetchAndListPolicies() {
  getGroupsSettings();
  getOrgUnits();
  const urlBase = "https://cloudidentity.googleapis.com/v1beta1/policies";
  const pageSize = 100;
  let nextPageToken = "";
  let hasNextPage = true;

  const params = {
    headers: {
      Authorization: `Bearer ${ScriptApp.getOAuthToken()}`,
    },
    muteHttpExceptions: true,
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Cloud Identity Policies");

  if (!sheet) {
    sheet = ss.insertSheet("Cloud Identity Policies");
  } else {
    sheet.clear();
  }

  setupSheetHeader(sheet);
  const policyMap = {};
  let rowCount = 0;
  additionalServicesData = [];

  while (hasNextPage) {
    try {
      const url = `${urlBase}?pageSize=${pageSize}${nextPageToken ? `&pageToken=${nextPageToken}` : ""}`;
      Logger.log(`Fetching URL: ${url}`);
      const response = UrlFetchApp.fetch(url, params);
      const jsonResponse = JSON.parse(response.getContentText());

      if (response.getResponseCode() !== 200) {
        throw new Error(`HTTP error ${response.getResponseCode()}: ${response.getContentText()}`);
      }

      const policies = jsonResponse.policies || [];
      nextPageToken = jsonResponse.nextPageToken || "";
      Logger.log(`Fetched ${policies.length} policies. NextPageToken: ${nextPageToken}`);

      policies.forEach(policy => {
        const policyData = processPolicy(policy);
        if (policyData) {
          const policyKey = `${policyData.category}-${policyData.policyName}-${policyData.orgUnitId}`;
          if (!policyMap[policyKey] || (policyData.type === 'ADMIN' && policyMap[policyKey].type !== 'ADMIN')) {
            policyMap[policyKey] = policyData;
            rowCount++;
          }
        }
      });

      hasNextPage = !!nextPageToken;
    } catch (error) {
      Logger.log(`Error fetching policies: ${error.message}`);
      hasNextPage = false;
    }
  }

  const rows = Object.values(policyMap).map(policyData => [policyData.category, policyData.policyName, policyData.orgUnitId, policyData.settingValue, policyData.policyQuery]);
  if (rows.length > 0) {
    Logger.log(`Writing ${rows.length} rows to the sheet.`);
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    sheet.hideColumn(sheet.getRange(1, 5));

    // Apply VLOOKUP after the data has been written
    applyVlookupToOrgUnitId(sheet);
  } else {
    Logger.log("No data to write to the sheet.");
  }

  finalizeCloudIdentitySheet(sheet);
  Logger.log("Policies listed successfully.");

  //Get the Cloud Identity Policies sheet
  const cloudIdentityPoliciesSheet = ss.getSheetByName("Cloud Identity Policies");
  if (cloudIdentityPoliciesSheet) {
    createOrUpdateNamedRange(cloudIdentityPoliciesSheet, "Policies", 1, 1, cloudIdentityPoliciesSheet.getLastRow(), 4);
  } else {
    Logger.log("Error: Cloud Identity Policies sheet not found for creating named range Policies");
  }

  createAdditionalServicesSheet(); 
  createWorkspaceSecurityChecklistSheet();
}

function processPolicy(policy) {
  try {
    let policyName = "";
    let orgUnitId = "";
    let settingValue = "";
    let policyQuery = "";
    let category = "";
    let type = "";

    if (policy.setting && policy.setting.type) {
      const settingType = policy.setting.type;
      const parts = settingType.split('/');
      if (parts.length > 1) {
        const categoryPart = parts[1];
        const dotIndex = categoryPart.indexOf('.');
        if (dotIndex !== -1) {
          category = categoryPart.substring(0, dotIndex);
        } else {
          category = categoryPart;
        }
      } else {
        category = "N/A";
      }
    } else {
      category = "N/A";
    }

    if (policy.setting && policy.setting.type) {
      let typeParts = policy.setting.type.split("/");
      if (typeParts.length > 1) {
        policyName = typeParts[typeParts.length - 1];
        const dotIndex = policyName.indexOf('.');
        if (dotIndex !== -1) {
          policyName = policyName.substring(dotIndex + 1);
        }
        policyName = policyName.replace(/_/g, " ");
      } else {
        policyName = policy.setting.type;
      }
    }

    if (policy.policyQuery && policy.policyQuery.query && policy.policyQuery.query.includes("groupId(")) {
      const groupIdRegex = /groupId\('([^']*)'\)/;
      const groupIdMatch = policy.policyQuery.query.match(groupIdRegex);
      if (groupIdMatch && groupIdMatch[1]) {
        orgUnitId = groupIdMatch[1];
      } else {
        orgUnitId = "Group ID not found in query";
      }
    } else if (policy.policyQuery && policy.policyQuery.orgUnit) {
      orgUnitId = getOrgUnitValue(policy.policyQuery.orgUnit);
    } else {
      orgUnitId = "SYSTEM";
    }

    if (policy.setting && policy.setting.value) {
      if (typeof policy.setting.value === 'object') {
        settingValue = formatObject(policy.setting.value);
      } else {
        settingValue = String(policy.setting.value);
      }
    } else {
      settingValue = "No setting value";
    }

    policyQuery = JSON.stringify(policy.policyQuery);
    type = policy.type;

    if (policyName === "service status") {
      additionalServicesData.push({
        service: category,
        orgUnit: orgUnitId,
        status: settingValue
      });
    }
    return { category, policyName, orgUnitId, settingValue, policyQuery, type };

  } catch (error) {
    Logger.log(`Error in processPolicy: ${error.message}`);
    return null;
  }
}
// New function to apply VLOOKUP
function applyVlookupToOrgUnitId(sheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const groupSheet = ss.getSheetByName('Group Settings'); // Corrected sheet name

  if (!groupSheet) {
    Logger.log("Error: 'Group Settings' sheet not found for VLOOKUP.");
    return;
  }

  const lastRow = groupSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("Error: No data found in Group Settings Sheet. Vlookup skipped.");
    return;
  }
  const range = groupSheet.getRange(2, 1, lastRow - 1, 3);
    //Removed this named range as it is being set in the getGroupsSettings() function
  //createOrUpdateNamedRange(groupSheet, "GroupID", 2, 1, lastRow, 3); //Modified line here

  const lastPolicyRow = sheet.getLastRow();
  for (let i = 2; i <= lastPolicyRow; i++) {
    const orgUnitIdCell = sheet.getRange(i, 3);
    let orgUnitId = orgUnitIdCell.getValue();

    if (typeof orgUnitId === 'string' && orgUnitId !== "SYSTEM") {
      orgUnitIdCell.setFormula(`=IFERROR(VLOOKUP("${orgUnitId}",GroupID, 3, false), "${orgUnitId}")`);
    }
  }
}

function createOrUpdateNamedRange(sheet, rangeName, startRow, startColumn, endRow, endColumn) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const range = sheet.getRange(startRow, startColumn, endRow - startRow + 1, endColumn - startColumn + 1); // Correct the row range
    let namedRange = ss.getRangeByName(rangeName);
    if (namedRange) {
      // Delete the old named range
      ss.removeNamedRange(rangeName);
    }
    // Create the named range with new range
    ss.setNamedRange(rangeName, range);
    Logger.log(`Named range "${rangeName}" created or updated successfully.`);

  } catch (error) {
    Logger.log(`Error creating or updating named range "${rangeName}": ${error.message}`);
  }
}

function formatObject(obj, indent = 0) {
  let formattedString = "";
  for (const key in obj) {
    if (obj.hasOwnProperty(key)) {
      const value = obj[key];
      const indentation = "  ".repeat(indent);
      formattedString += `${indentation}${key}: `;

      if (typeof value === 'object' && value !== null) {
        formattedString += '\n' + formatObject(value, indent + 1);
      } else {
        formattedString += value + '\n';
      }
    }
  }
  return formattedString;
}

function getOrgUnitValue(orgUnit) {
  if (orgUnit === "SYSTEM") return "SYSTEM";
  const orgUnitID = orgUnit.replace("orgUnits/", "");
  return `=IFERROR(VLOOKUP("${orgUnitID}", Org2ParentPath, 2, FALSE), VLOOKUP("${orgUnitID}", OrgID2Path, 3, FALSE))`;
}

function createAdditionalServicesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Additional Services");

  if (!sheet) {
    sheet = ss.insertSheet("Additional Services");
  } else {
    sheet.clear();
  }

  const header = ["Service", "OU", "Status"];
  const headerRange = sheet.getRange(1, 1, 1, header.length);

  headerRange.setValues([header]);
  headerRange.setFontWeight("bold")
    .setFontColor("#ffffff")
    .setFontFamily("Montserrat")
    .setBackground("#fc3165");
  sheet.setFrozenRows(1);
  if (!sheet.getFilter()) {
    sheet.getRange(1, 1, 1, header.length).createFilter();
  }

  // Set data from additionalServicesData
  if (additionalServicesData && additionalServicesData.length > 0) {
    const rows = additionalServicesData.map(data => [data.service, data.orgUnit, data.status]);
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
    applyVlookupToAdditionalServices(sheet, rows.length)
  }

  finalizeAdditionalServicesSheet(sheet);
}
function applyVlookupToAdditionalServices(sheet, numRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const groupSheet = ss.getSheetByName('Group Settings'); // Corrected sheet name

  if (!groupSheet) {
    Logger.log("Error: 'Group Settings' sheet not found for VLOOKUP.");
    return;
  }

  const lastRow = groupSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("Error: No data found in Group Settings Sheet. Vlookup skipped.");
    return;
  }
  const range = groupSheet.getRange(2, 1, lastRow - 1, 3);
    //Removed this named range as it is being set in the getGroupsSettings() function
  //createOrUpdateNamedRange(groupSheet, "GroupID", 2, 1, lastRow, 3); //Modified line here
  for (let i = 2; i <= numRows + 1; i++) {
    const orgUnitIdCell = sheet.getRange(i, 2);
    let orgUnitId = orgUnitIdCell.getValue();

    if (typeof orgUnitId === 'string') {
      orgUnitIdCell.setFormula(`=IFERROR(VLOOKUP("${orgUnitId}",GroupID, 3, false), "${orgUnitId}")`)
    }
  }
}
function finalizeAdditionalServicesSheet(sheet) {
  // Auto resize
  sheet.autoResizeColumns(1, 3);

  // Sort sheet by column A, A-Z
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) { // Check if there's data beyond the header row
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).sort({ column: 1, ascending: true });
  }
  
  // Delete excess columns
  const lastColumn = sheet.getLastColumn();
  if (lastColumn > 3) {
    sheet.deleteColumns(4, lastColumn - 3);
  }
}

function createWorkspaceSecurityChecklistSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Workspace Security Checklist");

  if (!sheet) {
    //If sheet doesn't exist, copy from template
    copyWorkspaceSecurityChecklistTemplate();
     sheet = ss.getSheetByName("Workspace Security Checklist"); // Get the sheet again
   if(!sheet) {
    Logger.log("Error: Workspace Security Checklist sheet not found even after copy. The script cannot be updated.");
    return;
  }
  }
  sheet.getRange('E4').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'two step verification enforcement\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'two step verification enforcement\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'two step verification enforcement\'")), CHAR(10)&CHAR(10), CHAR(10))))');
  sheet.getRange('F4').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'two step verification enforcement\'")), CHAR(10)&CHAR(10), CHAR(10)))');
  sheet.getRange('E5').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'two step verification enforcement factor\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'two step verification enforcement factor\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'two step verification enforcement factor\'")), CHAR(10)&CHAR(10), CHAR(10))))');
  sheet.getRange('F5').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'two step verification enforcement factor\'")), CHAR(10)&CHAR(10), CHAR(10)))');
  sheet.getRange('E6').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'advanced protection program\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'advanced protection program\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'advanced protection program\'")), CHAR(10)&CHAR(10), CHAR(10))))');
  sheet.getRange('F6').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10)&REPT("─", 20)&CHAR(10), QUERY(Policies, "select D where B = \'advanced protection program\'")), CHAR(10)&REPT("─", 20)&CHAR(10)&CHAR(10)&REPT("─", 20)&CHAR(10),CHAR(10)&REPT("─", 20)&CHAR(10)))');
  sheet.getRange('E10').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'password\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'password\'")), CHAR(10)&CHAR(10), CHAR(10))), "OU\'s with overridden policies:" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'password\'")), CHAR(10)&CHAR(10), CHAR(10))))');
  sheet.getRange('F10').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10)&REPT("─", 20)&CHAR(10), UNIQUE(QUERY(Policies, "select D where B = \'password\'"))), CHAR(10)&REPT("─", 20)&CHAR(10)&CHAR(10)&REPT("─", 20)&CHAR(10),CHAR(10)&REPT("─", 20)&CHAR(10) ))');
  sheet.getRange('E12').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where A = \'takeout\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where A = \'takeout\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where A = \'takeout\'")), CHAR(10)&CHAR(10), CHAR(10))))');
  sheet.getRange('F12').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where A = \'takeout\'")), CHAR(10)&CHAR(10), CHAR(10)))');
  sheet.getRange('E13').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'login challenges\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'login challenges\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'login challenges\'")), CHAR(10)&CHAR(10), CHAR(10))))');
  sheet.getRange('F13').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'login challenges\'")), CHAR(10)&CHAR(10), CHAR(10)))');
  sheet.getRange('E15').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'super admin account recovery\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'super admin account recovery\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'super admin account recovery\'")), CHAR(10)&CHAR(10), CHAR(10))))');
    sheet.getRange('F15').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'super admin account recovery\'")), CHAR(10)&CHAR(10), CHAR(10)))');
  sheet.getRange('E16').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'user account recovery\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'user account recovery\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'user account recovery\'")), CHAR(10)&CHAR(10), CHAR(10))))');
    sheet.getRange('F16').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'user account recovery\'")), CHAR(10)&CHAR(10), CHAR(10)))');
    sheet.getRange('E17').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'session controls\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'session controls\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'session controls\'")), CHAR(10)&CHAR(10), CHAR(10))))');
    sheet.getRange('F17').setFormula(`=IFERROR(LET(
    text_value, TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = 'session controls'")), CHAR(10)&CHAR(10), CHAR(10))),
    seconds, REGEXEXTRACT(text_value, "(\\d+)s$"),
    total_seconds,VALUE(seconds),
    days, INT(total_seconds / (24 * 3600)),
    remaining_seconds, MOD(total_seconds, (24 * 3600)),
    hours, INT(remaining_seconds / 3600),
    CONCATENATE(days, " days ", hours, " hours")), "")`);
  sheet.getRange('E19').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'cloud data sharing\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'cloud data sharing\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'cloud data sharing\'")), CHAR(10)&CHAR(10), CHAR(10))))');
    sheet.getRange('F19').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'cloud data sharing\'")), CHAR(10)&CHAR(10), CHAR(10)))');
// Set formulas for Third-party app access / Service account section
  sheet.getRange('E24').setFormula('=IF(ROWS(QUERY(Policies, "select C where B = \'apps access options\'")) = 0,"/",IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'apps access options\'"))) > 1,"OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'apps access options\'")), CHAR(10)&CHAR(10), CHAR(10))),TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'apps access options\'")), CHAR(10)&CHAR(10), CHAR(10)))))');
  sheet.getRange('F24').setFormula('=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'apps access options\'")), CHAR(10)&CHAR(10), CHAR(10))),"Allow users to install and run any app from the Marketplace")');
    sheet.getRange('E26').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'less secure apps\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'less secure apps\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'less secure apps\'")), CHAR(10)&CHAR(10), CHAR(10))))');
  sheet.getRange('F26').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'less secure apps\'")), CHAR(10)&CHAR(10), CHAR(10)))');
// Set formulas for Calendar section
  sheet.getRange('E43').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'primary calendar max allowed external sharing\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'primary calendar max allowed external sharing\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'primary calendar max allowed external sharing\'")), CHAR(10)&CHAR(10), CHAR(10))))');
  sheet.getRange('F43').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'primary calendar max allowed external sharing\'")), CHAR(10)&CHAR(10), CHAR(10)))');
  sheet.getRange('E45').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'external invitations\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'external invitations\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'external invitations\'")), CHAR(10)&CHAR(10), CHAR(10))))');
  sheet.getRange('F45').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'external invitations\'")), CHAR(10)&CHAR(10), CHAR(10)))');
  sheet.getRange('E46').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'secondary calendar max allowed external sharing\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'secondary calendar max allowed external sharing\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'secondary calendar max allowed external sharing\'")), CHAR(10)&CHAR(10), CHAR(10))))');
  sheet.getRange('F46').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'secondary calendar max allowed external sharing\'")), CHAR(10)&CHAR(10), CHAR(10)))');
  sheet.getRange('E49').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'appointment schedules\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'appointment schedules\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'appointment schedules\'")), CHAR(10)&CHAR(10), CHAR(10))))');
    sheet.getRange('F49').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'appointment schedules\'")), CHAR(10)&CHAR(10), CHAR(10)))');

// Set formulas for Meet section
    sheet.getRange('E51').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'video recording\'"))) > 1, "OUs with differences" & CHAR(10) & JOIN(CHAR(10), QUERY(Policies, "select C where B = \'video recording\'")), JOIN(CHAR(10), QUERY(Policies, "select C where B = \'video recording\'")))');
  sheet.getRange('F51').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'video recording\'")), CHAR(10)&CHAR(10), CHAR(10)))');
  sheet.getRange('E52').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'safety domain\'"))) > 1, "OUs with differences" & CHAR(10) & JOIN(CHAR(10), QUERY(Policies, "select C where B = \'safety domain\'")), JOIN(CHAR(10), QUERY(Policies, "select C where B = \'safety domain\'")))');
    sheet.getRange('F52').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'safety domain\'")), CHAR(10)&CHAR(10), CHAR(10)))');
    sheet.getRange('E53').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'safety access\'"))) > 1, "OUs with differences" & CHAR(10) & JOIN(CHAR(10), QUERY(Policies, "select C where B = \'safety access\'")), JOIN(CHAR(10), QUERY(Policies, "select C where B = \'safety access\'")))');
    sheet.getRange('F53').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'safety access\'")), CHAR(10)&CHAR(10), CHAR(10)))');
    sheet.getRange('E54').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'safety host management\'"))) > 1, "OUs with differences" & CHAR(10) & JOIN(CHAR(10), QUERY(Policies, "select C where B = \'safety host management\'")), JOIN(CHAR(10), QUERY(Policies, "select C where B = \'safety host management\'")))');
  sheet.getRange('F54').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'safety host management\'")), CHAR(10)&CHAR(10), CHAR(10)))');
    sheet.getRange('E55').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'safety external participants\'"))) > 1, "OUs with differences" & CHAR(10) & JOIN(CHAR(10), QUERY(Policies, "select C where B = \'safety external participants\'")), JOIN(CHAR(10), QUERY(Policies, "select C where B = \'safety external participants\'")))');
    sheet.getRange('F55').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'safety external participants\'")), CHAR(10)&CHAR(10), CHAR(10)))');

// Set formulas for Chat section
  sheet.getRange('E57').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'external chat restriction\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'external chat restriction\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'external chat restriction\'")), CHAR(10)&CHAR(10), CHAR(10))))');
  sheet.getRange('F57').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10)&REPT("─", 20)&CHAR(10), QUERY(Policies, "select D where B = \'external chat restriction\'")), CHAR(10)&REPT("─", 20)&CHAR(10)&CHAR(10)&REPT("─", 20)&CHAR(10),CHAR(10)&REPT("─", 20)&CHAR(10)))');
  sheet.getRange('E59').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'space history\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'space history\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'space history\'")), CHAR(10)&CHAR(10), CHAR(10))))');
    sheet.getRange('F59').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'space history\'")), CHAR(10)&CHAR(10), CHAR(10)))');
  sheet.getRange('E60').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'chat history\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'chat history\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'chat history\'")), CHAR(10)&CHAR(10), CHAR(10))))');
  sheet.getRange('F60').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10)&REPT("─", 20)&CHAR(10), QUERY(Policies, "select D where B = \'chat history\'")), CHAR(10)&REPT("─", 20)&CHAR(10)&CHAR(10)&REPT("─", 20)&CHAR(10),CHAR(10)&REPT("─", 20)&CHAR(10)))');
    sheet.getRange('E61').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'chat file sharing\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'chat file sharing\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'chat file sharing\'")), CHAR(10)&CHAR(10), CHAR(10))))');
    sheet.getRange('F61').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10)&REPT("─", 20)&CHAR(10), QUERY(Policies, "select D where B = \'chat file sharing\'")), CHAR(10)&REPT("─", 20)&CHAR(10)&CHAR(10)&REPT("─", 20)&CHAR(10),CHAR(10)&REPT("─", 20)&CHAR(10)))');
    sheet.getRange('E62').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'chat apps access\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'chat apps access\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'chat apps access\'")), CHAR(10)&CHAR(10), CHAR(10))))');
  sheet.getRange('F62').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10)&REPT("─", 20)&CHAR(10), QUERY(Policies, "select D where B = \'chat apps access\' and C = \'/\'")), CHAR(10)&REPT("─", 20)&CHAR(10)&CHAR(10)&REPT("─", 20)&CHAR(10),CHAR(10)&REPT("─", 20)&CHAR(10)))');
    
// Set formulas for Drive section
    sheet.getRange('E64').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'external sharing\'"))) > 1, "OUs with differences" & CHAR(10) &  TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'external sharing\'")), CHAR(10)&CHAR(10), CHAR(10))), "OU\'s with overridden policies" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'external sharing\'")), CHAR(10)&CHAR(10), CHAR(10))))');
  sheet.getRange('F64').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10), INDEX(QUERY(Policies, "select D where B = \'external sharing\' and C contains \'/\'"),1)), CHAR(10)&CHAR(10), CHAR(10)))');
  sheet.getRange('E65').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'general access default\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'general access default\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'general access default\'")), CHAR(10)&CHAR(10), CHAR(10))))');
  sheet.getRange('F65').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'general access default\'")), CHAR(10)&CHAR(10), CHAR(10)))');
  sheet.getRange('E71').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'shared drive creation\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'shared drive creation\'")), CHAR(10)&CHAR(10), CHAR(10))), "OU\'s with overridden policies:" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'shared drive creation\'")), CHAR(10)&CHAR(10), CHAR(10))))');
    sheet.getRange('F71').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10), INDEX(QUERY(Policies, "select D where B = \'shared drive creation\'"),1)), CHAR(10)&CHAR(10), CHAR(10)))');
  sheet.getRange('E73').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'drive for desktop\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'drive for desktop\'")), CHAR(10)&CHAR(10), CHAR(10))), "OU\'s with overridden policies:" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'drive for desktop\'")), CHAR(10)&CHAR(10), CHAR(10))))');
  sheet.getRange('F73').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10)&REPT("─", 20)&CHAR(10), QUERY(Policies, "select D where B = \'drive for desktop\'")), CHAR(10)&REPT("─", 20)&CHAR(10)&CHAR(10)&REPT("─", 20)&CHAR(10),CHAR(10)&REPT("─", 20)&CHAR(10)))');
    sheet.getRange('E74').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'drive sdk\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'drive sdk\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'drive sdk\'")), CHAR(10)&CHAR(10), CHAR(10))))');
  sheet.getRange('F74').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'drive sdk\'")), CHAR(10)&CHAR(10), CHAR(10)))');
    sheet.getRange('E76').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'dlp\'"))) > 1, "OUs and Groups with DLP rules" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'dlp\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'dlp\'")), CHAR(10)&CHAR(10), CHAR(10))))');
  sheet.getRange('E77').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'file security update\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'file security update\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'file security update\'")), CHAR(10)&CHAR(10), CHAR(10))))');
  sheet.getRange('F77').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10)&REPT("─", 20)&CHAR(10), QUERY(Policies, "select D where B = \'file security update\' and C = \'/\'")), CHAR(10)&REPT("─", 20)&CHAR(10)&CHAR(10)&REPT("─", 20)&CHAR(10),CHAR(10)&REPT("─", 20)&CHAR(10)))');
// Set formulas for Gmail section
 sheet.getRange('E81').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'imap access\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'imap access\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'imap access\'")), CHAR(10)&CHAR(10), CHAR(10))))');
  sheet.getRange('F81').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10)&REPT("─", 20)&CHAR(10), QUERY(Policies, "select D where B = \'imap access\'")), CHAR(10)&REPT("─", 20)&CHAR(10)&CHAR(10)&REPT("─", 20)&CHAR(10),CHAR(10)&REPT("─", 20)&CHAR(10)))');
    sheet.getRange('E82').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'pop access\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'pop access\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'pop access\'")), CHAR(10)&CHAR(10), CHAR(10))))');
  sheet.getRange('F82').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'pop access\'")), CHAR(10)&CHAR(10), CHAR(10)))');
   sheet.getRange('E83').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'auto forwarding\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'auto forwarding\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'auto forwarding\'")), CHAR(10)&CHAR(10), CHAR(10))))');
  sheet.getRange('F83').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'auto forwarding\'")), CHAR(10)&CHAR(10), CHAR(10)))');
    sheet.getRange('E84').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'per user outbound gateway\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'per user outbound gateway\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'per user outbound gateway\'")), CHAR(10)&CHAR(10), CHAR(10))))');
   sheet.getRange('F84').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'per user outbound gateway\'")), CHAR(10)&CHAR(10), CHAR(10)))');
  sheet.getRange('E87').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'spam override lists\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'spam override lists\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'spam override lists\'")), CHAR(10)&CHAR(10), CHAR(10))))');
   sheet.getRange('F87').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'spam override lists\' and C = \'/\'")), CHAR(10)&CHAR(10), CHAR(10)))');
  sheet.getRange('E89').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'enhanced pre delivery message scanning\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'enhanced pre delivery message scanning\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'enhanced pre delivery message scanning\'")), CHAR(10)&CHAR(10), CHAR(10))))');
  sheet.getRange('F89').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'enhanced pre delivery message scanning\'")), CHAR(10)&CHAR(10), CHAR(10)))');
  sheet.getRange('E92').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'email spam filter ip allowlist\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'email spam filter ip allowlist\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'email spam filter ip allowlist\'")), CHAR(10)&CHAR(10), CHAR(10))))');
  sheet.getRange('F92').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'email spam filter ip allowlist\'")), CHAR(10)&CHAR(10), CHAR(10)))');
    sheet.getRange('E95').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'email attachment safety\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'email attachment safety\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'email attachment safety\'")), CHAR(10)&CHAR(10), CHAR(10))))');
    sheet.getRange('F95').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'email attachment safety\' and C = \'/\'")), CHAR(10)&CHAR(10), CHAR(10)))');
     sheet.getRange('E96').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'links and external images\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'links and external images\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'links and external images\'")), CHAR(10)&CHAR(10), CHAR(10))))');
   sheet.getRange('F96').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'links and external images\' and C = \'/\'")), CHAR(10)&CHAR(10), CHAR(10)))');
     sheet.getRange('E97').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'spoofing and authentication\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'spoofing and authentication\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'spoofing and authentication\'")), CHAR(10)&CHAR(10), CHAR(10))))');
  sheet.getRange('F97').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'spoofing and authentication\' and C = \'/\'")), CHAR(10)&CHAR(10), CHAR(10)))');
  sheet.getRange('E102').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'content compliance\'"))) > 1, "OUs with content compliance rules" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'content compliance\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'content compliance\'")), CHAR(10)&CHAR(10), CHAR(10))))');
   sheet.getRange('E105').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'mail delegation\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'mail delegation\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'mail delegation\'")), CHAR(10)&CHAR(10), CHAR(10))))');
  sheet.getRange('F105').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10)&REPT("─", 20)&CHAR(10), QUERY(Policies, "select D where B = \'mail delegation\'")), CHAR(10)&REPT("─", 20)&CHAR(10)&CHAR(10)&REPT("─", 20)&CHAR(10),CHAR(10)&REPT("─", 20)&CHAR(10)))');
    sheet.getRange('E106').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'enhanced smime encryption\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'enhanced smime encryption\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'enhanced smime encryption\'")), CHAR(10)&CHAR(10), CHAR(10))))');
 sheet.getRange('F106').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'enhanced smime encryption\'")), CHAR(10)&CHAR(10), CHAR(10)))');
// Set formulas for Groups section
  sheet.getRange('E111').setFormula('=IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'groups sharing\'"))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'groups sharing\'")), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'groups sharing\'")), CHAR(10)&CHAR(10), CHAR(10))))');
  sheet.getRange('F111').setFormula('=TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'groups sharing\'")), CHAR(10)&CHAR(10), CHAR(10)))');
        SpreadsheetApp.getUi().alert('Inventory of settings completed');
}


function copyWorkspaceSecurityChecklistTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templateId = '1rbgKhzDYDmPDKuyx9_qR3CWpTX_ouacEKViuPwAUAf8';
  try {
    const templateSpreadsheet = SpreadsheetApp.openById(templateId);
    const templateSheet = templateSpreadsheet.getSheetByName("Workspace Security Checklist");
      if(!templateSheet) {
        Logger.log("Error: Workspace Security Checklist sheet not found in the template. Unable to continue.");
        return;
      }
      
    const newSheet = templateSheet.copyTo(ss);
    newSheet.setName("Workspace Security Checklist");
      
     //Delete Sheet1 if it exists
    let sheet1 = ss.getSheetByName("Sheet1");
    if(sheet1) {
        ss.deleteSheet(sheet1);
        Logger.log("Sheet1 deleted");
    }
      
    // Set the document title
    const domain = Session.getActiveUser().getEmail().split('@')[1];
    ss.rename('[' + domain + '] DoiT AdminPulse for Workspace');
     Logger.log("Document title set.");
      
    Logger.log("Workspace Security Checklist sheet copied from template.");
  } catch (error) {
    Logger.log(`Error copying Workspace Security Checklist template: ${error.message}`);
  }
}