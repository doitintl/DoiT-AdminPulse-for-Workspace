let additionalServicesData = [];

// =============================================
// Sheet Setup & Finalization Functions
// =============================================

function setupSheetHeader(sheet, header) {
  const filter = sheet.getFilter();
  if (filter) {
    filter.remove();
  }
  // Ensure header row exists before trying to clear/get range
  if (sheet.getMaxRows() < 1) {
      sheet.insertRowBefore(1);
  }
  sheet.getRange(1, 1, 1, header.length).clearContent();

  const headerRange = sheet.getRange(1, 1, 1, header.length);
  headerRange.setValues([header]);
  headerRange.setFontWeight("bold")
    .setFontColor("#ffffff")
    .setFontFamily("Montserrat")
    .setBackground("#fc3165");

  // Ensure frozen row setting doesn't exceed max rows
  if (sheet.getMaxRows() >= 1) {
      sheet.setFrozenRows(1);
  } else {
       Logger.log(`Skipping setFrozenRows for ${sheet.getName()} as it has 0 rows.`);
  }


  // Re-create filter, check if sheet has content rows first
  if (sheet.getLastRow() > 0) {
     try {
        // Apply filter to the range that includes potential data
        sheet.getRange(1, 1, sheet.getLastRow(), header.length).createFilter();
     } catch (e) {
        // Handle cases where filter might already exist somehow or other issues
        Logger.log(`Could not create filter on ${sheet.getName()}: ${e.message}`);
     }
  }
  Logger.log(`Header set for sheet: ${sheet.getName()}`);
}

function finalizeSheet(sheet, numColumns) {
  try {
    // Check if sheet object is valid
    if (!sheet || typeof sheet.getName !== 'function') {
        Logger.log("Error finalizing sheet: Invalid sheet object provided.");
        return;
    }
    const sheetName = sheet.getName();
    Logger.log(`Finalizing sheet: ${sheetName}...`);

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const maxRows = sheet.getMaxRows();

    // Auto-resize columns used
    if (lastCol > 0 && numColumns > 0 && lastRow > 0) { // Only resize if there's content
        try {
           sheet.autoResizeColumns(1, Math.min(lastCol, numColumns));
        } catch (e) {
            Logger.log(` Minor error during autoResizeColumns for ${sheetName}: ${e.message}`);
        }
    }

    // Delete unused columns
    if (lastCol > numColumns) {
      sheet.deleteColumns(numColumns + 1, lastCol - numColumns);
      Logger.log(`Deleted ${lastCol - numColumns} excess columns from ${sheetName}`);
    }

    // Delete unused rows
    if (lastRow >= 0 && lastRow < maxRows) { // Check lastRow >= 0
      const rowsToDelete = maxRows - Math.max(lastRow, 1); // Ensure at least 1 row remains if empty
      if (rowsToDelete > 0) {
         // Need to handle case where lastRow is 0 (only header)
         const startDeleteRow = Math.max(lastRow + 1, 2); // Start deleting from row 2 if sheet was empty
         if (startDeleteRow <= maxRows) {
             sheet.deleteRows(startDeleteRow, rowsToDelete);
             Logger.log(`Deleted ${rowsToDelete} excess rows from ${sheetName}`);
         }
      }
    }

    // Sort if data exists (rows > 1)
    if (lastRow > 1) {
      const sortLastCol = Math.min(sheet.getLastColumn(), numColumns); // Use current last column after deletes
       if (sortLastCol > 0) {
         try {
            sheet.getRange(2, 1, lastRow - 1, sortLastCol).sort({ column: 1, ascending: true });
            Logger.log(`Sorted data in ${sheetName}`);
         } catch (e) {
            Logger.log(`Error sorting data in ${sheetName}: ${e.message}`);
         }
       }
    }
     Logger.log(`Finalized sheet: ${sheetName}`);
  } catch (e) {
      Logger.log(`Error finalizing sheet ${sheet ? sheet.getName() : 'undefined'}: ${e.message} - Stack: ${e.stack}`);
  }
}


// =============================================
// Policy Fetching & Processing
// =============================================

function fetchAndListPolicies() {
  // --- Call Dependency Functions ---
  // Assumes getGroupsSettings() and getOrgUnits() exist elsewhere and
  // successfully create the required Named Ranges before this point.
  try {
      getGroupsSettings(); // Expected to create/update 'GroupID' named range
      getOrgUnits();       // Expected to create/update 'OrgID2Path', 'Org2ParentPath' named ranges
      Logger.log("Dependency functions getGroupsSettings() and getOrgUnits() executed.");
  } catch(e) {
       Logger.log(`CRITICAL ERROR: Failed running getGroupsSettings or getOrgUnits: ${e.message}. VLOOKUPs will fail.`);
       // Optionally alert user if running interactively
       // SpreadsheetApp.getUi().alert(`Error running setup functions: ${e.message}. Script cannot continue reliably.`);
       return; // Stop if dependencies fail
  }
  // --- End Dependency Calls ---

  const urlBase = "https://cloudidentity.googleapis.com/v1beta1/policies";
  const pageSize = 100;
  let nextPageToken = "";
  let hasNextPage = true;

  const params = {
    headers: {
      Authorization: `Bearer ${ScriptApp.getOAuthToken()}`,
    },
    muteHttpExceptions: true,
    method: 'get'
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "Cloud Identity Policies";
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    Logger.log(`Sheet '${sheetName}' created.`);
  } else {
    // Clear content below header, preserve header formatting
    if (sheet.getLastRow() > 1) {
       sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }
    Logger.log(`Cleared contents (below header) of sheet '${sheetName}'.`);
  }

  const header = ["Category", "Policy Name", "Org Unit ID", "Setting Value", "Policy Query"];
  setupSheetHeader(sheet, header); // Pass the header array

  const policyMap = {};
  additionalServicesData = [];

  Logger.log("Starting policy fetch loop...");
  let totalFetched = 0;
  let pageCount = 0;

  while (hasNextPage) {
    pageCount++;
    let url = `${urlBase}?pageSize=${pageSize}`;
    if (nextPageToken) {
      url += `&pageToken=${nextPageToken}`;
    }

    Logger.log(`Fetching Page ${pageCount}.`);

    try {
      const response = UrlFetchApp.fetch(url, params);
      const responseCode = response.getResponseCode();
      const responseBody = response.getContentText();

      if (responseCode !== 200) {
        let errorMessage = responseBody;
        try {
           const errorJson = JSON.parse(responseBody);
           errorMessage = errorJson.error ? JSON.stringify(errorJson.error) : responseBody;
        } catch (parseError) { /* Ignore */ }
        throw new Error(`HTTP error ${responseCode}: ${errorMessage}`);
      }

      const jsonResponse = JSON.parse(responseBody);
      const policies = jsonResponse.policies || [];
      nextPageToken = jsonResponse.nextPageToken || "";
      totalFetched += policies.length;
      Logger.log(`Fetched ${policies.length} policies (Page ${pageCount}). Total: ${totalFetched}. NextPage: ${!!nextPageToken}`);

      policies.forEach(policy => {
        const policyData = processPolicy(policy); // Returns object with raw GroupID or OU formula string in orgUnitId field
        if (policyData) {
          // Key uses the raw ID or formula string - this is ok for map uniqueness
          const policyKey = `${policyData.category}-${policyData.policyName}-${policyData.orgUnitId}`;
          if (!policyMap[policyKey] || (policyData.type === 'ADMIN' && policyMap[policyKey].type !== 'ADMIN')) {
            policyMap[policyKey] = policyData;
          }
        }
      });

      hasNextPage = !!nextPageToken;
       Utilities.sleep(100);

    } catch (error) {
      Logger.log(`ERROR during policy fetch (Page ${pageCount}): ${error.message}. Stopping pagination.`);
      hasNextPage = false;
    }
  }
  Logger.log("Policy fetch loop finished.");

  // Prepare rows for the sheet. Org Unit ID column contains raw IDs or formula strings.
  const rows = Object.values(policyMap).map(policyData => [
    policyData.category,
    policyData.policyName,
    policyData.orgUnitId,
    policyData.settingValue,
    policyData.policyQuery
  ]);

  if (rows.length > 0) {
    Logger.log(`Writing ${rows.length} rows to the sheet.`);
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows); // Write data including formulas/IDs

    // Apply VLOOKUP to ONLY the raw Group IDs AFTER writing
    applyVlookupToOrgUnitId(sheet);

    sheet.hideColumns(5); // Hide Policy Query column
    Logger.log("Applied Group VLOOKUPs (if any) and hid Policy Query column.");

  } else {
    Logger.log("No policy data to write to the sheet.");
  }

  // Finalize the sheet (resize, delete extra rows/cols, sort)
  finalizeSheet(sheet, 4); // Keep first 4 columns visible

  Logger.log("Cloud Identity Policies sheet processing completed.");

  // Update or create the "Policies" named range
  const lastPolicyRow = sheet.getLastRow();
   if (lastPolicyRow >= 1) { // Needs at least header row
       createOrUpdateNamedRange(sheet, "Policies", 1, 1, lastPolicyRow, 4); // Use 4 columns
   } else {
       Logger.log("Skipping 'Policies' named range creation as sheet is empty.");
   }


  // Create dependent sheets
  createAdditionalServicesSheet();
  createWorkspaceSecurityChecklistSheet(); // This will run last
}


function processPolicy(policy) {
  try {
    let policyName = "";
    let orgUnitId = ""; // Will hold raw Group ID, "SYSTEM", error string, or OU VLOOKUP formula string
    let settingValue = "";
    let policyQuery = "";
    let category = "";
    let type = "";

    // Category and Policy Name Extraction
    if (policy.setting && policy.setting.type) {
      const settingType = policy.setting.type;
      const parts = settingType.split('/');
      if (parts.length > 1) {
        const categoryPart = parts[1];
        const dotIndex = categoryPart.indexOf('.');
        category = (dotIndex !== -1) ? categoryPart.substring(0, dotIndex) : categoryPart;
        policyName = (dotIndex !== -1) ? categoryPart.substring(dotIndex + 1) : categoryPart;
        policyName = policyName.replace(/_/g, " ");
      } else {
        category = "N/A"; policyName = settingType;
      }
    } else {
      category = "N/A"; policyName = "Unknown";
    }

    // Determine Target (Org Unit or Group) - Using Original Logic Pattern
    if (policy.policyQuery && policy.policyQuery.query && policy.policyQuery.query.includes("groupId(")) {
      const groupIdRegex = /groupId\('([^']*)'\)/;
      const groupIdMatch = policy.policyQuery.query.match(groupIdRegex);
      orgUnitId = (groupIdMatch && groupIdMatch[1]) ? groupIdMatch[1] : "Group ID not found in query"; // RAW Group ID
    } else if (policy.policyQuery && policy.policyQuery.orgUnit) {
      orgUnitId = getOrgUnitValue(policy.policyQuery.orgUnit); // Get OU VLOOKUP FORMULA string
    } else {
      orgUnitId = "SYSTEM";
    }

    // Setting Value Processing
    if (policy.setting && policy.setting.hasOwnProperty('value')) {
      if (typeof policy.setting.value === 'object' && policy.setting.value !== null) {
        settingValue = formatObject(policy.setting.value);
      } else {
        settingValue = String(policy.setting.value);
      }
    } else {
      settingValue = "No setting value";
    }

    // Store Query and Type
    policyQuery = policy.policyQuery ? JSON.stringify(policy.policyQuery) : "{}";
    type = policy.type || 'UNKNOWN';

    // Populate Additional Services Data
    if (policyName === "service status") {
      additionalServicesData.push({
        service: category,
        orgUnit: orgUnitId, // Contains raw GroupID or OU Formula
        status: settingValue
      });
    }

    // Return processed data including the raw GroupID or the OU formula string
    return { category, policyName, orgUnitId, settingValue, policyQuery, type };

  } catch (error) {
    Logger.log(`Error in processPolicy for policy ${policy ? policy.name : 'undefined'}: ${error.message} - Stack: ${error.stack}`);
    return null;
  }
}

// =============================================
// VLOOKUP and Data Formatting Functions
// =============================================

/**
 * Generates the VLOOKUP formula string to find the Org Unit Path.
 * Relies on named ranges OrgID2Path and Org2ParentPath.
 * THIS MATCHES THE ORIGINAL SCRIPT.
 * @param {string} orgUnit The raw org unit string (e.g., "orgUnits/123abc456")
 * @return {string} The formula string or "SYSTEM".
 */
function getOrgUnitValue(orgUnit) {
  if (orgUnit === "SYSTEM" || !orgUnit) {
     return "SYSTEM";
  }
  const orgUnitID = orgUnit.replace("orgUnits/", "");
  if (!orgUnitID) {
      Logger.log(`Warning: Could not extract valid Org Unit ID from "${orgUnit}"`);
      return orgUnit;
  }
  // Original formula logic using the named ranges
  // Ensure your named ranges OrgID2Path (Col C=Path), Org2ParentPath (Col B=Path) are correct
  return `=IFERROR(VLOOKUP("${orgUnitID}", OrgID2Path, 3, FALSE), IFERROR(VLOOKUP("${orgUnitID}", Org2ParentPath, 2, FALSE),"${orgUnitID} (OU Lookup Failed)"))`;
}


/**
 * Applies VLOOKUP formula ONLY to cells containing raw Group IDs
 * in the Org Unit ID column (Col 3). Ignores existing formulas and specific strings.
 * THIS MATCHES THE ORIGINAL SCRIPT'S INTENT.
 * @param {Sheet} sheet The "Cloud Identity Policies" sheet object.
 */
function applyVlookupToOrgUnitId(sheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const groupSheet = ss.getSheetByName('Group Settings');
  const groupIDRange = ss.getRangeByName('GroupID');

  if (!groupSheet || !groupIDRange) {
    Logger.log("WARNING: 'Group Settings' sheet or 'GroupID' named range not found. Group VLOOKUPs skipped in Policies sheet.");
    return;
  }

  const lastPolicyRow = sheet.getLastRow();
  if (lastPolicyRow < 2) {
      Logger.log("No data rows in Policies sheet to apply Group VLOOKUP.");
      return;
  }

  Logger.log(`Applying Group VLOOKUPs where needed in ${sheet.getName()}...`);
  let groupLookupsApplied = 0;
  const orgUnitIdCol = 3; // Column C

  // Loop through rows individually (as per original script)
  for (let i = 2; i <= lastPolicyRow; i++) {
    const orgUnitIdCell = sheet.getRange(i, orgUnitIdCol);
    let currentCellValue = orgUnitIdCell.getValue();

    // --- Apply Group VLOOKUP only if it's a raw Group ID string ---
    // Check if it's a string, not SYSTEM, not an error, AND not already a formula
    if (typeof currentCellValue === 'string' &&
        currentCellValue !== "SYSTEM" &&
        !currentCellValue.includes(" Lookup Failed)") && // Avoid re-applying on failures
        !currentCellValue.includes(" not found in query") && // Avoid applying on errors
        !currentCellValue.startsWith('=')) // Ignore existing formulas (OU lookups)
    {
        // Assume it's a raw Group ID if it doesn't start with '=' and isn't SYSTEM/Error
        const groupIdToLookup = currentCellValue;
        // Ensure GroupID named range has group name in column 3
        const formula = `=IFERROR(VLOOKUP("${groupIdToLookup}", GroupID, 3, FALSE), "${groupIdToLookup}")`; // Original fallback logic
        try {
           orgUnitIdCell.setFormula(formula);
           groupLookupsApplied++;
        } catch (e) {
           Logger.log(`Error setting Group VLOOKUP formula in cell C${i}: ${e.message}`);
           orgUnitIdCell.setValue(`${groupIdToLookup} (Formula Error)`);
        }
    }
    // --- End Group VLOOKUP check ---
  }
  Logger.log(`Finished applying Group VLOOKUPs to Policies. Formulas applied: ${groupLookupsApplied}`);
}


/**
 * Applies VLOOKUP formula ONLY to cells containing raw Group IDs
 * in the OU/Group column (Col 2) of the Additional Services sheet.
 * THIS MATCHES THE ORIGINAL SCRIPT'S INTENT.
 * @param {Sheet} sheet The "Additional Services" sheet object.
 * @param {number} numRows The number of data rows (excluding header).
 */
function applyVlookupToAdditionalServices(sheet, numRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const groupSheet = ss.getSheetByName('Group Settings');
  const groupIDRange = ss.getRangeByName('GroupID');

  if (!groupSheet || !groupIDRange) {
    Logger.log("WARNING: 'Group Settings' sheet or 'GroupID' named range not found. Group VLOOKUPs skipped in Additional Services.");
    return;
  }

  if (numRows < 1) {
      Logger.log("No data rows in Additional Services sheet to apply VLOOKUP.");
      return;
  }

  Logger.log(`Applying Group VLOOKUPs where needed in ${sheet.getName()}...`);
  let groupLookupsApplied = 0;
  const orgUnitIdCol = 2; // Column B

  // Loop through rows individually
  for (let i = 2; i <= numRows + 1; i++) { // +1 because numRows is count, loop needs to go to last row index
    const orgUnitIdCell = sheet.getRange(i, orgUnitIdCol);
    let currentCellValue = orgUnitIdCell.getValue();

    // --- Apply Group VLOOKUP only if it's a raw Group ID string ---
    if (typeof currentCellValue === 'string' &&
        currentCellValue !== "SYSTEM" &&
        !currentCellValue.includes(" Lookup Failed)") &&
        !currentCellValue.includes(" not found in query") &&
        !currentCellValue.startsWith('='))
    {
        const groupIdToLookup = currentCellValue;
        const formula = `=IFERROR(VLOOKUP("${groupIdToLookup}", GroupID, 3, FALSE), "${groupIdToLookup}")`; // Original fallback logic
        try {
           orgUnitIdCell.setFormula(formula);
           groupLookupsApplied++;
        } catch (e) {
           Logger.log(`Error setting Group VLOOKUP formula in cell B${i}: ${e.message}`);
           orgUnitIdCell.setValue(`${groupIdToLookup} (Formula Error)`);
        }
    }
    // --- End Group VLOOKUP check ---
  }
  Logger.log(`Finished applying Group VLOOKUPs to Additional Services. Formulas applied: ${groupLookupsApplied}`);
}


function createOrUpdateNamedRange(sheet, rangeName, startRow, startColumn, endRow, endColumn) {
   if (!sheet || typeof sheet.getRange !== 'function' || endRow < startRow || endColumn < startColumn) {
     Logger.log(`Skipping named range "${rangeName}" creation/update due to invalid sheet or dimensions (EndRow: ${endRow}, StartRow: ${startRow}).`);
     return;
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const numRows = endRow - startRow + 1;
    const numCols = endColumn - startColumn + 1;
    const range = sheet.getRange(startRow, startColumn, numRows, numCols);

    let namedRange = ss.getRangeByName(rangeName);
    if (namedRange) {
      ss.removeNamedRange(rangeName);
    }
    ss.setNamedRange(rangeName, range);
    Logger.log(`Named range "${rangeName}" created/updated for range ${range.getA1Notation()} on sheet ${sheet.getName()}.`);

  } catch (error) {
    Logger.log(`Error creating or updating named range "${rangeName}": ${error.message}`);
  }
}

function formatObject(obj, indent = 0) {
  let formattedString = "";
  const maxIndent = 5;

  if (indent > maxIndent) {
      return "  ".repeat(indent) + "... (Depth Exceeded)\n";
  }

  if (Array.isArray(obj)) {
       if (obj.length === 0) return "[]\n";
       formattedString += "[\n";
       obj.forEach((item, index) => {
           formattedString += "  ".repeat(indent + 1) + `[${index}]: `;
           if (typeof item === 'object' && item !== null) {
               formattedString += "\n" + formatObject(item, indent + 2);
           } else {
               formattedString += item + "\n";
           }
       });
       formattedString += "  ".repeat(indent) + "]\n";
       return formattedString;
   }

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
  return formattedString.replace(/\n$/, ""); // Remove trailing newline
}


// =============================================
// Additional Services Sheet
// =============================================

function createAdditionalServicesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "Additional Services";
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
     Logger.log(`Sheet '${sheetName}' created.`);
  } else {
     // Clear content below header
     if (sheet.getLastRow() > 1) {
       sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
     }
     Logger.log(`Cleared contents (below header) of sheet '${sheetName}'.`);
  }

  const header = ["Service", "OU / Group", "Status"];
  setupSheetHeader(sheet, header);

  if (additionalServicesData && additionalServicesData.length > 0) {
    // Data now contains raw GroupIDs or OU Formula strings
    const rows = additionalServicesData.map(data => [data.service, data.orgUnit, data.status]);
    Logger.log(`Writing ${rows.length} rows to ${sheetName}.`);
    sheet.getRange(2, 1, rows.length, header.length).setValues(rows);

    // Apply VLOOKUP only to raw Group IDs AFTER writing
    applyVlookupToAdditionalServices(sheet, rows.length);
  } else {
      Logger.log(`No data found in additionalServicesData for sheet ${sheetName}.`);
  }

  finalizeSheet(sheet, header.length);
  Logger.log(`${sheetName} sheet processing completed.`);
}

// =============================================
// Workspace Security Checklist Sheet
// =============================================

function createWorkspaceSecurityChecklistSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "Workspace Security Checklist";
  let sheet = ss.getSheetByName(sheetName);

  const policiesRange = ss.getRangeByName("Policies");
  if (!policiesRange) {
      Logger.log("ERROR: Named range 'Policies' not found. Cannot apply formulas to Workspace Security Checklist.");
      // Optionally alert if interactive
      // SpreadsheetApp.getUi().alert("Error: Could not find policy data (Named Range 'Policies'). Checklist formulas cannot be applied.");
      return;
  }

  if (!sheet) {
    Logger.log(`Sheet '${sheetName}' not found. Attempting to copy from template.`);
    const copySuccess = copyWorkspaceSecurityChecklistTemplate();
    if (!copySuccess) {
        Logger.log(`ERROR: Failed to copy template for ${sheetName}. Cannot apply formulas.`);
        return;
    }
    sheet = ss.getSheetByName(sheetName);
    if(!sheet) {
        Logger.log(`ERROR: ${sheetName} sheet still not found after attempting copy. Cannot apply formulas.`);
        return;
    }
  } else {
      Logger.log(`Sheet '${sheetName}' found. Formulas will be applied. Existing content in E:F for non-matched rows will be preserved.`);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 1) {
      Logger.log(`${sheetName} sheet seems empty or headerless. Cannot apply formulas.`);
      return;
  }
  const columnCValues = sheet.getRange(1, 3, lastRow, 1).getValues();

  // Formula map using original logic where possible
  const formulaMap = {
    "Require 2-Step Verification for users": [
      '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'two step verification enforcement\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'two step verification enforcement\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'two step verification enforcement\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
      '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'two step verification enforcement\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "Setting Not Found")'
    ],
    "Enforce security keys for admins and high-value accounts": [
      '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'two step verification enforcement factor\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'two step verification enforcement factor\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'two step verification enforcement factor\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
      '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'two step verification enforcement factor\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "Setting Not Found")'
    ],
    "Allow users to enroll in the Advanced Protection Program": [
      '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'advanced protection program\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'advanced protection program\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'advanced protection program\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
      '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10)&REPT("─", 20)&CHAR(10), QUERY(Policies, "select D where B = \'advanced protection program\'", 0)), CHAR(10)&REPT("─", 20)&CHAR(10)&CHAR(10)&REPT("─", 20)&CHAR(10),CHAR(10)&REPT("─", 20)&CHAR(10))), "Setting Not Found")'
    ],
     "Use unique passwords": [
      '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'password\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'password\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "OU\'s with overridden policies:" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'password\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
      '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10)&REPT("─", 20)&CHAR(10), UNIQUE(QUERY(Policies, "select D where B = \'password\'", 0))), CHAR(10)&REPT("─", 20)&CHAR(10)&CHAR(10)&REPT("─", 20)&CHAR(10),CHAR(10)&REPT("─", 20)&CHAR(10) )),"Setting Not Found")'
    ],
    "Turn off Google data download as needed (Takeout)": [
      '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where A = \'takeout\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where A = \'takeout\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where A = \'takeout\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
      '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where A = \'takeout\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "Setting Not Found")'
    ],
    "Add user login challenges, add Employee ID as login challenge": [
       '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'login challenges\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'login challenges\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'login challenges\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
       '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'login challenges\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "Setting Not Found")'
    ],
    "Do not allow Super Admins to recover their own accounts": [
       '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'super admin account recovery\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'super admin account recovery\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'super admin account recovery\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
       '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'super admin account recovery\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "Setting Not Found")'
    ],
    "Do not allow users recover their own accounts": [
       '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'user account recovery\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'user account recovery\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'user account recovery\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
       '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'user account recovery\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "Setting Not Found")'
    ],
    "Configure Google session control to strengthen session expiration": [
      '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'session controls\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'session controls\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'session controls\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
      '=IFERROR(LET(text_value, TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'session controls\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), seconds, REGEXEXTRACT(text_value, "(\\d+)s$"), total_seconds,VALUE(seconds), days, INT(total_seconds / (24 * 3600)), remaining_seconds, MOD(total_seconds, (24 * 3600)), hours, INT(remaining_seconds / 3600), CONCATENATE(days, " days ", hours, " hours")), IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'session controls\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "Setting Not Found"))'
    ],
     "Share data with Google Cloud services": [
       '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'cloud data sharing\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'cloud data sharing\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'cloud data sharing\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
       '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'cloud data sharing\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "Setting Not Found")'
    ],
     "Set whether users can install Marketplace apps": [
       '=IFERROR(IF(ROWS(QUERY(Policies, "select C where B = \'apps access options\'", 0)) = 0,"/",IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'apps access options\'", 0))) > 1,"OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'apps access options\'", 0)), CHAR(10)&CHAR(10), CHAR(10))),TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'apps access options\'", 0)), CHAR(10)&CHAR(10), CHAR(10))))),"Policy Not Found")',
       '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'apps access options\'", 0)), CHAR(10)&CHAR(10), CHAR(10))),"Allow users to install and run any app from the Marketplace")'
    ],
      "Review third-party app access to core services": [
      '=IFERROR(IF(ROWS(QUERY(Policies, "select C where B = \'unconfigured third party apps\'", 0)) = 0,"/",IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'unconfigured third party apps\'", 0))) > 1,"OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'unconfigured third party apps\'", 0)), CHAR(10)&CHAR(10), CHAR(10))),TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'unconfigured third party apps\'", 0)), CHAR(10)&CHAR(10), CHAR(10))))),"Policy Not Found")',
      '=IFERROR(TRIM(JOIN(CHAR(10) & REPT("─", 20) & CHAR(10),QUERY(Policies, "select D where B = \'unconfigured third party apps\'", 0))),"(Default) Allow users to access any third-party apps.")'
    ],
    "Block access to less secure apps": [
       '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'less secure apps\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'less secure apps\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'less secure apps\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
       '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'less secure apps\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "Setting Not Found")'
    ],
    "Control access to Google core services": [
    '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'google services\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'google services\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'google services\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
    '=IFERROR(TRIM(JOIN(CHAR(10) & REPT("─", 20) & CHAR(10), QUERY(Policies, "select D where B = \'google services\'", 0))),"All Google APIs are Unrestricted")'
  ],
    "Limit external calendar sharing of primary calendars": [
       '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'primary calendar max allowed external sharing\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'primary calendar max allowed external sharing\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'primary calendar max allowed external sharing\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
       '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'primary calendar max allowed external sharing\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "Setting Not Found")'
    ],
    "Warn users when they invite external guests": [
       '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'external invitations\'", 0))) = 0, "/", IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'external invitations\'", 0))) > 1, "OUs with differences"&CHAR(10)&TEXTJOIN(CHAR(10),TRUE,UNIQUE(QUERY(Policies, "select C where B = \'external invitations\'", 0))), TEXTJOIN(CHAR(10),TRUE,UNIQUE(QUERY(Policies, "select C where B = \'external invitations\'", 0))))), "Policy Not Found")',
       '=IFERROR(IF(ROWS(QUERY(Policies, "select D where B = \'external invitations\'", 0)) = 0, "warnOnInvite: true", TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'external invitations\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))), "warnOnInvite: true")'
    ],
    "Limit external calendar sharing of seconary calendars": [
       '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'secondary calendar max allowed external sharing\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'secondary calendar max allowed external sharing\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'secondary calendar max allowed external sharing\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
       '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'secondary calendar max allowed external sharing\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "Setting Not Found")'
    ],
    "Set a policy for users to required payment for appointment schedules": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'appointment schedules\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'appointment schedules\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'appointment schedules\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'appointment schedules\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "Setting Not Found")'
    ],
    "Let people record their meetings": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'video recording\'", 0))) > 1, "OUs with differences" & CHAR(10) & JOIN(CHAR(10), QUERY(Policies, "select C where B = \'video recording\'", 0)), JOIN(CHAR(10), QUERY(Policies, "select C where B = \'video recording\'", 0))),"Policy Not Found")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'video recording\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "Setting Not Found")'
    ],
    "Set a policy for Who can join meetings created by your organization": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'safety domain\'", 0))) > 1, "OUs with differences" & CHAR(10) & JOIN(CHAR(10), QUERY(Policies, "select C where B = \'safety domain\'", 0)), JOIN(CHAR(10), QUERY(Policies, "select C where B = \'safety domain\'", 0))),"Policy Not Found")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'safety domain\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "Setting Not Found")'
    ],
    "Set a policy for Which meetings or calls users in the organization can join": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'safety access\'", 0))) > 1, "OUs with differences" & CHAR(10) & JOIN(CHAR(10), QUERY(Policies, "select C where B = \'safety access\'", 0)), JOIN(CHAR(10), QUERY(Policies, "select C where B = \'safety access\'", 0))),"Policy Not Found")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'safety access\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "Setting Not Found")'
    ],
    "Set a policy for Default host management": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'safety host management\'", 0))) > 1, "OUs with differences" & CHAR(10) & JOIN(CHAR(10), QUERY(Policies, "select C where B = \'safety host management\'", 0)), JOIN(CHAR(10), QUERY(Policies, "select C where B = \'safety host management\'", 0))),"Policy Not Found")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'safety host management\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "Setting Not Found")'
    ],
    "Warn for external Meet participants": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'safety external participants\'", 0))) > 1, "OUs with differences" & CHAR(10) & JOIN(CHAR(10), QUERY(Policies, "select C where B = \'safety external participants\'", 0)), JOIN(CHAR(10), QUERY(Policies, "select C where B = \'safety external participants\'", 0))),"Policy Not Found")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'safety external participants\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "Setting Not Found")'
    ],
    "Limit who can chat externally": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'external chat restriction\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'external chat restriction\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'external chat restriction\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10)&REPT("─", 20)&CHAR(10), QUERY(Policies, "select D where B = \'external chat restriction\'", 0)), CHAR(10)&REPT("─", 20)&CHAR(10)&CHAR(10)&REPT("─", 20)&CHAR(10),CHAR(10)&REPT("─", 20)&CHAR(10))), "Setting Not Found")'
    ],
    "Set a chat history policy for spaces": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'space history\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'space history\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'space history\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'space history\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "Setting Not Found")'
    ],
    "Set a policy for chat history and if users can turn off chat history": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'chat history\'", 0))) = 0, "/", IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'chat history\'", 0))) > 1, "OUs with differences"&CHAR(10)&TEXTJOIN(CHAR(10),TRUE,UNIQUE(QUERY(Policies, "select C where B = \'chat history\'", 0))), TEXTJOIN(CHAR(10),TRUE,UNIQUE(QUERY(Policies, "select C where B = \'chat history\'", 0))))), "Policy Not Found")',
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select D where B = \'chat history\'", 0))) = 0, "historyOnByDefault: true"&CHAR(10)&"allowUserModification: true", IF(ROWS(UNIQUE(QUERY(Policies, "select D where B = \'chat history\'", 0))) > 1, "OUs with differences"&CHAR(10)&TEXTJOIN(CHAR(10),TRUE,UNIQUE(QUERY(Policies, "select D where B = \'chat history\'", 0))), TEXTJOIN(CHAR(10),TRUE,UNIQUE(QUERY(Policies, "select D where B = \'chat history\'", 0))))), "historyOnByDefault: true"&CHAR(10)&"allowUserModification: true")'
    ],
     "Set a policy for chat file sharing": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'chat file sharing\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'chat file sharing\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'chat file sharing\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10)&REPT("─", 20)&CHAR(10), QUERY(Policies, "select D where B = \'chat file sharing\'", 0)), CHAR(10)&REPT("─", 20)&CHAR(10)&CHAR(10)&REPT("─", 20)&CHAR(10),CHAR(10)&REPT("─", 20)&CHAR(10))), "Setting Not Found")'
    ],
     "Set a policy for Chat Apps": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'chat apps access\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'chat apps access\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'chat apps access\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10)&REPT("─", 20)&CHAR(10), QUERY(Policies, "select D where B = \'chat apps access\' and C = \'/\'", 0)), CHAR(10)&REPT("─", 20)&CHAR(10)&CHAR(10)&REPT("─", 20)&CHAR(10),CHAR(10)&REPT("─", 20)&CHAR(10))), "Root Setting Not Found")'
    ],
     "Set sharing options for your domain": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'external sharing\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'external sharing\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "OU\'s with overridden policies" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'external sharing\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), INDEX(QUERY(Policies, "select D where B = \'external sharing\' and C contains \'/\'", 0),1)), CHAR(10)&CHAR(10), CHAR(10))), "Root Setting Not Found")'
    ],
    "Set the default for link sharing": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where LOWER(B) = \'general access default\'", 0))) = 0, "/", IF(ROWS(UNIQUE(QUERY(Policies, "select C where LOWER(B) = \'general access default\'", 0))) > 1, "OUs with general access differences" & CHAR(10) & TEXTJOIN(CHAR(10), TRUE, UNIQUE(QUERY(Policies, "select C where LOWER(B) = \'general access default\'", 0))), TEXTJOIN(CHAR(10), TRUE, UNIQUE(QUERY(Policies, "select C where LOWER(B) = \'general access default\'", 0))))), "Policy Not Found")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'general access default\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "Setting Not Found")'
    ],
     "Control content sharing in new shared drives": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'shared drive creation\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'shared drive creation\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "OU\'s with overridden policies:" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'shared drive creation\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), INDEX(QUERY(Policies, "select D where B = \'shared drive creation\'", 0),1)), CHAR(10)&CHAR(10), CHAR(10))), "Setting Not Found")'
    ],
     "Disable desktop access to Drive": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'drive for desktop\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'drive for desktop\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "OU\'s with overridden policies:" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'drive for desktop\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10)&REPT("─", 20)&CHAR(10), QUERY(Policies, "select D where B = \'drive for desktop\'", 0)), CHAR(10)&REPT("─", 20)&CHAR(10)&CHAR(10)&REPT("─", 20)&CHAR(10),CHAR(10)&REPT("─", 20)&CHAR(10))), "Setting Not Found")'
    ],
     "Set a policy for SDK app access to Google Drive for users": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'drive sdk\'", 0))) = 0, "/", IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'drive sdk\'", 0))) > 1, "OUs with differences" & CHAR(10) & TEXTJOIN(CHAR(10), TRUE, UNIQUE(QUERY(Policies, "select C where B = \'drive sdk\'", 0))), TEXTJOIN(CHAR(10), TRUE, UNIQUE(QUERY(Policies, "select C where B = \'drive sdk\'", 0))))), "Policy Not Found")',
        '=IFERROR(IF(ROWS(QUERY(Policies, "select D where B = \'drive sdk\'", 0)) = 0, "enableDriveSdkApiAccess: true", TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'drive sdk\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))), "enableDriveSdkApiAccess: true")'
    ],
    "Block or warn on sharing files with sensitive data": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'dlp\'", 0))) > 1, "OUs and Groups with DLP rules" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'dlp\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'dlp\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
        '"See Admin Console for details"'
    ],
    "Apply the Security update for files": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'file security update\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'file security update\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'file security update\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10)&REPT("─", 20)&CHAR(10), QUERY(Policies, "select D where B = \'file security update\' and C = \'/\'", 0)), CHAR(10)&REPT("─", 20)&CHAR(10)&CHAR(10)&REPT("─", 20)&CHAR(10),CHAR(10)&REPT("─", 20)&CHAR(10))), "Root Setting Not Found")'
    ],
    "Disable IMAP access": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'imap access\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'imap access\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'imap access\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10)&REPT("─", 20)&CHAR(10), QUERY(Policies, "select D where B = \'imap access\'", 0)), CHAR(10)&REPT("─", 20)&CHAR(10)&CHAR(10)&REPT("─", 20)&CHAR(10),CHAR(10)&REPT("─", 20)&CHAR(10))), "Setting Not Found")'
    ],
    "Disable POP access": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'pop access\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'pop access\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'pop access\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'pop access\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "Setting Not Found")'
    ],
    "Disable automatic forwarding": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'auto forwarding\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'auto forwarding\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'auto forwarding\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'auto forwarding\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "Setting Not Found")'
    ],
    "Do not allow per-user outbound gateways": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'per user outbound gateway\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'per user outbound gateway\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'per user outbound gateway\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'per user outbound gateway\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "Setting Not Found")'
    ],
     "Don't bypass spam filters for internal senders": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'spam override lists\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'spam override lists\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'spam override lists\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'spam override lists\' and C = \'/\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "Root Setting Not Found")'
    ],
    "Enable enhanced pre-delivery message scanning": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'enhanced pre delivery message scanning\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'enhanced pre delivery message scanning\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'enhanced pre delivery message scanning\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'enhanced pre delivery message scanning\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "Setting Not Found")'
    ],
    "Don't add IP addresses to your allowlist": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'email spam filter ip allowlist\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'email spam filter ip allowlist\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'email spam filter ip allowlist\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'email spam filter ip allowlist\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "Setting Not Found")'
    ],
    "Enable additional attachment protection": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'email attachment safety\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'email attachment safety\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'email attachment safety\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'email attachment safety\' and C = \'/\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "Root Setting Not Found")'
    ],
    "Enable additional link and external content protection": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'links and external images\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'links and external images\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'links and external images\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'links and external images\' and C = \'/\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "Root Setting Not Found")'
    ],
    "Enable additional spoofing protection": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'spoofing and authentication\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'spoofing and authentication\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'spoofing and authentication\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'spoofing and authentication\' and C = \'/\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "Root Setting Not Found")'
    ],
    "Scan and block emails with sensitive data": [
        '=IFERROR(IF(ROWS(QUERY(Policies, "select C where LOWER(B) = \'content compliance\'", 0)) = 0, "No Gmail Content Compliance rules found", IF(ROWS(UNIQUE(QUERY(Policies, "select C where LOWER(B) = \'content compliance\'", 0))) > 1, "OUs with content compliance rules" & CHAR(10) & TEXTJOIN(CHAR(10), TRUE, UNIQUE(QUERY(Policies, "select C where LOWER(B) = \'content compliance\'", 0))), TEXTJOIN(CHAR(10), TRUE, UNIQUE(QUERY(Policies, "select C where LOWER(B) = \'content compliance\'", 0))))), "Error querying rules")',
        '"See Admin Console for details"'
    ],
     "Disable Mail delegation, unless required": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'mail delegation\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'mail delegation\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'mail delegation\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'mail delegation\' and C = \'/\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "Root Setting Not Found")'
    ],
    "Turn on hosted S/MIME for message encryption": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'enhanced smime encryption\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'enhanced smime encryption\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'enhanced smime encryption\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'enhanced smime encryption\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "Setting Not Found")'
    ],
    "Limit group creation to admins": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'groups sharing\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'groups sharing\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'groups sharing\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'groups sharing\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "Setting Not Found")'
    ],
    "Block sharing sites outside the domain": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'sites creation and modification\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'sites creation and modification\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'sites creation and modification\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"Policy Not Found")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'sites creation and modification\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "Setting Not Found")'
    ]
  };


  Logger.log(`Applying formulas to ${sheetName}...`);
  let formulasApplied = 0;
  let formulaErrors = 0;
  let keyNotFound = 0;

  // Loop through the rows from the sheet values
  for (let i = 0; i < columnCValues.length; i++) {
    const rowNumber = i + 1;
    const policyText = columnCValues[i][0] ? String(columnCValues[i][0]).trim() : "";

    // Skip header row and empty rows in Column C
    if (rowNumber === 1 || !policyText) {
      continue;
    }

    // Check if key exists in the map
    if (formulaMap.hasOwnProperty(policyText)) {
      const formulas = formulaMap[policyText];
      if (formulas && Array.isArray(formulas) && formulas.length === 2) {
        try {
           // Set Column E formula
           sheet.getRange(rowNumber, 5).setFormula(formulas[0]);

           // Set Column F formula or value
           if (String(formulas[1]).startsWith('=')) {
              sheet.getRange(rowNumber, 6).setFormula(formulas[1]);
           } else {
              // Set as plain text (remove surrounding quotes if present)
              const cellValue = String(formulas[1]).replace(/^"|"$/g, '');
              sheet.getRange(rowNumber, 6).setValue(cellValue);
           }
           formulasApplied++;
        } catch (e) {
           Logger.log(`Error setting formula/value at row ${rowNumber} ('${policyText}'): ${e.message}`);
        }
      } else {
          // Log only once per key if definition is bad
          if (!formulaMap[`_logged_error_${policyText}`]) {
             Logger.log(`WARNING: Formula definition incomplete or malformed array for key '${policyText}' in formulaMap.`);
             formulaMap[`_logged_error_${policyText}`] = true;
          }
          formulaErrors++;
      }
    } else {
      // Key not found in map - DO NOTHING to the cell
      keyNotFound++;
      // Logging for missing keys is suppressed
    }
  } // --- End for loop ---

  Logger.log(`Finished applying formulas. Applied/Attempted: ${formulasApplied}, Map Errors: ${formulaErrors}, Keys Not Found (Ignored): ${keyNotFound}`);

  SpreadsheetApp.flush();
  Logger.log(`${sheetName} processing complete.`);
  SpreadsheetApp.getUi().alert('Workspace Policy Check is completed.');

}


// =============================================
// Template Copy Function
// =============================================

function copyWorkspaceSecurityChecklistTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templateId = '1rbgKhzDYDmPDKuyx9_qR3CWpTX_ouacEKViuPwAUAf8';
  const targetSheetName = "Workspace Security Checklist";

  try {
    Logger.log(`Attempting to open template spreadsheet with ID: ${templateId}`);
    const templateSpreadsheet = SpreadsheetApp.openById(templateId);
    Logger.log(`Template spreadsheet "${templateSpreadsheet.getName()}" opened.`);

    const templateSheet = templateSpreadsheet.getSheetByName(targetSheetName);
    if(!templateSheet) {
      const message = `Error: Sheet named "${targetSheetName}" not found in the template spreadsheet (ID: ${templateId}). Unable to copy.`;
      Logger.log(message);
      return false; // Indicate failure
    }
    Logger.log(`Found sheet "${targetSheetName}" in template.`);

    let existingSheet = ss.getSheetByName(targetSheetName);
    if (existingSheet) {
        Logger.log(`Sheet "${targetSheetName}" already exists. Deleting it before copying.`);
        ss.deleteSheet(existingSheet);
    }

    Logger.log(`Copying sheet "${targetSheetName}" to active spreadsheet...`);
    const newSheet = templateSheet.copyTo(ss);
    newSheet.setName(targetSheetName);
    ss.setActiveSheet(newSheet);
    Logger.log(`Sheet "${targetSheetName}" copied and renamed successfully.`);

    let sheet1 = ss.getSheetByName("Sheet1");
    if(sheet1 && ss.getSheets().length > 1) {
        ss.deleteSheet(sheet1);
        Logger.log("Default 'Sheet1' deleted.");
    }

    const domain = Session.getActiveUser().getEmail().split('@')[1];
    const expectedTitle = `[${domain}] DoiT AdminPulse for Workspace`;
    if (ss.getName() !== expectedTitle) {
        ss.rename(expectedTitle);
        Logger.log(`Spreadsheet title set to: "${expectedTitle}"`);
    } else {
         Logger.log("Spreadsheet title already set correctly.");
    }

    Logger.log("Workspace Security Checklist template copy process completed successfully.");
    return true;

  } catch (error) {
    const errorMessage = `Error during template copy process: ${error.message}. Check template ID, permissions, and sheet name.`;
    Logger.log(`${errorMessage} - Stack: ${error.stack}`);
    SpreadsheetApp.getUi().alert(errorMessage);
    return false;
  }
}
