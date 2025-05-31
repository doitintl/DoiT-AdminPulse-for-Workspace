let additionalServicesData = [];
var orgUnitMap = new Map();
var customerRootOuId = null;
var actualCustomerId = null;

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

function finalizeSheet(sheet, numVisibleColumns) { // Parameter is num *visible* columns
  try {
    if (!sheet || typeof sheet.getName !== 'function') {
      Logger.log("Error finalizing sheet: Invalid sheet object provided."); return;
    }
    const sheetName = sheet.getName();
    Logger.log(`Finalizing sheet: ${sheetName}...`);

    const lastRowWithContent = sheet.getLastRow();
    const currentLastCol = sheet.getLastColumn(); // Current last column, might be > numVisibleColumns if query col was just hidden
    const maxRowsInGrid = sheet.getMaxRows();
    const frozenRows = sheet.getFrozenRows();

    // Auto-resize visible columns that have content
    if (lastRowWithContent > frozenRows && numVisibleColumns > 0) {
        try { sheet.autoResizeColumns(1, numVisibleColumns); }
        catch (e) { Logger.log(`Minor error autoResizeColumns for ${sheetName}: ${e}`); }
    }

    // Delete unused columns (those beyond the *original* intended number of columns before hiding)
    // If Policy Query was column 5, and we hide it, numVisibleColumns is 4.
    // But the sheet might still have 5 columns. This needs careful thought if columns are hidden *before* this.
    // Assuming numVisibleColumns is the count of columns that SHOULD remain.
    if (currentLastCol > numVisibleColumns) {
        // If columns were hidden, getLastColumn might be less than actual data extent.
        // It's safer to use sheet.getMaxColumns() if we want to clear everything to the right.
        // For now, let's assume numVisibleColumns means actual final count.
        const actualDataLastCol = sheet.getRange("A1").offset(0, sheet.getMaxColumns()-1).getColumn(); // True last possible col
        if (actualDataLastCol > numVisibleColumns) {
             sheet.deleteColumns(numVisibleColumns + 1, actualDataLastCol - numVisibleColumns);
             Logger.log(`Deleted ${actualDataLastCol - numVisibleColumns} excess columns from ${sheetName}`);
        }
    }


    // Delete unused rows at the end
    if (lastRowWithContent < maxRowsInGrid) {
        const firstBlankRow = lastRowWithContent + 1;
        if (lastRowWithContent === frozenRows && frozenRows > 0 && firstBlankRow === frozenRows + 1) {
            if (maxRowsInGrid > firstBlankRow) { // If there's more than one non-frozen row
                sheet.deleteRows(firstBlankRow +1, maxRowsInGrid - firstBlankRow); // Keep one blank row after header
                Logger.log(`Deleted ${maxRowsInGrid - firstBlankRow} excess rows from ${sheetName} (had only frozen rows, kept one blank).`);
            } else {
                Logger.log(`Sheet ${sheetName} has only frozen rows and possibly one blank. No excess rows to delete.`);
            }
        } else if (firstBlankRow <= maxRowsInGrid) {
            const numRowsToDelete = maxRowsInGrid - firstBlankRow + 1;
            if (numRowsToDelete > 0) {
                sheet.deleteRows(firstBlankRow, numRowsToDelete);
                Logger.log(`Deleted ${numRowsToDelete} excess rows from ${sheetName}.`);
            }
        }
    }

    // Sort if data exists (more than just a header row)
    if (lastRowWithContent > Math.max(1, frozenRows)) {
      const sortLastCol = Math.min(sheet.getLastColumn(), numVisibleColumns);
      if (sortLastCol > 0) {
        try {
          const firstDataRow = frozenRows + 1;
          const numDataRows = lastRowWithContent - frozenRows;
          if (numDataRows > 0) {
            sheet.getRange(firstDataRow, 1, numDataRows, sortLastCol).sort({ column: 1, ascending: true });
            Logger.log(`Sorted data in ${sheetName}`);
          }
        } catch (e) { Logger.log(`Error sorting data in ${sheetName}: ${e}`); }
      }
    }
    Logger.log(`Finalized sheet: ${sheetName}`);
  } catch (e) { Logger.log(`Error finalizing sheet ${sheet ? sheet.getName() : 'undefined'}: ${e.message} - Stack: ${e.stack}`); }
}



// =============================================
// NEW FUNCTION to build map from sheet
// =============================================
// =============================================
// Org Unit Map Population (CORRECTED for key consistency & #ERROR! logging)
// =============================================
function populateOrgUnitMapFromSheet() {
  orgUnitMap.clear(); // Assumes orgUnitMap is a global Map instance
  Logger.log("Attempting to populate global orgUnitMap from 'Org Units' sheet...");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orgUnitsSheet = ss.getSheetByName("Org Units");

  if (!orgUnitsSheet) {
    Logger.log("ERROR: 'Org Units' sheet not found. Cannot populate orgUnitMap from sheet.");
    return false;
  }

  const lastRow = orgUnitsSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("WARNING: 'Org Units' sheet has no data beyond header. Global orgUnitMap will be empty.");
    return true; // Map is empty, but the function didn't fail to execute.
  }

  // Assuming Col A (index 0 in 'values' array): Org Unit ID (Raw)
  // Assuming Col C (index 2 in 'values' array): OrgUnit Path
  const idColumnIndex = 0;
  const pathColumnIndex = 2;
  const maxColToRead = Math.max(idColumnIndex, pathColumnIndex) + 1; // 1-based for getRange

  const range = orgUnitsSheet.getRange(2, 1, lastRow - 1, maxColToRead);
  const values = range.getValues();

  let entriesAdded = 0;
  values.forEach((row, index) => {
    const rawOrgId = row[idColumnIndex];
    const orgPathValue = row[pathColumnIndex]; // Get the value from the sheet

    // Check if rawOrgId is valid and orgPathValue is not null/undefined
    // (typeof orgPathValue === 'string' was good, but let's handle cases where it might be a Sheets error object directly)
    if (rawOrgId && String(rawOrgId).trim() !== "") {
      let mapKey = String(rawOrgId).trim();
      // Ensure all keys in orgUnitMap have the 'id:' prefix
      if (!mapKey.startsWith("id:")) {
        mapKey = "id:" + mapKey;
      }

      let cleanOrgPath;
      if (orgPathValue === null || orgPathValue === undefined) {
        cleanOrgPath = "PATH_MISSING_IN_SHEET"; // Or some other indicator
        Logger.log(`WARNING: Path is missing/null for OU ID '${mapKey}' from 'Org Units' sheet, row ${index + 2}. Using placeholder.`);
      } else {
        cleanOrgPath = String(orgPathValue).trim(); // Convert to string and trim
      }
      
      // Check for "#ERROR!" string explicitly after converting to string
      if (cleanOrgPath.toUpperCase().includes("#ERROR!") || cleanOrgPath.toUpperCase().includes("#N/A") || cleanOrgPath.toUpperCase().includes("#VALUE!") || cleanOrgPath.toUpperCase().includes("#REF!") || cleanOrgPath.toUpperCase().includes("#DIV/0!") || cleanOrgPath.toUpperCase().includes("#NUM!") || cleanOrgPath.toUpperCase().includes("#NAME?") || cleanOrgPath.toUpperCase().includes("#NULL!")) {
          Logger.log(`WARNING: Reading problematic path value ('${cleanOrgPath}') for OU ID '${mapKey}' from 'Org Units' sheet, row ${index + 2}. This will likely cause issues downstream.`);
      }
      
      orgUnitMap.set(mapKey, cleanOrgPath);
      entriesAdded++;
    } else {
      // Logger.log(`Skipping row ${index + 2} in 'Org Units' sheet during map population (missing ID or Path was not a string). ID: '${rawOrgId}', Path: '${orgPathValue}'`);
    }
  }); // Correctly closes the forEach callback

  Logger.log(`Global orgUnitMap populated from sheet with ${entriesAdded} entries. Final map size: ${orgUnitMap.size}`);

  // Enhanced Debug for root ID
  // Ensure customerRootOuId is the raw ID string like "C0xxxxxxx" (defined globally)
  const testRootOuApiId = customerRootOuId ? "id:" + customerRootOuId : "id:YOUR_FALLBACK_ROOT_OU_ID_HERE"; // Use a relevant fallback if customerRootOuId might be empty
  if (orgUnitMap.has(testRootOuApiId)) {
    const mappedPath = orgUnitMap.get(testRootOuApiId);
    Logger.log(`populateOrgUnitMapFromSheet: Global map CONTAINS root ID "${testRootOuApiId}" with path: "${mappedPath}" (Expected path should be: "/")`);
    if (mappedPath !== "/") {
        Logger.log(`WARNING: Root ID "${testRootOuApiId}" is in map, but its path is NOT "/". Path found: "${mappedPath}". This will affect root policy mapping.`);
    }
  } else {
    Logger.log(`populateOrgUnitMapFromSheet: WARNING - Global map DOES NOT CONTAIN the expected root ID "${testRootOuApiId}" after reading from sheet. This will significantly affect root policy mapping.`);
    // Check if any key maps to "/" as a fallback to identify a potential root
    let pathSlashExists = false;
    let keyForSlashPath = "";
    for (const [key, value] of orgUnitMap.entries()) {
        if (value === "/") {
            pathSlashExists = true;
            keyForSlashPath = key;
            break;
        }
    }
    if (pathSlashExists) {
        Logger.log(`INFO: Found a path "/" in orgUnitMap associated with key: "${keyForSlashPath}". This might be the root OU if the expected ID ("${testRootOuApiId}") was not found or misconfigured.`);
    } else {
        Logger.log('CRITICAL WARNING: No entry in orgUnitMap has the path "/" for the root OU. Root policy resolution will fail.');
    }
  }
  return true;
}

// =============================================
// Policy Fetching & Processing
// =============================================

function fetchAndListPolicies() {
  additionalServicesData = [];

  try {
    Logger.log("Executing getGroupsSettings()...");
    getGroupsSettings(); // Assumed to populate "Group Settings" sheet and "GroupID" named range

    Logger.log("Executing getOrgUnits() to populate 'Org Units' sheet...");
    getOrgUnits(); // Assumed to populate "Org Units" sheet correctly

    Logger.log("Executing populateOrgUnitMapFromSheet() to build global map...");
    const mapPopulatedSuccess = populateOrgUnitMapFromSheet();

    if (!mapPopulatedSuccess || orgUnitMap.size === 0) {
      Logger.log("CRITICAL ERROR: Global orgUnitMap could not be populated or is empty. Aborting policy fetch.");
      // SpreadsheetApp.getUi().alert("Error: OU information could not be loaded. Policy report may be incomplete.");
      return;
    }
    Logger.log(`Global orgUnitMap is ready with ${orgUnitMap.size} entries.`);

  } catch (e) {
    Logger.log(`CRITICAL ERROR during dependency functions or map population: ${e.toString()} - Stack: ${e.stack}`);
    return;
  }

  const urlBase = "https://cloudidentity.googleapis.com/v1beta1/policies";
  const pageSize = 100; // Max allowed by API is 100 for policies.list
  let nextPageToken = "";
  let hasNextPage = true;

  const params = {
    headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
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
    if (sheet.getLastRow() > 1) { // Clear content below header
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getMaxColumns()).clearContent();
    }
    Logger.log(`Cleared contents (below header) of sheet '${sheetName}'.`);
  }

  const header = ["Category", "Policy Name", "OU Path / Group ID", "Setting Value", "Policy Query"];
  setupSheetHeader(sheet, header);

  const policyMap = {}; // Using a map to handle potential duplicates/overrides
  Logger.log("Starting policy fetch loop (this will fetch ALL pages)...");
  let totalFetchedPolicies = 0;
  let pageCount = 0;
  let apiCallErrors = 0;

  while (hasNextPage) {
    pageCount++;
    // The query URL for listing policies generally does not take a 'query' parameter itself.
    // It lists all policies. The 'policyQuery' is a field *within* the policy object.
    // An "empty query" to get everything is achieved by not adding specific filter parameters to the list call.
    let url = `${urlBase}?pageSize=${pageSize}${nextPageToken ? '&pageToken=' + nextPageToken : ''}`;
    Logger.log(`Fetching Page ${pageCount}. URL: ${url.substring(0, url.indexOf('?') + 15)}... (pageSize & token)`);

    let response;
    let responseBody;
    try {
      response = UrlFetchApp.fetch(url, params);
      responseBody = response.getContentText();
      const responseCode = response.getResponseCode();

      if (responseCode === 200) {
        const jsonResponse = JSON.parse(responseBody);
        const policies = jsonResponse.policies || [];
        nextPageToken = jsonResponse.nextPageToken || "";
        hasNextPage = !!nextPageToken;
        totalFetchedPolicies += policies.length;
        Logger.log(`Fetched ${policies.length} policies (Page ${pageCount}). Total cumulative: ${totalFetchedPolicies}. NextPage: ${hasNextPage}`);

        policies.forEach(policy => {
          // Log raw policy for deeper debugging if settings are still missing
          // Logger.log(`Raw policy from API: ${JSON.stringify(policy)}`);
          const policyData = processPolicy(policy);

          if (policyData && typeof policyData.orgUnitId === 'string') { // Check orgUnitId is a string
            const policyKey = `${policyData.category}-${policyData.policyName}-${policyData.orgUnitId}`;
            const existingPolicyInMap = policyMap[policyKey];

            // Prefer 'ADMIN' type policies. If types are same or new is not ADMIN, keep first encountered.
            if (!existingPolicyInMap || (policyData.type === 'ADMIN' && existingPolicyInMap.type !== 'ADMIN')) {
              policyMap[policyKey] = policyData;
            }
          } else if (policyData && policyData.orgUnitId === undefined) {
            Logger.log(`WARNING: processPolicy returned data but orgUnitId is undefined. Raw policy setting type: ${policy.setting ? policy.setting.type : 'N/A'}, Policy Query: ${JSON.stringify(policy.policyQuery)}`);
          } else if (!policyData) {
            Logger.log(`WARNING: processPolicy returned null. Raw policy setting type: ${policy.setting ? policy.setting.type : 'N/A'}, Policy Query: ${JSON.stringify(policy.policyQuery)}`);
          }
        });
      } else {
        Logger.log(`ERROR: API call failed for Page ${pageCount}. HTTP ${responseCode}. Response: ${responseBody.substring(0, 500)}`);
        apiCallErrors++;
        if (apiCallErrors > 3) { // Arbitrary limit to prevent infinite loops on persistent errors
             Logger.log("Too many API errors. Aborting policy fetch loop.");
             hasNextPage = false; // Stop trying
        }
        // Consider a small delay before retrying the same pageToken, or just break
        Utilities.sleep(2000); // Wait 2s before potential next attempt (if hasNextPage was true due to previous token)
        // If error on first page, nextPageToken is still "", loop might retry first page.
        // If error on subsequent page, it will retry with same nextPageToken.
        // This simple retry might not be robust enough for all scenarios.
      }
    } catch (e) {
      Logger.log(`EXCEPTION during UrlFetchApp or JSON.parse for Page ${pageCount}: ${e.message}. Response (first 500 chars): ${responseBody ? responseBody.substring(0,500) : 'N/A'}. Stack: ${e.stack}`);
      apiCallErrors++;
      if (apiCallErrors > 3) {
           Logger.log("Too many critical errors. Aborting policy fetch loop.");
           hasNextPage = false;
      }
      Utilities.sleep(2000);
    }
     if (pageCount > 200 && hasNextPage) { // Safety break for very large number of pages / runaway loop
        Logger.log("WARNING: Exceeded 200 pages. Breaking loop as a precaution.");
        hasNextPage = false;
    }
  } // End while hasNextPage

  Logger.log(`Policy fetch loop completed. Total policies processed into map (pre-filtering): ${Object.keys(policyMap).length}`);

  const rows = Object.values(policyMap).map(data => [
    data.category,
    data.policyName,
    data.orgUnitId, // This is the OU Path string or Group ID string
    data.settingValue,
    data.policyQuery
  ]);

  if (rows.length > 0) {
    Logger.log(`Writing ${rows.length} unique/prioritized policies to sheet '${sheetName}'.`);
    // Example debug for the first row's target
    // Logger.log(`First policy row to be written - Target (Col C): '${rows[0][2]}', Category: '${rows[0][0]}', Name: '${rows[0][1]}'`);
    sheet.getRange(2, 1, rows.length, header.length).setValues(rows); // Use header.length for num columns
    
    Logger.log("Applying VLOOKUPs for Group names in 'Cloud Identity Policies' sheet...");
    applyVlookupToOrgUnitId(sheet); // This function targets column C for Group ID lookups
    
    const policyQueryColumnIndex = header.indexOf("Policy Query") + 1;
    if (policyQueryColumnIndex > 0) {
        sheet.hideColumns(policyQueryColumnIndex);
        Logger.log("Hid 'Policy Query' column.");
    } else {
        Logger.log("Could not find 'Policy Query' column to hide.");
    }
  } else {
    Logger.log("No policy data to write to the 'Cloud Identity Policies' sheet.");
  }

  finalizeSheet(sheet, header.length - 1); // -1 because one column ("Policy Query") is hidden

  Logger.log("'Cloud Identity Policies' sheet processing completed.");
  const lastPolicyRow = sheet.getLastRow();
  if (lastPolicyRow >= 1) {
    createOrUpdateNamedRange(sheet, "Policies", 1, 1, lastPolicyRow, header.length -1); // Use visible columns for named range
  } else {
    Logger.log("Skipping 'Policies' named range creation as 'Cloud Identity Policies' sheet is empty or header only.");
  }

  createAdditionalServicesSheet();
  createWorkspaceSecurityChecklistSheet();
  Logger.log("All processing finished for fetchAndListPolicies.");
}

function processPolicy(policy) {
   let policyName = "Unknown Policy Name";
  let targetOuOrGroup = "Unknown Target"; // This will become the value for 'orgUnit' in additionalServicesData
  let settingValue = "No setting value";
  let policyQueryString = policy.policyQuery ? JSON.stringify(policy.policyQuery) : "{}";
  let category = "general";
  let type = policy.type || 'UNKNOWN_POLICY_TYPE'; // e.g., 'ADMIN', 'USER'

  const rawSettingType = policy.setting && policy.setting.type ? policy.setting.type : "UnknownSettingType";

  try {
    let typeToParse = rawSettingType;

    // 1. Strip known prefixes
    if (typeToParse.startsWith("settings/")) {
      typeToParse = typeToParse.substring("settings/".length);
    }
    const genericPrefixMatch = typeToParse.match(/^(booleans|strings|integers|listValue|enumValue|double)\/(.+)/);
    if (genericPrefixMatch && genericPrefixMatch[2]) {
      typeToParse = genericPrefixMatch[2];
    }

    // 2. Extract Category and Policy Name
    const firstDot = typeToParse.indexOf('.');
    const firstSlash = typeToParse.indexOf('/');

    if (firstDot !== -1 && (firstSlash === -1 || firstDot < firstSlash)) { // Dot exists and is primary
        category = typeToParse.substring(0, firstDot);
        policyName = typeToParse.substring(firstDot + 1);
    } else if (firstSlash !== -1) { // No dot before slash, or no dot at all
        category = typeToParse.substring(0, firstSlash);
        policyName = typeToParse.substring(firstSlash + 1);
    } else { // No dot, no slash, or simple type
        if (typeToParse.toLowerCase().startsWith("chrome.")) {
            category = "chrome";
            policyName = typeToParse.substring("chrome.".length);
        } else if (typeToParse.toLowerCase().includes("chrome")) {
            category = "chrome";
            policyName = typeToParse;
        } else {
            category = "general"; // Fallback category
            policyName = typeToParse; // Use the remaining string as policy name
        }
    }

    // 3. Clean up names
    category = String(category).replace(/_/g, " ").trim() || "general";
    policyName = String(policyName).replace(/[._]/g, " ").trim() || rawSettingType; // Fallback to raw type if name is empty

    // Logger.log(`Parsed: RawType='${rawSettingType}' -> Category='${category}', PolicyName='${policyName}'`);

    // 4. Determine Target OU or Group
    if (policy.policyQuery && policy.policyQuery.query && policy.policyQuery.query.includes("groupId(")) {
      const groupIdRegex = /groupId\('([^']*)'\)/;
      const groupIdMatch = policy.policyQuery.query.match(groupIdRegex);
      targetOuOrGroup = (groupIdMatch && groupIdMatch[1]) ? groupIdMatch[1] : `Group ID Parse Error: ${policy.policyQuery.query}`;
    } else if (policy.policyQuery && policy.policyQuery.orgUnit) {
      const rawOuTargetString = policy.policyQuery.orgUnit;
      
      if (rawOuTargetString.toLowerCase() === "orgunits/customer/my_customer") {
        let rootPathFound = false;
        // Prioritize mapping via customerRootOuId if available and path is "/" in orgUnitMap
        // Assumes customerRootOuId is the raw customer ID like "C0xxxxxxx"
        // Assumes orgUnitMap keys are prefixed like "id:C0xxxxxxx"
        if (customerRootOuId && orgUnitMap.has("id:" + customerRootOuId) && orgUnitMap.get("id:" + customerRootOuId) === "/") {
            targetOuOrGroup = "/";
            rootPathFound = true;
        } else { // Fallback: find any ID in orgUnitMap that points to path "/"
            for (const [ou_id_in_map, ou_path_in_map] of orgUnitMap.entries()) {
                if (ou_path_in_map === "/") {
                    targetOuOrGroup = "/";
                    rootPathFound = true;
                    // Logger.log(`Mapped 'orgunits/customer/my_customer' to path "/" via fallback map entry: ${ou_id_in_map} -> ${ou_path_in_map}`);
                    break; 
                }
            }
        }
        if (!rootPathFound) {
            targetOuOrGroup = "/ (Root OU Not Mapped or Path Incorrect in orgUnitMap)";
            Logger.log(`WARNING: Could not map 'orgunits/customer/my_customer' to path "/". CustomerRootOuId: '${customerRootOuId}'. Map lookup for 'id:${customerRootOuId}': ${orgUnitMap.get("id:"+customerRootOuId)}`);
        }
      } else {
        let idPart = rawOuTargetString.replace("orgUnits/", ""); // e.g., "id:03ph8a2z1a2gsfw" or "03ph8a2z1a2gsfw"
        // Ensure 'id:' prefix for lookup, consistent with how populateOrgUnitMapFromSheet stores keys
        let finalIdToLookup = idPart.startsWith("id:") ? idPart : "id:" + idPart;

        if (orgUnitMap.has(finalIdToLookup)) {
          targetOuOrGroup = orgUnitMap.get(finalIdToLookup); // This is where #ERROR! from the map would be retrieved
        } else {
          targetOuOrGroup = finalIdToLookup; // Fallback to the (prefixed) ID if not in map
          Logger.log(`WARNING: OU ID '${finalIdToLookup}' (from API target '${rawOuTargetString}') not found in orgUnitMap. Using raw ID as target.`);
        }
      }
    } else {
      targetOuOrGroup = "SYSTEM_OR_UNSPECIFIED_TARGET"; // Default if no OU or Group query
      // Logger.log(`Policy has no specific OU/Group target in policyQuery. Raw policy: ${JSON.stringify(policy)}`);
    }

    // 5. Get Setting Value
    if (policy.setting && policy.setting.hasOwnProperty('value')) {
      settingValue = (typeof policy.setting.value === 'object' && policy.setting.value !== null) ?
                     formatObject(policy.setting.value) : String(policy.setting.value);
    } else {
      settingValue = "NO_EXPLICIT_VALUE"; // Indicates the policy might be structural or value is elsewhere
    }

    // 6. Populate additionalServicesData for "service status" policies
    if (String(policyName).toLowerCase().trim() === "service status") {
      // *** ADDED DEBUG LOGGING HERE for targetOuOrGroup issues ***
      if (typeof targetOuOrGroup === 'string') {
        const upperTarget = targetOuOrGroup.toUpperCase();
        if (upperTarget.includes("#ERROR!") || upperTarget.includes("#N/A") || upperTarget.includes("#VALUE!") || upperTarget.includes("#REF!") || upperTarget.includes("#DIV/0!") || upperTarget.includes("#NUM!") || upperTarget.includes("#NAME?") || upperTarget.includes("#NULL!")) {
            Logger.log(`WARNING: Pushing problematic target ('${targetOuOrGroup}') for service '${category}' into additionalServicesData. This likely originated from the orgUnitMap.`);
        }
      } else if (targetOuOrGroup === null || targetOuOrGroup === undefined) {
         Logger.log(`WARNING: Target for service '${category}' is null/undefined before pushing to additionalServicesData.`);
      }

      additionalServicesData.push({
        service: category,        // e.g., "Gmail", "Drive" (parsed from policy type)
        orgUnit: targetOuOrGroup, // Resolved OU Path string (potentially "#ERROR!") or Group ID string
        status: settingValue      // e.g., "ON", "OFF", or structured value
      });
      Logger.log(`Added to additionalServicesData: Service='${category}', Target='${targetOuOrGroup}', Status='${settingValue}'`);
    }

    // 7. Return structured policy data
    return {
      category: category,
      policyName: policyName,
      orgUnitId: targetOuOrGroup, // This is the resolved OU Path or Group ID for the "Cloud Identity Policies" sheet
      settingValue: settingValue,
      policyQuery: policyQueryString,
      type: type
    };

  } catch (error) {
    Logger.log(`ERROR in processPolicy for rawSettingType '${rawSettingType}': ${error.message} - Policy (first 500 chars): ${JSON.stringify(policy).substring(0,500)}... - Stack: ${error.stack}`);
    return null; // Return null if processing fails catastrophically
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
 * in the "OU Path / Group ID" column (Col 3). Ignores resolved OU names and specific strings.
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
  if (lastPolicyRow < 2) { // Need at least one data row
    Logger.log("No data rows in Policies sheet to apply Group VLOOKUP.");
    return;
  }

  Logger.log(`Applying Group VLOOKUPs (if needed) to Col C in ${sheet.getName()} from row 2 to ${lastPolicyRow}...`);
  let groupLookupsApplied = 0;
  const targetIdCol = 3; // Column C holds "OU Path / Group ID"

  // Iterate through each relevant cell in the target column
  for (let r = 2; r <= lastPolicyRow; r++) {
    const cell = sheet.getRange(r, targetIdCol);
    // Use getDisplayValue() to get exactly what's visible, helps with non-string types or formatting
    const currentCellValue = cell.getDisplayValue();
    const currentCellFormula = cell.getFormula();

    // Logger.log(`Row ${r}, Col C - DisplayValue: '${currentCellValue}', Formula: '${currentCellFormula}'`); // UNCOMMENT FOR INTENSE DEBUG

    let isLikelyGroupID = false;
    if (typeof currentCellValue === 'string' && currentCellValue.trim() !== "") {
      const trimmedValue = currentCellValue.trim();
      if (
          trimmedValue !== "SYSTEM" &&
          !trimmedValue.startsWith('/') &&           // <<<< CRUCIAL: Skips OU Paths
          !trimmedValue.includes("(Group Not Found)") && // Avoid re-processing failed lookups
          !trimmedValue.includes("(OU Lookup Failed)") && // Avoid processing previous OU lookup failures
          !trimmedValue.includes("not found in query") && // Avoid processing error strings
          !trimmedValue.startsWith('id:') &&        // Skips raw OU IDs that weren't mapped
          currentCellFormula === ""                 // Only apply if cell doesn't already have a formula
                                                    // This also means if it's a path (value), formula is "", so other conditions must prevent it
      ) {
        // Further check: Does it look like an email? (Simple check for '@')
        // Or is it a numerical-like ID that Groups might use but OU paths don't?
        // This depends on your group ID format. If group IDs are emails, this helps.
        if (trimmedValue.includes('@') || /^[a-zA-Z0-9]+$/.test(trimmedValue)) { // Looks like an email or an alphanumeric ID
             isLikelyGroupID = true;
        }
      }
    }

    if (isLikelyGroupID) {
      const groupIdToLookup = currentCellValue.trim();
      Logger.log(`Row ${r}: Applying Group VLOOKUP for potential Group ID '${groupIdToLookup}'.`);
      const formula = `=IFERROR(VLOOKUP("${groupIdToLookup}", GroupID, 3, FALSE), "${groupIdToLookup} (Group Not Found)")`;
      try {
        cell.setFormula(formula);
        groupLookupsApplied++;
      } catch (e) {
        Logger.log(`Error setting Group VLOOKUP for '${groupIdToLookup}' in cell C${r}: ${e.message}`);
        cell.setValue(`${groupIdToLookup} (Formula Error)`); // Fallback on error
      }
    } else {
      // Logger.log(`Row ${r}: Skipping Group VLOOKUP for Col C value '${currentCellValue}'.`); // UNCOMMENT FOR INTENSE DEBUG
    }
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
  const groupIDRange = ss.getRangeByName('GroupID'); // Assumes GroupID: Col A=ID, Col C=Name

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
  const orgUnitIdCol = 2; // Column B: "OU / Group"

  // Get all values in the column at once
  const range = sheet.getRange(2, orgUnitIdCol, numRows, 1);
  const values = range.getValues();
  const formulas = range.getFormulasR1C1(); // Get existing formulas to preserve them

  for (let i = 0; i < values.length; i++) {
    let currentCellValue = values[i][0];

    // Check if it's a string, not "SYSTEM", not an OU path (doesn't start with "/"),
    // not an error, AND not already a formula.
    if (typeof currentCellValue === 'string' &&
        currentCellValue !== "SYSTEM" &&
        !currentCellValue.startsWith('/') && // OU Paths/Names will start with /
        !currentCellValue.includes(" Lookup Failed)") &&
        !currentCellValue.includes(" not found in query") &&
        !currentCellValue.startsWith('id:') && // Ignore raw OU IDs that weren't mapped
        !currentCellValue.startsWith('=')) // Check if the *value* starts with =, not just if getFormula is non-empty
    {
        const groupIdToLookup = currentCellValue;
        // Ensure GroupID named range has group name in column 3
        formulas[i][0] = `=IFERROR(VLOOKUP("${groupIdToLookup}", GroupID, 3, FALSE), "${groupIdToLookup} (Group Not Found)")`;
        groupLookupsApplied++;
    }
    // If it doesn't meet criteria, formulas[i][0] will retain its original formula or be empty if it was a value.
    // This needs to be handled carefully if mixing setValues and setFormulas.
    // Better to rebuild the formulas array for setFormulas.
  }

  if (groupLookupsApplied > 0) {
      range.setFormulasR1C1(formulas); // Set all formulas (new and existing)
  }

  Logger.log(`Finished applying Group VLOOKUPs to Additional Services. Formulas potentially modified/set: ${groupLookupsApplied}`);
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
