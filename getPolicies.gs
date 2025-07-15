// ---------------------------------------------
// Part 1: Configuration & Main Controller
// ---------------------------------------------

const SCRIPT_NAME = "Workspace Policy Check";
const TRIGGER_FUNCTION_NAME = "continuePolicyFetchAndProcess";
const MAX_RUNTIME_MINUTES = 28;


/**
 * Main function. Initializes, runs dependencies, and starts the policy fetch.
 */
function runFullPolicyCheck() {
  const scriptLock = LockService.getScriptLock();
  if (!scriptLock.tryLock(1000)) {
    Logger.log("A policy check is already running. Exiting.");
    SpreadsheetApp.getActiveSpreadsheet().toast("A policy check is already in progress. Please wait for it to complete.", "Info", 10);
    return;
  }

  try {
    const startTime = new Date();
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet(); 

    Logger.log(`============================================================`);
    Logger.log(`▶️ START: Running '${SCRIPT_NAME}' at ${startTime.toLocaleString()}`);
    Logger.log(`============================================================`);

    deleteTriggers();
    PropertiesService.getScriptProperties().deleteAllProperties();
    Logger.log("Cleaned up old triggers and properties for a fresh run.");

    try {
      Logger.log("--- Calling Dependency: getGroupsSettings() ---");
      ss.toast('Updating Group data...', SCRIPT_NAME, -1); 
      getGroupsSettings();
      SpreadsheetApp.flush(); // Force the spreadsheet to update

      Logger.log("--- Calling Dependency: getOrgUnits() ---");
      ss.toast('Updating OU data...', SCRIPT_NAME, -1);
      getOrgUnits();
      SpreadsheetApp.flush(); // Force the spreadsheet to update

    } catch (e) {
      const errorMessage = `A dependency script ('getGroupsSettings' or 'getOrgUnits') failed to run. Error: ${e.message}. The script cannot continue.`;
      Logger.log(errorMessage + `
Stack: ${e.stack}`);
      ss.toast(e.message, '❌ Dependency Error', 30);
      ui.alert(errorMessage);
      return;
    }
    
    if (!ss.getRangeByName('GroupID') || !ss.getRangeByName('OrgID2Path')) {
        const errorMessage = `VALIDATION FAILED: Required named ranges ('GroupID', 'OrgID2Path') were not found after dependency scripts ran. Please ensure they create these ranges correctly. The script cannot continue.`;
        Logger.log(errorMessage);
        ss.toast(errorMessage, '❌ Validation Error', 30);
        ui.alert(errorMessage);
        return;
    }
    
    Logger.log("✅ All dependency scripts ran and outputs were validated.");
    ss.toast('Dependencies validated. Starting policy fetch...', SCRIPT_NAME, 10);
    SpreadsheetApp.flush();

    PropertiesService.getScriptProperties().setProperty('startTime', startTime.getTime());
    continuePolicyFetchAndProcess();

  } finally {
    scriptLock.releaseLock();
    Logger.log("Script lock released.");
  }
}''


// ---------------------------------------------
// Part 2: Timeout, Continuation, and Processing Logic
// ---------------------------------------------

function isTimeUp(startTime) {
  const maxRuntimeSeconds = MAX_RUNTIME_MINUTES * 60;
  return (new Date().getTime() - Number(startTime)) / 1000 >= maxRuntimeSeconds;
}

function deleteTriggers() {
  try {
    ScriptApp.getProjectTriggers().forEach(trigger => {
      if (trigger.getHandlerFunction() === TRIGGER_FUNCTION_NAME) {
        ScriptApp.deleteTrigger(trigger);
      }
    });
  } catch (e) { /* Ignore */ }
}

function continuePolicyFetchAndProcess() {
  // --- FIX #1: Define 'ss' within this function's scope ---
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const scriptProperties = PropertiesService.getScriptProperties();
  const properties = scriptProperties.getProperties();

  const startTime = Number(properties.startTime || new Date().getTime());
  let nextPageToken = properties.nextPageToken || "";
  let policyMap = properties.policyMap ? JSON.parse(properties.policyMap) : {};
  let additionalServicesData = properties.additionalServicesData ? JSON.parse(properties.additionalServicesData) : [];
  
  // --- FIX #2: Define 'initialPolicyCount' before it is used ---
  const initialPolicyCount = Object.keys(policyMap).length;
  Logger.log(`Resuming process. Policies collected so far: ${initialPolicyCount}.`);
  
  // Now this toast will work correctly
  ss.toast(`Fetching policies... (${initialPolicyCount} collected so far)`, SCRIPT_NAME, 20);
  SpreadsheetApp.flush();

  const urlBase = "https://cloudidentity.googleapis.com/v1beta1/policies";
  const pageSize = 100;
  let hasNextPage = true;
  const params = {
    headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
    muteHttpExceptions: true,
    method: 'get'
  };

  Logger.log("Starting policy fetch batch...");
  do {
    if (isTimeUp(startTime)) {
      Logger.log("Approaching time limit. Pausing execution.");
      // This toast will also work now
      ss.toast('Pausing to avoid timeout. Will resume in 1 minute.', SCRIPT_NAME, 60);      
      hasNextPage = true;
      break;
    }

    let url = `${urlBase}?pageSize=${pageSize}${nextPageToken ? `&pageToken=${nextPageToken}` : ''}`;

    try {
      const response = UrlFetchApp.fetch(url, params);
      const responseCode = response.getResponseCode();
      if (responseCode !== 200) throw new Error(`HTTP ${responseCode}: ${response.getContentText()}`);
      
      const jsonResponse = JSON.parse(response.getContentText());
      const policies = jsonResponse.policies || [];
      nextPageToken = jsonResponse.nextPageToken || "";
      Logger.log(`Fetched ${policies.length} policies. NextPage: ${!!nextPageToken}`);

      policies.forEach(policy => {
        const policyData = processPolicy(policy, additionalServicesData); 
        if (policyData) {
          const policyKey = `${policyData.category}-${policyData.policyName}-${policyData.orgUnitId}`;
          if (!policyMap[policyKey] || (policyData.type === 'ADMIN' && policyMap[policyKey].type !== 'ADMIN')) {
            policyMap[policyKey] = policyData;
          }
        }
      });
      
      // And this toast will work
      ss.toast(`Fetching policies... (${Object.keys(policyMap).length} collected)`, SCRIPT_NAME, 20);
      SpreadsheetApp.flush(); // Flush to see updates during the loop

      hasNextPage = !!nextPageToken;
      if (hasNextPage) Utilities.sleep(100);

    } catch (error) {
      let userMessage = `An unexpected error occurred during the policy fetch: ${error.message}. The script cannot continue.`;
      let errorTitle = '❌ Network Error';

      // Check specifically for a 403 Permission Denied error
      if (error.message.includes("HTTP 403") || error.message.includes("Permission denied")) {
        userMessage = "Permission Denied: Could not fetch policies. This report requires Super Administrator privileges to run successfully.";
        errorTitle = '❌ Permission Error';
      }

      Logger.log(`FATAL ERROR during policy fetch: ${error.message}. Halting execution.`);
      Logger.log(`Stack: ${error.stack}`);

      ss.toast(userMessage, errorTitle, 60);
      ui.alert(errorTitle, userMessage, ui.ButtonSet.OK);

      deleteTriggers();
      PropertiesService.getScriptProperties().deleteAllProperties();
      
      // CRITICAL: Halt the function completely.
      return; 
    }
  } while (hasNextPage);

  if (hasNextPage) {
    Logger.log("Saving state and setting trigger to continue...");
    scriptProperties.setProperties({
        'nextPageToken': nextPageToken,
        'policyMap': JSON.stringify(policyMap),
        'additionalServicesData': JSON.stringify(additionalServicesData),
        'startTime': startTime
    });
    deleteTriggers(); 
    ScriptApp.newTrigger(TRIGGER_FUNCTION_NAME).timeBased().after(60 * 1000).create();
    Logger.log(`⏸️ PAUSING. Trigger set to continue.`);
  } else {
    Logger.log("✅ All policies fetched. Finalizing sheets.");
    finalizeAllSheets(policyMap, additionalServicesData, startTime);
  }
}

function processPolicy(policy, servicesDataArray) {
  try {
    let policyName = "", orgUnitId = "", settingValue = "", policyQuery = "", category = "", type = "";
    if (policy.setting && policy.setting.type) {
      const parts = policy.setting.type.split('/');
      if (parts.length > 1) {
        const categoryPart = parts[1];
        const dotIndex = categoryPart.indexOf('.');
        category = (dotIndex !== -1) ? categoryPart.substring(0, dotIndex) : categoryPart;
        policyName = (dotIndex !== -1) ? categoryPart.substring(dotIndex + 1) : categoryPart;
        policyName = policyName.replace(/_/g, " ");
      } else { category = "N/A"; policyName = policy.setting.type; }
    } else { category = "N/A"; policyName = "Unknown"; }
    
    if (policy.policyQuery && policy.policyQuery.query && policy.policyQuery.query.includes("groupId(")) {
      const groupIdMatch = policy.policyQuery.query.match(/groupId\('([^']*)'\)/);
      orgUnitId = (groupIdMatch && groupIdMatch[1]) ? groupIdMatch[1] : "Group ID not found";
    } else if (policy.policyQuery && policy.policyQuery.orgUnit) {
      orgUnitId = getOrgUnitValue(policy.policyQuery.orgUnit); 
    } else { orgUnitId = "SYSTEM"; }
    
    if (policy.setting && policy.setting.hasOwnProperty('value')) {
      settingValue = (typeof policy.setting.value === 'object' && policy.setting.value !== null)
        ? formatObject(policy.setting.value) : String(policy.setting.value);
    } else { settingValue = "No setting value"; }
    
    policyQuery = policy.policyQuery ? JSON.stringify(policy.policyQuery) : "{}";
    type = policy.type || 'UNKNOWN';
    
    if (policyName === "service status") {
      servicesDataArray.push({ service: category, orgUnit: orgUnitId, status: settingValue });
    }
    return { category, policyName, orgUnitId, settingValue, policyQuery, type };
  } catch (e) {
    Logger.log(`Error processing policy ${policy ? policy.name : 'undefined'}: ${e.message}`);
    return null;
  }
}


// ---------------------------------------------
// Part 3: Finalization and Sheet Utilities
// ---------------------------------------------

function finalizeAllSheets(policyMap, additionalServicesData, startTime) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Show toasts for each major step of the finalization process
    ss.toast('Finalizing... Writing policies to sheet.', SCRIPT_NAME, -1); // Use -1 to keep visible until the next toast
    SpreadsheetApp.flush();

    const policySheetName = "Cloud Identity Policies";
    let policySheet = ss.getSheetByName(policySheetName);
    if (!policySheet) policySheet = ss.insertSheet(policySheetName);
    else if (policySheet.getLastRow() > 1) policySheet.getRange(2, 1, policySheet.getLastRow() - 1, policySheet.getMaxColumns()).clearContent();
    
    const policyHeader = ["Category", "Policy Name", "Org Unit ID", "Setting Value", "Policy Query"];
    setupSheetHeader(policySheet, policyHeader); 
    
    const policyRows = Object.values(policyMap).map(p => [p.category, p.policyName, p.orgUnitId, p.settingValue, p.policyQuery]);
    if (policyRows.length > 0) {
      policySheet.getRange(2, 1, policyRows.length, policyRows[0].length).setValues(policyRows);
      applyVlookupToOrgUnitId(policySheet);
      policySheet.hideColumns(5);
    }
    
    finalizeSheet(policySheet, 4);
    
    if (policySheet.getLastRow() >= 1) {
       createOrUpdateNamedRange(policySheet, "Policies", 1, 1, policySheet.getLastRow(), 4);
    }

    ss.toast('Creating Additional Services sheet...', SCRIPT_NAME, 15);
    SpreadsheetApp.flush();
    createAdditionalServicesSheet(additionalServicesData);

    ss.toast('Creating Workspace Security Checklist...', SCRIPT_NAME, 20);
    SpreadsheetApp.flush();
    createWorkspaceSecurityChecklistSheet();

    // Clean up properties and triggers
    deleteTriggers();
    PropertiesService.getScriptProperties().deleteAllProperties();

    // --- MODIFIED SECTION ---
    // Calculate final duration and create the success message
    const endTime = new Date();
    const totalDuration = (endTime.getTime() - startTime) / 1000 / 60;
    const successMessage = `'${SCRIPT_NAME}' completed successfully in ${totalDuration.toFixed(2)} minutes.`;
    
    // Log the final status and end time for debugging purposes
    Logger.log(`============================================================`);
    Logger.log(`✅ FINISH: Script completed at ${endTime.toLocaleString()}`);
    Logger.log(successMessage);
    Logger.log(`============================================================`);

    // Add the completion timestamp to the Workspace Security Checklist sheet.
    try {
      const checklistSheet = ss.getSheetByName("Workspace Security Checklist");
      if (checklistSheet) {
        const timestampMessage = "Last policy inventory completed at: " + endTime.toLocaleString();
        checklistSheet.getRange("F1").setValue(timestampMessage)
                                   .setFontStyle("Montserrat")
                                   .setFontColor("#ffffff")
                                   .setHorizontalAlignment("left");
        Logger.log(`Updated timestamp on 'Workspace Security Checklist' sheet.`);
      }
    } catch (e) {
      // If this fails, it's not critical, so just log it.
      Logger.log(`Warning: Could not set completion timestamp. Error: ${e.message}`);
    }
    // Show a final success toast for 10 seconds instead of a blocking alert
    ss.toast(successMessage, '✅ Complete!', 10);
}

function setupSheetHeader(sheet, header) {
  if (sheet.getFilter()) sheet.getFilter().remove();
  if (sheet.getMaxRows() < 1) sheet.insertRowBefore(1);
  sheet.getRange(1, 1, 1, header.length).clearContent();
  const headerRange = sheet.getRange(1, 1, 1, header.length);
  headerRange.setValues([header]).setFontWeight("bold").setFontColor("#ffffff").setFontFamily("Montserrat").setBackground("#fc3165");
  if (sheet.getMaxRows() >= 1) sheet.setFrozenRows(1);
  if (sheet.getLastRow() > 1) {
    try { sheet.getRange(1, 1, sheet.getLastRow(), header.length).createFilter(); } catch (e) { /* ignore */ }
  }
}

function finalizeSheet(sheet, numColumns) {
  try {
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const maxRows = sheet.getMaxRows();
    if (lastRow > 1 && lastCol > 0) sheet.autoResizeColumns(1, Math.min(lastCol, numColumns));
    if (lastCol > numColumns) sheet.deleteColumns(numColumns + 1, lastCol - numColumns);
    const rowsToDelete = maxRows - Math.max(lastRow, 1);
    if (rowsToDelete > 0) {
      const startDeleteRow = Math.max(lastRow + 1, 2);
      if (startDeleteRow <= maxRows) sheet.deleteRows(startDeleteRow, rowsToDelete);
    }
    if (lastRow > 1) {
      const sortLastCol = Math.min(sheet.getLastColumn(), numColumns);
      if (sortLastCol > 0) sheet.getRange(2, 1, lastRow - 1, sortLastCol).sort({ column: 1, ascending: true });
    }
  } catch (e) {
    Logger.log(`Error finalizing sheet ${sheet.getName()}: ${e.message}`);
  }
}

function createOrUpdateNamedRange(sheet, rangeName, startRow, startColumn, endRow, endColumn) {
   if (!sheet || endRow < startRow || endColumn < startColumn) return;
   const ss = SpreadsheetApp.getActiveSpreadsheet();
   const range = sheet.getRange(startRow, startColumn, (endRow - startRow + 1), (endColumn - startColumn + 1));
   if (ss.getRangeByName(rangeName)) ss.removeNamedRange(rangeName);
   ss.setNamedRange(rangeName, range);
}

function getOrgUnitValue(orgUnit) {
  if (orgUnit === "SYSTEM" || !orgUnit) {
     return "SYSTEM";
  }
  
  const orgUnitID = orgUnit.replace("orgUnits/", "");
  
  // --- FIX FOR ROOT OU ---
  // Get the saved root OU ID from the dependency script's run
  const rootOuId = PropertiesService.getScriptProperties().getProperty('customerRootOuId');
  // If the ID we are looking up is the root ID, just return "/"
  if (rootOuId && orgUnitID === rootOuId.replace('id:', '')) {
      return "/";
  }

  if (!orgUnitID) {
      Logger.log(`Warning: Could not extract valid Org Unit ID from "${orgUnit}"`);
      return orgUnit;
  }
  
  // This formula is for all non-root OUs
  return `=IFERROR(VLOOKUP("${orgUnitID}", OrgID2Path, 3, FALSE), IFERROR(VLOOKUP("${orgUnitID}", Org2ParentPath, 2, FALSE),"/"))`;
}

/**
 * CORRECTED: Formats a JavaScript object into a clean, human-readable string.
 * - For simple objects like {"key": "value"}, it produces "key: value".
 * - For more complex objects, it produces an indented, multi-line string.
 * - Handles nested objects and arrays.
 */
function formatObject(obj, indent = 0) {
  // If the object is not really an object, return it as a string.
  if (obj === null || typeof obj !== 'object') {
    return String(obj);
  }

  // --- NEW: Special handling for 2-Step Verification "Not Enforced" state ---
  // If the policy object is for 2SV and has the epoch timestamp, return "Not Enforced".
  if (obj.hasOwnProperty('enforcedFrom') && obj.enforcedFrom === '1970-01-01T00:00:00Z') {
    return "Not Enforced";
  }

  // Special case for simple {"key": "value"} objects
  const keys = Object.keys(obj);
  if (keys.length === 1 && typeof obj[keys[0]] !== 'object') {
    return `${keys[0]}: ${obj[keys[0]]}`;
  }

  // Handle arrays
  if (Array.isArray(obj)) {
    if (obj.length === 0) return "[]";
    return obj.map(item => formatObject(item, indent + 1)).join(', ');
  }

  // Handle more complex objects with indentation
  let formattedString = "";
  let first = true;
  for (const key of keys) {
    if (obj.hasOwnProperty(key)) {
      if (!first) {
        formattedString += "\n" + "  ".repeat(indent);
      }
      const value = obj[key];
      formattedString += `${key}: `;
      if (typeof value === 'object' && value !== null) {
        // Add a newline and indent before printing the nested object
        formattedString += "\n" + "  ".repeat(indent + 1) + formatObject(value, indent + 1);
      } else {
        formattedString += value;
      }
      first = false;
    }
  }
  return formattedString;
}

function applyVlookupToOrgUnitId(sheet) {
  if (!SpreadsheetApp.getActiveSpreadsheet().getRangeByName('GroupID')) return;
  const lastPolicyRow = sheet.getLastRow();
  if (lastPolicyRow < 2) return;
  const range = sheet.getRange(2, 3, lastPolicyRow - 1, 1);
  const values = range.getValues();
  const formulas = range.getFormulas();
  for (let i = 0; i < values.length; i++) {
    // Only apply formula if the cell is a raw ID (not a formula, not SYSTEM)
    if (values[i][0] && !formulas[i][0] && values[i][0] !== 'SYSTEM') {
      formulas[i][0] = `=IFERROR(VLOOKUP("${values[i][0]}", GroupID, 3, FALSE), "${values[i][0]}")`;
    }
  }
  range.setFormulas(formulas);
}

function applyVlookupToAdditionalServices(sheet, numRows) {
  if (!SpreadsheetApp.getActiveSpreadsheet().getRangeByName('GroupID')) return;
  if (numRows < 1) return;
  const range = sheet.getRange(2, 2, numRows, 1);
  const values = range.getValues();
  const formulas = range.getFormulas();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] && !formulas[i][0] && values[i][0] !== 'SYSTEM') {
      formulas[i][0] = `=IFERROR(VLOOKUP("${values[i][0]}", GroupID, 3, FALSE), "${values[i][0]}")`;
    }
  }
  range.setFormulas(formulas);
}

function createAdditionalServicesSheet(additionalServicesData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "Additional Services";
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  else if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getMaxColumns()).clearContent();
  const header = ["Service", "OU / Group", "Status"];
  setupSheetHeader(sheet, header);
  if (additionalServicesData && additionalServicesData.length > 0) {
    const rows = additionalServicesData.map(data => [data.service, data.orgUnit, data.status]);
    sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
    applyVlookupToAdditionalServices(sheet, rows.length);
  }
  finalizeSheet(sheet, header.length);
}


// ---------------------------------------------
// Part 4: Checklist Creation
// ---------------------------------------------

function createWorkspaceSecurityChecklistSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "Workspace Security Checklist";
  let sheet = ss.getSheetByName(sheetName);

  if (!ss.getRangeByName("Policies")) {
      Logger.log("ERROR: Named range 'Policies' not found. Cannot create checklist.");
      return;
  }
  if (!sheet) {
    if (!copyWorkspaceSecurityChecklistTemplate()) {
        Logger.log(`ERROR: Failed to copy template for ${sheetName}.`);
        return;
    }
    sheet = ss.getSheetByName(sheetName);
  }
  if (!sheet || sheet.getLastRow() < 1) return;
  
  const columnCValues = sheet.getRange(1, 3, sheet.getLastRow(), 1).getValues();

  // Formula map
  const formulaMap = {
    "Require 2-Step Verification for users": [
      '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'two step verification enforcement\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'two step verification enforcement\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'two step verification enforcement\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
      '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'two step verification enforcement\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "(Default) Users allowed to enroll in 2SV and enforcement is off")'
    ],
    "Enforce security keys for admins and high-value accounts": [
      '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'two step verification enforcement factor\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'two step verification enforcement factor\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'two step verification enforcement factor\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
      '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'two step verification enforcement factor\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "(Default) Not enforcing security keys")'
    ],
    "Allow users to enroll in the Advanced Protection Program": [
      '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'advanced protection program\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'advanced protection program\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'advanced protection program\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
      '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10)&REPT("─", 20)&CHAR(10), QUERY(Policies, "select D where B = \'advanced protection program\'", 0)), CHAR(10)&REPT("─", 20)&CHAR(10)&CHAR(10)&REPT("─", 20)&CHAR(10),CHAR(10)&REPT("─", 20)&CHAR(10))), "(Default) enableAdvancedProtectionSelfEnrollment: true"&CHAR(10)&"securityCodeOption: ALLOWED_WITHOUT_REMOTE_ACCESS")'
    ],
     "Use unique passwords": [
      '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'password\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'password\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "OUs with overridden policies:" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'password\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
      '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10)&REPT("─", 20)&CHAR(10), UNIQUE(QUERY(Policies, "select D where B = \'password\'", 0))), CHAR(10)&REPT("─", 20)&CHAR(10)&CHAR(10)&REPT("─", 20)&CHAR(10),CHAR(10)&REPT("─", 20)&CHAR(10) )),"(Default) allowedStrength: STRONG"&CHAR(10)&"minimumLength: 8"&CHAR(10)&"maximumLength: 100"&CHAR(10)&"enforceRequirementsAtLogin: false"&CHAR(10)&"allowReuse: false"&CHAR(10)&"expirationDuration: 0s")'
    ],
    "Turn off Google data download as needed (Takeout)": [
      '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where A = \'takeout\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where A = \'takeout\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where A = \'takeout\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
      '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where A = \'takeout\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "(Default) Users allowed to use Takeout")'
    ],
    "Add user login challenges, add Employee ID as login challenge": [
       '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'login challenges\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'login challenges\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'login challenges\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
       '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'login challenges\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "(Default) enableEmployeeIdChallenge: false")'
    ],
    "Do not allow Super Admins to recover their own accounts": [
       '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'super admin account recovery\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'super admin account recovery\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'super admin account recovery\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
       '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'super admin account recovery\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "(Default) Super Admin account recovery is ON")'
    ],
    "Do not allow users to recover their own accounts": [
       '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'user account recovery\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'user account recovery\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'user account recovery\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
       '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'user account recovery\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "(Default) User account recovery is OFF")'
    ],
    "Configure Google session control to strengthen session expiration": [
      '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'session controls\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'session controls\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'session controls\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
      '=IFERROR(LET(data, FILTER(\'Cloud Identity Policies\'!D:D, \'Cloud Identity Policies\'!B:B = "session controls"), IF(COUNTA(data)=0, "Setting Not Found", TEXTJOIN(CHAR(10), TRUE, MAP(data, LAMBDA(cell, IF(ISBLANK(cell), "", IFERROR(LET(total_seconds, VALUE(REGEXEXTRACT(TO_TEXT(cell), "(\\d+)")), days, INT(total_seconds / 86400), hours, INT(MOD(total_seconds, 86400) / 3600), minutes, INT(MOD(total_seconds, 3600) / 60), IF(total_seconds = 0, "0 minutes", TEXTJOIN(" ", TRUE, IF(days > 0, days & IF(days = 1, " day", " days"), ""), IF(hours > 0, hours & IF(hours = 1, " hour", " hours"), ""), IF(minutes > 0, minutes & IF(minutes = 1, " minute", " minutes"), "")))), cell))))))))'
    ],
     "Share data with Google Cloud services": [
       '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'cloud data sharing\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'cloud data sharing\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'cloud data sharing\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
       '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'cloud data sharing\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "(Default) sharingOptions: ENABLED")'
    ],
     "Set whether users can install Marketplace apps": [
       '=IFERROR(IF(ROWS(QUERY(Policies, "select C where B = \'apps access options\'", 0)) = 0,"/",IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'apps access options\'", 0))) > 1,"OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'apps access options\'", 0)), CHAR(10)&CHAR(10), CHAR(10))),TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'apps access options\'", 0)), CHAR(10)&CHAR(10), CHAR(10))))),"/")',
       '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'apps access options\'", 0)), CHAR(10)&CHAR(10), CHAR(10))),"(Default) Allow users to install and run any app from the Marketplace")'
    ],
      "Review third-party app access to core services": [
      '=IFERROR(IF(ROWS(QUERY(Policies, "select C where B = \'unconfigured third party apps\'", 0)) = 0,"/",IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'unconfigured third party apps\'", 0))) > 1,"OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'unconfigured third party apps\'", 0)), CHAR(10)&CHAR(10), CHAR(10))),TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'unconfigured third party apps\'", 0)), CHAR(10)&CHAR(10), CHAR(10))))),"/")',
      '=IFERROR(TRIM(JOIN(CHAR(10) & REPT("─", 20) & CHAR(10),QUERY(Policies, "select D where B = \'unconfigured third party apps\'", 0))),"(Default) Allow users to access any third-party apps.")'
    ],
    "Block access to less secure apps": [
       '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'less secure apps\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'less secure apps\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'less secure apps\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
       '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'less secure apps\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "(Default) Disable access to less secure apps (Recommended)")'
    ],
    "Control access to Google core services": [
    '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'google services\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'google services\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'google services\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
    '=IFERROR(TRIM(JOIN(CHAR(10) & REPT("─", 20) & CHAR(10), QUERY(Policies, "select D where B = \'google services\'", 0))),"(Default) All Google APIs are Unrestricted")'
  ],
    "Limit external calendar sharing of primary calendars": [
       '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'primary calendar max allowed external sharing\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'primary calendar max allowed external sharing\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'primary calendar max allowed external sharing\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
       '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'primary calendar max allowed external sharing\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "(Default) Only free/busy information (hide event details)")'
    ],
    "Warn users when they invite external guests": [
       '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'external invitations\'", 0))) = 0, "/", IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'external invitations\'", 0))) > 1, "OUs with differences"&CHAR(10)&TEXTJOIN(CHAR(10),TRUE,UNIQUE(QUERY(Policies, "select C where B = \'external invitations\'", 0))), TEXTJOIN(CHAR(10),TRUE,UNIQUE(QUERY(Policies, "select C where B = \'external invitations\'", 0))))), "/")',
       '=IFERROR(IF(ROWS(QUERY(Policies, "select D where B = \'external invitations\'", 0)) = 0, "(Default) warnOnInvite: true", TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'external invitations\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))), "(Default) warnOnInvite: true")'
    ],
    "Limit external calendar sharing of secondary calendars": [
       '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'secondary calendar max allowed external sharing\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'secondary calendar max allowed external sharing\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'secondary calendar max allowed external sharing\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
       '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'secondary calendar max allowed external sharing\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "(Default) Share all information, but outsiders cannot change calendars")'
    ],
    "Set a policy for users to require payment for appointment schedules": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'appointment schedules\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'appointment schedules\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'appointment schedules\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'appointment schedules\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "(Default) enablePayments: false")'
    ],
    "Let people record their meetings": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'video recording\'", 0))) > 1, "OUs with differences" & CHAR(10) & JOIN(CHAR(10), QUERY(Policies, "select C where B = \'video recording\'", 0)), JOIN(CHAR(10), QUERY(Policies, "select C where B = \'video recording\'", 0))),"/")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'video recording\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "(Default) Meet recording is enabled (if on Business Standard or Above)")'
    ],
    "Set a policy for who can join meetings created by your organization": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'safety domain\'", 0))) > 1, "OUs with differences" & CHAR(10) & JOIN(CHAR(10), QUERY(Policies, "select C where B = \'safety domain\'", 0)), JOIN(CHAR(10), QUERY(Policies, "select C where B = \'safety domain\'", 0))),"/")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'safety domain\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "(Default) usersAllowedToJoin: ALL")'
    ],
    "Set a policy for which meetings or calls users in the organization can join": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'safety access\'", 0))) > 1, "OUs with differences" & CHAR(10) & JOIN(CHAR(10), QUERY(Policies, "select C where B = \'safety access\'", 0)), JOIN(CHAR(10), QUERY(Policies, "select C where B = \'safety access\'", 0))),"/")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'safety access\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "(Default) meetingsAllowedToJoin: ALL")'
    ],
    "Set a policy for Default host management": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'safety host management\'", 0))) > 1, "OUs with differences" & CHAR(10) & JOIN(CHAR(10), QUERY(Policies, "select C where B = \'safety host management\'", 0)), JOIN(CHAR(10), QUERY(Policies, "select C where B = \'safety host management\'", 0))),"/")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'safety host management\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "(Default) enableHostManagement: false")'
    ],
    "Warn for external Meet participants": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'safety external participants\'", 0))) > 1, "OUs with differences" & CHAR(10) & JOIN(CHAR(10), QUERY(Policies, "select C where B = \'safety external participants\'", 0)), JOIN(CHAR(10), QUERY(Policies, "select C where B = \'safety external participants\'", 0))),"/")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'safety external participants\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "(Default) enableExternalLabel: true")'
    ],
    "Limit who can chat externally": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'external chat restriction\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'external chat restriction\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'external chat restriction\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10)&REPT("─", 20)&CHAR(10), QUERY(Policies, "select D where B = \'external chat restriction\'", 0)), CHAR(10)&REPT("─", 20)&CHAR(10)&CHAR(10)&REPT("─", 20)&CHAR(10),CHAR(10)&REPT("─", 20)&CHAR(10))), "(Default) Allow users to send messages externally in chats and spaces.")'
    ],
    "Set a chat history policy for spaces": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'space history\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'space history\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'space history\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'space history\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "(Default) historyState: DEFAULT_HISTORY_ON")'
    ],
    "Set a policy for chat history and if users can turn off chat history": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'chat history\'", 0))) = 0, "/", IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'chat history\'", 0))) > 1, "OUs with differences"&CHAR(10)&TEXTJOIN(CHAR(10),TRUE,UNIQUE(QUERY(Policies, "select C where B = \'chat history\'", 0))), TEXTJOIN(CHAR(10),TRUE,UNIQUE(QUERY(Policies, "select C where B = \'chat history\'", 0))))), "/")',
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select D where B = \'chat history\'", 0))) = 0, "historyOnByDefault: true"&CHAR(10)&"allowUserModification: true", IF(ROWS(UNIQUE(QUERY(Policies, "select D where B = \'chat history\'", 0))) > 1, "OUs with differences"&CHAR(10)&TEXTJOIN(CHAR(10),TRUE,UNIQUE(QUERY(Policies, "select D where B = \'chat history\'", 0))), TEXTJOIN(CHAR(10),TRUE,UNIQUE(QUERY(Policies, "select D where B = \'chat history\'", 0))))), "historyOnByDefault: true"&CHAR(10)&"allowUserModification: true")'
    ],
     "Set a policy for chat file sharing": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'chat file sharing\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'chat file sharing\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'chat file sharing\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10)&REPT("─", 20)&CHAR(10), QUERY(Policies, "select D where B = \'chat file sharing\'", 0)), CHAR(10)&REPT("─", 20)&CHAR(10)&CHAR(10)&REPT("─", 20)&CHAR(10),CHAR(10)&REPT("─", 20)&CHAR(10))), "(Default) externalFileSharing: ALL_FILES"&CHAR(10)&"internalFileSharing: ALL_FILES")'
    ],
     "Set a policy for Chat Apps": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'chat apps access\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'chat apps access\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'chat apps access\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10)&REPT("─", 20)&CHAR(10), QUERY(Policies, "select D where B = \'chat apps access\' and C = \'/'\'", 0)), CHAR(10)&REPT("─", 20)&CHAR(10)&CHAR(10)&REPT("─", 20)&CHAR(10),CHAR(10)&REPT("─", 20)&CHAR(10))), "(Default) enableApps: true"&CHAR(10)&"enableWebhooks: true")'
    ],
     "Set sharing options for your domain": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'external sharing\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'external sharing\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "OUs with overridden policies" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'external sharing\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), INDEX(QUERY(Policies, "select D where B = \'external sharing\' and C contains \'/'\'", 0),1)), CHAR(10)&CHAR(10), CHAR(10))), "(Default) External Sharing is ON"&CHAR(10)&"Warn for external sharing is ON"&CHAR(10)&"Visitor Sharing is ON"&CHAR(10)&"Users can publish files to anyone with link")'
    ],
    "Set the default for link sharing": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where LOWER(B) = \'general access default\'", 0))) = 0, "/", IF(ROWS(UNIQUE(QUERY(Policies, "select C where LOWER(B) = \'general access default\'", 0))) > 1, "OUs with general access differences" & CHAR(10) & TEXTJOIN(CHAR(10), TRUE, UNIQUE(QUERY(Policies, "select C where LOWER(B) = \'general access default\'", 0))), TEXTJOIN(CHAR(10), TRUE, UNIQUE(QUERY(Policies, "select C where LOWER(B) = \'general access default\'", 0))))), "/")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'general access default\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "(Default) Private to owner")'
    ],
     "Control content sharing in new shared drives": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'shared drive creation\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'shared drive creation\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "OUs with overridden policies:" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'shared drive creation\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), INDEX(QUERY(Policies, "select D where B = \'shared drive creation\'", 0),1)), CHAR(10)&CHAR(10), CHAR(10))), "(Default) allowSharedDriveCreation: false"&CHAR(10)&"orgUnitForNewSharedDrives: CUSTOM_ORG_UNIT"&CHAR(10)&"customOrgUnit:"&CHAR(10)&"allowManagersToOverrideSettings: true"&CHAR(10)&"allowExternalUserAccess: true"&CHAR(10)&"allowNonMemberAccess: true"&CHAR(10)&"allowedPartiesForDownloadPrintCopy: ALL"&CHAR(10)&"allowContentManagersToShareFolders: true")'
    ],
     "Disable desktop access to Drive": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'drive for desktop\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'drive for desktop\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "OUs with overridden policies:" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'drive for desktop\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10)&REPT("─", 20)&CHAR(10), QUERY(Policies, "select D where B = \'drive for desktop\'", 0)), CHAR(10)&REPT("─", 20)&CHAR(10)&CHAR(10)&REPT("─", 20)&CHAR(10),CHAR(10)&REPT("─", 20)&CHAR(10))), "(Default) allowDriveForDesktop: true"&CHAR(10)&"restrictToAuthorizedDevices: false"&CHAR(10)&"showDownloadLink: true"&CHAR(10)&"allowRealTimePresence: true")'
    ],
     "Set a policy for SDK app access to Google Drive for users": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'drive sdk\'", 0))) = 0, "/", IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'drive sdk\'", 0))) > 1, "OUs with differences" & CHAR(10) & TEXTJOIN(CHAR(10), TRUE, UNIQUE(QUERY(Policies, "select C where B = \'drive sdk\'", 0))), TEXTJOIN(CHAR(10), TRUE, UNIQUE(QUERY(Policies, "select C where B = \'drive sdk\'", 0))))), "/")',
        '=IFERROR(IF(ROWS(QUERY(Policies, "select D where B = \'drive sdk\'", 0)) = 0, "enableDriveSdkApiAccess: true", TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'drive sdk\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))), "(Default) enableDriveSdkApiAccess: true")'
    ],
    "Block or warn on sharing files with sensitive data": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'dlp\'", 0))) > 1, "OUs and Groups with DLP rules" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'dlp\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'dlp\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
        'See Admin Console for details'
    ],
    "Apply the Security update for files": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'file security update\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'file security update\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'file security update\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10)&REPT("─", 20)&CHAR(10), QUERY(Policies, "select D where B = \'file security update\' and C = '/'", 0)), CHAR(10)&REPT("─", 20)&CHAR(10)&CHAR(10)&REPT("─", 20)&CHAR(10),CHAR(10)&REPT("─", 20)&CHAR(10))), "(Default) securityUpdate: APPLY_TO_IMPACTED_FILES"&CHAR(10)&"allowUsersToManageUpdate: true")'
    ],
    "Disable IMAP access": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'imap access\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'imap access\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'imap access\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10)&REPT("─", 20)&CHAR(10), QUERY(Policies, "select D where B = \'imap access\'", 0)), CHAR(10)&REPT("─", 20)&CHAR(10)&CHAR(10)&REPT("─", 20)&CHAR(10),CHAR(10)&REPT("─", 20)&CHAR(10))), "(Default) IMAP access is allowed for all users")'
    ],
    "Disable POP access": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'pop access\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'pop access\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'pop access\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'pop access\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "(Default) POP access is allowed for all users")'
    ],
    "Disable automatic forwarding": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'auto forwarding\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'auto forwarding\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'auto forwarding\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'auto forwarding\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "(Default) Automatic forwarding is allowed")'
    ],
    "Do not allow per-user outbound gateways": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'per user outbound gateway\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'per user outbound gateway\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'per user outbound gateway\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'per user outbound gateway\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "(Default) Per-user outbound gateways are disabled")'
    ],
     "Don\'t bypass spam filters for internal senders": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'spam override lists\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'spam override lists\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'spam override lists\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'spam override lists\' and C = \'/\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "(Default) No spam bypass for internal senders")'
    ],
    "Enable enhanced pre-delivery message scanning": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'enhanced pre delivery message scanning\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'enhanced pre delivery message scanning\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'enhanced pre delivery message scanning\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'enhanced pre delivery message scanning\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "(Default) enableImprovedSuspiciousContentDetection: true")'
    ],
    "Don\'t add IP addresses to your allowlist": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'email spam filter ip allowlist\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'email spam filter ip allowlist\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'email spam filter ip allowlist\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'email spam filter ip allowlist\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "(Default) No IPs in Allowlist")'
    ],
    "Enable additional attachment protection": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'email attachment safety\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'email attachment safety\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'email attachment safety\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'email attachment safety\' and C = \'/\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "(Default) enableEncryptedAttachmentProtection: true"&CHAR(10)&"encryptedAttachmentProtectionConsequence: WARNING"&CHAR(10)&"enableAttachmentWithScriptsProtection: true"&CHAR(10)&"attachmentWithScriptsProtectionConsequence: WARNING"&CHAR(10)&"enableAnomalousAttachmentProtection: false"&CHAR(10)&"anomalousAttachmentProtectionConsequence: WARNING"&CHAR(10)&"allowedAnomalousAttachmentFiletypes:"&CHAR(10)&"[]"&CHAR(10)&"applyFutureRecommendedSettingsAutomatically: true"&CHAR(10)&"encryptedAttachmentProtectionQuarantineId: 0"&CHAR(10)&"attachmentWithScriptsProtectionQuarantineId: 0"&CHAR(10)&"anomalousAttachmentProtectionQuarantineId: 0")'
    ],
    "Enable additional link and external content protection": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'links and external images\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'links and external images\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'links and external images\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'links and external images\' and C = \'/\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "(Default) enableShortenerScanning: true"&CHAR(10)&"enableExternalImageScanning: true"&CHAR(10)&"enableAggressiveWarningsOnUntrustedLinks: true"&CHAR(10)&"applyFutureSettingsAutomatically: true")'
    ],
    "Enable additional spoofing protection": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'spoofing and authentication\'", 0))) > 0, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'spoofing and authentication\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'spoofing and authentication\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'spoofing and authentication\' and C = \'/\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "(Default) detectDomainNameSpoofing: true"&CHAR(10)&"detectEmployeeNameSpoofing: true"&CHAR(10)&"detectDomainSpoofingFromUnauthenticatedSenders: true"&CHAR(10)&"detectUnauthenticatedEmails: false"&CHAR(10)&"domainNameSpoofingConsequence: WARNING"&CHAR(10)&"employeeNameSpoofingConsequence: WARNING"&CHAR(10)&"domainSpoofingConsequence: WARNING"&CHAR(10)&"unauthenticatedEmailConsequence: WARNING"&CHAR(10)&"detectGroupsSpoofing: false"&CHAR(10)&"groupsSpoofingVisibilityType: PRIVATE_GROUPS_ONLY"&CHAR(10)&"groupsSpoofingConsequence: WARNING"&CHAR(10)&"applyFutureSettingsAutomatically: true"&CHAR(10)&"domainNameSpoofingQuarantineId: 0"&CHAR(10)&"employeeNameSpoofingQuarantineId: 0"&CHAR(10)&"domainSpoofingQuarantineId: 0"&CHAR(10)&"unauthenticatedEmailQuarantineId: 0"&CHAR(10)&"groupsSpoofingQuarantineId: 0")'
    ],
    "Scan and block emails with sensitive data": [
        '=IFERROR(IF(ROWS(QUERY(Policies, "select C where LOWER(B) = \'content compliance\'", 0)) = 0, "No Gmail Content Compliance rules found", IF(ROWS(UNIQUE(QUERY(Policies, "select C where LOWER(B) = \'content compliance\'", 0))) > 1, "OUs with content compliance rules" & CHAR(10) & TEXTJOIN(CHAR(10), TRUE, UNIQUE(QUERY(Policies, "select C where LOWER(B) = \'content compliance\'", 0))), TEXTJOIN(CHAR(10), TRUE, UNIQUE(QUERY(Policies, "select C where LOWER(B) = \'content compliance\'", 0))))), "Error querying rules")',
        '"See Admin Console for details"'
    ],
     "Disable Mail delegation, unless required": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'mail delegation\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'mail delegation\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'mail delegation\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'mail delegation\' and C = \'/\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "(Default) Email delegation is OFF")'
    ],
    "Turn on hosted S/MIME for message encryption": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'enhanced smime encryption\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'enhanced smime encryption\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'enhanced smime encryption\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'enhanced smime encryption\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "(Default) allowUserToUploadCertificates: false"&CHAR(10)&"Requires Frontline Plus or Enterprise Plus")'
    ],
    "Limit group creation to admins": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'groups sharing\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'groups sharing\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'groups sharing\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'groups sharing\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "(Default) Anyone in the organization can create groups")'
    ],
    "Block sharing sites outside the domain": [
        '=IFERROR(IF(ROWS(UNIQUE(QUERY(Policies, "select C where B = \'sites creation and modification\'", 0))) > 1, "OUs with differences" & CHAR(10) & TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'sites creation and modification\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select C where B = \'sites creation and modification\'", 0)), CHAR(10)&CHAR(10), CHAR(10)))),"/")',
        '=IFERROR(TRIM(SUBSTITUTE(JOIN(CHAR(10), QUERY(Policies, "select D where B = \'sites creation and modification\'", 0)), CHAR(10)&CHAR(10), CHAR(10))), "(Default) allowSitesCreation: true"&CHAR(10)&"allowSitesModification: true")'
    ]
  };

  Logger.log(`Applying formulas to ${sheetName}...`);
  for (let i = 0; i < columnCValues.length; i++) {
    const rowNumber = i + 1;
    const policyText = columnCValues[i][0] ? String(columnCValues[i][0]).trim() : "";
    if (rowNumber === 1 || !policyText) continue;
    if (formulaMap.hasOwnProperty(policyText)) {
      const formulas = formulaMap[policyText];
      try {
        sheet.getRange(rowNumber, 5).setFormula(formulas[0]);
        if (String(formulas[1]).startsWith('=')) {
           sheet.getRange(rowNumber, 6).setFormula(formulas[1]);
        } else {
           sheet.getRange(rowNumber, 6).setValue(String(formulas[1]).replace(/^['"]|['"]$/g, ''));
        }
      } catch (e) { /* ignore */ }
    }
  }
  SpreadsheetApp.flush();
  Logger.log(`${sheetName} processing complete.`);
}

function copyWorkspaceSecurityChecklistTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templateId = '1rbgKhzDYDmPDKuyx9_qR3CWpTX_ouacEKViuPwAUAf8';
  const targetSheetName = "Workspace Security Checklist";
  try {
    const templateSheet = SpreadsheetApp.openById(templateId).getSheetByName(targetSheetName);
    if(!templateSheet) return false;
    if (ss.getSheetByName(targetSheetName)) ss.deleteSheet(ss.getSheetByName(targetSheetName));
    const newSheet = templateSheet.copyTo(ss).setName(targetSheetName);
    ss.setActiveSheet(newSheet);
    const sheet1 = ss.getSheetByName("Sheet1");
    if(sheet1 && ss.getSheets().length > 1) ss.deleteSheet(sheet1);
    const domain = Session.getActiveUser().getEmail().split('@')[1];
    if (ss.getName() !== `[${domain}] DoiT AdminPulse for Workspace`) {
        ss.rename(`[${domain}] DoiT AdminPulse for Workspace`);
    }
    return true;
  } catch (e) {
    Logger.log(`Error during template copy: ${e.message}`);
    return false;
  }
}