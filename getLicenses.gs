/**
 * Logs the license assignments, including the product ID and the SKU ID, for
 * the users in the domain. Notice the use of page tokens to access the full
 * list of results.
 * 
 */
function getLicenseAssignments() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const lastSheetIndex = sheets.length;

  // Check if "Licenses" sheet exists, delete it if it does
  let licenseSheet = spreadsheet.getSheetByName("Licenses");
  if (licenseSheet) {
    spreadsheet.deleteSheet(licenseSheet);
  }

  // Create the "Licenses" sheet at the last index
  licenseSheet = spreadsheet.insertSheet("Licenses", lastSheetIndex);

  // Add headers with Montserrat font, fill color, and freeze header row
  const headerRange = licenseSheet.getRange("A1:C1");
  headerRange.setFontFamily("Montserrat").setBackground("#fc3165").setFontWeight("bold").setFontColor("#ffffff").setValues([["Email", "License", "Suspended"]]);
  licenseSheet.setFrozenRows(1);

  // Delete columns D-Z
  licenseSheet.deleteColumns(4, 23);

  const productIds = [
    "Google-Apps",
    "101031",
    "101037",
    "101034",
    "101047",
    "101038",
    "Google-Vault",
    "101001",
    "101005",
    "101033",
  ];

  var skuNameMapping = {
    "Google-Apps-Unlimited": "G Suite Business",
    "Google-Apps-For-Business": "G Suite Basic",
    "Google-Apps-Lite": "G Suite Lite",
    "Google-Apps-For-Postini": "Google Apps Message Security",
    1010340004: "Google Workspace Enterprise Standard - Archived User",
    1010340001: "Google Workspace Enterprise Plus - Archived User",
    1010340005: "Google Workspace Business Starter - Archived User",
    1010340006: "Google Workspace Business Standard - Archived User",
    1010340003: "Google Workspace Business Plus - Archived User",
    1010340002: "G Suite Business - Archived User",
    1010470003: "Gemini Business",
    1010470001: "Gemini Enterprise",
    1010470006: "AI Security",
    1010470007: "AI Meetings and Messaging",
    1010380001: "AppSheet Core",
    1010380002: "AppSheet Enterprise Standard",
    1010380003: "AppSheet Enterprise Plus",
    "Google-Vault": "Google Vault",
    "Google-Vault-Former-Employee": "Google Vault - Former Employee",
    1010010001: "Cloud Identity Free",
    1010050001: "Cloud Identity Premium",
    1010020027: "Google Workspace Business Starter",
    1010020028: "Google Workspace Business Standard",
    1010020025: "Google Workspace Business Plus",
    1010060003: "Google Workspace Enterprise Essentials",
    1010020029: "Google Workspace Enterprise Starter",
    1010020026: "Google Workspace Enterprise Standard",
    1010020020: "Google Workspace Enterprise Plus",
    1010060001: "Google Workspace Essentials",
    1010060005: "Google Workspace Enterprise Essentials Plus",
    1010020030: "Google Workspace Frontline Starter",
    1010020031: "Google Workspace Frontline Standard",
    1010330003: "Google Voice Starter",
    1010330004: "Google Voice Standard",
    1010330002: "Google Voice Premier",
    "Google-Apps-For-Education": "Google Workspace for Education Fundamentals",
    1010310005: "Google Workspace for Education Standard",
    1010310006: "Google Workspace for Education Standard (Staff)",
    1010310007: "Google Workspace for Education Standard (Extra Student)",
    1010310008: "Google Workspace for Education Plus",
    1010310009: "Google Workspace for Education Plus (Staff)",
    1010310010: "Google Workspace for Education Plus (Extra Student)",
    1010370001: "Google Workspace for Education: Teaching and Learning Upgrade",
    1010310002: "Google Workspace for Education Plus - Legacy",
    1010310003: "Google Workspace for Education Plus - Legacy (Student)",
  };

  const userEmail = Session.getActiveUser().getEmail();
  const domain = userEmail.split("@").pop();
  const customerId = domain;

  // Accumulate licenses for each user
  const userLicensesMap = {};

  for (const productId of productIds) {
    const assignments = getAllLicenseAssignments(productId, customerId);

    // Accumulate licenses for each user
    for (const assignment of assignments) {
      userLicensesMap[assignment.userId] =
        userLicensesMap[assignment.userId] || [];

      // Translate SKU to common name using skuNameMapping
      const commonName = skuNameMapping[assignment.skuId] || assignment.skuId;
      userLicensesMap[assignment.userId].push(commonName);
    }
  }

  const data = Object.entries(userLicensesMap).map(([userId, licenses]) => [
    userId,
    licenses.join(", "),
  ]);
  writeDataToSheet(licenseSheet, data);

  const lastRow = licenseSheet.getLastRow();
  applyVLookupFormula(licenseSheet, lastRow);
  sortSheetByColumn(licenseSheet, 2, true);

  // Auto resize columns
  licenseSheet.autoResizeColumns(1, 3);

  // Add conditional formatting rules
  const conditionalFormatRules = licenseSheet.getConditionalFormatRules();
  const trueRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("TRUE")
    .setBackground("#FFCDD2")
    .setRanges([licenseSheet.getRange("C2:C")])
    .build();

  const falseRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("FALSE")
    .setBackground("#B7E1CD")
    .setRanges([licenseSheet.getRange("C2:C")])
    .build();

  conditionalFormatRules.push(trueRule);
  conditionalFormatRules.push(falseRule);

  licenseSheet.setConditionalFormatRules(conditionalFormatRules);
}

// Helper function to get all license assignments for a product
function getAllLicenseAssignments(productId, customerId) {
  let assignments = [];
  let pageToken = null;

  do {
    const response = AdminLicenseManager.LicenseAssignments.listForProduct(
      productId,
      customerId,
      {
        maxResults: 500,
        pageToken: pageToken,
      },
    );

    assignments = assignments.concat(response.items);
    pageToken = response.nextPageToken;
  } while (pageToken);

  return assignments;
}

// Helper function to write data to the sheet
function writeDataToSheet(sheet, data) {
  if (data.length > 0) {
    const range = sheet.getRange(2, 1, data.length, data[0].length);
    range.setValues(data);
  }
}

// Helper function to apply VLOOKUP formula to a range
function applyVLookupFormula(sheet, lastRow) {
  const formulaRange = sheet.getRange(2, 3, lastRow - 1, 1);
  formulaRange.setFormula(`=VLOOKUP(A2:A${lastRow}, UserStatus, 4, FALSE)`);
}

// Helper function to sort sheet by a specific column
function sortSheetByColumn(sheet, column, ascending) {
  const range = sheet.getDataRange();
  const numRows = range.getNumRows();
  const sortRange = sheet.getRange(2, 1, numRows - 1, sheet.getLastColumn());

  if (sheet.getFilter()) {
    const filter = sheet.getFilter();
    filter.remove();
  }

  sortRange.sort({ column: column, ascending: ascending });
}
