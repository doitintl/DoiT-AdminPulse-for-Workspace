/**
 * Logs the license assignments for all users in the domain.
 * This function now automatically checks for and runs the getUsersList
 * dependency if the required 'UserStatus' named range is missing.
 */
function getLicenseAssignments() {
  const functionName = 'getLicenseAssignments';
  const startTime = new Date();
  Logger.log(`-- Starting ${functionName} at: ${startTime.toLocaleString()}`);

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // --- DEPENDENCY CHECK ---
    // Check if the 'UserStatus' named range exists. This is created by getUsersList.
    const userStatusRange = spreadsheet.getRangeByName('UserStatus');
    if (!userStatusRange) {
      // If the range doesn't exist, inform the user and run the dependency.
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "Required user data not found. Running user list update first. This may take a moment...",
        "Dependency Update", -1 // A negative duration means the toast stays until dismissed or replaced.
      );
      getUsersList(); // Run the function that creates the named range.
      SpreadsheetApp.getActiveSpreadsheet().toast("User list updated. Continuing with license report.", "Update Complete", 5);
    }
    // --- END DEPENDENCY CHECK ---

    let licenseSheet = spreadsheet.getSheetByName("Licenses");
    if (licenseSheet) {
      spreadsheet.deleteSheet(licenseSheet);
    }
    licenseSheet = spreadsheet.insertSheet("Licenses", spreadsheet.getNumSheets());

    licenseSheet.getRange("A1:C1").setValues([["Email", "License", "Suspended"]])
      .setFontFamily("Montserrat").setBackground("#fc3165").setFontWeight("bold").setFontColor("#ffffff");
    licenseSheet.setFrozenRows(1);

    const productIds = [
      "Google-Apps", "101031", "101034", "101037", "101047", "Google-Vault",
      "101001", "101005", "101033", "101038"
    ];

    const skuNameMapping = {
      "Google-Apps-Unlimited": "G Suite Business", "Google-Apps-For-Business": "G Suite Basic",
      "Google-Apps-Lite": "G Suite Lite", "1010020027": "Google Workspace Business Starter",
      "1010020028": "Google Workspace Business Standard", "1010020025": "Google Workspace Business Plus",
      "1010060003": "Google Workspace Enterprise Essentials", "1010020029": "Google Workspace Enterprise Starter",
      "1010020026": "Google Workspace Enterprise Standard", "1010020020": "Google Workspace Enterprise Plus",
      "1010060001": "Google Workspace Essentials", "1010060005": "Google Workspace Enterprise Essentials Plus",
      "1010020030": "Google Workspace Frontline Starter", "1010020031": "Google Workspace Frontline Standard",
      "1010020034": "Google Workspace Frontline Plus", "Google-Apps-For-Education": "Google Workspace for Education - Fundamentals",
      "1010070001": "Google Workspace for Education Fundamentals", "1010070004": "Google Workspace for Education Gmail Only",
      "1010310005": "Google Workspace for Education Standard", "1010310006": "Google Workspace for Education Standard (Staff)",
      "1010310007": "Google Workspace for Education Standard (Extra Student)", "1010310008": "Google Workspace for Education Plus",
      "1010310009": "Google Workspace for Education Plus (Staff)", "1010310010": "Google Workspace for Education Plus (Extra Student)",
      "1010370001": "Google Workspace for Education: Teaching and Learning Upgrade", "1010310002": "Google Workspace for Education Plus - Legacy",
      "1010310003": "Google Workspace for Education Plus - Legacy (Student)", "1010470003": "Gemini Business",
      "1010470001": "Gemini Enterprise", "1010470004": "Gemini Education", "1010470005": "Gemini Education Premium",
      "1010470006": "AI Security", "1010470007": "AI Meetings and Messaging", "1010380001": "AppSheet Core",
      "1010380002": "AppSheet Enterprise Standard", "1010380003": "AppSheet Enterprise Plus", "Google-Vault": "Google Vault",
      "Google-Vault-Former-Employee": "Google Vault - Former Employee", "1010010001": "Cloud Identity",
      "1010050001": "Cloud Identity Premium", "1010330003": "Google Voice Starter", "1010330004": "Google Voice Standard",
      "1010330002": "Google Voice Premier", "1010340004": "Google Workspace Enterprise Standard - Archived User",
      "1010340001": "Google Workspace Enterprise Plus - Archived User", "1010340005": "Google Workspace Business Starter - Archived User",
      "1010340006": "Google Workspace Business Standard - Archived User", "1010340003": "Google Workspace Business Plus - Archived User",
      "1010340002": "G Suite Business - Archived User", "1010340007": "Google Workspace for Education Fundamentals - Archived User"
    };

    const userEmail = Session.getActiveUser().getEmail();
    const domain = userEmail.split("@").pop();
    const customerId = domain;

    const userLicensesMap = {};

    for (const productId of productIds) {
      let pageToken = null;
      try {
        do {
          const response = AdminLicenseManager.LicenseAssignments.listForProduct(
            productId, customerId, { maxResults: 500, pageToken: pageToken }
          );
          if (response.items) {
            for (const assignment of response.items) {
              userLicensesMap[assignment.userId] = userLicensesMap[assignment.userId] || [];
              const commonName = skuNameMapping[assignment.skuId] || assignment.skuId;
              userLicensesMap[assignment.userId].push(commonName);
            }
          }
          pageToken = response.nextPageToken;
        } while (pageToken);
      } catch (e) {
        if (e.message.indexOf('Not Found') === -1 && e.message.indexOf('invalid') === -1) {
           Logger.log(`Warning: API call for product '${productId}' failed with an unexpected error: ${e.message}`);
        }
      }
    }

    const data = Object.entries(userLicensesMap).map(([userId, licenses]) => [userId, licenses.join(", ")]);

    if (data.length > 0) {
      licenseSheet.getRange(2, 1, data.length, data[0].length).setValues(data);
      const lastRow = licenseSheet.getLastRow();
      
      SpreadsheetApp.flush();
      licenseSheet.getRange(2, 3, lastRow - 1, 1).setFormulaR1C1("=IFERROR(VLOOKUP(RC[-2], UserStatus, 4, FALSE), \"N/A\")");

      const dataRange = licenseSheet.getRange(2, 1, lastRow - 1, licenseSheet.getLastColumn());
      if (licenseSheet.getFilter()) {
        licenseSheet.getFilter().remove();
      }
      dataRange.sort({ column: 2, ascending: true });
      licenseSheet.getDataRange().createFilter();

      const suspendedRange = licenseSheet.getRange("C2:C" + lastRow);
      const trueRule = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("TRUE").setBackground("#FFCDD2").setRanges([suspendedRange]).build();
      const falseRule = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("FALSE").setBackground("#B7E1CD").setRanges([suspendedRange]).build();
      licenseSheet.setConditionalFormatRules([trueRule, falseRule]);
      
      licenseSheet.autoResizeColumns(1, 3);
      if (licenseSheet.getMaxColumns() > 3) {
        licenseSheet.deleteColumns(4, licenseSheet.getMaxColumns() - 3);
      }
    } else {
      licenseSheet.getRange("A2").setValue("No license assignments found.");
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