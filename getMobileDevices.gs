/**
 * This script lists all mobile devices in a Google Workspace environment.
 * 
 **/

function getMobileDevices() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const lastSheetIndex = sheets.length;

  // Check if "Mobile Report" sheet exists, delete it if it does, and create it at the last index
  let mobileDeviceSheet = spreadsheet.getSheetByName("Mobile Report");
  if (mobileDeviceSheet !== null) {
    spreadsheet.deleteSheet(mobileDeviceSheet);
  }
  mobileDeviceSheet = spreadsheet.insertSheet("Mobile Report", lastSheetIndex);
  // Add headers
  const headers = ["Full Name", "Email", "Device Id", "Model", "Type", "Status", "Last Sync"];
  mobileDeviceSheet.appendRow(headers);

  // Apply formatting to header row
  const headerRange = mobileDeviceSheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight("bold").setFontColor("#ffffff").setBackground("#fc3165");

  // Freeze the header row
  mobileDeviceSheet.setFrozenRows(1);

  // Apply conditional formatting
  const range = mobileDeviceSheet.getRange("F2:F1000");
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("Approved")
    .setBackground("#b7e1cd")
    .setRanges([range])
    .build();
  const rules = mobileDeviceSheet.getConditionalFormatRules();
  rules.push(rule);
  mobileDeviceSheet.setConditionalFormatRules(rules);

  const lastRow = mobileDeviceSheet.getLastRow();
  const filterRange = mobileDeviceSheet.getRange('B1:G' + lastRow);  // Filter columns B through G (including header)
  filterRange.createFilter();

  const customerId = "my_customer";

  let rows = [];
  let pageToken = "";

  do {
    const page = AdminDirectory.Mobiledevices.list(customerId, {
      orderBy: "OS",
      maxResults: 100,
      pageToken: pageToken,
      projection: "Full",
    });

    // Check if the 'mobiledevices' property exists in the response
    if (page && page.mobiledevices) { 
      const devices = page.mobiledevices;

      devices.forEach((device) => {
        rows.push([
          device.name,
          device.email,
          device.deviceId,
          device.model,
          device.type,
          device.status,
          device.lastSync,
        ]);
      });
    }

    pageToken = page.nextPageToken;

  } while (pageToken);

  // Auto resize columns based on content
  if (rows.length > 0) {
    mobileDeviceSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    mobileDeviceSheet.autoResizeColumns(1, rows[0].length);
  }

  // Delete unnecessary columns
  const columnsToDelete = 26 - 7; // H to Z is 26 columns in total, subtracting 7 columns (A to G)
  if (columnsToDelete > 0) {
    mobileDeviceSheet.deleteColumns(8, columnsToDelete);
  }
}