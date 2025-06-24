/**
 * Retrieves mobile device data using the Admin SDK Directory API and populates a Google Sheet.
 *
 * This function fetches all mobile devices for the domain, including health statuses,
 * writes them to a dedicated sheet, and applies formatting and filters.
 */
function getMobileDevices() {
  const functionName = 'getMobileDevices';
  const startTime = new Date();
  Logger.log(`-- Starting ${functionName} at: ${startTime.toLocaleString()}`);

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let mobileDeviceSheet = spreadsheet.getSheetByName("Mobile Report");

    if (mobileDeviceSheet !== null) {
      spreadsheet.deleteSheet(mobileDeviceSheet);
    }
    mobileDeviceSheet = spreadsheet.insertSheet("Mobile Report", spreadsheet.getNumSheets());

    const headers = [
      "Full Name", "Email", "Device Id", "OS", "Model", "Type",
      "Status", "Compromised", "Password Set", "Encrypted", "Last Sync"
    ];
    mobileDeviceSheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight("bold")
      .setFontColor("#ffffff")
      .setBackground("#fc3165")
      .setFontFamily("Montserrat");
    mobileDeviceSheet.setFrozenRows(1);

    // Fetch all device data
    const customerId = "my_customer";
    let rows = [];
    let pageToken = "";

    do {
      const page = AdminDirectory.Mobiledevices.list(customerId, {
        orderBy: "email",
        maxResults: 100,
        pageToken: pageToken,
        projection: "Full",
      });

      if (page && page.mobiledevices) {
        const devices = page.mobiledevices;
        devices.forEach((device) => {
          // Check for the epoch start time (Dec 31, 1969 or Jan 1, 1970) which indicates "Never" synced
          let formattedLastSync;
          if (!device.lastSync || device.lastSync === "1970-01-01T00:00:00.000Z") {
            formattedLastSync = "Never";
          } else {
            formattedLastSync = new Date(device.lastSync).toLocaleString();
          }

          rows.push([
            device.name,
            device.email,
            device.deviceId,
            device.os,
            device.model,
            device.type,
            device.status,
            device.deviceCompromisedStatus,
            device.devicePasswordStatus,
            device.encryptionStatus,
            formattedLastSync,
          ]);
        });
      }
      pageToken = page.nextPageToken;
    } while (pageToken);


    if (rows.length > 0) {
      mobileDeviceSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
      const lastRow = mobileDeviceSheet.getLastRow();

      
      mobileDeviceSheet.autoResizeColumns(1, 3); 
      for (let i = 4; i <= 11; i++) { 
        mobileDeviceSheet.autoResizeColumn(i);
      }

      const rules = [];

      // Rule for Column G (Status)
      const statusRange = mobileDeviceSheet.getRange("G2:G" + lastRow);
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains("Approved").setBackground("#b7e1cd").setRanges([statusRange]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains("Pending").setBackground("#fff2cc").setRanges([statusRange]).build());

      // Rule for Column H (Compromised)
      const compromisedRange = mobileDeviceSheet.getRange("H2:H" + lastRow);
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("COMPROMISED").setBackground("#f4cccc").setRanges([compromisedRange]).build());
      
      // Rule for Column I (Password Set)
      const passwordRange = mobileDeviceSheet.getRange("I2:I" + lastRow);
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("SET").setBackground("#b7e1cd").setRanges([passwordRange]).build());

      // Rule for Column K (Last Sync)
      const lastSyncRange = mobileDeviceSheet.getRange("K2:K" + lastRow);
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Never").setBackground("#f4cccc").setRanges([lastSyncRange]).build());

      mobileDeviceSheet.setConditionalFormatRules(rules); // Apply all rules at once

      // Apply filter to the full data range
      const filterRange = mobileDeviceSheet.getRange(1, 1, lastRow, headers.length);
      if (filterRange.getFilter()) {
        filterRange.getFilter().remove();
      }
      filterRange.createFilter();

    } else {
      mobileDeviceSheet.getRange("A2").setValue("No mobile devices found.");
    }

    if (mobileDeviceSheet.getMaxColumns() > headers.length) {
      mobileDeviceSheet.deleteColumns(headers.length + 1, mobileDeviceSheet.getMaxColumns() - headers.length);
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