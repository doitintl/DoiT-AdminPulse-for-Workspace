/**
 * This script lists all mobile devices in a Google Workspace environment.
 * @OnlyCurrentDoc
 **/

function getMobileDevices() {
  const spreadsheet = SpreadsheetApp.getActive();
  const ss = spreadsheet.getSheetByName("Mobile Report");

  // Check if there is data in row 2 and clear the sheet contents accordingly
  const dataRange = ss.getRange(2, 1, 1, ss.getLastColumn());
  const isDataInRow2 = dataRange.getValues().flat().some(Boolean);

  if (isDataInRow2) {
    ss.getRange(2, 1, ss.getLastRow() - 1, ss.getLastColumn()).clearContent();
  }

  const customerId = "my_customer";

  const rows = [];
  let pageToken = "";

  do {
    const page = AdminDirectory.Mobiledevices.list(customerId, {
      orderBy: "OS",
      maxResults: 100,
      pageToken: pageToken,
      projection: "Full",
    });

    const devices = page.mobiledevices;

    if (devices.length > 0) {
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
    } else {
      console.log("No devices found.");
    }

    pageToken = page.nextPageToken || "";
  } while (pageToken);

  if (rows.length > 0) {
    ss.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
}
