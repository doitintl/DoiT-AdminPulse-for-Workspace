/** 
 * This script lists all mobile devices in a Google Workspace environment. 
 * @OnlyCurrentDoc
 **/

function getMobileDevices() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = spreadsheet.getSheetByName('Mobile Report');

  // Check if there is data in row 2 and clear the sheet contents accordingly
  var dataRange = ss.getRange(2, 1, 1, ss.getLastColumn());
  var isDataInRow2 = dataRange.getValues().flat().some(Boolean);

  if (isDataInRow2) {
    ss.getRange(2, 1, ss.getLastRow() - 1, ss.getLastColumn()).clearContent();
  }

  var customerId = 'my_customer';

  var rows = [];
  var pageToken;
  
  do {
    var page = AdminDirectory.Mobiledevices.list(customerId, {
      orderBy: 'OS',
      maxResults: 100,
      pageToken: pageToken,
      projection:'Full'
    });
    
    var devices = page.mobiledevices;

    if (devices.length > 0) {
      for (var i = 0; i < devices.length; i++) {
        var device = devices[i];
        rows.push([device.name, device.email, device.deviceId, device.model, device.type, device.status, device.lastSync]);
      }
    } else {
      Logger.log('No devices found.');
    }

    pageToken = page.nextPageToken;
  } while (pageToken);
  
  if (rows.length > 0) {
    ss.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
}
