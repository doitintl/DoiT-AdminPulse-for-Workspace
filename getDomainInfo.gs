function getDomainInfo() {
  // Create a Customers list query request
  const customerDomain = "my_customer";
  const domainInfo = AdminDirectory.Customers.get(customerDomain);

  // Get the active spreadsheet
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Delete the "General Account Settings" sheet if it exists
  const existingSheet = spreadsheet.getSheetByName('General Account Settings');
  if (existingSheet) {
    spreadsheet.deleteSheet(existingSheet);
  }

  // Create the "General Account Settings" sheet
  const generalSheet = spreadsheet.insertSheet('General Account Settings');

  // Set up the sheet with headers, formatting, and column sizes
  generalSheet.getRange('A1:M1').setValues([['Customer Workspace ID', 'Primary Domain', 'Organization Name', 'Language', 'Customer Contact', 'Address1', 'Address2', 'Postal Code', 'Country Code', 'Region', 'Locality', 'Phone number', 'Alternate Email']]);
  generalSheet.getRange('A1:M1').setFontWeight('bold').setBackground('#fc3165').setFontColor('#ffffff');
  generalSheet.autoResizeColumns(1, 13);
  generalSheet.setColumnWidth(1, 150);
  generalSheet.setColumnWidth(2, 186);
  generalSheet.setColumnWidth(6, 150);
  generalSheet.setColumnWidth(7, 186);
  generalSheet.getRange('2:2').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  // Append a new row with customer data
  generalSheet.appendRow([
    domainInfo.id,
    domainInfo.customerDomain,
    domainInfo.postalAddress.organizationName,
    domainInfo.language,
    domainInfo.postalAddress.contactName,
    domainInfo.postalAddress.addressLine1,
    domainInfo.postalAddress.addressLine2,
    domainInfo.postalAddress.postalCode,
    domainInfo.postalAddress.countryCode,
    domainInfo.postalAddress.region,
    domainInfo.postalAddress.locality,
    domainInfo.phoneNumber,
    domainInfo.alternateEmail,
  ]);

  // Delete cells N-Z and rows 3-1000
  generalSheet.deleteColumns(14, 13);
  generalSheet.deleteRows(3, 998);  
  generalSheet.autoResizeColumns(1, generalSheet.getLastColumn());
}
