/**
 * Script will query busniess address associated with Google Workspace, as well as language, and contact infomation.
 * @OnlyCurrentDoc
 */

function getDomainInfo() {
  var customerDomain = 'my_customer';

  // Create a Customers list query request
  var domainInfo = AdminDirectory.Customers.get(customerDomain);

  // Get the active spreadsheet and the "Workspace Account Settings" sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("General Account Settings");

  // Clear the existing data in the target range (A2:M2)
  sheet.getRange('A2:M2').clearContent();

  // Append a new row with customer data
  sheet.appendRow([
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
}
