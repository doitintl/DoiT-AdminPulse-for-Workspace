/**  Created with script from https://developers.cloudflare.com/1.1.1.1/other-ways-to-use-1.1.1.1/dns-in-google-sheets/ 
*
*/
function getDomainList(customer) {
  // Retrieve domain information from Admin Directory API
  var domainList = [];
  var pageToken = null;
  var customer = 'my_customer';
  
  // Clear sheet contents starting from row 2
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Domains/DNS') || ss.insertSheet('Domains/DNS', 5);
  sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();

  do {
    try {
      var response = AdminDirectory.Domains.list(customer, { pageToken: pageToken });
      pageToken = response.nextPageToken;

      response.domains.forEach(function (domain) {
        domainList.push([domain.domainName, domain.verified, domain.isPrimary]);
      });
    } catch (e) {
      console.error('Error retrieving domains:', e);
      break;
    }
  } while (pageToken);

  // Retrieve domain alias information from Admin Directory API
  pageToken = null;

  do {
    try {
      var response = AdminDirectory.DomainAliases.list(customer, { pageToken: pageToken });
      pageToken = response.nextPageToken;

      if (response.domainAliases) {
        response.domainAliases.forEach(function (domainAlias) {
          domainList.push([domainAlias.domainAliasName, domainAlias.verified, 'False']);
        });
      }
    } catch (e) {
      console.error('Error retrieving domain aliases:', e);
      break;
    }
  } while (pageToken);

  // Write domain information to spreadsheet
  sheet.getRange(2, 1, domainList.length, domainList[0].length).setValues(domainList);

  // Add formulas to columns D, E, F, and G starting from row 2
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var formulasD = '=IFERROR(NSLookup($D$1,A3), empty)';
    var formulasE = '=IFERROR(NSLookup("txt", A3), empty)';
    var formulasF = '=IFERROR(NSLookup("TXT","google._domainkey."&A3), empty)';
    var formulasG = '=IFERROR(NSLookup("TXT","_dmarc."&A3), empty)';

    sheet.getRange(2, 4, lastRow - 1, 1).setFormula(formulasD);
    sheet.getRange(2, 5, lastRow - 1, 1).setFormula(formulasE);
    sheet.getRange(2, 6, lastRow - 1, 1).setFormula(formulasF);
    sheet.getRange(2, 7, lastRow - 1, 1).setFormula(formulasG);
  }
}


function NSLookup(type, domain) { //Function takes DNS record type and domain as input

  if (typeof type == 'undefined') { //Validation for record type and domain
    throw new Error('Missing parameter 1 dns type');
  }

  if (typeof domain == 'undefined') {
    throw new Error('Missing parameter 2 domain name');
  }

  type = type.toUpperCase(); //Convert record type to uppercase

  var url = 'https://cloudflare-dns.com/dns-query?name=' + encodeURIComponent(domain) + '&type=' + encodeURIComponent(type); //Concatenate URL query 

  var options = {
    muteHttpExceptions: true,
    headers: {
      accept: "application/dns-json"
    }
  };

  var result = UrlFetchApp.fetch(url, options);
  var rc = result.getResponseCode();
  var resultText = result.getContentText();

  if (rc !== 200) {
    throw new Error(rc);
  }

  var errors = [
    { name: "NoError", description: "No Error"}, // 0
    { name: "FormErr", description: "Format Error"}, // 1
    { name: "ServFail", description: "Server Failure"}, // 2
    { name: "NXDomain", description: "Non-Existent Domain"}, // 3
    { name: "NotImp", description: "Not Implemented"}, // 4
    { name: "Refused", description: "Query Refused"}, // 5
    { name: "YXDomain", description: "Name Exists when it should not"}, // 6
    { name: "YXRRSet", description: "RR Set Exists when it should not"}, // 7
    { name: "NXRRSet", description: "RR Set that should exist does not"}, // 8
    { name: "NotAuth", description: "Not Authorized"} // 9
  ];

  var response = JSON.parse(resultText);

  if (response.Status !== 0) {
    return errors[response.Status].name;
  }

  var outputData = [];

  for (var i in response.Answer) {
    outputData.push(response.Answer[i].data);
  }

  var outputString = outputData.join('\n');

  return outputString;
}