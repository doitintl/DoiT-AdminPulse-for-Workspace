/**This script leverages CloudFlare DNS in Google Sheets. https://github.com/cloudflare/cloudflare-docs/blob/production/content/1.1.1.1/other-ways-to-use-1.1.1.1/dns-in-google-sheets.md 

The MIT License (MIT)

Copyright (c) 2021 Cloudflare, Inc. 

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software *without restriction, 
including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, 
and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT  NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR *PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,  WHETHER IN AN ACTION OF CONTRACT, TORT *OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
**/

function getDomainList(customer) {
  // Retrieve domain information from Admin Directory API
  const domainList = [];
  let pageToken = null;
  const customerDomain = 'my_customer';
  
const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
let sheet = spreadsheet.getSheetByName('Domains/DNS');

// Check if "Domains/DNS" sheet exists, delete it if it does
if (sheet) {
  spreadsheet.deleteSheet(sheet);
}

// Get the index for the new sheet (as the last sheet in the workbook)
const newSheetIndex = spreadsheet.getNumSheets();

// Create a new sheet at the last index
sheet = spreadsheet.insertSheet('Domains/DNS', newSheetIndex);

// Set "google.com" in cell A2 as a placeholder
sheet.getRange('A2').setValue('google.com');
  
  // Set the headers
  const headers = ["Domains", "Verified", "Primary", "MX", "SPF", "DKIM", "DMARC"];
  const headerRange = sheet.getRange("A1:G1");
  headerRange.setValues([headers]);
  headerRange.setFontColor("#ffffff");
  headerRange.setFontSize(10);
  headerRange.setFontFamily("Montserrat");
  headerRange.setBackground("#fc3165");
  headerRange.setFontWeight("bold");

  // Set notes for cells D1 - F1
  sheet.getRange("D1").setNote("Mail Exchange");
  sheet.getRange("E1").setNote("Sender Policy Framework (SPF)");
  sheet.getRange("F1").setNote("DomainKeys Identified Mail, MXDomain = No Google DKIM record");
  sheet.getRange("G1").setNote("Domain-based Message Authentication, Reporting, and Conformance (DMARC)");

  // Retrieve domain information
  do {
    try {
      const response = AdminDirectory.Domains.list(customerDomain, { pageToken: pageToken });
      pageToken = response.nextPageToken;

      response.domains.forEach(function (domain) {
        domainList.push([domain.domainName, domain.verified, domain.isPrimary]);
      });
    } catch (e) {
      console.error('Error retrieving domains:', e);
      break;
    }
  } while (pageToken);

  // Retrieve domain alias information
  pageToken = null;

  do {
    try {
      const response = AdminDirectory.DomainAliases.list(customerDomain, { pageToken: pageToken });
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
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const formulasD = '=IFERROR(NSLookup($D$1,A2), empty)';
    const formulasE = '=IFERROR(NSLookup("txt", A2), empty)';
    const formulasF = '=IFERROR(NSLookup("TXT","google._domainkey."&A2), empty)';
    const formulasG = '=IFERROR(NSLookup("TXT","_dmarc."&A2), empty)';

    sheet.getRange(2, 4, lastRow - 1, 1).setFormula(formulasD);
    sheet.getRange(2, 5, lastRow - 1, 1).setFormula(formulasE);
    sheet.getRange(2, 6, lastRow - 1, 1).setFormula(formulasF);
    sheet.getRange(2, 7, lastRow - 1, 1).setFormula(formulasG);
  }

  // Delete columns H-Z
  sheet.deleteColumns(8, 19);

  // Delete empty rows
  const dataRange = sheet.getDataRange();
  const allData = dataRange.getValues();
  for (let i = allData.length - 1; i >= 0; i--) {
    if (allData[i].every(value => value === '')) {
      sheet.deleteRow(i + 1);
    }
  }

  // Set column sizes
  sheet.autoResizeColumn(1);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 150);
  sheet.setColumnWidth(6, 150);
  sheet.setColumnWidth(7, 150);

  // Apply conditional formatting rules
  const rangeD = sheet.getRange("D2:D" + lastRow);
  const ruleD = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("google")
    .setBackground("#b7e1cd")
    .setRanges([rangeD])
    .build();

  const rangeE = sheet.getRange("E2:E" + lastRow);
  const ruleE = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("_spf.google.com")
    .setBackground("#b7e1cd")
    .setRanges([rangeE])
    .build();

  const rangeF = sheet.getRange("F2:F" + lastRow);
  const ruleF = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("v=dkim1;")
    .setBackground("#b7e1cd")
    .setRanges([rangeF])
    .build();

  const rangeG = sheet.getRange("G2:G" + lastRow);
  const ruleG = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("v=dmarc")
    .setBackground("#b7e1cd")
    .setRanges([rangeG])
    .build();

  const ruleDRed = SpreadsheetApp.newConditionalFormatRule()
    .whenTextDoesNotContain("google")
    .setBackground("#ffb6c1")
    .setRanges([rangeD])
    .build();

  const ruleERed = SpreadsheetApp.newConditionalFormatRule()
    .whenTextDoesNotContain("_spf.google.com")
    .setBackground("#ffb6c1")
    .setRanges([rangeE])
    .build();

  const ruleFRed = SpreadsheetApp.newConditionalFormatRule()
    .whenTextDoesNotContain("v=dkim1;")
    .setBackground("#ffb6c1")
    .setRanges([rangeF])
    .build();

  const ruleGRed = SpreadsheetApp.newConditionalFormatRule()
    .whenTextDoesNotContain("v=dmarc")
    .setBackground("#ffb6c1")
    .setRanges([rangeG])
    .build();

  const rules = [ruleD, ruleE, ruleF, ruleG, ruleDRed, ruleERed, ruleFRed, ruleGRed];
  sheet.setConditionalFormatRules(rules);

  // --- Add Persistent Toast Notification ---
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "If you use a third-party mail gateway or a SPF flattener, records may be highlighted red and should be manually inspected.", 
    "Instructions", 
    -1 // Persistent toast
  );
// --- Add Filter View ---
const filterRange = sheet.getRange('B1:G' + lastRow);  // Define the filter range
filterRange.createFilter();

  // --- Freeze Row 1 ---
  sheet.setFrozenRows(1); 
}

function NSLookup(type, domain) { //Function takes DNS record type and domain as input

  if (typeof type == 'undefined') { //Validation for record type and domain
    throw new Error('Missing parameter 1 dns type');
  }

  if (typeof domain == 'undefined') {
    throw new Error('Missing parameter 2 domain name');
  }

  type = type.toUpperCase(); //Convert record type to uppercase

  const url = 'https://cloudflare-dns.com/dns-query?name=' + encodeURIComponent(domain) + '&type=' + encodeURIComponent(type); //Concatenate URL query 

  const options = {
    muteHttpExceptions: true,
    headers: {
      accept: "application/dns-json"
    }
  };

  const result = UrlFetchApp.fetch(url, options);
  const rc = result.getResponseCode();
  const resultText = result.getContentText();

  if (rc !== 200) {
    throw new Error(rc);
  }

  const errors = [
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

  const response = JSON.parse(resultText);

  if (response.Status !== 0) {
    return errors[response.Status].name;
  }

  const outputData = [];

  for (const i in response.Answer) {
    outputData.push(response.Answer[i].data);
  }

  const outputString = outputData.join('\n');

  return outputString;
}