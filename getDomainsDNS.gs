/**
 * This script retrieves domain information from Admin Directory API and performs
 * DNS lookups using Google Public DNS. It was initially based on a script
 * that used Cloudflare DNS
 * (https://github.com/cloudflare/cloudflare-docs/blob/production/content/1.1.1.1/other-ways-to-use-1.1.1.1/dns-in-google-sheets.md),
 * but has been significantly modified to remove Cloudflare DNS and integrate
 * with Google Public DNS and Google Admin API.
 */
function getDomainList(customer) {
  const customerDomain = 'my_customer'; // Replace with your customer ID

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName('Domains/DNS');

  if (sheet) {
    spreadsheet.deleteSheet(sheet);
  }

  sheet = spreadsheet.insertSheet('Domains/DNS', spreadsheet.getNumSheets());

  // Set Headers and Notes (do this only once)
  const headers = ["Domains", "Verified", "Primary", "MX", "SPF", "DKIM", "DMARC"];
  sheet.getRange("A1:G1").setValues([headers])
    .setFontColor("#ffffff")
    .setFontSize(10)
    .setFontFamily("Montserrat")
    .setBackground("#fc3165")
    .setFontWeight("bold");
  sheet.getRange("D1").setNote("Mail Exchange");
  sheet.getRange("E1").setNote("Sender Policy Framework");
  sheet.getRange("F1").setNote("DomainKeys Identified Mail, MXDomain = No Google DKIM record");
  sheet.getRange("G1").setNote("Domain-based Message Authentication, Reporting, and Conformance");

  //Get the Domain
  const domainList = getDomainInformation_(customerDomain);

  // Write domain information to spreadsheet in one batch
  sheet.getRange(2, 1, domainList.length, domainList[0].length).setValues(domainList);

  // Get DNS Records
  const dnsResults = getDnsRecords(domainList);

  // Get the last row
  const lastRow = sheet.getLastRow();

  sheet.getRange(2,4,lastRow -1 , 4).setValues(dnsResults);

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
  const filterRange = sheet.getRange('B1:G' + sheet.getLastRow());  // Define the filter range
  filterRange.createFilter();

  // --- Freeze Row 1 ---
  sheet.setFrozenRows(1);
}

function getDomainInformation_(customerDomain) {
  const domainList = [];
  let pageToken = null;

  // Retrieve domain information
  do {
    try {
      const response = AdminDirectory.Domains.list(customerDomain, { pageToken: pageToken });
      pageToken = response.nextPageToken;

      if (!response || !response.domains) {
        console.warn('No domains found or error retrieving domains.');
        break;
      }

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

      if (response && response.domainAliases) {
        response.domainAliases.forEach(function (domainAlias) {
          domainList.push([domainAlias.domainAliasName, domainAlias.verified, 'False']);
        });
      }
    } catch (e) {
      console.error('Error retrieving domain aliases:', e);
      break;
    }
  } while (pageToken);

  return domainList;
}

function getDnsRecords(domainList) {
  const dnsResults = [];
  for (let i = 0; i < domainList.length; i++) {
    const domain = domainList[i][0];

    let mxRecords = performGoogleDNSLookup("MX", domain);
    let spfRecords = performGoogleDNSLookup("TXT", domain);
    let dkimRecords = performGoogleDNSLookup("TXT", "google._domainkey." + domain);
    let dmarcRecords = performGoogleDNSLookup("TXT", "_dmarc." + domain);

    // Add notes if no records are found
    if (!mxRecords) {
      mxRecords = "No MX records found";
    }
    if (!spfRecords) {
      spfRecords = "No SPF records found";
    }
    if (!dkimRecords) {
      dkimRecords = "No DKIM records found";
    }
    if (!dmarcRecords) {
      dmarcRecords = "No DMARC records found";
    }
    dnsResults.push([mxRecords,spfRecords, dkimRecords, dmarcRecords]);
  }
  return dnsResults;
}

/**
 * Performs a DNS lookup using Google Public DNS.
 * @param {string} type The DNS record type (e.g., "MX", "TXT").
 * @param {string} domain The domain to lookup.
 * @return {string} The DNS data, or an empty string if no records found or an error occurred.
 */
function performGoogleDNSLookup(type, domain) {
  const url = `https://dns.google/resolve?name=${encodeURIComponent(domain)}&type=${encodeURIComponent(type)}`;

  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const httpCode = response.getResponseCode();
    const json = response.getContentText();
    const data = JSON.parse(json);

    if (httpCode !== 200) {
      Logger.log(`Google Public DNS HTTP Error ${httpCode} for ${domain} (${type}): ${json}`);
      return ""; // HTTP Error: treated as no record found.  Return an empty string
    }

    if (data.Answer) {
      const outputData = data.Answer.map(record => record.data);
      return outputData.join('\n');
    } else {
      // No Answer Section
      return ""; // Return empty string for no record found
    }
  } catch (error) {
    Logger.log(`Error fetching DNS data from Google Public DNS for ${domain} (${type}): ${error}`);
    return ""; //Treat as a no-record-found. Return an empty string.
  }
}