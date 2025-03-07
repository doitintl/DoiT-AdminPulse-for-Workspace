/**
 * This script retrieves domain information from Admin Directory API and performs
 * DNS lookups using Google Public DNS. It was initially based on a script
 * that used Cloudflare DNS
 * (https://github.com/cloudflare/cloudflare-docs/blob/production/content/1.1.1.1/other-ways-to-use-1.1.1.1/dns-in-google-sheets.md),
 * but has been significantly modified to remove Cloudflare DNS and integrate
 * with Google Public DNS and Google Admin API.
 * */

function getDomainList() {
  const customerDomain = 'my_customer'; // Replace with your customer ID

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName('Domains/DNS');

  try {
    if (sheet) {
      spreadsheet.deleteSheet(sheet);
    }

    sheet = spreadsheet.insertSheet('Domains/DNS', spreadsheet.getNumSheets());

    const headers = ["Domains", "Verified", "Primary", "MX", "SPF", "DKIM", "DMARC", "DNS Status"];
    sheet.getRange("A1:H1").setValues([headers])
      .setFontColor("#ffffff")
      .setFontSize(10)
      .setFontFamily("Montserrat")
      .setBackground("#fc3165")
      .setFontWeight("bold");
    sheet.getRange("D1").setNote("Mail Exchange, Red cells indicate Google MX records were not found.");
    sheet.getRange("E1").setNote("Sender Policy Framework, Red cells indicate Google SPF record was not found.");
    sheet.getRange("F1").setNote("DomainKeys Identified Mail, Red cells indicate the default DKIM selector for Google was not found.");
    sheet.getRange("G1").setNote("Domain-based Message Authentication, Reporting, and Conformance, Red cells indicate no DMARC records found.");
    sheet.getRange("H1").setNote("Status of DNS Lookups");

    // Get the Domain
    const domainList = getDomainInformation_(customerDomain);

    // Write domain information to spreadsheet in one batch
    sheet.getRange(2, 1, domainList.length, domainList[0].length).setValues(domainList);

    // Get DNS Records
    const dnsResultsWithStatus = getDnsRecords(domainList); // Modified to return DNS results AND status

    // Get the last row
    const lastRow = sheet.getLastRow();

    //Extract DNS results
    const dnsResults = dnsResultsWithStatus.map(item => [item.mxRecords.data, item.spfRecords.data, item.dkimRecords.data, item.dmarcRecords.data]);
    sheet.getRange(2, 4, lastRow - 1, 4).setValues(dnsResults);

    //Write the status message to the sheet
    const statusMessages = dnsResultsWithStatus.map(item => {
      let overallStatus = "";
      if (item.mxRecords.status !== "Lookup Complete" ||
        item.spfRecords.status !== "Lookup Complete" ||
        item.dkimRecords.status !== "Lookup Complete" ||
        item.dmarcRecords.status !== "Lookup Complete") {
        //Something is not complete.
        let issues = [];
        if (item.mxRecords.status !== "Lookup Complete") {
          issues.push(`MX: ${item.mxRecords.status}`);
        }
        if (item.spfRecords.status !== "Lookup Complete") {
          issues.push(`SPF: ${item.spfRecords.status}`);
        }
        if (item.dkimRecords.status !== "Lookup Complete") {
          issues.push(`DKIM: ${item.dkimRecords.status}`);
        }
        if (item.dmarcRecords.status !== "Lookup Complete") {
          issues.push(`DMARC: ${item.dmarcRecords.status}`);
        }
        overallStatus = "Issues Found:\n" + issues.join("\n"); // Added line breaks

      } else {
        overallStatus = "Lookup Complete";
      }
      return [overallStatus];
    });
    sheet.getRange(2, 8, lastRow - 1, 1).setValues(statusMessages); // Write status messages to column H


    // Delete columns I-Z
    sheet.deleteColumns(9, 18);

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
    sheet.setColumnWidth(7, 164); // Adjusted for Status Message
    sheet.setColumnWidth(8, 300); // Adjusted for Status Message

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

    const rangeDRed = sheet.getRange("D2:D" + lastRow);
    const ruleDRed = SpreadsheetApp.newConditionalFormatRule()
      .whenTextDoesNotContain("google")
      .setBackground("#ffb6c1")
      .setRanges([rangeD])
      .build();

    const rangeERed = sheet.getRange("E2:E" + lastRow);
    const ruleERed = SpreadsheetApp.newConditionalFormatRule()
      .whenTextDoesNotContain("_spf.google.com")
      .setBackground("#ffb6c1")
      .setRanges([rangeE])
      .build();

    const rangeFRed = sheet.getRange("F2:F" + lastRow);
    const ruleFRed = SpreadsheetApp.newConditionalFormatRule()
      .whenTextDoesNotContain("v=dkim1;")
      .setBackground("#ffb6c1")
      .setRanges([rangeF])
      .build();

    const rangeGRed = sheet.getRange("G2:G" + lastRow);
    const ruleGRed = SpreadsheetApp.newConditionalFormatRule()
      .whenTextDoesNotContain("v=dmarc")
      .setBackground("#ffb6c1")
      .setRanges([rangeG])
      .build();

    const rules = [ruleD, ruleE, ruleF, ruleG, ruleDRed, ruleERed, ruleFRed, ruleGRed];
    sheet.setConditionalFormatRules(rules);

    // --- Add Filter View ---
    const filterRange = sheet.getRange('B1:H' + sheet.getLastRow());
    filterRange.createFilter();

    // --- Freeze Row 1 ---
    sheet.setFrozenRows(1);

    SpreadsheetApp.getActiveSpreadsheet().toast('Domain and DNS check completed successfully!', 'Success', 3);

  } catch (e) {
    // Display error message to the user
    SpreadsheetApp.getActiveSpreadsheet().toast(`An error occurred: ${e.message}`, 'Error', 5);
    Logger.log(e);
  }
}

function getDomainInformation_(customerDomain) {
  const domainList = [];
  let pageToken = null;
  const maxRetries = 3;

  // Retrieve domain information
  do {
    let retryCount = 0;
    let success = false;
    while (!success && retryCount < maxRetries) {
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
        success = true;
      } catch (e) {
        console.error(`Error retrieving domains (retry ${retryCount + 1}/${maxRetries}):`, e);
        retryCount++;
        Utilities.sleep(1000 * Math.pow(2, retryCount));
      }
    }
    if (!success) {
      console.error("Failed to retrieve domains after multiple retries.");
      break;
    }
  } while (pageToken);

  // Retrieve domain alias information
  pageToken = null;
  do {
    let retryCount = 0;
    let success = false;
    while (!success && retryCount < maxRetries) {
      try {
        const response = AdminDirectory.DomainAliases.list(customerDomain, { pageToken: pageToken });
        pageToken = response.nextPageToken;

        if (response && response.domainAliases) {
          response.domainAliases.forEach(function (domainAlias) {
            domainList.push([domainAlias.domainAliasName, domainAlias.verified, 'False']);
          });
        }
        success = true;
      } catch (e) {
        console.error(`Error retrieving domain aliases (retry ${retryCount + 1}/${maxRetries}):`, e);
        retryCount++;
        Utilities.sleep(1000 * Math.pow(2, retryCount));
      }
    }
    if (!success) {
      console.error("Failed to retrieve domain aliases after multiple retries.");
      break;
    }
  } while (pageToken);

  return domainList;
}

function getDnsRecords(domainList) {
  const dnsResults = [];
  const delayBetweenCalls = 100; // milliseconds

  for (let i = 0; i < domainList.length; i++) {
    const domain = domainList[i][0];

    let mxRecords = performGoogleDNSLookup("MX", domain);
    Utilities.sleep(delayBetweenCalls);

    let spfRecords = performGoogleDNSLookup("TXT", domain);
    Utilities.sleep(delayBetweenCalls);

    let dkimRecords = performGoogleDNSLookup("TXT", "google._domainkey." + domain);
    Utilities.sleep(delayBetweenCalls);

    let dmarcRecords = performGoogleDNSLookup("TXT", "_dmarc." + domain);
    Utilities.sleep(delayBetweenCalls);


    mxRecords.data = mxRecords.data || "No MX records found";
    spfRecords.data = spfRecords.data || "No SPF records found";
    dkimRecords.data = dkimRecords.data || "No DKIM records found";
    dmarcRecords.data = dmarcRecords.data || "No DMARC records found";

    dnsResults.push({
      mxRecords: mxRecords,
      spfRecords: spfRecords,
      dkimRecords: dkimRecords,
      dmarcRecords: dmarcRecords
    });
  }
  return dnsResults;
}

/**
 * Performs a DNS lookup using Google Public DNS.
 * @param {string} type The DNS record type (e.g., "MX", "TXT").
 * @param {string} domain The domain to lookup.
 * @return {object} An object with the DNS data and a status message.
 **/

function performGoogleDNSLookup(type, domain) {
  const url = `https://dns.google/resolve?name=${encodeURIComponent(domain)}&type=${encodeURIComponent(type)}`;
  const maxRetries = 3;
  let status = ""; // Intialize status

  for (let retry = 0; retry <= maxRetries; retry++) {
    try {
      const options = {
        muteHttpExceptions: true,
        followRedirects: true // Important to handle redirects.
      };
      const response = UrlFetchApp.fetch(url, options);
      const httpCode = response.getResponseCode();
      const contentText = response.getContentText(); // Get content for error logging.

      if (httpCode === 200) {
        const data = JSON.parse(contentText);
        if (data.Answer) {
          const outputData = data.Answer.map(record => record.data);
          return {
            data: outputData.join('\n'),
            status: "Lookup Complete" // Set successful status for this record type
          };
        } else {
          // No Answer Section
          return {
            data: "",
            status: "No records found" // Set status for no records found
          }; // Return empty string for no record found
        }
      } else if (httpCode === 429) {
        // Handle Too Many Requests
        status = `Too Many Requests`; // Status message
        Logger.warn(`Google Public DNS 429 Too Many Requests for ${domain} (${type}).  Manual check recommended after a cool-down period.`);
        const retryAfter = response.getHeaders()['Retry-After'];
        let waitTime = retryAfter ? parseInt(retryAfter, 10) : 60; // Default to 60 seconds if header is missing
        waitTime = Math.min(waitTime, 300); // Limit wait to 5 minutes.
        Logger.log(`Waiting ${waitTime} seconds before retrying...`);
        Utilities.sleep(waitTime * 1000); // Wait in milliseconds
        // Do NOT return here.  Let the retry happen.
      } else if (httpCode === 500 || httpCode === 502) {
        // Handle Internal Server Error and Bad Gateway (Retry)
        status = `Internal Error`; // Status message
        Logger.warn(`Google Public DNS HTTP Error ${httpCode} for ${domain} (${type}).  Manual check recommended.`);
        Utilities.sleep(1000 * Math.pow(2, retry));  // Exponential Backoff
        // Do NOT return here.  Let the retry happen.
      } else if (httpCode === 400 || httpCode === 413 || httpCode === 414 || httpCode === 415) {
        // Handle Permanent Errors (Don't Retry)
        status = `Bad Request`; // Status message
        Logger.error(`Google Public DNS HTTP Error ${httpCode} for ${domain} (${type}): ${contentText}`);
        return { data: "", status: status }; // Don't retry bad request, payload too large etc.
      } else if (httpCode === 301 || httpCode === 308) {
        //Handle redirects and log.
        status = `Redirect`; // Status message
        Logger.log(`Google Public DNS HTTP Redirect ${httpCode} for ${domain} (${type}).`);
        return { data: "", status: status }; //Return empty string since fetch is following redirects.
      }
      else {
        // Unhandled HTTP Error
        status = `Other Error`; // Status message
        Logger.error(`Google Public DNS Unhandled HTTP Error ${httpCode} for ${domain} (${type}): ${contentText}`);
        return { data: "", status: status };
      }
    } catch (error) {
      status = `Exception`; // Status message
      Logger.error(`Error fetching DNS data from Google Public DNS for ${domain} (${type}): ${error}`);
      return { data: "", status: status };
    }
  } // End Retry Loop
  status = `Multiple Retries Failed`; // Status Message
  Logger.error(`Failed to retrieve DNS record for ${domain} (${type}) after multiple retries.`);
  return { data: "", status: status };
}