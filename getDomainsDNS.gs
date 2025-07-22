/**
 * This script retrieves domain information from Admin Directory API and performs
 * DNS lookups using Google Public DNS. It was initially based on a script
 * that used Cloudflare DNS
 * (https://github.com/cloudflare/cloudflare-docs/blob/production/content/1.1.1.1/other-ways-to-use-1.1.1.1/dns-in-google-sheets.md),
 * but has been significantly modified to remove Cloudflare DNS and integrate
 * with Google Public DNS and Google Admin API.
 * */
function getDomainList() {
  const functionName = 'getDomainList';
  const startTime = new Date();
  Logger.log(`-- Starting ${functionName} at: ${startTime.toLocaleString()}`);

  try {
    const customerDomain = 'my_customer';
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName('Domains/DNS');

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
    sheet.getRange("D1").setNote("Mail Exchange, Green cells indicate Google MX records were found.");
    sheet.getRange("E1").setNote("Sender Policy Framework, Green cells indicate Google SPF record was found.");
    sheet.getRange("F1").setNote("DomainKeys Identified Mail, Green cells indicate the default DKIM selector for Google was found.");
    sheet.getRange("G1").setNote("Domain-based Message Authentication, Reporting, and Conformance, Green cells indicate DMARC records were found.");
    sheet.getRange("H1").setNote("Status of DNS Lookups");

    // Get the Domain
    const domainList = getDomainInformation_(customerDomain);

    if (domainList.length === 0) {
      sheet.getRange("A2").setValue("No domains found for this customer.");
      Logger.log(`In ${functionName}: No domains found. Aborting further processing.`);
      return; 
    }

    // Write domain information to spreadsheet in one batch
    sheet.getRange(2, 1, domainList.length, domainList[0].length).setValues(domainList);

    // Get DNS Records
    const dnsResultsWithStatus = getDnsRecords(domainList);

    // Get the last row
    const lastRow = sheet.getLastRow();

    //Extract DNS results
    const dnsResults = dnsResultsWithStatus.map(item => [item.mxRecords.data, item.spfRecords.data, item.dkimRecords.data, item.dmarcRecords.data]);
    sheet.getRange(2, 4, lastRow - 1, 4).setValues(dnsResults);

    //Write the status message to the sheet
    const statusCells = sheet.getRange(2, 8, lastRow - 1, 1);
    const richTextValues = dnsResultsWithStatus.map(item => {
      let builder = SpreadsheetApp.newRichTextValue();
      let baseText = "";
      let warnings = [];

      if (item.spfRecords.status.includes("Warning: Multiple SPF records found")) {
        warnings.push("Warning: Multiple SPF records found");
      }

      if (item.mxRecords.status.startsWith("Lookup Complete") && 
          item.spfRecords.status.startsWith("Lookup Complete") && 
          item.dkimRecords.status.startsWith("Lookup Complete") && 
          item.dmarcRecords.status.startsWith("Lookup Complete")) {
        baseText = "Lookup Complete";
      } else {
        let issues = [];
        if (!item.mxRecords.status.startsWith("Lookup Complete")) issues.push(`MX: ${item.mxRecords.status.split(";")[0]}`);
        if (!item.spfRecords.status.startsWith("Lookup Complete")) issues.push(`SPF: ${item.spfRecords.status.split(";")[0]}`);
        if (!item.dkimRecords.status.startsWith("Lookup Complete")) issues.push(`DKIM: ${item.dkimRecords.status.split(";")[0]}`);
        if (!item.dmarcRecords.status.startsWith("Lookup Complete")) issues.push(`DMARC: ${item.dmarcRecords.status.split(";")[0]}`);
        baseText = "Issues Found:\n" + issues.join("\n");
      }

      builder.setText(baseText);

      if (warnings.length > 0) {
        let warningText = "\n" + warnings.join("\n");
        let boldStyle = SpreadsheetApp.newTextStyle().setBold(true).build();
        builder.setText(baseText + warningText);
        builder.setTextStyle(baseText.length, baseText.length + warningText.length, boldStyle);
      }
      
      return [builder.build()];
    });
    statusCells.setRichTextValues(richTextValues);


    // Delete columns I-Z
    if (sheet.getMaxColumns() > 8) {
        sheet.deleteColumns(9, sheet.getMaxColumns() - 8);
    }

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
    sheet.setColumnWidth(7, 164);
    sheet.setColumnWidth(8, 300);

    // Apply conditional formatting rules
    const rules = [];
    const ranges = {
        D: "google",
        E: "_spf.google.com",
        F: "v=dkim1;",
        G: "v=dmarc"
    };

    for (const col in ranges) {
        const range = sheet.getRange(`${col}2:${col}${lastRow}`);
        const text = ranges[col];
        
        const greenRule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextContains(text)
            .setBackground("#b7e1cd")
            .setRanges([range])
            .build();
        
        const redRule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextDoesNotContain(text)
            .setBackground("#ffb6c1")
            .setRanges([range])
            .build();

        rules.push(greenRule, redRule);
    }
    sheet.setConditionalFormatRules(rules);

    // --- Add Filter View ---
    sheet.getRange('A1:H' + sheet.getLastRow()).createFilter();

    // --- Freeze Row 1 ---
    sheet.setFrozenRows(1);

  } catch (e) {
    // --- START OF MODIFICATION ---
    let message = e.message;
    let title = '❌ Script Error';

    // Check for our custom permission error message
    if (message.includes("Permission Denied")) {
        title = '❌ Permission Error';
    }

    // Log the fatal error for debugging
    Logger.log(`!! FATAL ERROR in ${functionName}: ${message}\nStack: ${e.stack}`);
    
    // Display a blocking alert to the user.
    SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
  } finally {
    const endTime = new Date();
    const duration = (endTime.getTime() - startTime.getTime()) / 1000; // Duration in seconds
    Logger.log(`-- Finished ${functionName} at: ${endTime.toLocaleString()} (Duration: ${duration.toFixed(2)}s)`);
  }
}

// --- ALL HELPER FUNCTIONS (getDomainInformation_, getDnsRecords, performGoogleDNSLookup) REMAIN THE SAME ---
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

        if (response && response.domains) {
          response.domains.forEach(function (domain) {
            domainList.push([domain.domainName, domain.verified, domain.isPrimary]);
          });
        }
        success = true;
      } catch (e) {
        // If it's a permission error, don't retry. Throw it immediately.
        if (e.details && e.details.code === 403) {
          Logger.log(`Permission error detected in getDomainInformation_. Halting.`);
          throw new Error("Permission Denied: Could not access domain information. Please run as a Super Administrator.");
        }
        // For other errors, use the original retry logic.
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
        if (e.details && e.details.code === 403) {
          Logger.log(`Permission error detected in getDomainInformation_ for aliases. Halting.`);
          throw new Error("Permission Denied: Could not access domain information. Please run as a Super Administrator.");
        }
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

    // Filter for SPF records and handle multiple policies.
    const spfPolicies = (spfRecords.data || "").match(/^.*v=spf1.*$/gim);
    if (!spfPolicies) {
      spfRecords.data = "No SPF record found";
    } else {
      if (spfPolicies.length > 1) {
        // If multiple SPF records are found, append a warning to the status.
        // This is a configuration error that admins should be aware of.
        spfRecords.status += "; Warning: Multiple SPF records found";
      }
      // Join all found SPF records, separated by a newline.
      spfRecords.data = spfPolicies.join('\n');
    }

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