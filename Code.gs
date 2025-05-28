const spreadsheetId = "YOUR_SPREADSHEET_ID_HERE"; // The ID must be changed to the spredsheet ID used for DMARC tool (https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID_HERE/edit?)

/**
 * Main function: Process DMARC reports from Gmail, parse attachments,
 * append data to spreadsheet, update summary with charts and colors,
 * and export monthly CSV archive.
 */
function processDMARCReports() {
  const labelName = "DMARC";
  const processedLabelName = "DMARC/Processed";
  const sheetName = "DMARC Reports";
  const thresholdFailures = 3;

  try {
    const processedLabel = getOrCreateLabel(processedLabelName);
    const ss = SpreadsheetApp.openById(spreadsheetId);
    Logger.log("Spreadsheet loaded: " + (ss ? "yes" : "no"));
    if (!ss) {
      throw new Error("Could not open spreadsheet. Check spreadsheetId and permissions.");
    }
    const sheet = getOrCreateSheet(ss, sheetName, [
      "Message ID", "Reporter", "Source IP", "Disposition",
      "DKIM", "SPF", "Domain", "Header From", "Count", "Processed Date"
    ]);

    // Get existing message IDs for deduplication
    const lastRow = sheet.getLastRow();
    let existingMessageIds = [];
    if (lastRow > 1) {
      existingMessageIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    }

    // Search Gmail threads with DMARC label and attachments
    const threads = GmailApp.search(`label:${labelName} has:attachment`);
    Logger.log(`Found ${threads.length} threads with label:${labelName} and attachments`);
    const alerts = [];

    for (const thread of threads) {
      Logger.log(`Processing thread: ${thread.getId()}`);
      for (const msg of thread.getMessages()) {
        const msgId = msg.getId();
        Logger.log(`  Message ID: ${msgId}`);
        if (existingMessageIds.includes(msgId)) {
          Logger.log(`    Skipping already processed message: ${msgId}`);
          continue;
        }

        const attachments = msg.getAttachments();
        Logger.log(`    Found ${attachments.length} attachments`);
        for (const attachment of attachments) {
          try {
            const filename = attachment.getName().toLowerCase();
            Logger.log(`      Attachment: ${filename}`);
            let xmlBlobs = [];

            if (filename.endsWith(".zip")) {
              xmlBlobs = Utilities.unzip(attachment.copyBlob());
              Logger.log(`      Unzipped to ${xmlBlobs.length} blobs`);
            } else if (filename.endsWith(".gz")) {
              xmlBlobs = [Utilities.ungzip(attachment.copyBlob())];
              Logger.log(`      GZipped blob processed`);
            } else if (filename.endsWith(".xml")) {
              xmlBlobs = [attachment.copyBlob()];
              Logger.log(`      XML blob processed`);
            } else {
              Logger.log(`      Skipping unsupported file type: ${filename}`);
              continue; // unsupported file type
            }

            for (const blob of xmlBlobs) {
              try {
                Logger.log(`        Attempting XML parse for blob of size ${blob.getBytes().length}`);
                const xml = XmlService.parse(blob.getDataAsString());
                const root = xml.getRootElement();
                const reportMeta = root.getChild("report_metadata");
                const orgName = reportMeta ? reportMeta.getChildText("org_name") : "";
                const records = root.getChildren("record");
                Logger.log(`        Found ${records.length} <record> elements`);
                for (const record of records) {
                  const row = record.getChild("row");
                  const ip = row ? row.getChildText("source_ip") : "";
                  const count = row ? row.getChildText("count") : "";
                  const policy = row ? row.getChild("policy_evaluated") : null;
                  const disposition = policy ? policy.getChildText("disposition") : "";
                  const dkim = policy ? policy.getChildText("dkim") : "";
                  const spf = policy ? policy.getChildText("spf") : "";
                  const identifiers = record.getChild("identifiers");
                  const headerFrom = identifiers ? identifiers.getChildText("header_from") : "";
                  const authResults = record.getChild("auth_results");
                  const dkimDomain = authResults && authResults.getChild("dkim") ? authResults.getChild("dkim").getChildText("domain") : "";
                  const spfDomain = authResults && authResults.getChild("spf") ? authResults.getChild("spf").getChildText("domain") : "";

                  Logger.log(`        Appending row: [${msgId}, ${orgName}, ${ip}, ${disposition}, ${dkim}, ${spf}, ${dkimDomain || spfDomain}, ${headerFrom}, ${count}, <date>]`);
                  // Append row with current timestamp
                  sheet.appendRow([
                    msgId, orgName, ip, disposition, dkim, spf,
                    dkimDomain || spfDomain, headerFrom, count,
                    new Date()
                  ]);

                  // Alert if failed DKIM or SPF exceeds threshold
                  if ((dkim === "fail" || spf === "fail") && parseInt(count) >= thresholdFailures) {
                    alerts.push(`⚠️ ${orgName} - IP: ${ip} failed DKIM/SPF ${count} times.`);
                  }
                }
              } catch (e) {
                Logger.log(`        XML parse failed: ${e}`);
                continue;
              }
            }

          } catch (err) {
            Logger.log("Error parsing attachment: " + err);
          }
        }

        // Move processed messages to processed label and remove original
        thread.addLabel(processedLabel);
        thread.removeLabel(GmailApp.getUserLabelByName(labelName));
      }
      thread.moveToArchive();
    }

    // Send alert email if failures detected
    if (alerts.length > 0) {
      MailApp.sendEmail({
        to: Session.getActiveUser().getEmail(),
        subject: "DMARC Alert: SPF/DKIM Failures Detected",
        body: alerts.join("\n")
      });
    }

    // Update summary sheet with charts and formatting
    updateDMARCSummary(ss);

    // Color the tabs for visibility
    try {
      ss.getSheetByName("DMARC Reports").setTabColor("#4285F4"); // Google blue
    } catch (e) {}
    try {
      ss.getSheetByName("Summary").setTabColor("#34A853"); // Google green
    } catch (e) {}

    // Export monthly CSV archive
    exportMonthlyCSV(ss, sheetName);

  } catch (err) {
    Logger.log("Error in processDMARCReports: " + err);
  }
}

/**
 * Get or create Gmail label by name
 */
function getOrCreateLabel(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

/**
 * Get or create a sheet with headers; clears if exists
 */
function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  // Always check and set headers
  const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  let needsHeader = false;
  for (let i = 0; i < headers.length; i++) {
    if (firstRow[i] !== headers[i]) {
      needsHeader = true;
      break;
    }
  }
  if (needsHeader) {
    sheet.clear();
    sheet.appendRow(headers);
  }
  // Always ensure header formatting is present
  sheet.getRange(1, 1, 1, headers.length).setBackground("#b7e1cd").setFontWeight("bold");
  return sheet;
}

/**
 * Update or create the summary sheet with aggregated data,
 * colored formatting and charts for better visualization
 */
function updateDMARCSummary(ss) {
  if (!ss) return;

  const rawSheet = ss.getSheetByName("DMARC Reports");
  if (!rawSheet) return;

  let summarySheet = ss.getSheetByName("Summary");
  if (!summarySheet) {
    summarySheet = ss.insertSheet("Summary");
  } else {
    summarySheet.clear();
    const charts = summarySheet.getCharts();
    charts.forEach(chart => summarySheet.removeChart(chart));
  }

  const data = rawSheet.getDataRange().getValues();
  if (data.length < 2) {
    // Still create headers and style if no data
    summarySheet.getRange("A1:B1").setValues([["Reporting Org", "Report Count"]]);
    summarySheet.getRange("A1:B1").setBackground("#b7e1cd").setFontWeight("bold");
    summarySheet.getRange("D1").setValue("No DMARC data available");
    return;
  }

  const headers = data[0];
  const orgIndex = headers.indexOf("Reporter");
  const ipIndex = headers.indexOf("Source IP");
  const dkimIndex = headers.indexOf("DKIM");
  const spfIndex = headers.indexOf("SPF");

  // Aggregate counts by org and failing IPs
  const orgMap = {};
  const failMap = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const org = row[orgIndex];
    const ip = row[ipIndex];
    const dkim = row[dkimIndex];
    const spf = row[spfIndex];

    orgMap[org] = (orgMap[org] || 0) + 1;

    if (dkim !== "pass" || spf !== "pass") {
      failMap[ip] = (failMap[ip] || 0) + 1;
    }
  }

  // Write Reporting Orgs summary
  summarySheet.getRange("A1:B1").setValues([["Reporting Org", "Report Count"]]);
  const orgEntries = Object.entries(orgMap).sort((a, b) => b[1] - a[1]);
  if (orgEntries.length) {
    summarySheet.getRange(2, 1, orgEntries.length, 2).setValues(orgEntries);
    // Add borders to org summary
    summarySheet.getRange(1, 1, orgEntries.length + 1, 2).setBorder(true, true, true, true, true, true);
  }
  // Color header row
  summarySheet.getRange("A1:B1").setBackground("#b7e1cd").setFontWeight("bold");

  // Write Failing IP summary
  const startRow = orgEntries.length + 4;
  summarySheet.getRange(startRow, 1, 1, 2).setValues([["Failing IP", "Failure Count"]]);
  const failEntries = Object.entries(failMap).sort((a, b) => b[1] - a[1]);
  if (failEntries.length) {
    summarySheet.getRange(startRow + 1, 1, failEntries.length, 2).setValues(failEntries);
    // Add borders to fail summary
    summarySheet.getRange(startRow, 1, failEntries.length + 1, 2).setBorder(true, true, true, true, true, true);
  }
  summarySheet.getRange(startRow, 1, 1, 2).setBackground("#f4cccc").setFontWeight("bold");

  // Place charts as floating objects above the grid, do not expand any row/column
  const pieChartCellRow = 2;
  const pieChartCellCol = 5; // Column E
  const barChartCellRow = 16;
  const barChartCellCol = 5; // Column E
  summarySheet.getRange(pieChartCellRow, pieChartCellCol).setValue("Report Counts by Org").setFontWeight("bold").setFontSize(12);
  summarySheet.getRange(barChartCellRow, barChartCellCol).setValue("Top Failing IPs").setFontWeight("bold").setFontSize(12);

  if (orgEntries.length > 0) {
    const pieChartRange = summarySheet.getRange("A1:B" + (orgEntries.length + 1));
    const pieChart = summarySheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(pieChartRange)
      .setPosition(pieChartCellRow + 1, pieChartCellCol, 0, 0) // E3
      .setOption('title', '')
      .build();
    summarySheet.insertChart(pieChart);
    // Do NOT set row/column height
  }
  if (failEntries.length > 0) {
    const barChartRange = summarySheet.getRange(startRow, 1, Math.min(6, failEntries.length + 1), 2);
    const barChart = summarySheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(barChartRange)
      .setPosition(barChartCellRow + 1, barChartCellCol, 0, 0) // E17
      .setOption('title', '')
      .build();
    summarySheet.insertChart(barChart);
    // Do NOT set row/column height
  }
}

/**
 * Export the current month's DMARC report data as a CSV file
 * stored in a 'DMARC Archives' folder in Google Drive
 */
function exportMonthlyCSV(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;

  const folderName = "DMARC Archives";
  let folder;
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(folderName);
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return; // No data to export

  // Filter data for current month
  const currentDate = new Date();
  const currentMonth = currentDate.getMonth();
  const currentYear = currentDate.getFullYear();

  const filteredData = data.filter((row, i) => {
    if (i === 0) return true; // headers
    const date = new Date(row[row.length - 1]); // last column = Processed Date
    return date.getMonth() === currentMonth && date.getFullYear() === currentYear;
  });

  if (filteredData.length < 2) return; // No data for current month

  // Convert to CSV
  const csvContent = filteredData.map(row =>
    row.map(cell => `"${(cell + "").replace(/"/g, '""')}"`).join(",")
  ).join("\r\n");

  const fileName = `DMARC_Report_${currentYear}_${(currentMonth + 1).toString().padStart(2, "0")}.csv`;
  // Create or overwrite existing file
  const existingFiles = folder.getFilesByName(fileName);
  while (existingFiles.hasNext()) {
    existingFiles.next().setTrashed(true);
  }

  folder.createFile(fileName, csvContent, MimeType.PLAIN_TEXT);
}

/**
 * Add custom menu to spreadsheet UI to manually trigger DMARC processing
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("DMARC Tools")
    .addItem("Process DMARC Reports", "processDMARCReports")
    .addToUi();
}

/**
 * Auto-label DMARC reports in the last 7 days from common senders
 */
function autoLabelDMARCReports() {
  // Search for emails from common DMARC senders with .xml/.zip/.gz attachments in the last 7 days
  var threads = GmailApp.search(
    'newer_than:7d (subject:"Report domain:" OR subject:"DMARC" OR subject:"Aggregate report") has:attachment'
  );
  var label = getOrCreateLabel("DMARC");
  threads.forEach(function(thread) {
    var messages = thread.getMessages();
    var hasDMARC = messages.some(function(msg) {
      var attachments = msg.getAttachments();
      return attachments.some(function(att) {
        var name = att.getName().toLowerCase();
        return name.endsWith('.xml') || name.endsWith('.zip') || name.endsWith('.gz');
      });
    });
    var threadLabels = thread.getLabels().map(function(l) { return l.getName(); });
    if (hasDMARC && threadLabels.indexOf(label.getName()) === -1) {
      thread.addLabel(label);
    }
  });
}

/**
 * Combined function: Auto-label and process DMARC reports in one go.
 * Run this function on a daily (or more frequent) trigger for full automation.
 */
function autoLabelAndProcessDMARCReports() {
  autoLabelDMARCReports();
  processDMARCReports();
}

/**
 * Test function to color the tabs of the spreadsheet
 */
function colorTabsTest() {
  var ss = SpreadsheetApp.openById(spreadsheetId);
  ss.getSheetByName("DMARC Reports").setTabColor("#4285F4");
  ss.getSheetByName("Summary").setTabColor("#34A853");
}

/**
 * List all sheet names in the active spreadsheet
 */
function listSheetNames() {
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sheets = ss.getSheets();
  var names = sheets.map(function(sheet) { return sheet.getName(); });
  Logger.log(names);
}

/**
 * Delete DMARC/Processed emails older than 7 days
 * Run this on a daily trigger to keep mailbox clean
 */
function deleteOldProcessedDMARCEmails() {
  var threads = GmailApp.search('label:"DMARC/Processed" older_than:7d');
  threads.forEach(function(thread) {
    thread.moveToTrash();
  });
}
