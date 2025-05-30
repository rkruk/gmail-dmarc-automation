const spreadsheetId = "YOUR_SPREADSHEET_ID_HERE";

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
    // --- Always ensure all setup and branding is up to date ---
    setupConfigSheet(); // Ensures Config exists and is correct
    setupHelpSheet();   // Ensures Help tab exists
    setupDashboardSheet(); // Ensures Dashboard exists and is up to date

    const processedLabel = getOrCreateLabel(processedLabelName);
    const ss = SpreadsheetApp.openById(spreadsheetId);
    Logger.log("Spreadsheet loaded: " + (ss ? "yes" : "no"));
    if (!ss) {
      throw new Error("Could not open spreadsheet. Check spreadsheetId and permissions.");
    }

    // Automatically archive last month's data before processing new reports
    archiveLastMonthDMARCData(ss, sheetName);

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

    // Export monthly CSV archive
    exportMonthlyCSV(ss, sheetName);

    // Enrich DMARC Reports with Country and Failure Reason columns
    enrichDMARCReportsWithGeoAndReason();

    // Purge old data according to Config
    purgeOldDMARCData();

    // Add drill-down links to summary
    addDrillDownLinksToSummary();

    // --- Always apply branding and logo automatically ---
    applyBranding();

    // --- Scheduled report (if on trigger) ---
    sendScheduledDMARCReport();

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
  // Always ensure header formatting is present for all columns (including new ones)
  var headerCols = sheet.getLastColumn();
  var headerRange = sheet.getRange(1, 1, 1, headerCols);
  headerRange.clearFormat(); // Remove all previous formatting
  headerRange.setBackground("#b7e1cd");
  headerRange.setFontWeight("bold");
  headerRange.setFontColor("#000000");
  headerRange.setFontSize(10);

  // Ensure all data columns (including new ones) have consistent number formatting and alignment
  for (var col = 1; col <= sheet.getLastColumn(); col++) {
    sheet.setColumnWidth(col, 120); // Set a reasonable default width for all columns
    sheet.getRange(1, col, sheet.getLastRow()).setHorizontalAlignment("left");
    sheet.getRange(1, col, sheet.getLastRow()).setVerticalAlignment("middle");
    // Optionally, auto-resize columns for content
    sheet.autoResizeColumn(col);
  }
  return sheet;
}

/**
 * Archive last month's DMARC data to a new sheet and keep only current month's data in DMARC Reports
 * Call this at the start of each month (e.g. in your main trigger)
 */
function archiveLastMonthDMARCData(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return; // No data
  const headers = data[0];
  const dateCol = headers.length - 1;
  const now = new Date();
  const lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const lastMonthNum = lastMonth.getMonth();
  const lastMonthYear = lastMonth.getFullYear();
  // Filter last month's data
  const lastMonthRows = data.filter(function(row, i) {
    if (i === 0) return false;
    const date = new Date(row[dateCol]);
    return date.getMonth() === lastMonthNum && date.getFullYear() === lastMonthYear;
  });
  if (lastMonthRows.length === 0) return;
  // Create or get archive sheet
  const archiveSheetName = `${lastMonthYear}-${String(lastMonthNum + 1).padStart(2, "0")}`;
  let archiveSheet = ss.getSheetByName(archiveSheetName);
  if (!archiveSheet) {
    archiveSheet = ss.insertSheet(archiveSheetName);
    archiveSheet.appendRow(headers);
  }
  // Append last month's rows to archive sheet
  archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, lastMonthRows.length, headers.length)
    .setValues(lastMonthRows);
  // Remove last month's rows from main sheet
  for (let i = data.length - 1; i > 0; i--) {
    const date = new Date(data[i][dateCol]);
    if (date.getMonth() === lastMonthNum && date.getFullYear() === lastMonthYear) {
      sheet.deleteRow(i + 1);
    }
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

/**
 * Aggregate DMARC data from all archive sheets and the current sheet for enterprise-level reporting
 * Returns an array of all rows (with headers)
 */
function getAllDMARCData(ss, mainSheetName) {
  const sheets = ss.getSheets();
  let allData = [];
  let headers = null;
  sheets.forEach(function(sheet) {
    const name = sheet.getName();
    if (name === mainSheetName || /^\d{4}-\d{2}$/.test(name)) {
      const data = sheet.getDataRange().getValues();
      if (data.length < 2) return;
      if (!headers) headers = data[0];
      allData = allData.concat(data.slice(1));
    }
  });
  return headers ? [headers].concat(allData) : [];
}

/**
 * Update or create the summary sheet with aggregated data,
 * colored formatting and charts for better visualization
 * Allows filtering by date range if D1 cell is set (format: YYYY-MM-DD:YYYY-MM-DD)
 * If D2 cell is set to 'ALL', aggregates across all archive sheets and current
 */
function updateDMARCSummary(ss) {
  if (!ss) return;
  let summarySheet = ss.getSheetByName("Summary");
  if (!summarySheet) {
    summarySheet = ss.insertSheet("Summary");
  } else {
    summarySheet.clear();
    const charts = summarySheet.getCharts();
    charts.forEach(function(chart) { summarySheet.removeChart(chart); });
  }

  // --- Data Preparation ---
  // Date range filter (optional)
  let dateRange = summarySheet.getRange("C2").getValue();
  let useAll = summarySheet.getRange("C3").getValue();
  let data;
  if (useAll && useAll.toString().toUpperCase() === 'ALL') {
    data = getAllDMARCData(ss, "DMARC Reports");
  } else {
    const rawSheet = ss.getSheetByName("DMARC Reports");
    if (!rawSheet) return;
    data = rawSheet.getDataRange().getValues();
  }
  if (!data || data.length < 2) {
    summarySheet.getRange("B5:C5").setValues([["Reporting Org", "Report Count"]]);
    summarySheet.getRange("B5:C5").setBackground("#b7e1cd").setFontWeight("bold");
    summarySheet.getRange("C2").setValue("No DMARC data available");
    return;
  }
  // Filter by date range if set
  if (dateRange && typeof dateRange === "string" && dateRange.includes(":")) {
    const [start, end] = dateRange.split(":");
    const startDate = new Date(start);
    const endDate = new Date(end);
    const dateCol = data[0].length - 1;
    data = [data[0]].concat(data.slice(1).filter(function(row) {
      const d = new Date(row[dateCol]);
      return d >= startDate && d <= endDate;
    }));
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
  const orgEntries = Object.entries(orgMap).sort(function(a, b) { return b[1] - a[1]; });
  const failEntries = Object.entries(failMap).sort(function(a, b) { return b[1] - a[1]; });

  // --- Additional Analysis Section ---
  // 1. Top sending domains (from 'Domain' column)
  const domainIndex = headers.indexOf("Domain");
  const domainMap = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const domain = row[domainIndex];
    if (domain) domainMap[domain] = (domainMap[domain] || 0) + 1;
  }
  const domainEntries = Object.entries(domainMap).sort((a, b) => b[1] - a[1]);

  // 2. Pass/fail rates (DKIM, SPF)
  let dkimPass = 0, dkimFail = 0, spfPass = 0, spfFail = 0;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[dkimIndex] === "pass") dkimPass++; else dkimFail++;
    if (row[spfIndex] === "pass") spfPass++; else spfFail++;
  }

  // --- Professional Summary Layout ---
  // Section: Controls (move to top left, but do not display as data)
  summarySheet.getRange("B1").setValue("Controls:").setFontWeight("bold").setFontSize(11);
  summarySheet.getRange("B2").setValue("Date Range (YYYY-MM-DD:YYYY-MM-DD)").setFontWeight("bold").setBackground("#e3e3e3");
  summarySheet.getRange("C2").setValue("").setNote("Enter a date range here, e.g. 2025-05-01:2025-05-30. Leave blank for all dates.");
  summarySheet.getRange("B3").setValue("Type 'ALL' to aggregate all months").setFontWeight("bold").setBackground("#e3e3e3");
  summarySheet.getRange("C3").setValue("").setNote("Type ALL to aggregate all months of data, or leave blank for current month only.");
  summarySheet.getRange("B1:C3").setBorder(true, true, true, true, true, true).setBackground("#f9f9f9");

  // Section: Reporting Org Table
  let rowPtr = 5;
  summarySheet.getRange(rowPtr, 2, 1, 2).setValues([["Reporting Org", "Report Count"]]);
  summarySheet.getRange(rowPtr, 2, 1, 2).setBackground("#b7e1cd").setFontWeight("bold");
  if (orgEntries.length) {
    summarySheet.getRange(rowPtr + 1, 2, orgEntries.length, 2).setValues(orgEntries);
    summarySheet.getRange(rowPtr, 2, orgEntries.length + 1, 2).setBorder(true, true, true, true, true, true);
  }
  rowPtr += orgEntries.length + 3;

  // Section: Failing IP Table
  summarySheet.getRange(rowPtr, 2, 1, 2).setValues([["Failing IP", "Failure Count"]]);
  summarySheet.getRange(rowPtr, 2, 1, 2).setBackground("#f4cccc").setFontWeight("bold");
  if (failEntries.length) {
    summarySheet.getRange(rowPtr + 1, 2, failEntries.length, 2).setValues(failEntries);
    summarySheet.getRange(rowPtr, 2, failEntries.length + 1, 2).setBorder(true, true, true, true, true, true);
  }
  rowPtr += failEntries.length + 3;

  // Section: Top Sending Domains
  if (domainEntries.length) {
    summarySheet.getRange(rowPtr, 2, 1, 2).setValues([["Top Sending Domain", "Count"]]);
    summarySheet.getRange(rowPtr, 2, 1, 2).setBackground("#cfe2f3").setFontWeight("bold");
    summarySheet.getRange(rowPtr + 1, 2, domainEntries.length, 2).setValues(domainEntries);
    summarySheet.getRange(rowPtr, 2, domainEntries.length + 1, 2).setBorder(true, true, true, true, true, true);
    rowPtr += domainEntries.length + 3;
  }

  // Section: DKIM/SPF Pass/Fail Table
  summarySheet.getRange(rowPtr, 2, 1, 4).setValues([["DKIM Pass", "DKIM Fail", "SPF Pass", "SPF Fail"]]);
  summarySheet.getRange(rowPtr, 2, 1, 4).setBackground("#ffe599").setFontWeight("bold");
  summarySheet.getRange(rowPtr + 1, 2, 1, 4).setValues([[dkimPass, dkimFail, spfPass, spfFail]]);
  summarySheet.getRange(rowPtr, 2, 2, 4).setBorder(true, true, true, true, true, true);
  rowPtr += 4;

  // --- Chart Placement (dynamic, non-overlapping, right side) ---
  // Use only row-based positioning for titles and charts, not pixel offsets, to guarantee correct stacking in Google Sheets.
  let chartCol = 7; // Column G for charts
  let chartRow = 2;
  const chartPadding = 8; // rows to skip after each chart for guaranteed separation
  function placeChartWithTitle(title, chartRange, chartType, chartRows, chartCols) {
    // Place the title in the current row
    summarySheet.getRange(chartRow, chartCol, 1, chartCols || 2).merge().setValue(title).setFontWeight("bold").setFontSize(12);
    // Place the chart directly below the title, using row-based positioning
    const chart = summarySheet.newChart()
      .setChartType(chartType)
      .addRange(chartRange)
      .setPosition(chartRow + 1, chartCol, 0, 0)
      .setOption('title', '')
      .build();
    summarySheet.insertChart(chart);
    // Move chartRow pointer down by chartRows (height of chart) + title + padding
    chartRow += chartRows + 1 + chartPadding;
  }
  if (orgEntries.length > 0) {
    const pieChartRange = summarySheet.getRange(5, 2, orgEntries.length + 1, 2);
    placeChartWithTitle("Report Counts by Org", pieChartRange, Charts.ChartType.PIE, Math.max(orgEntries.length + 8, 12));
  }
  if (failEntries.length > 0) {
    const barChartRange = summarySheet.getRange(5 + orgEntries.length + 3, 2, Math.min(6, failEntries.length + 1), 2);
    placeChartWithTitle("Top Failing IPs", barChartRange, Charts.ChartType.BAR, Math.max(Math.min(6, failEntries.length + 1) + 8, 12));
  }
  if (domainEntries.length > 0) {
    const domainChartRange = summarySheet.getRange(rowPtr - domainEntries.length - 3, 2, Math.min(6, domainEntries.length + 1), 2);
    placeChartWithTitle("Top Sending Domains", domainChartRange, Charts.ChartType.BAR, Math.max(Math.min(6, domainEntries.length + 1) + 8, 12));
  }
  // Pass/fail pie chart
  const pfChartRange = summarySheet.getRange(rowPtr - 3, 2, 1, 4);
  placeChartWithTitle("DKIM/SPF Pass/Fail", pfChartRange, Charts.ChartType.PIE, 12, 4);

  // Auto-resize columns and rows for best fit (with extra width for long headers)
  summarySheet.autoResizeColumns(1, summarySheet.getMaxColumns());
  // Manually set minimum width for key columns to ensure full header visibility
  summarySheet.setColumnWidth(2, 140); // B: e.g. 'Reporting Org', 'Failing IP', etc.
  summarySheet.setColumnWidth(3, 120); // C: e.g. 'Report Count', 'Failure Count', etc.
  summarySheet.setColumnWidth(4, 120); // D: e.g. 'SPF Pass', etc.
  summarySheet.setColumnWidth(5, 120); // E: e.g. 'SPF Fail', etc.
  summarySheet.autoResizeRows(1, summarySheet.getMaxRows());

  // Always set tab colors after summary update (use American English: setTabColor)
  try {
    const dmarcSheet = ss.getSheetByName("DMARC Reports");
    if (dmarcSheet) dmarcSheet.setTabColor("#4285F4");
  } catch (e) {}
  try {
    if (summarySheet) summarySheet.setTabColor("#34A853");
  } catch (e) {}
}

/**
 * Add drill-down hyperlinks in the Summary sheet to jump to filtered data in DMARC Reports
 */
function addDrillDownLinksToSummary() {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const summary = ss.getSheetByName('Summary');
  const reports = ss.getSheetByName('DMARC Reports');
  if (!summary || !reports) return;
  const data = summary.getDataRange().getValues();
  // Add links for Reporting Org and Failing IP tables
  for (let row = 6; row < data.length; row++) {
    // Reporting Org links (col 2)
    const org = data[row][1];
    if (org && typeof org === 'string' && org !== '' && org !== 'Reporting Org') {
      summary.getRange(row + 1, 2).setFormula(`=HYPERLINK("#gid=${reports.getSheetId()}&filter=Reporter:${org}", "${org}")`);
    }
    // Failing IP links (col 2, after org table)
    if (data[row][1] && data[row][0] && data[row][0].match(/\d+\.\d+\.\d+\.\d+/)) {
      const ip = data[row][0];
      summary.getRange(row + 1, 2).setFormula(`=HYPERLINK("#gid=${reports.getSheetId()}&filter=Source IP:${ip}", "${ip}")`);
    }
  }
}

/**
 * Create a Config sheet for settings (email recipients, retention, etc.)
 */
function setupConfigSheet() {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  let configSheet = ss.getSheetByName('Config');
  if (!configSheet) {
    configSheet = ss.insertSheet('Config');
    configSheet.getRange(1, 1, 1, 2).setValues([["Setting", "Value"]]);
    configSheet.getRange(2, 1, 1, 2).setValues([["Report Recipients (comma separated)", Session.getActiveUser().getEmail()]]);
    configSheet.getRange(3, 1, 1, 2).setValues([["Retention Months", 12]]);
    configSheet.getRange(1, 1, 1, 2).setBackground("#b7e1cd").setFontWeight("bold");
    configSheet.setColumnWidths(1, 2, 260);
    configSheet.setFrozenRows(1);
    configSheet.setTabColor("#333333");
  }
}

/**
 * Purge/archive DMARC data older than the retention period set in Config sheet
 */
function purgeOldDMARCData() {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const configSheet = ss.getSheetByName('Config');
  if (!configSheet) return;
  const retentionMonths = parseInt(configSheet.getRange(3, 2).getValue(), 10) || 12;
  const mainSheet = ss.getSheetByName('DMARC Reports');
  if (!mainSheet) return;
  const data = mainSheet.getDataRange().getValues();
  if (data.length < 2) return;
  const headers = data[0];
  const dateCol = headers.indexOf('Processed Date');
  const now = new Date();
  const cutoff = new Date(now.getFullYear(), now.getMonth() - retentionMonths, now.getDate());
  for (let i = data.length - 1; i > 0; i--) {
    const rowDate = new Date(data[i][dateCol]);
    if (rowDate < cutoff) {
      mainSheet.deleteRow(i + 1);
    }
  }
}

// --- Add IP Country and Failure Reason columns to DMARC Reports sheet ---
function enrichDMARCReportsWithGeoAndReason() {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName("DMARC Reports");
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;
  const headers = data[0];
  let ipCol = headers.indexOf("Source IP");
  let dispCol = headers.indexOf("Disposition");
  let dkimCol = headers.indexOf("DKIM");
  let spfCol = headers.indexOf("SPF");
  let countCol = headers.indexOf("Count");
  // Add columns if not present
  let countryCol = headers.indexOf("Country");
  let reasonCol = headers.indexOf("Failure Reason");
  let needHeaderUpdate = false;
  if (countryCol === -1) { headers.push("Country"); countryCol = headers.length - 1; needHeaderUpdate = true; }
  if (reasonCol === -1) { headers.push("Failure Reason"); reasonCol = headers.length - 1; needHeaderUpdate = true; }
  if (needHeaderUpdate) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    // Always style the entire header row (including new columns) to match: light green, bold
    var headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    headerRange.setBackground(null); // Clear any previous background
    headerRange.setFontWeight("normal"); // Clear any previous font weight
    headerRange.setBackground("#b7e1cd");
    headerRange.setFontWeight("bold");

    // Ensure all data columns (including new ones) have consistent number formatting and alignment
    for (var col = 1; col <= sheet.getLastColumn(); col++) {
      sheet.setColumnWidth(col, 120); // Set a reasonable default width for all columns
      sheet.getRange(1, col, sheet.getLastRow()).setHorizontalAlignment("left");
      sheet.getRange(1, col, sheet.getLastRow()).setVerticalAlignment("middle");
      // Optionally, auto-resize columns for content
      sheet.autoResizeColumn(col);
    }
  }
  // Prepare to update rows
  let updates = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const ip = row[ipCol];
    const disp = row[dispCol];
    const dkim = row[dkimCol];
    const spf = row[spfCol];
    const count = row[countCol];
    // Failure Reason logic
    let reason = "";
    if (disp === "reject") {
      if (dkim === "fail" && spf === "fail") reason = "Both DKIM and SPF failed. Message rejected.";
      else if (dkim === "fail") reason = "DKIM failed. Message rejected.";
      else if (spf === "fail") reason = "SPF failed. Message rejected.";
      else reason = "Rejected for other policy reason.";
    } else if (disp === "none") {
      if (dkim === "fail" && spf === "fail") reason = "Both DKIM and SPF failed, but policy is 'none'. No action taken.";
      else if (dkim === "fail") reason = "DKIM failed, but policy is 'none'. No action taken.";
      else if (spf === "fail") reason = "SPF failed, but policy is 'none'. No action taken.";
      else reason = "Passed authentication, no action taken.";
    } else {
      reason = `Disposition: ${disp}, DKIM: ${dkim}, SPF: ${spf}`;
    }
    // GeoIP lookup (ip-api.com, free, but rate-limited)
    let country = row[countryCol] || "";
    if (!country && ip && /^\d+\.\d+\.\d+\.\d+$/.test(ip)) {
      try {
        const response = UrlFetchApp.fetch(`http://ip-api.com/json/${ip}?fields=country`, {muteHttpExceptions:true, timeout:5});
        const geo = JSON.parse(response.getContentText());
        country = geo && geo.country ? geo.country : "Unknown";
      } catch (e) { country = "Unknown"; }
    }
    // Prepare update
    let updateRow = row.slice();
    updateRow[countryCol] = country;
    updateRow[reasonCol] = reason;
    updates.push(updateRow);
  }
  // Write back enriched data
  if (updates.length) {
    sheet.getRange(2, 1, updates.length, headers.length).setValues(updates);
    // Auto-resize new columns and all rows for visibility
    sheet.autoResizeColumn(countryCol + 1);
    sheet.autoResizeColumn(reasonCol + 1);
    sheet.autoResizeRows(1, sheet.getLastRow());
  }
}

/**
 * Add a Documentation/Help sheet with usage, contact, and glossary
 */
function setupHelpSheet() {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  let helpSheet = ss.getSheetByName('Help');
  if (!helpSheet) helpSheet = ss.insertSheet('Help');
  helpSheet.clear();
  helpSheet.getRange(1, 1).setValue('DMARC Reporting Tool - Help & Documentation').setFontWeight('bold').setFontSize(14);
  helpSheet.getRange(3, 1).setValue('Usage Instructions:').setFontWeight('bold');
  helpSheet.getRange(4, 1).setValue('1. DMARC reports are processed automatically from your Gmail.');
  helpSheet.getRange(5, 1).setValue('2. The "DMARC Reports" sheet contains all parsed data.');
  helpSheet.getRange(6, 1).setValue('3. The "Summary" and "Dashboard" sheets provide visual analytics.');
  helpSheet.getRange(7, 1).setValue('4. The "Config" sheet lets you set report recipients, retention, and logo.');
  helpSheet.getRange(8, 1).setValue('5. Data older than the retention period is purged automatically.');
  helpSheet.getRange(10, 1).setValue('Contact:').setFontWeight('bold');
  helpSheet.getRange(11, 1).setValue('For support, contact your IT administrator or security team.');
  helpSheet.getRange(13, 1).setValue('Glossary:').setFontWeight('bold');
  helpSheet.getRange(14, 1).setValue('DMARC: Domain-based Message Authentication, Reporting & Conformance');
  helpSheet.getRange(15, 1).setValue('DKIM: DomainKeys Identified Mail');
  helpSheet.getRange(16, 1).setValue('SPF: Sender Policy Framework');
  helpSheet.getRange(17, 1).setValue('Disposition: The action taken on a message (none, reject, quarantine)');
  helpSheet.setColumnWidth(1, 600);
  helpSheet.setTabColor('#FFD700');
}

/**
 * Add a Dashboard sheet with high-level KPIs and trendlines
 */
function setupDashboardSheet() {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  let dashboard = ss.getSheetByName('Dashboard');
  if (!dashboard) dashboard = ss.insertSheet('Dashboard');
  dashboard.clear();
  dashboard.setTabColor('#000000');
  dashboard.getRange(1, 1).setValue('DMARC Dashboard').setFontWeight('bold').setFontSize(16).setFontFamily('Arial').setFontColor('#000000').setBackground('#FFFFFF');
  dashboard.getRange(2, 1).setValue('Key Metrics').setFontWeight('bold').setFontSize(12).setFontFamily('Arial').setFontColor('#000000');
  // Pull summary stats from DMARC Reports
  const reports = ss.getSheetByName('DMARC Reports');
  if (!reports) return;
  const data = reports.getDataRange().getValues();
  if (data.length < 2) return;
  const headers = data[0];
  const countCol = headers.indexOf('Count');
  const dkimCol = headers.indexOf('DKIM');
  const spfCol = headers.indexOf('SPF');
  let totalMsgs = 0, dkimFails = 0, spfFails = 0;
  for (let i = 1; i < data.length; i++) {
    totalMsgs += parseInt(data[i][countCol], 10) || 0;
    if (data[i][dkimCol] === 'fail') dkimFails++;
    if (data[i][spfCol] === 'fail') spfFails++;
  }
  dashboard.getRange(3, 1, 3, 2).setValues([
    ['Total Messages', totalMsgs],
    ['DKIM Failures', dkimFails],
    ['SPF Failures', spfFails]
  ]);
  dashboard.getRange(3, 1, 3, 1).setFontWeight('bold').setFontFamily('Arial').setFontColor('#000000');
  dashboard.getRange(3, 2, 3, 1).setFontFamily('Arial').setFontColor('#000000');
  dashboard.getRange(1, 1, 6, 2).setBackground('#FFFFFF');
  // Trendline chart for failures over time
  const dateCol = headers.indexOf('Processed Date');
  let trendData = {};
  for (let i = 1; i < data.length; i++) {
    const date = new Date(data[i][dateCol]);
    const key = date.toISOString().slice(0, 10);
    if (!trendData[key]) trendData[key] = { dkim: 0, spf: 0 };
    if (data[i][dkimCol] === 'fail') trendData[key].dkim++;
    if (data[i][spfCol] === 'fail') trendData[key].spf++;
  }
  const trendRows = Object.keys(trendData).sort().map(date => [date, trendData[date].dkim, trendData[date].spf]);
  if (trendRows.length) {
    dashboard.getRange(8, 1, 1, 3).setValues([["Date", "DKIM Failures", "SPF Failures"]]);
    dashboard.getRange(8, 1, 1, 3).setFontWeight('bold').setFontFamily('Arial').setFontColor('#000000').setBackground('#e3e3e3');
    dashboard.getRange(9, 1, trendRows.length, 3).setValues(trendRows);
    dashboard.getRange(9, 1, trendRows.length, 3).setFontFamily('Arial').setFontColor('#000000');
    // Add chart with legend and axis titles
    const chart = dashboard.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(dashboard.getRange(8, 1, trendRows.length + 1, 3))
      .setPosition(2, 4, 0, 0)
      .setOption('title', 'Failures Over Time')
      .setOption('legend', { position: 'right' })
      .setOption('hAxis', { title: 'Date' })
      .setOption('vAxis', { title: 'Failure Count' })
      .build();
    dashboard.insertChart(chart);
    // Add a clear explanation above the chart
    dashboard.getRange(6, 4).setValue('Chart: Blue = DKIM Failures, Red = SPF Failures').setFontColor('#1565c0').setFontSize(10).setFontWeight('bold');
  }
  dashboard.setColumnWidths(1, 4, 140);
  dashboard.setFrozenRows(1);
}

function applyBranding() {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const dashboard = ss.getSheetByName('Dashboard');
  const summary = ss.getSheetByName('Summary');
  // Branding for Dashboard
  if (dashboard) {
    dashboard.getRange(1, 1, dashboard.getMaxRows(), dashboard.getMaxColumns())
      .setFontFamily('Arial').setFontColor('#000000').setBackground('#FFFFFF');
    // Remove all images from Dashboard
    const dashboardImages = dashboard.getImages();
    dashboardImages.forEach(function(img) { img.remove(); });
    // Clear any previous logo/error message
    dashboard.getRange(1, 7).clearContent();
  }
  // Branding for Summary
  if (summary) {
    summary.getRange(1, 1, summary.getMaxRows(), summary.getMaxColumns())
      .setFontFamily('Arial').setFontColor('#000000').setBackground('#FFFFFF');
    // Remove all images from Summary
    const summaryImages = summary.getImages();
    summaryImages.forEach(function(img) { img.remove(); });
    // Clear any previous logo/error message
    summary.getRange(1, 7).clearContent();
  }
}

/**
 * Scheduled email report: send PDF summary to recipients from Config
 */
function sendScheduledDMARCReport() {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const configSheet = ss.getSheetByName('Config');
  if (!configSheet) return;
  const recipients = configSheet.getRange(2, 2).getValue();
  const summarySheet = ss.getSheetByName('Summary');
  if (!summarySheet) return;
  // Export summary as PDF
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=pdf&gid=${summarySheet.getSheetId()}&portrait=false&size=A4&fitw=true&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false`;
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });
  const blob = response.getBlob().setName('DMARC_Summary.pdf');
  // Send email
  MailApp.sendEmail({
    to: recipients,
    subject: 'Scheduled DMARC Report',
    body: 'Please find attached the latest DMARC summary report.',
    attachments: [blob]
  });
}
