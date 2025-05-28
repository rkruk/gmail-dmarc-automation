<h1 align="center">Google Apps Script:</h1>
<h1 align="center">Fully Automated DMARC Report Processor for Gmail & Google Sheets</h1>

## Overview

This Google Apps Script automates the entire workflow of collecting, parsing, storing, summarizing, and archiving DMARC aggregate reports received via email. It is designed for domain administrators, security teams, and anyone who needs to process DMARC XML reports at scale, with zero manual intervention.

**Key Features:**
- Automatically labels and processes all new DMARC report emails in Gmail
- Robustly extracts and parses DMARC XML data (including from ZIP/GZ attachments)
- Appends processed data to a Google Sheet
- Generates a professional summary sheet with charts and tables
- Exports monthly CSV archives to Google Drive
- Cleans up processed emails after 7 days
- Fully automated via time-driven triggers

---

## Why Use This Script?

- **No manual steps:** All DMARC report handling is fully automated.
- **Handles real-world DMARC formats:** Supports .xml, .zip, and .gz attachments from all major DMARC senders.
- **Deduplication:** Never processes the same report twice.
- **Professional reporting:** Summary sheet with charts for quick insights.
- **Mailbox hygiene:** Automatically deletes processed emails after 7 days.
- **Easy to deploy:** Just copy, configure, and set up triggers.

---

## Prerequisites

- A Google account with access to Gmail and Google Drive
- A Google Sheet to store DMARC data (create a new one or use an existing one)
- Basic familiarity with Google Apps Script: .gs (for setup)

---

## Setup Instructions

### 1. Prepare Your Google Sheet
- Create a new Google Sheet (or use an existing one).
- Note the Sheet's ID (the long string in the URL between `/d/` and `/edit`).

### 2. Add the Script
- Open your Google Sheet.
- Go to **Extensions → Apps Script**.
- Paste the entire script into the editor, replacing any existing code.
- At the top of the script, set the `spreadsheetId` variable to your Sheet's ID (The ID must be changed to the spredsheet ID used for DMARC tool: [https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID_HERE/edit?](https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID_HERE/edit?):
  ```js
  const spreadsheetId = "YOUR_SPREADSHEET_ID_HERE";
  ```

### 3. Set Up Gmail Filter (Recommended)
- In Gmail, create a filter to label all incoming DMARC reports:
  - Go to **Settings → Filters and Blocked Addresses → Create a new filter**.
  - In the "Has the words" field, enter:
    ```
    "Report domain:" OR DMARC OR "Aggregate report"
    ```
  - Click "Create filter".
  - Check "Apply the label" and select/create a label named `DMARC`.
  - (Optional) Check "Skip the Inbox (Archive it)" to keep your inbox clean.

### 4. Set Up Triggers
- In the Apps Script editor, click the clock icon (Triggers) in the left sidebar.
- Add two time-driven triggers:
  1. **autoLabelAndProcessDMARCReports** (daily):
     - Function: `autoLabelAndProcessDMARCReports`
     - Event source: Time-driven
     - Type: Day timer (choose a time)
  2. **deleteOldProcessedDMARCEmails** (daily):
     - Function: `deleteOldProcessedDMARCEmails`
     - Event source: Time-driven
     - Type: Day timer (choose a time)

---

## How It Works

1. **Labeling:**
   - The script auto-labels new DMARC emails (with .xml/.zip/.gz attachments) as `DMARC`.
2. **Processing:**
   - It processes all labeled emails, extracts and parses DMARC XML data, and appends the results to the "DMARC Reports" sheet.
   - Deduplication ensures each report is only processed once.
   - After processing, emails are moved to the `DMARC/Processed` label and archived.
3. **Summary & Charts:**
   - The script updates a "Summary" sheet with aggregated data and professional charts.
4. **CSV Export:**
   - Each month, a CSV archive of that month's data is saved to a `DMARC Archives` folder in Google Drive.
5. **Cleanup:**
   - Processed emails older than 7 days are automatically deleted from Gmail.

---

## Script Functions

- `autoLabelDMARCReports()`: Labels new DMARC emails in Gmail.
- `processDMARCReports()`: Processes all labeled DMARC emails, parses attachments, appends data, updates summary, and exports CSV.
- `autoLabelAndProcessDMARCReports()`: Runs both labeling and processing in one go (set this as your main daily trigger).
- `deleteOldProcessedDMARCEmails()`: Deletes processed DMARC emails older than 7 days (set as a daily trigger).
- `onOpen()`: Adds a custom menu to the Google Sheet for manual processing.
- Utility/test functions: `colorTabsTest()`, `listSheetNames()`.

---

## Customization

- **Change the DMARC label name:** Edit the `labelName` variable in the script if you use a different label.
- **Change the processed label name:** Edit the `processedLabelName` variable.
- **Change the threshold for DKIM/SPF failure alerts:** Edit the `thresholdFailures` variable.
- **Change the retention period:** Edit the `older_than:7d` in `deleteOldProcessedDMARCEmails()` if you want a different retention period.

---

## Security & Privacy

- The script only processes emails with the specified label and attachments.
- All extracted data is stored in your Google Sheet and Drive; nothing is sent externally.
- You can review and audit all code and logs in Apps Script.

---

## Troubleshooting

- **Script not running?**
  - Check that triggers are set up correctly.
  - Make sure the Sheet ID is correct and you have permission.
- **No DMARC emails found?**
  - Check your Gmail filter and label setup.
- **Errors in parsing?**
  - Some malformed DMARC reports may not parse; these are skipped and logged.
- **Permission errors?**
  - Make sure you have authorized the script to access Gmail, Drive, and Sheets.

---

## Quotas & Limits

Google Apps Script and related Google services (Gmail, Sheets, Drive) have daily quotas and limits. For most users, these are sufficient for typical DMARC report processing, but heavy use or large volumes may hit these limits. If a quota is exceeded, the script will pause until the next day.

**Key quotas (as of May 2025, subject to change):**
- **Gmail:**
  - Read/modify: ~1,500 messages/day (consumer), higher for Workspace
  - Sending emails: 100/day (consumer), 1,500/day (Workspace)
- **Google Sheets:**
  - Write cells: 50,000,000 cells/day
  - Read cells: 100,000,000 cells/day
  - Calls to SpreadsheetApp: 90 minutes/day
- **Google Drive:**
  - Create files: 2100/day
  - Write operations: 10,000/day
- **Apps Script total runtime:**
  - 6 minutes per execution (consumer), 30 minutes (Workspace)
  - 90 minutes total script runtime per day (consumer), 6 hours (Workspace)

See the [official quotas documentation](https://developers.google.com/apps-script/guides/services/quotas) for the latest details and Workspace-specific limits.

**Note:**
- If you exceed a quota, the script will stop and resume the next day.
- For most domains, DMARC report volume is well within free quotas.
- Large organizations or high-volume domains may need a Google Workspace account for higher limits.
  
---

## Credits

Inspired by the need for a truly hands-off, robust DMARC automation tool for Google Workspace users.

---

## License

This script is provided as-is, without warranty. You are free to use, modify, and share it. Attribution appreciated.

---

## Feedback & Contributions

If you find this script useful or have suggestions, please open an issue or submit a pull request!
