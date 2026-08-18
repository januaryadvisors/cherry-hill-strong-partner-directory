// Send Log webhook for the Cherry Hill Strong Partner Directory.
//
// Deploy this bound to the SAME spreadsheet the directory reads from
// (the one behind SHEET_CSV_URL in index.html).
//
// Setup:
//   1. Open the spreadsheet.
//   2. Extensions > Apps Script.
//   3. Delete any starter code, paste this file's contents in.
//   4. Deploy > New deployment > type "Web app".
//      - Execute as: Me
//      - Who has access: Anyone
//   5. Copy the Web app URL it gives you.
//   6. Paste that URL into SEND_LOG_WEBHOOK_URL in index.html.
//
// Each send from the directory appends one row per organization to a
// "Send Log" tab (created automatically on first send if it doesn't exist).

const LOG_SHEET_NAME = 'Send Log';

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const sheet = getOrCreateLogSheet_();
    const timestamp = new Date();
    const orgs = Array.isArray(payload.orgs) ? payload.orgs : [];

    orgs.forEach(org => {
      sheet.appendRow([
        timestamp,
        payload.sender || '',
        payload.type || '',
        org.name || '',
        org.contactEmail || '',
      ]);
    });

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCreateLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Sent By', 'Type', 'Organization', 'Contact Email']);
  }
  return sheet;
}
