/**
 * Kasun & Piumi Wedding Invitation – RSVP → Google Sheets
 *
 * Paste this into Extensions → Apps Script of the spreadsheet, then
 * Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone)
 * and put the /exec URL into SCRIPT_URL in src/App.tsx.
 */

const SPREADSHEET_ID = "14axc9Ebxy20pPUgvEKe32roAUsNesGSRCUioe6W_i0I";
const SHEET_NAME = "RSVP";
const HEADERS = ["Timestamp", "Name", "Guests", "Attending", "Wish"];

function doPost(e) {
  return handleRequest(e);
}

function doGet(e) {
  return handleRequest(e);
}

// Run once from the Apps Script editor to create the RSVP sheet with headers
function setup() {
  getOrCreateSheet();
}

function getOrCreateSheet() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  const existing = headerRange.getValues()[0];
  const headersOk = HEADERS.every(function (h, i) {
    return existing[i] === h;
  });

  if (!headersOk) {
    headerRange
      .setValues([HEADERS])
      .setFontWeight("bold")
      .setBackground("#6b21a8")
      .setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, HEADERS.length);
  }

  return sheet;
}

function handleRequest(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    const params = (e && e.parameter) || {};
    if (params.formName !== "rsvp") {
      return json({ result: "error", error: "Invalid formName" });
    }

    const sheet = getOrCreateSheet();
    params["Attending"] = params["Guests"] === "0" ? "No" : "Yes";

    const row = HEADERS.map(function (h, i) {
      return i === 0 ? new Date() : params[h] || "";
    });
    sheet.appendRow(row);

    return json({ result: "success", row: sheet.getLastRow() });
  } catch (error) {
    return json({ result: "error", error: error.message });
  } finally {
    lock.releaseLock();
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
