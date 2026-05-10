/**
 * -----------------------------------------------------------------------------
 * ORCHID MANAGEMENT SYSTEM - WATERING LOGIC (WateringLogic.gs)
 * -----------------------------------------------------------------------------
 * Version: 2.6.0
 * Features: Modeless UI, Multi-Range List Support, Row 56 Ledger Transcription.
 * -----------------------------------------------------------------------------
 */

/**
 * Opens the HTML dialog as MODELESS.
 * Allows interaction with the spreadsheet (scrolling/selecting) while open.
 */
function showWateringModal() {
  const html = HtmlService.createHtmlOutputFromFile('WateringDialog')
      .setWidth(360)
      .setHeight(420);
  
  SpreadsheetApp.getUi().showModelessDialog(html, 'Log Bulk Watering');
}

/**
 * Helper: Pulls every selected row, even non-adjacent ones (e.g., 38, 47, 51).
 */
function getCurrentRangeString() {
  const selection = SpreadsheetApp.getActiveRangeList();
  if (!selection) return "";
  
  let rows = [];
  selection.getRanges().forEach(range => {
    const start = range.getRow();
    const end = range.getLastRow();
    for (let r = start; r <= end; r++) {
      if (r >= 3) rows.push(r); // Skip headers
    }
  });
  
  // Sort numerically and remove duplicates
  return [...new Set(rows)].sort((a, b) => a - b).join(", ");
}

/**
 * Main Bridge: Processes the range string and updates all relevant sheets.
 */
function processBulkWatering(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Maintenance Log");
  const now = new Date();
  
  if (!sheet) throw new Error("Maintenance Log sheet not found.");

  // 1. Parse the input (Handles "38, 47" or "38:40, 55")
  let rows = [];
  const parts = data.range.split(/[\s,]+/);
  
  parts.forEach(part => {
    if (part.includes(':')) {
      const bounds = part.split(':');
      const start = parseInt(bounds[0]);
      const end = parseInt(bounds[1]);
      for (let r = start; r <= end; r++) { if (!isNaN(r)) rows.push(r); }
    } else {
      const r = parseInt(part);
      if (!isNaN(r)) rows.push(r);
    }
  });

  // 2. Filter for unique rows and skip headers
  rows = [...new Set(rows)].filter(r => r >= 3);
  if (rows.length === 0) throw new Error("No valid rows selected.");

  // 3. Update the sheets
  rows.forEach(currentRow => {
    const orchidID = sheet.getRange(currentRow, 1).getValue(); // Col A
    
    // Update Main Maintenance Log
    sheet.getRange(currentRow, 3).setValue(data.method); // Col C
    sheet.getRange(currentRow, 4).setValue(now);         // Col D
    sheet.getRange(currentRow, 5).setValue(data.notes);  // Col E

    // Direct Transcription to individual Orchid ID Sheet
    if (orchidID) {
      const orchidSheet = ss.getSheetByName(String(orchidID).trim());
      if (orchidSheet) {
        // Find next empty row starting at ledger line 56
        let targetRow = orchidSheet.getRange("A100").getNextDataCell(SpreadsheetApp.Direction.UP).getRow() + 1;
        if (targetRow < 56) targetRow = 56;

        orchidSheet.getRange(targetRow, 1).setValue(now);
        orchidSheet.getRange(targetRow, 2).setValue(data.method);
        orchidSheet.getRange(targetRow, 3).setValue(data.notes || "");
      }
    }
  });

  return "Success! Updated " + rows.length + " orchid(s).";
}