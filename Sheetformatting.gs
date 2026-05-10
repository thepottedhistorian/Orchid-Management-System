/**
 * -----------------------------------------------------------------------------
 * ORCHID INVENTORY - SHEET FORMATING UTILITIES (SheetFormatting.gs)
 * -----------------------------------------------------------------------------
 * Purpose: Consolidated formatting tools to ensure aesthetic and data 
 * consistency across all individual orchid ledger sheets.
 * Logic:
 * 1. Batch-copies styles (colors, borders, fonts) from a "Master" look.
 * 2. Optimizes row heights for readability.
 * 3. Standardizes date formatting (MMMM d, yyyy) for "History-First" sync.
 * -----------------------------------------------------------------------------
 */

/**
 * NEW: Copy specific layout/styles from a source sheet to targets.
 * Use this when you've updated the look of one sheet and want to "push" 
 * that look to others (colors, borders, fonts).
 */
function applyBatchFormatting() {
  const ss = SpreadsheetApp.openById(MASTER_ID);
  const sourceSheet = ss.getSheetByName('27'); // Change '27' to your current "Gold Standard" sheet
  
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert("Source sheet '27' not found. Please check Globals.");
    return;
  }
  
  const sourceRange = sourceSheet.getRange('A1:P40'); 

  // Target a specific range of sheets (Update these IDs as your collection grows)
  const targetSheets = ['37', '38', '39', '40'];

  targetSheets.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const targetRange = sheet.getRange('A1:P40');
    
    // Copy only the formatting (colors, borders, etc.), not the text/data
    sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  });
  
  SpreadsheetApp.getUi().alert("Layout formatting copied to sheets: " + targetSheets.join(", "));
}

/**
 * Bulk Format Numbered Sheets (2 to 40).
 * Performance optimized: sets row heights and alignments in a single pass.
 */
function bulkFormatSheets_2to40() {
  const ss = SpreadsheetApp.openById(MASTER_ID);

  for (let i = 2; i <= 40; i++) {
    const sheet = ss.getSheetByName(i.toString());
    if (!sheet) continue;

    const range = sheet.getDataRange();
    range.setWrap(true);
    range.setVerticalAlignment("middle");

    // PERFORMANCE FIX: Sets all rows to height 1 first to clear "ghost" heights
    const maxRows = sheet.getMaxRows();
    sheet.setRowHeights(1, maxRows, 1);

    // Reset standard heights for specific headers (Adjust based on your layout)
    sheet.setRowHeight(7, 21);
    sheet.setRowHeight(14, 21);
    
    // Expand rows to fit text content automatically
    sheet.autoResizeRows(1, sheet.getLastRow());
  }
  
  SpreadsheetApp.getUi().alert("Bulk row-height formatting complete.");
}

/**
 * Fix timestamps in A19-A33 (Blooms) and G19-G68 (Watering) on the active sheet.
 * Use this if you've pasted data that didn't automatically format.
 */
function fixAllPhasesSequentially() {
  const ss = SpreadsheetApp.openById(MASTER_ID);
  const sheet = SpreadsheetApp.getActiveSheet();
  
  // Apply "Long Date" format for consistency with sync logic
  fixDateRange_(sheet.getRange(19, 1, 15, 1), "mmmm d, yyyy");
  fixDateRange_(sheet.getRange(19, 7, 50, 1), "mmmm d, yyyy");

  SpreadsheetApp.getUi().alert("Timestamp formatting fixed for: " + sheet.getName());
}

/**
 * INTERNAL ENGINE: Cleans, parses, and formats date ranges.
 * Removes non-printing characters and invisible whitespace.
 */
function fixDateRange_(range, format) {
  const values = range.getValues();
  let changed = false;

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[0].length; c++) {
      const v = values[r][c];
      if (v instanceof Date) continue;

      if (typeof v === "string") {
        // Clean invisible ghost characters and trim whitespace
        const cleaned = v
          .replace(/[^\S\r\n]+/g, " ")
          .replace(/[\u200B-\u200D\uFEFF]/g, "")
          .trim();

        if (!cleaned) continue;

        const parsed = new Date(cleaned);
        if (!isNaN(parsed.getTime())) {
          values[r][c] = parsed;
          changed = true;
        }
      }
    }
  }

  if (changed) range.setValues(values);
  range.setHorizontalAlignment("left").setNumberFormat(format);
}