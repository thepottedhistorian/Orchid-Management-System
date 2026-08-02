/**
 * -----------------------------------------------------------------------------
 * ORCHID MANAGEMENT SYSTEM - DATE STANDARDIZATION UTILITIES (DateUtils.gs)
 * -----------------------------------------------------------------------------
 * Version: 2.6.0 (Zero Number Format Calls)
 * Updated: 2026-07-28
 * Project: The Satyrion Chronicles / Orchid Tracker
 * -----------------------------------------------------------------------------
 * SUMMARY:
 * Standardizes date/timestamp entries into "Month day, year" (e.g., "January 2, 2026")
 * across designated sheets (Inventory, Archived Orchids, and ID sheets 1-65).
 * - 100% safe for Google Sheets Tables & Typed Columns (no setNumberFormat calls).
 * - Leaves 4-digit standalone years (e.g., "2026") untouched.
 * - Leaves blank cells and non-date text intact.
 */

/**
 * Standardizes dates within the user's currently selected range.
 */
function standardizeSelectedDates() {
  const range = SpreadsheetApp.getActiveRange();
  if (!range) {
    SpreadsheetApp.getUi().alert('Please select a range of cells first.');
    return;
  }
  const count = processDateRange(range);
  SpreadsheetApp.getUi().alert(`Date Standardization Complete!\nUpdated ${count} date entries in selection.`);
}

/**
 * Standardizes dates on the currently focused active sheet.
 */
function standardizeActiveSheetDates() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getDataRange();
  const count = processDateRange(range);
  SpreadsheetApp.getUi().alert(`Date Standardization Complete!\nUpdated ${count} date entries on sheet: "${sheet.getName()}".`);
}

/**
 * Batch standardizes dates across the Inventory sheet, Archived Orchids sheet,
 * and all Orchid ID sheets in the range 1 through 65.
 */
function standardizeAllOrchidInventoryDates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    'Batch Date Standardization',
    'This will process the "Inventory" sheet, "Archived Orchids" sheet, and Orchid ID sheets 1 through 65.\n\nDo you wish to proceed?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  let totalUpdated = 0;
  let processedSheetsCount = 0;

  const explicitSheetNames = ['Inventory', 'Archived Orchids'];
  
  explicitSheetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      const range = sheet.getDataRange();
      totalUpdated += processDateRange(range);
      processedSheetsCount++;
    }
  });

  for (let i = 1; i <= 65; i++) {
    const sheetName = String(i);
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      const range = sheet.getDataRange();
      totalUpdated += processDateRange(range);
      processedSheetsCount++;
    }
  }

  ui.alert(
    'System Batch Complete',
    `Successfully standardized dates across ${processedSheetsCount} sheets.\nTotal date entries updated: ${totalUpdated}`,
    ui.ButtonSet.OK
  );
}
/**
 * Core processing function that safely converts short dates and native date objects
 * into strict plain-text formatted strings across both standard ranges and typed table columns.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Range} range - The range to process.
 * @return {number} Count of formatted date cells.
 */
function processDateRange(range) {
  const values = range.getValues();
  const displayValues = range.getDisplayValues();
  const timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const sheet = range.getSheet();
  
  const startRow = range.getRow();
  const startCol = range.getColumn();
  const rows = values.length;
  const cols = values[0].length;
  
  let formattedCount = 0;

  const monthNames = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];

  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const cellValue = values[r][c];
      const displayText = String(displayValues[r][c]).replace(/\u00A0/g, ' ').trim();

      // Skip empty cells, header rows, or standalone 4-digit years (e.g., "2026")
      if (displayText === '' || /^\d{4}$/.test(displayText)) {
        continue;
      }

      let formattedDateString = null;

      // 1. Direct Regex match for short dates (e.g., "5/12/2026", "4/26/2025")
      const shortDateMatch = displayText.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
      if (shortDateMatch) {
        const monthIdx = parseInt(shortDateMatch[1], 10) - 1;
        const day = parseInt(shortDateMatch[2], 10);
        const year = shortDateMatch[3];

        if (monthIdx >= 0 && monthIdx < 12) {
          formattedDateString = `${monthNames[monthIdx]} ${day}, ${year}`;
        }
      }

      // 2. Fallback for native Date objects or ISO/timestamp strings
      if (!formattedDateString) {
        let dateObj = null;
        if (cellValue instanceof Date && !isNaN(cellValue.getTime())) {
          dateObj = cellValue;
        } else {
          const parsed = Date.parse(displayText);
          if (!isNaN(parsed)) {
            dateObj = new Date(parsed);
          }
        }

        if (dateObj) {
          formattedDateString = Utilities.formatDate(dateObj, timeZone, "MMMM d, yyyy");
        }
      }

      // Perform update if display string doesn't match target "Month day, year" format
      if (formattedDateString && displayText !== formattedDateString) {
        const cell = sheet.getRange(startRow + r, startCol + c);
        
        // Prefixing with a single quote (') forces Google Sheets to treat the value 
        // strictly as text, breaking out of default short-date display formatting
        cell.setValue("'" + formattedDateString);
        formattedCount++;
      }
    }
  }

  return formattedCount;
}