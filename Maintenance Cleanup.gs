/**
 * -----------------------------------------------------------------------------
 * ORCHID INVENTORY - MAINTENANCE CLEANUP (Maintenance.gs)
 * -----------------------------------------------------------------------------
 * Purpose: Synchronization utility to back-fill the Maintenance Log with 
 * repotting dates stored on individual orchid ledger sheets (Cell A38).
 * Logic:
 * 1. Scans all numeric orchid sheets for a date in cell A38.
 * 2. Locates the most recent entry for that ID in the 'Maintenance Log'.
 * 3. Overwrites Column H (Repotted On) with the A38 date.
 * -----------------------------------------------------------------------------
 */

/**
 * Syncs individual ledger baseline dates (A38) back to the master Maintenance Log.
 */
function cleanDuplicateOrchidEntriesByDate() {
  const ss = SpreadsheetApp.openById(MASTER_ID);
  const sheets = ss.getSheets();
  let totalDeleted = 0;

  sheets.forEach(sheet => {
    const sheetName = sheet.getName().trim();
    
    // Only run on sheets named with numbers (Orchid IDs)
    if (/^\d+$/.test(sheetName)) {
      const lastRow = sheet.getLastRow();
      
      // We need at least two rows of data (56 and 57) to have a potential duplicate
      if (lastRow < 57) return; 

      // .getDisplayValues() gets the text exactly as formatted in the sheet
      // Range starts at row 56, column 1, height is lastRow - 55, width is 3 (A, B, C)
      const range = sheet.getRange(56, 1, lastRow - 55, 3);
      const displayValues = range.getDisplayValues(); 
      
      /**
       * PROCESS BOTTOM-UP
       * Essential for row deletion logic so the loop index doesn't get offset.
       */
      for (let i = displayValues.length - 1; i > 0; i--) {
        const currentRow = displayValues[i];
        const previousRow = displayValues[i - 1];

        // Check Column A (Date Text) and Column B (Method Text)
        const dateMatch = currentRow[0] === previousRow[0];
        const methodMatch = currentRow[1] === previousRow[1];

        /**
         * DUPLICATE CRITERIA:
         * If Date and Method are identical to the row above, delete the current row.
         * We ignore Column C (Notes) to catch cases where one note might be missing 
         * but the event is clearly a duplicate.
         */
        if (dateMatch && methodMatch && currentRow[0] !== "") {
          sheet.deleteRow(i + 56);
          totalDeleted++;
        }
      }
    }
  });

  // UI Notification for the administrator
  try {
    SpreadsheetApp.getUi().alert(`Cleanup Complete: Removed ${totalDeleted} duplicate rows based on matching display text.`);
  } catch (e) {
    console.log(`Cleanup Complete: Removed ${totalDeleted} duplicate rows.`);
  }
}