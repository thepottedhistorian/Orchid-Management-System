/**
 * -----------------------------------------------------------------------------
 * ORCHID INVENTORY - MASTER DATA AGGREGATION (Orchid Tracker.gs)
 * -----------------------------------------------------------------------------
 * Purpose: Provides collection-wide statistics for the Master Inventory workbook.
 * These functions calculate the latest maintenance activity, total population,
 * and current blooming status.
 * -----------------------------------------------------------------------------
 */

/**
 * GET_MAIN_LATEST_WATERING
 * Iterates through all numbered ID sheets in the Master workbook to find the 
 * single most recent date recorded in the watering/maintenance logs.
 */
function GET_MAIN_LATEST_WATERING() {
  // MASTER_ID is pulled from Globals.gs
  const ss = SpreadsheetApp.openById(MASTER_ID);
  const sheets = ss.getSheets();
  let latestDate = new Date(0); 

  sheets.forEach(sheet => {
    const name = sheet.getName().trim();
    
    // Only process numeric ID sheets (Orchid ID Ledgers)
    if (/^\d+$/.test(name)) {
      const lastRow = sheet.getRange("A" + sheet.getMaxRows()).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
      
      // Standard maintenance logs begin at Row 55 (Template Baseline)
      if (lastRow >= 55) {
        const dateValue = sheet.getRange(lastRow, 1).getValue();
        const checkDate = new Date(dateValue);
        
        if (!isNaN(checkDate.getTime())) {
          if (checkDate > latestDate) {
            latestDate = checkDate;
          }
        }
      }
    }
  });

  return latestDate.getTime() === 0 ? "No Dates Found" : latestDate;
}

/**
 * COUNT_MAIN_ORCHIDS
 * Looks at Column B (NAME_COL + 1) of the "Inventory" sheet to count specimens.
 */
function COUNT_MAIN_ORCHIDS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inventorySheet = ss.getSheetByName(INVENTORY_SHEET); 
  
  if (!inventorySheet) return 0;
  
  const lastRow = inventorySheet.getLastRow();
  if (lastRow < 2) return 0;

  // Uses Column B (Index 2) as the source of truth for plant presence
  const names = inventorySheet.getRange(2, NAME_COL + 1, lastRow - 1, 1).getValues();
  
  let count = 0;
  for (let i = 0; i < names.length; i++) {
    // Counts any row where the Name field is not empty
    if (names[i][0].toString().trim() !== "") {
      count++;
    }
  }
  
  return count;
}

/**
 * COUNT_MAIN_BLOOMS
 * Scans the Status column (BLOOM_STATUS_COL) of the "Inventory" sheet.
 */
function COUNT_MAIN_BLOOMS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inventorySheet = ss.getSheetByName(INVENTORY_SHEET); 
  
  if (!inventorySheet) return 0;

  const lastRow = inventorySheet.getLastRow();
  if (lastRow < 2) return 0;

  // Uses BLOOM_STATUS_COL from Globals.gs (Index 10 / Column K)
  const statusValues = inventorySheet.getRange(2, BLOOM_STATUS_COL + 1, lastRow - 1, 1).getValues(); 
  
  let bloomCount = 0;
  
  statusValues.forEach(row => {
    const status = row[0].toString().toLowerCase().trim();
    // Strict check: only count if it specifically matches "in bloom"
    if (status === "in bloom") { 
      bloomCount++;
    }
  });

  return bloomCount;
}