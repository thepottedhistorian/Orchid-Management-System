/**
 * ------------------------------------------------------------
 * ORCHID INVENTORY - DATA DELIVERY & CLEANUP (Delivery.gs)
 * ------------------------------------------------------------
 * Purpose: Finalizes the repotting sync by pushing calculated 
 * dates from the Master Inventory to individual plant sheets.
 * ------------------------------------------------------------
 */

/**
 * DELIVERY SCRIPT: Inventory [U] -> Orchid ID Sheet [D10]
 * Pushes the "Next Repot Due" date to the specific ledger sheet.
 * Updated: 2026-03-30 to use shifted NEXT_COL (Index 20).
 */
function pushInventoryTToOrchidD10() {
  const ss = SpreadsheetApp.openById(MASTER_ID);
  const inv = ss.getSheetByName(INVENTORY_SHEET);
  if (!inv) return;

  const invData = inv.getDataRange().getValues();
  let updateCount = 0;

  for (let i = 1; i < invData.length; i++) {
    const id = invData[i][ID_COL];      // Column A (Index 0)
    const nextDue = invData[i][NEXT_COL]; // Now Column U (Index 20)

    if (id && nextDue instanceof Date && !isNaN(nextDue.getTime())) {
      const orchidSheet = ss.getSheetByName(String(id));
      if (orchidSheet) {
        const cellD10 = orchidSheet.getRange("D10");
        
        // Clear old content and set the new standardized date
        cellD10.clearContent();
        cellD10.setValue(nextDue);
        cellD10.setNumberFormat('MMMM d, yyyy');
        cellD10.setHorizontalAlignment("center");
        cellD10.setVerticalAlignment("middle");
        
        updateCount++;
        console.log(`Pushed Date to ID ${id}: ${nextDue}`);
      }
    }
  }

  try {
    SpreadsheetApp.getUi().alert(`Success: ${updateCount} orchid sheets updated. D10 now matches Inventory Column U.`);
  } catch (e) {
    console.log("Push complete. " + updateCount + " sheets updated.");
  }
}

/**
 * CLEANUP: Wipes D11 on all Orchid ID sheets.
 * Use this to clear out old legacy formatting or "ghost" data.
 */
function clearAllD11s() {
  const ss = SpreadsheetApp.openById(MASTER_ID);
  const sheets = ss.getSheets();
  let count = 0;

  sheets.forEach(sheet => {
    const name = sheet.getName().trim();
    // Only target numbered orchid sheets (e.g., "101")
    if (/^\d+$/.test(name)) {
      sheet.getRange("D11").clearContent().clearFormat();
      count++;
    }
  });

  try {
    SpreadsheetApp.getUi().alert(`Cleaned D11 on ${count} orchid sheets.`);
  } catch (e) {
    console.log("Cleanup complete. " + count + " sheets cleared.");
  }
}