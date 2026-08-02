/**
 * -----------------------------------------------------------------------------
 * ORCHID INVENTORY - STOPLIGHT UTILITIES (Stoplights.gs)
 * -----------------------------------------------------------------------------
 * Purpose: UI Logic to color-code the Inventory and individual sheets based 
 * on maintenance proximity (Red = Overdue, Amber = <30 days, Green = OK).
 * -----------------------------------------------------------------------------
 */

/**
 * Main Entry Point: Refreshes every orchid in the system.
 */
function applyAllSystemStoplights() {
  const ss = SpreadsheetApp.openById(MASTER_ID);
  const inv = ss.getSheetByName(INVENTORY_SHEET);
  if (!inv) return;

  const data = inv.getDataRange().getValues();
  const today = new Date();
  let count = 0;

  for (let i = 1; i < data.length; i++) {
    const row = i + 1;
    const id = data[i][ID_COL];
    
    // Skip hidden/archived rows or rows without IDs
    if (inv.isRowHiddenByUser(row) || !id) continue;

    const nextDue = data[i][NEXT_COL];
    let color = "#10b981"; // Default Green (Emerald 500)

    if (nextDue instanceof Date && !isNaN(nextDue.getTime())) {
      const diffDays = (nextDue - today) / (1000 * 60 * 60 * 24);

      if (diffDays < 0) {
        color = "#ef4444"; // Red (Overdue)
      } else if (diffDays < 30) {
        color = "#f59e0b"; // Amber (Warning)
      }
    }

    // Apply color to the ID cell in Inventory for quick visual reference
    inv.getRange(row, ID_COL + 1).setBackground(color).setFontColor("white");

    // Push the same color-coding to the individual orchid ledger if it exists
    const orchidSheet = ss.getSheetByName(String(id));
    if (orchidSheet) {
      applyIndividualStoplight_(orchidSheet, color);
      count++;
    }
  }
  
  try {
    SpreadsheetApp.getUi().alert(`Stoplight Sync Complete: ${count} ledgers updated.`);
  } catch(e) {
    console.log(`Updated ${count} stoplights.`);
  }
}

/**
 * INTERNAL HELPER: Updates repotting visual indicator on a single plant ledger.
 * Targets ONLY the "Next Repot Due" cell D10.
 */
function applyIndividualStoplight_(sheet, color) {
  // Target only Repot Date (Cell D10) - DO NOT touch D5 (Bloom Status)
  sheet.getRange("D10").setBackground(color).setFontColor("white");
}