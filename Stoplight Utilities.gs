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

    // TARGET SHIFTED COLUMN: NEXT_COL is now Index 20 (Column U)
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
 * INTERNAL HELPER: Updates the visual indicators on a single plant ledger.
 * Targets the status header and the "Next Due" cell D10.
 */
function applyIndividualStoplight_(sheet, color) {
  // Update the Status Background (Usually Row 5 or the Header)
  sheet.getRange("D5").setBackground(color).setFontColor("white");
  
  // Update the Repot Date Background
  sheet.getRange("D10").setBackground(color).setFontColor("white");
}

/**
 * AUTOMATED FORMATTING: Handles the bloom icons and conditional colors.
 * Called by onEdit and Provisioning.
 */
function applyBloomStatusFormatting(sheet) {
  const statusRange = sheet.getRange("D5");
  const status = statusRange.getValue().toString().toLowerCase();
  
  if (status.includes("in bloom")) {
    statusRange.setValue("🌸 In Bloom");
  } else if (status.includes("bud") || status.includes("spike")) {
    statusRange.setValue("✨ " + status.charAt(0).toUpperCase() + status.slice(1));
  } else if (status.includes("alive") || status.includes("dormant")) {
    // Keeps current text but ensures formatting is clean
    statusRange.setFontWeight("bold").setHorizontalAlignment("center");
  }
}

/**
 * UTILITY: Refresh all bloom icons across the entire collection.
 */
function refreshAllBloomStatusFormatting() {
  const ss = SpreadsheetApp.openById(MASTER_ID);
  const sheets = ss.getSheets();
  
  sheets.forEach(sheet => {
    const name = sheet.getName().trim();
    if (/^\d+$/.test(name)) {
      applyBloomStatusFormatting(sheet);
    }
  });
}