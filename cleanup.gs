/**
 * FORCE CLEANUP UTILITY (Dash-Aware Version):
 * Evaluates A20:B33 for active blooms. Ignores hyphens/dashes (-) in Column B 
 * so legacy records don't trigger false "In Bloom" statuses.
 */
function forceCleanupBloomStatuses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let corrected = [];

  sheets.forEach(sheet => {
    const sheetName = sheet.getName().trim();
    
    // Target only numerical Orchid ID sheets
    if (/^\d+$/.test(sheetName)) {
      const logValues = sheet.getRange("A20:B33").getValues();
      let hasActiveBloom = false;

      // Check rows in Bloom Log
      for (let i = 0; i < logValues.length; i++) {
        const startVal = String(logValues[i][0]).trim();
        const endVal = String(logValues[i][1]).trim();
        
        const hasStart = startVal !== "" && startVal !== "null" && startVal !== "-";
        
        // A row is CLOSED if endVal has a date, OR if endVal contains a dash/hyphen (-)
        const isClosed = endVal !== "" && endVal !== "null"; // Any character (date or dash) closes it!

        if (hasStart && !isClosed) {
          hasActiveBloom = true;
          break;
        }
      }

      // If no active bloom row exists in the log, force cell D5 to "Not in Bloom"
      if (!hasActiveBloom) {
        const d5Cell = sheet.getRange("D5");
        
        d5Cell.setValue("Not in Bloom");
        
        // Apply visual formatting (Lighter Red background + ⬜ icon + White text)
        d5Cell.setBackground("#fca5a5");
        d5Cell.setFontColor("#ffffff");
        d5Cell.setFontWeight("bold");
        d5Cell.setHorizontalAlignment("center");
        d5Cell.setValue("⬜ Not in Bloom");
        
        corrected.push(sheetName);
      }
    }
  });

  const msg = `Force cleanup finished. Reset status to 'Not in Bloom' (Light Red) on Orchid IDs: ${corrected.join(", ")}`;
  console.log(msg);
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {}
}