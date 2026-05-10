/**
 * -----------------------------------------------------------------------------
 * ORCHID INVENTORY - BLOOM VISUAL UTILITIES (Bloom Utilities.gs)
 * -----------------------------------------------------------------------------
 * Purpose: Applies conditional formatting and iconography (🌸, 🌱, 🌼) to 
 * individual orchid sheets based on their current bloom status.
 * -----------------------------------------------------------------------------
 */

/**
 * Updates cell D5 on a specific orchid sheet with icons and background colors.
 */
function applyBloomStatusFormatting(sheet) {
  const cell = sheet.getRange("D5");
  let val = cell.getValue().toString().trim();

  // Remove existing icons if they are already there to prevent "🌸 🌸 Status"
  let cleanStatus = val.replace(/^[^a-zA-Z0-9]+/, "").trim();
  const statusLower = cleanStatus.toLowerCase();

  let icon = "⬜"; 
  let color = "#E0E0E0"; // Default Grey

  if (statusLower.includes("not in bloom")) {
    icon = "⬜"; color = "#E0E0E0";
  } else if (statusLower.includes("in spike") || statusLower.includes("spiking")) {
    icon = "🌱"; color = "#B6E3A8"; // Soft Green
  } else if (statusLower.includes("in bud")) {
    icon = "🌼"; color = "#D5B4FF"; // Soft Purple
  } else if (statusLower.includes("in bloom")) {
    icon = "🌸"; color = "#A7C7FF"; // Soft Blue
  }

  cell.setBackground(color);
  cell.setFontWeight("bold");
  cell.setHorizontalAlignment("center");
  
  const newString = `${icon} ${cleanStatus}`;
  
  // Only update if the value actually changed to avoid infinite loops
  if (val !== newString) {
    cell.setValue(newString);
  }
}

/**
 * Batch utility to refresh formatting across every orchid in the collection.
 */
function refreshAllBloomStatusFormatting() {
  // Pulls MASTER_ID automatically from Globals.gs
  const ss = SpreadsheetApp.openById(MASTER_ID);
  const sheets = ss.getSheets();
  
  sheets.forEach(sheet => {
    const name = sheet.getName().trim();
    // Only run on sheets named with numbers (Orchid IDs)
    if (/^\d+$/.test(name)) {
      applyBloomStatusFormatting(sheet);
    }
  });
}