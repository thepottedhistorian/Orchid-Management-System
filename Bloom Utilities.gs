/**
 * -----------------------------------------------------------------------------
 * ORCHID INVENTORY - BLOOM VISUAL UTILITIES (Bloom Utilities.gs)
 * -----------------------------------------------------------------------------
 * Purpose: Applies conditional formatting and iconography (🌸, 🌱, 🌼, ⬜) to 
 * individual orchid sheets based on their current bloom status.
 * -----------------------------------------------------------------------------
 */

/**
 * Updates cell D5 on a specific orchid sheet with icons, text colors, and background colors.
 */
function applyBloomStatusFormatting(sheet) {
  const cell = sheet.getRange("D5");
  let val = cell.getValue().toString().trim();

  // Strip existing emojis/icons to prevent "🌸 🌸 Status" accumulation
  let cleanStatus = val.replace(/^[^a-zA-Z0-9]+/, "").trim();
  const statusLower = cleanStatus.toLowerCase();

  let icon = "⬜"; 
  let bgColor = "#fca5a5"; // Lighter Red for "Not in Bloom"
  let fontColor = "#ffffff"; // White text for clean contrast

  if (statusLower.includes("not in bloom")) {
    icon = "⬜"; 
    bgColor = "#fca5a5"; // Soft / Light Red
    fontColor = "#ffffff";
  } else if (statusLower.includes("in spike") || statusLower.includes("spiking")) {
    icon = "🌱"; 
    bgColor = "#bbf7d0"; // Soft Light Green
    fontColor = "#065f46";
  } else if (statusLower.includes("in bud")) {
    icon = "🌼"; 
    bgColor = "#e9d5ff"; // Soft Purple
    fontColor = "#581c87";
  } else if (statusLower.includes("in bloom")) {
    icon = "🌸"; 
    bgColor = "#10b981"; // Vibrant Emerald Green
    fontColor = "#ffffff";
  }

  cell.setBackground(bgColor);
  cell.setFontColor(fontColor);
  cell.setFontWeight("bold");
  cell.setHorizontalAlignment("center");
  
  const newString = `${icon} ${cleanStatus}`;
  
  if (val !== newString) {
    cell.setValue(newString);
  }
}

/**
 * Batch utility to refresh formatting across every orchid in the collection.
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