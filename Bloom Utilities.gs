/**
 * -----------------------------------------------------------------------------
 * ORCHID INVENTORY - BLOOM VISUAL & DURATION UTILITIES (Bloom Utilities.gs)
 * -----------------------------------------------------------------------------
 * Purpose: 
 * 1. Applies conditional formatting and iconography (🌸, 🌱, 🌼, ⬜) to cell D5.
 * 2. Recalculates and populates bloom cycle durations (Column C) across ID sheets.
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

/**
 * -----------------------------------------------------------------------------
 * BLOOM DURATION CALCULATORS
 * -----------------------------------------------------------------------------
 */

/**
 * Recalculates and populates bloom duration for a SINGLE Orchid ID sheet.
 * Checks rows 20 to 35 for "In Bloom" (Col A) and "Out of Bloom" (Col B) dates.
 */
function updateSheetBloomDurations(sheet) {
  if (!sheet) return;

  const range = sheet.getRange("A20:C35");
  const values = range.getValues();

  for (let i = 0; i < values.length; i++) {
    const rawStart = values[i][0]; // Column A
    const rawEnd = values[i][1];   // Column B

    if (rawStart && rawEnd) {
      const startDate = (rawStart instanceof Date) ? rawStart : new Date(rawStart);
      const endDate = (rawEnd instanceof Date) ? rawEnd : new Date(rawEnd);

      if (!isNaN(startDate.getTime()) && !isNaN(endDate.getTime())) {
        const durationStr = calculateDurationString_(startDate, endDate);
        sheet.getRange(20 + i, 3).setValue(durationStr); // Column C
      }
    }
  }
}

/**
 * MASTER BATCH RUNNER: Scans ALL Orchid ID sheets across the collection
 * and fixes/updates bloom durations in Column C.
 */
function fixAllMissingBloomDurations() {
  const ss = SpreadsheetApp.openById(MASTER_ID);
  const sheets = ss.getSheets();
  let updatedSheetsCount = 0;

  sheets.forEach(sheet => {
    const name = sheet.getName().trim();
    if (/^\d+$/.test(name)) {
      updateSheetBloomDurations(sheet);
      updatedSheetsCount++;
    }
  });

  try {
    SpreadsheetApp.getUi().alert(`Bloom Log Sync Complete: Audited and updated bloom durations across ${updatedSheetsCount} Orchid ID sheets.`);
  } catch (e) {
    console.log(`Bloom Log Sync Complete: Audited ${updatedSheetsCount} sheets.`);
  }
}

/**
 * HELPER MATH ENGINE: Formats duration between two Date objects into human-readable text.
 */
function calculateDurationString_(startDate, endDate) {
  if (!(startDate instanceof Date) || isNaN(startDate.getTime()) ||
      !(endDate instanceof Date) || isNaN(endDate.getTime()) || endDate < startDate) {
    return "";
  }

  let start = new Date(startDate.getTime());
  let end = new Date(endDate.getTime());

  let years = end.getFullYear() - start.getFullYear();
  let months = end.getMonth() - start.getMonth();
  let days = end.getDate() - start.getDate();

  if (days < 0) {
    months--;
    const prevMonthLastDay = new Date(end.getFullYear(), end.getMonth(), 0).getDate();
    days += prevMonthLastDay;
  }

  if (months < 0) {
    years--;
    months += 12;
  }

  let totalMonths = years * 12 + months;
  let parts = [];

  if (totalMonths > 0) parts.push(totalMonths + " month" + (totalMonths > 1 ? "s" : ""));
  if (days > 0) parts.push(days + " day" + (days > 1 ? "s" : ""));
  if (parts.length === 0) parts.push("0 days");

  return parts.join(" ");
}