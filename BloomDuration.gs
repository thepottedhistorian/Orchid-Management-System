/**
 * ==============================================================================
 * ORCHID MANAGEMENT SYSTEM - BLOOM DURATION CALCULATOR (BloomDuration.gs)
 * ==============================================================================
 * Version: 1.0.0
 * Purpose: Scans Bloom Log rows (A20:B35) across Orchid ID sheets (1-65).
 *          Calculates exact duration between "In Bloom" (Col A) and 
 *          "Out of Bloom" (Col B) dates, writing the formatted text 
 *          (e.g., "4 months 14 days") into Column C.
 * 
 * Safely handles:
 * - Hyphens/dashes ("-", "—") or blank values in either column.
 * - Active blooms (where Column B is blank or "-").
 * ==============================================================================
 */

/**
 * Recalculates and populates bloom durations across all Orchid ID sheets (1-65).
 */
function updateAllBloomDurations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let totalRowsUpdated = 0;
  let processedSheets = 0;

  sheets.forEach(sheet => {
    const sheetName = sheet.getName().trim();

    // Target only numerical Orchid ID sheets (1 to 65)
    if (/^\d+$/.test(sheetName)) {
      const updatedOnSheet = updateSheetBloomDurations(sheet);
      if (updatedOnSheet > 0) {
        totalRowsUpdated += updatedOnSheet;
      }
      processedSheets++;
    }
  });

  const msg = `Bloom Duration Sync Complete!\nProcessed ${processedSheets} ID sheets. Updated ${totalRowsUpdated} duration records in Column C.`;
  Logger.log(msg);
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    // Graceful fallback for automated trigger context
  }
}

/**
 * Updates Column C bloom durations for a single specimen worksheet.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The target sheet instance.
 * @return {number} Count of duration cells updated.
 */
function updateSheetBloomDurations(sheet) {
  // Bloom Log starts at Row 20 and spans through Row 35
  const startRow = 20;
  const numRows = 16; 
  const range = sheet.getRange(startRow, 1, numRows, 2); // Pull Cols A & B
  const values = range.getValues();
  const displayValues = range.getDisplayValues();

  let updatedCount = 0;

  for (let i = 0; i < values.length; i++) {
    const targetRow = startRow + i;
    const startDateVal = values[i][0];
    const endDateVal = values[i][1];

    const startStr = String(displayValues[i][0]).replace(/\u00A0/g, ' ').trim();
    const endStr = String(displayValues[i][1]).replace(/\u00A0/g, ' ').trim();

    // Calculate duration string
    const durationText = computeDurationString(startDateVal, endDateVal, startStr, endStr);

    // Update Column C (Column index 3)
    const cCell = sheet.getRange(targetRow, 3);
    const currentCText = cCell.getDisplayValue().trim();

    if (currentCText !== durationText) {
      cCell.setValue(durationText);
      cCell.setHorizontalAlignment("center");
      updatedCount++;
    }
  }

  return updatedCount;
}

/**
 * Helper function: Computes difference in months and days between two date entries.
 * Returns an empty string "" if dates are missing, blank, or invalid (clears out stray dashes).
 */
function computeDurationString(startDateVal, endDateVal, startStr, endStr) {
  // 1. Return empty string if start or end date is missing, blank, "Unknown", or a placeholder hyphen
  const invalidTokens = ['', '-', '—', 'null', 'unknown'];
  if (!startStr || invalidTokens.includes(startStr.toLowerCase())) {
    return "";
  }
  if (!endStr || invalidTokens.includes(endStr.toLowerCase())) {
    return "";
  }

  // 2. Parse Date objects and normalize to midnight UTC to prevent timezone drift
  let d1 = parseToUtcDate(startDateVal, startStr);
  let d2 = parseToUtcDate(endDateVal, endStr);

  if (!d1 || !d2 || d2 < d1) {
    return "";
  }

  // 3. Calculate year, month, and day differences
  let years = d2.getUTCFullYear() - d1.getUTCFullYear();
  let months = d2.getUTCMonth() - d1.getUTCMonth();
  let days = d2.getUTCDate() - d1.getUTCDate();

  if (days < 0) {
    months -= 1;
    // Get total days in the previous month of d2 (UTC)
    const prevMonthLastDay = new Date(Date.UTC(d2.getUTCFullYear(), d2.getUTCMonth(), 0)).getUTCDate();
    days += prevMonthLastDay;
  }

  if (months < 0) {
    years -= 1;
    months += 12;
  }

  const totalMonths = (years * 12) + months;

  // 4. Build clean output string
  const parts = [];
  if (totalMonths > 0) {
    parts.push(totalMonths + (totalMonths === 1 ? " month" : " months"));
  }
  if (days > 0) {
    parts.push(days + (days === 1 ? " day" : " days"));
  }

  return parts.length > 0 ? parts.join(" ") : "0 days";
}

/**
 * Helper: Normalizes native JS Dates or date text strings to UTC Midnight.
 */
function parseToUtcDate(dateVal, dateStr) {
  if (dateVal instanceof Date && !isNaN(dateVal.getTime())) {
    return new Date(Date.UTC(dateVal.getFullYear(), dateVal.getMonth(), dateVal.getDate()));
  }
  
  const parsed = Date.parse(dateStr);
  if (!isNaN(parsed)) {
    const d = new Date(parsed);
    return new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  }
  
  return null;
}