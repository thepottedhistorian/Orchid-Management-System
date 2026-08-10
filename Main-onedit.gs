/**
 * -----------------------------------------------------------------------------
 * ORCHID MANAGEMENT SYSTEM - CENTRAL EVENT HANDLER (main-onEdit.gs)
 * -----------------------------------------------------------------------------
 * Version: 11.1.0 - Fixed Column L Auto-Timestamp & Single-Row Bloom Logging
 * -----------------------------------------------------------------------------
 */

function forceTriggerLink() {
  const ss = SpreadsheetApp.openById(MASTER_ID);
  console.log("Connection verified to: " + ss.getName());
}

function onEdit(e) {
  if (!e) return;
  const ss = e.source;
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName().trim();
  
  const startRow = range.getRow();
  const numRows = range.getNumRows();
  const startCol = range.getColumn();
  const lastCol = range.getLastColumn();

  // 1. MAINTENANCE LOG SYNC
  if (sheetName === MAINT_LOG_SHEET) {
    const ML_ID_COL = 1; 
    const ML_METHOD_COL = 3;
    const ML_DATE_COL = 4;
    const ML_NOTES_COL = 5;
    const ML_REPOT_DATE = 8;
    const ML_DUE_COL = 7;

    for (let i = 0; i < numRows; i++) {
      const currentRow = startRow + i;
      if (currentRow < 2) continue; 

      const orchidID = sheet.getRange(currentRow, ML_ID_COL).getValue();
      if (!orchidID) continue;
      const orchidSheet = ss.getSheetByName(String(orchidID));

      if (startCol <= ML_NOTES_COL && lastCol >= ML_METHOD_COL) {
        const methodVal = sheet.getRange(currentRow, ML_METHOD_COL).getValue();
        const notesVal = sheet.getRange(currentRow, ML_NOTES_COL).getValue();
        
        if (methodVal) {
          const now = new Date();
          sheet.getRange(currentRow, ML_DATE_COL).setValue(now);
          
          if (orchidSheet) {
            let targetRow = orchidSheet.getRange("A100").getNextDataCell(SpreadsheetApp.Direction.UP).getRow() + 1;
            if (targetRow < 56) targetRow = 56;
            
            orchidSheet.getRange(targetRow, 1).setValue(now);       // Date
            orchidSheet.getRange(targetRow, 2).setValue(methodVal); // Method
            orchidSheet.getRange(targetRow, 3).setValue(notesVal);  // Notes
          }
        }
      }

      if (startCol <= 9 && lastCol >= ML_REPOT_DATE) {
        const repotDate = sheet.getRange(currentRow, ML_REPOT_DATE).getValue();
        if (repotDate instanceof Date && orchidSheet) {
          orchidSheet.getRange("A38").setValue(repotDate);
          const invSheet = ss.getSheetByName(INVENTORY_SHEET);
          const invData = invSheet.getDataRange().getValues();
          
          for (let j = 1; j < invData.length; j++) {
            if (invData[j][ID_COL] == orchidID) {
              const freqRaw = invData[j][FREQ_COL];
              const parsedMonths = (typeof parseFrequencyToMonths === 'function')
                ? parseFrequencyToMonths(freqRaw)
                : null;

              // A blank or nonnumeric interval must never be treated as zero
              // months.  `Date#setMonth(month + null)` leaves the date on the
              // repot day, which was the cause of same-day due dates.
              // Use the application's established one-year fallback instead.
              const months = (typeof parsedMonths === 'number' && parsedMonths > 0)
                ? parsedMonths
                : 12;
              const nextDate = new Date(repotDate);
              nextDate.setMonth(nextDate.getMonth() + months);
              
              sheet.getRange(currentRow, ML_DUE_COL).setValue(nextDate);
              invSheet.getRange(j + 1, NEXT_COL + 1).setValue(nextDate);
              orchidSheet.getRange("D10").setValue(nextDate); 
              break;
            }
          }
        }
      }
    }
  }

  // 2. INVENTORY MASTER SYNC
  if (sheetName === INVENTORY_SHEET) {
    for (let i = 0; i < numRows; i++) {
      const currentRow = startRow + i;
      const orchidID = sheet.getRange(currentRow, ID_COL + 1).getValue();
      if (!orchidID) continue;
      const orchidSheet = ss.getSheetByName(String(orchidID));

      if (startCol === BLOOM_STATUS_COL + 1 && orchidSheet) {
        const rawValue = range.getValue().toString();
        orchidSheet.getRange("D5").setValue(rawValue); // Update Side Card (Bloom Status)
        
        const cleanVal = rawValue.replace(/[^\w\s]/gi, '').trim().toLowerCase();
        const now = new Date();
        
        if (cleanVal === "in bloom") {
          // 2a. Update Column L (Last Bloom Date) on Inventory sheet immediately
          sheet.getRange(currentRow, 12).setValue(now);
          
          // 2b. Log the bloom start using current timestamp
          logBloomEvent_(orchidSheet, "start", now);
        } 
        else if (cleanVal === "not in bloom") {
          const inventoryBloomDate = sheet.getRange(currentRow, 12).getValue();
          const targetEndDate = (inventoryBloomDate instanceof Date && !isNaN(inventoryBloomDate.getTime())) ? inventoryBloomDate : now;
          logBloomEvent_(orchidSheet, "end", targetEndDate);
        }
        
        if (typeof applyBloomStatusFormatting === 'function') {
          applyBloomStatusFormatting(orchidSheet);
        }
      }
    }
  }
}

/**
 * Helper: Record bloom history on individual sheet using explicit Inventory dates.
 * PAIRING LOGIC: Auto-closes open bloom cycles and calculates duration in Column C.
 */
function logBloomEvent_(orchidSheet, type, targetDate) {
  const logRange = orchidSheet.getRange("A20:C33");
  const logData = logRange.getValues(); 
  let openRowIndices = [];
  let emptyRowIndex = -1;

  for (let i = 0; i < logData.length; i++) {
    const hasStart = logData[i][0] !== "" && logData[i][0] !== null;
    const hasEnd = logData[i][1] !== "" && logData[i][1] !== null;

    if (hasStart && !hasEnd) {
      openRowIndices.push(i);
    }
    if (!hasStart && !hasEnd && emptyRowIndex === -1) {
      emptyRowIndex = i;
    }
  }

  const dateToApply = (targetDate instanceof Date && !isNaN(targetDate.getTime())) ? targetDate : new Date();

  if (type === "start") {
    // Check if the latest row is already logged for today to prevent duplicates
    const lastRowIndex = emptyRowIndex > 0 ? emptyRowIndex - 1 : (logData.length - 1);
    const existingDate = logData[lastRowIndex][0];
    if (existingDate instanceof Date && existingDate.toDateString() === dateToApply.toDateString() && logData[lastRowIndex][1] === "") {
      return; // Already logged today, do nothing
    }

    if (emptyRowIndex !== -1) {
      orchidSheet.getRange(20 + emptyRowIndex, 1).setValue(dateToApply);
    }
  } 
  else if (type === "end") {
    // Close all unclosed cycles
    if (openRowIndices.length > 0) {
      openRowIndices.forEach(idx => {
        const startDate = logData[idx][0] instanceof Date ? logData[idx][0] : new Date(logData[idx][0]);
        orchidSheet.getRange(20 + idx, 2).setValue(dateToApply);
        const durationStr = calculateDurationString_(startDate, dateToApply);
        if (durationStr) orchidSheet.getRange(20 + idx, 3).setValue(durationStr);
      });
    }
  }
}

/**
 * Helper: Formats duration between two Date objects into human-readable text.
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

/**
 * Automatically recalculates bloom durations when an "Out of Bloom" 
 * date is edited in Column B (Rows 20-35) on an ID sheet.
 */
function checkBloomDurationOnEdit(e) {
  if (!e || !e.range) return;
  
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName().trim();

  // Check if edit happened on a numerical ID sheet (1-65), in Column B, Row 20+
  if (/^\d+$/.test(sheetName) && range.getColumn() === 2 && range.getRow() >= 20) {
    updateSheetBloomDurations(sheet);
  }
}
