/**
 * -----------------------------------------------------------------------------
 * ORCHID MANAGEMENT SYSTEM - CENTRAL EVENT HANDLER (main-onEdit.gs)
 * -----------------------------------------------------------------------------
 * Updated: 2026-04-17 - Finalized Row Pairing & Sequence Logic.
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
    // REPLACEMENT LOGIC START:
    // We target Column A starting at your maintenance ledger (Row 56).
    // This finds the last entry in that specific block and moves down one.
    let targetRow = orchidSheet.getRange("A100").getNextDataCell(SpreadsheetApp.Direction.UP).getRow() + 1;
    
    // Safety: If the ledger is totally empty, start at row 56.
    if (targetRow < 56) {
      targetRow = 56;
    }
    
    // Write the data to the Individual Orchid Sheet
    orchidSheet.getRange(targetRow, 1).setValue(now);       // Date
    orchidSheet.getRange(targetRow, 2).setValue(methodVal); // Method (e.g., Watering)
    orchidSheet.getRange(targetRow, 3).setValue(notesVal);  // Notes
    // REPLACEMENT LOGIC END
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
              const months = (typeof parseFrequencyToMonths === 'function') ? parseFrequencyToMonths(freqRaw) : 12;
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
        orchidSheet.getRange("D5").setValue(rawValue);
        
        const cleanVal = rawValue.replace(/[^\w\s]/gi, '').trim().toLowerCase();
        
        if (cleanVal === "in bloom") {
          sheet.getRange(currentRow, LAST_BLOOM_COL + 1).setValue(new Date());
          logBloomEvent_(orchidSheet, "start");
        } 
        else if (cleanVal === "not in bloom") {
          logBloomEvent_(orchidSheet, "end");
        }
        
        if (typeof applyBloomStatusFormatting === 'function') {
          applyBloomStatusFormatting(orchidSheet);
        }
      }
    }
  }
}

/**
 * Helper: Record bloom history on individual sheet.
 * PAIRING LOGIC: Treats Row (A+B) as a single event.
 */
function logBloomEvent_(orchidSheet, type) {
  const logRange = orchidSheet.getRange("A20:B33");
  const logData = logRange.getValues(); 
  let targetRowIndex = -1;

  if (type === "start") {
    // 1. Check if the VERY last entry is an "open" bloom (Start date exists, End date is empty)
    // If so, we DON'T start a new row. We assume the user is toggling/correcting.
    for (let i = logData.length - 1; i >= 0; i--) {
      if (logData[i][0] && !logData[i][1]) {
        return; // Exit. An event is already open.
      }
    }

    // 2. Otherwise, find the first row that is COMPLETELY empty
    for (let i = 0; i < logData.length; i++) {
      if (!logData[i][0] && !logData[i][1]) { 
        targetRowIndex = i;
        break;
      }
    }
    if (targetRowIndex !== -1) {
      orchidSheet.getRange(20 + targetRowIndex, 1).setValue(new Date());
    }
  } 
  else if (type === "end") {
    // Find the LAST row that has a start (A) but is missing an end (B)
    for (let i = logData.length - 1; i >= 0; i--) {
      // We check if it's a valid date or a non-empty string in A, and truly empty in B
      if (logData[i][0] !== "" && logData[i][1] === "") { 
        targetRowIndex = i;
        break;
      }
    }
    
    if (targetRowIndex !== -1) {
      orchidSheet.getRange(20 + targetRowIndex, 2).setValue(new Date());
    }
  }
}