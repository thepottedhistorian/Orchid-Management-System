/**
 * -----------------------------------------------------------------------------
 * ORCHID INVENTORY - BATCH REPOT ENGINE (batch-tools.gs)
 * -----------------------------------------------------------------------------
 * Handles:
 * 1. Base date selection (Last Repot Date from Log/A38 over Acquisition Date).
 * 2. Repot frequency calculation using the UPPER END (e.g., 1-2 years = 2 years).
 * 3. Triple-sync to Inventory [Col U], Maintenance Log [Col G], and ID Sheet [D10].
 * -----------------------------------------------------------------------------
 */

/**
 * 🔄 MASTER RUNNER: Recalculates and synchronizes all repot dates collection-wide.
 */
function recalcAllPlantRepotDueDates() {
  const ss = SpreadsheetApp.openById(MASTER_ID);
  const inv = ss.getSheetByName(INVENTORY_SHEET);
  const log = ss.getSheetByName(MAINT_LOG_SHEET);
  if (!inv) return;

  const invData = inv.getDataRange().getValues();
  let updatedCount = 0;

  for (let i = 1; i < invData.length; i++) {
    const id = invData[i][ID_COL]; 
    if (!id || id.toString().trim() === "" || isNaN(id)) continue;

    const acqRaw = invData[i][ACQ_COL];             // Column I (Acquisition)
    const freqRaw = invData[i][FREQ_COL];           // Column T (Frequency)
    const freqMonths = parseFrequencyToMonths(freqRaw); // Takes upper range number

    // 1. Determine base date: Last Repot Date takes priority over Acquisition Date
    const lastRepot = getLastRepotDateForID(id);
    const acqDate = (acqRaw instanceof Date) ? acqRaw : new Date(acqRaw);
    
    // 2. Calculate next repot due date
    const nextDate = calculateNextRepotDate_(acqDate, lastRepot, freqMonths);

    if (nextDate) {
      // Step A: Update Inventory Column U (Next Repot Due)
      inv.getRange(i + 1, NEXT_COL + 1).setValue(nextDate); 

      // Step B: Update Maintenance Log Column G
      if (log) {
        const logVals = log.getDataRange().getValues();
        for (let r = logVals.length - 1; r >= 1; r--) {
          if (logVals[r][0] == id) {
            log.getRange(r + 1, 7).setValue(nextDate); // Column G
            break; 
          }
        }
      }

      // Step C: Update Individual Orchid ID Sheet Cell D10
      const orchidSheet = ss.getSheetByName(String(id));
      if (orchidSheet) {
        const cellD10 = orchidSheet.getRange("D10");
        cellD10.clearContent();
        cellD10.setValue(nextDate);
        cellD10.setNumberFormat('MMMM d, yyyy');
        cellD10.setHorizontalAlignment("center");
        cellD10.setVerticalAlignment("middle");
      }
      updatedCount++;
    }
  }

  try {
    SpreadsheetApp.getUi().alert(`Success: Recalculated repot dates for ${updatedCount} orchids across Inventory [U], Maintenance Log [G], and ID Sheets [D10].`);
  } catch (e) {
    console.log(`Recalculation Complete: ${updatedCount} records updated.`);
  }
}

/**
 * STEP 1 FUNCTION: Syncs A38 from ID sheets into Maintenance Log Column H.
 */
function transferA38ToLogColumnH() {
  const ss = SpreadsheetApp.openById(MASTER_ID);
  const log = ss.getSheetByName(MAINT_LOG_SHEET);
  if (!log) return;

  const logData = log.getDataRange().getValues();
  let updatedCount = 0;

  for (let i = 1; i < logData.length; i++) {
    const id = logData[i][0];
    if (!id || isNaN(id)) continue;

    const lastRepot = getLastRepotDateForID(id);
    if (lastRepot instanceof Date && !isNaN(lastRepot.getTime())) {
      log.getRange(i + 1, 8).setValue(lastRepot); // Column H (Index 8)
      updatedCount++;
    }
  }

  try {
    SpreadsheetApp.getUi().alert(`Step 1 Complete: Transferred last repot dates to Column H for ${updatedCount} orchids.`);
  } catch (e) {
    console.log(`Step 1 Complete: Transferred dates for ${updatedCount} orchids.`);
  }
}

/**
 * STEP 2 FUNCTION: Recalculates Maintenance Log Column G.
 */
function calculateOnlyMaintenanceLogG() {
  recalcAllPlantRepotDueDates(); // Invokes full calculation loop
}

/**
 * 🔍 HELPER: GET LAST REPOT DATE FOR ID
 * Checks individual sheet A38:A55 first, falls back to Maintenance Log Column H.
 */
function getLastRepotDateForID(id) {
  const ss = SpreadsheetApp.openById(MASTER_ID);
  let latest = null;

  // 1. Check Individual Orchid Sheet (A38:A55)
  const orchidSheet = ss.getSheetByName(String(id));
  if (orchidSheet) {
    const sheetData = orchidSheet.getRange("A38:A55").getValues();
    for (let j = 0; j < sheetData.length; j++) {
      let d = sheetData[j][0];
      if (d && d instanceof Date && !isNaN(d.getTime())) {
        if (!latest || d > latest) latest = d;
      }
    }
  }

  // 2. Fall back to Maintenance Log Column H
  if (!latest) {
    const log = ss.getSheetByName(MAINT_LOG_SHEET);
    if (log) {
      const logData = log.getDataRange().getValues();
      for (let k = 1; k < logData.length; k++) {
        if (logData[k][0] == id && logData[k][7]) {
          let d = logData[k][7];
          if (!(d instanceof Date)) d = new Date(d);
          if (!isNaN(d.getTime())) {
            if (!latest || d > latest) latest = d;
          }
        }
      }
    }
  }
  return latest;
}

/**
 * 🧮 INTERNAL MATH: CALCULATE NEXT REPOT DATE
 * Uses Last Repot Date if present; otherwise defaults to Acquisition Date.
 */
function calculateNextRepotDate_(acqDate, lastRepotDate, freqMonths) {
  const base = lastRepotDate ? new Date(lastRepotDate) : new Date(acqDate);
  if (isNaN(base.getTime()) || !freqMonths) return null;
  const next = new Date(base);
  next.setMonth(next.getMonth() + freqMonths);
  return next;
}

/**
 * 🛠️ UTILITY: PARSE FREQUENCY TO MONTHS (UPPER RANGE HIGHEST NUMBER)
 * Extracts the HIGHEST number found in a string:
 * - "1-2 years"  -> 2 years  -> 24 months
 * - "12-18 months" -> 18 months -> 18 months
 * - "1 year"     -> 1 year   -> 12 months
 */
function parseFrequencyToMonths(freqRaw) {
  if (!freqRaw) return null;
  const text = freqRaw.toString().toLowerCase().trim();
  
  const matches = text.match(/(\d+(\.\d+)?)/g);
  if (!matches) return null;
  
  // Takes the last (highest) number matched in the range expression
  const num = parseFloat(matches[matches.length - 1]);
  
  return text.includes("year") ? num * 12 : num;
}