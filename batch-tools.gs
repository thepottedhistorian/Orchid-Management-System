/**
 * -----------------------------------------------------------------------------
 * ORCHID INVENTORY - BATCH REPOT ENGINE (batch-tools.gs)
 * -----------------------------------------------------------------------------
 * Handles:
 * 1. Base date selection (Last Repot Date from Log/A38 over Acquisition Date).
 * 2. Repot frequency calculation using the UPPER END (e.g., 1-2 years = 2 years).
 * 3. Triple-sync to Inventory [Col U], Maintenance Log [Col G], and ID Sheet [D10].
 * 4. HTML Preview Modal Data Generation (previewRecalcRepotDates).
 * -----------------------------------------------------------------------------
 */

/**
 * Helper to reliably get active spreadsheet reference
 */
function getTargetSpreadsheet_() {
  if (typeof MASTER_ID !== 'undefined' && MASTER_ID) {
    try {
      return SpreadsheetApp.openById(MASTER_ID);
    } catch (e) {
      console.warn("Could not open by MASTER_ID, falling back to Active Spreadsheet.");
    }
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * 🖥️ UI LAUNCHER: Opens the Repot Date Preview Modal.
 */
function showRepotPreviewModal() {
  const htmlTemplate = HtmlService.createTemplateFromFile('repot-preview');
  htmlTemplate.previewData = previewRecalcRepotDates();
  
  const html = htmlTemplate.evaluate()
    .setWidth(900)
    .setHeight(650);
    
  SpreadsheetApp.getUi().showModalDialog(html, 'Repot Date Preview');
}

/**
 * 📊 PREVIEW DATA GENERATOR
 */
function previewRecalcRepotDates() {
  const ss = getTargetSpreadsheet_();
  const inv = ss.getSheetByName(INVENTORY_SHEET);
  if (!inv) throw new Error("Inventory sheet '" + INVENTORY_SHEET + "' not found.");

  const invData = inv.getDataRange().getValues();
  const previewData = [];

  for (let i = 1; i < invData.length; i++) {
    const rowNum = i + 1;
    const id = invData[i][ID_COL]; 
    if (!id || id.toString().trim() === "" || isNaN(id)) continue;

    const plantName = invData[i][NAME_COL] || "Unknown Orchid";
    const acqRaw = invData[i][ACQ_COL];
    const freqRaw = invData[i][FREQ_COL];
    const freqMonths = parseFrequencyToMonths(freqRaw) || 12;

    const lastRepot = getLastRepotDateForID(id);
    let acqDate = null;
    if (acqRaw instanceof Date && !isNaN(acqRaw.getTime())) {
      acqDate = acqRaw;
    } else if (acqRaw) {
      const parsed = new Date(acqRaw);
      if (!isNaN(parsed.getTime())) acqDate = parsed;
    }
    
    let baseDate = lastRepot || acqDate;
    let baseUsed = lastRepot ? "Last Repot" : (acqDate ? "Acquisition" : "None");

    const nextDate = calculateNextRepotDate_(acqDate, lastRepot, freqMonths);
    const status = nextDate ? "OK" : "MISSING BASE DATE";

    previewData.push({
      row: rowNum,
      id: id,
      name: plantName,
      acquisition: acqDate ? Utilities.formatDate(acqDate, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd") : "N/A",
      lastRepot: lastRepot ? Utilities.formatDate(lastRepot, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd") : "None",
      baseUsed: baseUsed,
      frequency: freqRaw || "1 Year",
      months: freqMonths,
      next: nextDate ? Utilities.formatDate(nextDate, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd") : "N/A",
      status: status
    });
  }

  return previewData;
}

/**
 * 🔄 MASTER RUNNER: Recalculates and synchronizes all repot dates.
 */
function recalcAllPlantRepotDueDates() {
  const ss = getTargetSpreadsheet_();
  const inv = ss.getSheetByName(INVENTORY_SHEET);
  const log = ss.getSheetByName(MAINT_LOG_SHEET);
  if (!inv) throw new Error("Inventory sheet '" + INVENTORY_SHEET + "' not found.");

  const invData = inv.getDataRange().getValues();
  let updatedCount = 0;

  for (let i = 1; i < invData.length; i++) {
    const id = invData[i][ID_COL]; 
    if (!id || id.toString().trim() === "" || isNaN(id)) continue;

    const acqRaw = invData[i][ACQ_COL];
    const freqRaw = invData[i][FREQ_COL];
    const freqMonths = parseFrequencyToMonths(freqRaw) || 12;

    const lastRepot = getLastRepotDateForID(id);
    let acqDate = null;
    if (acqRaw instanceof Date && !isNaN(acqRaw.getTime())) {
      acqDate = acqRaw;
    } else if (acqRaw) {
      const parsed = new Date(acqRaw);
      if (!isNaN(parsed.getTime())) acqDate = parsed;
    }
    
    const nextDate = calculateNextRepotDate_(acqDate, lastRepot, freqMonths);

    if (nextDate) {
      // Step A: Update Inventory Column U (Next Repot Due)
      inv.getRange(i + 1, NEXT_COL + 1).setValue(nextDate); 

      // Step B: Update Maintenance Log Column G
      if (log) {
        const logVals = log.getDataRange().getValues();
        for (let r = logVals.length - 1; r >= 1; r--) {
          if (logVals[r][0] == id) {
            log.getRange(r + 1, 7).setValue(nextDate);
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

  console.log(`Recalculation Complete: ${updatedCount} records updated.`);
  return updatedCount;
}

/**
 * 🔍 HELPER: GET LAST REPOT DATE FOR ID
 */
function getLastRepotDateForID(id) {
  const ss = getTargetSpreadsheet_();
  let latest = null;

  // 1. Check Individual Orchid Sheet (A38:A55)
  const orchidSheet = ss.getSheetByName(String(id));
  if (orchidSheet) {
    const sheetData = orchidSheet.getRange("A38:A55").getValues();
    for (let j = 0; j < sheetData.length; j++) {
      let d = sheetData[j][0];
      if (d) {
        if (!(d instanceof Date)) d = new Date(d);
        if (!isNaN(d.getTime())) {
          if (!latest || d > latest) latest = d;
        }
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
 */
function calculateNextRepotDate_(acqDate, lastRepotDate, freqMonths) {
  let base = null;
  
  if (lastRepotDate && lastRepotDate instanceof Date && !isNaN(lastRepotDate.getTime())) {
    base = new Date(lastRepotDate);
  } else if (acqDate && acqDate instanceof Date && !isNaN(acqDate.getTime())) {
    base = new Date(acqDate);
  }

  if (!base || isNaN(base.getTime()) || !freqMonths) return null;
  
  const next = new Date(base);
  next.setMonth(next.getMonth() + freqMonths);
  return next;
}

/**
 * 🛠️ UTILITY: PARSE FREQUENCY TO MONTHS
 */
function parseFrequencyToMonths(freqRaw) {
  if (!freqRaw) return null;
  const text = freqRaw.toString().toLowerCase().trim();
  
  const matches = text.match(/(\d+(\.\d+)?)/g);
  if (!matches) return null;
  
  const num = parseFloat(matches[matches.length - 1]);
  return text.includes("year") ? num * 12 : num;
}


/**
 * 🧪 DIAGNOSTIC TOOL: Run this directly from the Apps Script Editor toolbar.
 * Inspects row 67 (ID 66) and prints findings to the Execution Log below.
 */
function debugRow67() {
  const ss = getTargetSpreadsheet_();
  const inv = ss.getSheetByName(INVENTORY_SHEET);
  const invData = inv.getDataRange().getValues();
  
  const rowIndex = 66; // Row 67 in Google Sheets (0-indexed array)
  
  if (invData.length <= rowIndex) {
    console.log("❌ Row 67 was not found in the Inventory range!");
    return;
  }
  
  const rowData = invData[rowIndex];
  const id = rowData[ID_COL];
  const name = rowData[NAME_COL];
  const acqRaw = rowData[ACQ_COL];
  const freqRaw = rowData[FREQ_COL];
  
  console.log("=== ROW 67 DIAGNOSTIC RESULTS ===");
  console.log("ID Read:", id);
  console.log("Plant Name:", name);
  console.log("Acquisition Raw Value:", acqRaw, "| Is Date Object?:", acqRaw instanceof Date);
  console.log("Frequency Raw Value:", freqRaw);
  
  const lastRepot = getLastRepotDateForID(id);
  console.log("Last Repot Date Found (from Sheet 66 / Log):", lastRepot);
  
  const freqMonths = parseFrequencyToMonths(freqRaw);
  console.log("Parsed Frequency (in Months):", freqMonths);
  
  let acqDate = (acqRaw instanceof Date && !isNaN(acqRaw.getTime())) ? acqRaw : (acqRaw ? new Date(acqRaw) : null);
  const nextDate = calculateNextRepotDate_(acqDate, lastRepot, freqMonths || 12);
  
  console.log("Calculated Next Repot Due Date:", nextDate);
}