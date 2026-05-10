/**
 * -----------------------------------------------------------------------------
 * ORCHID INVENTORY - BATCH UTILITIES (batch-tools.gs)
 * -----------------------------------------------------------------------------
 * Purpose: Advanced calculation engine for collection-wide repotting cycles.
 * Updated: 2026-03-30 to handle Column O shift and Globals sync.
 * -----------------------------------------------------------------------------
 */

/**
 * 👁️ PREVIEW RECALC REPOT DATES
 * Generates an HTML modal to preview the system-wide repotting schedule
 * before committing changes to the spreadsheet.
 */
function previewRecalcRepotDates() {
  const ss = SpreadsheetApp.openById(MASTER_ID);
  const inv = ss.getSheetByName(INVENTORY_SHEET);
  if (!inv) return;

  const range = inv.getDataRange();
  const values = range.getValues(); 
  const displayValues = range.getDisplayValues(); 
  const preview = [];

  for (let i = 1; i < values.length; i++) {
    const rowNum = i + 1;
    const id = values[i][ID_COL];
    const name = values[i][NAME_COL];
    
    // Skip empty or invalid ID rows
    if (!id || id.toString().trim() === "" || isNaN(id)) continue;
    if (!name || name.toString().trim() === "") continue;

    // Pulling from Globals: ACQ_COL (Index 8) and FREQ_COL (Now Index 19)
    const rawAcq = values[i][ACQ_COL];
    const displayAcq = displayValues[i][ACQ_COL];
    const freqRaw = displayValues[i][FREQ_COL]; // Now correctly looks at Column T

    let acqDate = null;
    if (rawAcq instanceof Date && !isNaN(rawAcq.getTime())) {
      acqDate = rawAcq;
    } else if (displayAcq) {
      acqDate = new Date(displayAcq);
    }
    
    // Logic: Extracts the maximum month count from the frequency string
    const freqMonths = parseFrequencyToMonths(freqRaw);
    const lastRepot = getLastRepotDateForID(id);
    const baseUsed = lastRepot ? "Last Repot" : "Acquisition";

    let nextDate = null;
    let status = "OK";

    if (!acqDate || isNaN(acqDate.getTime())) {
      status = "Missing/Bad Acq Date";
    } else if (!freqMonths) {
      status = "Missing Frequency";
    } else {
      nextDate = calculateNextRepotDate_(acqDate, lastRepot, freqMonths);
    }

    preview.push({
      row: rowNum,
      id: id,
      name: name,
      acquisition: (acqDate && !isNaN(acqDate.getTime())) ? Utilities.formatDate(acqDate, ss.getSpreadsheetTimeZone(), "MMM d, yyyy") : "N/A",
      lastRepot: lastRepot ? Utilities.formatDate(lastRepot, ss.getSpreadsheetTimeZone(), "MMM d, yyyy") : "N/A",
      baseUsed: baseUsed,
      frequency: freqRaw || "N/A",
      months: freqMonths || 0,
      next: nextDate ? Utilities.formatDate(nextDate, ss.getSpreadsheetTimeZone(), "MMM d, yyyy") : "N/A",
      status: status
    });
  }

  // Ensure the 'repot-preview' HTML file exists in your project
  const html = HtmlService.createTemplateFromFile("repot-preview");
  html.previewData = preview;
  SpreadsheetApp.getUi().showModalDialog(html.evaluate().setWidth(1100).setHeight(850), "Repot Date Preview");
}

/**
 * 🔄 RECALC ALL PLANT REPOT DUE DATES
 * Commits the calculated dates to Inventory [U], Maintenance Log [G], and individual ID Sheets [D10].
 */
function recalcAllPlantRepotDueDates() {
  const ss = SpreadsheetApp.openById(MASTER_ID);
  const inv = ss.getSheetByName(INVENTORY_SHEET);
  const log = ss.getSheetByName(MAINT_LOG_SHEET);
  if (!inv) return;

  // Clear maintenance log Column G to prevent stale overlaps
  if (log && log.getLastRow() > 1) {
    log.getRange(2, 7, log.getLastRow() - 1, 1).clearContent();
  }

  const data = inv.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const id = data[i][ID_COL]; 
    if (!id || id.toString().trim() === "" || isNaN(id)) continue;

    const acqRaw = data[i][ACQ_COL]; 
    const freqRaw = data[i][FREQ_COL]; // Corrected via Globals
    const freqMonths = parseFrequencyToMonths(freqRaw);

    const lastRepot = getLastRepotDateForID(id);
    let acqDate = (acqRaw instanceof Date) ? acqRaw : new Date(acqRaw);
    
    const nextDate = calculateNextRepotDate_(acqDate, lastRepot, freqMonths);

    if (nextDate) {
      // 1. Update Inventory [Column U] - Now uses NEXT_COL + 1 (Index 20 + 1)
      inv.getRange(i + 1, NEXT_COL + 1).setValue(nextDate); 

      // 2. Update Maintenance Log [Column G]
      if (log) {
        const logVals = log.getDataRange().getValues();
        for (let r = logVals.length - 1; r >= 1; r--) {
          if (logVals[r][0] == id) {
            log.getRange(r + 1, 7).setValue(nextDate); 
            break; 
          }
        }
      }

      // 3. Update Individual Orchid Sheet [D10]
      const orchidSheet = ss.getSheetByName(String(id));
      if (orchidSheet) {
        const cellD10 = orchidSheet.getRange("D10");
        cellD10.setValue(nextDate);
        cellD10.setNumberFormat('MMMM d, yyyy');
        cellD10.setBackground("#d4edda"); // Success highlight
      }
    }
  }
  SpreadsheetApp.getUi().alert("Recalculation Complete: All system dates synchronized.");
}

/**
 * 🔍 HELPER: GET LAST REPOT DATE FOR ID
 */
function getLastRepotDateForID(id) {
  const ss = SpreadsheetApp.openById(MASTER_ID);
  let latest = null;

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

  if (!latest) {
    const log = ss.getSheetByName(MAINT_LOG_SHEET);
    if (log) {
      const logData = log.getDataRange().getValues();
      for (let k = 1; k < logData.length; k++) {
        // Checking Column A for ID and Column H (Index 7) for Repot Date
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
  const base = lastRepotDate ? new Date(lastRepotDate) : new Date(acqDate);
  if (isNaN(base.getTime()) || !freqMonths) return null;
  const next = new Date(base);
  next.setMonth(next.getMonth() + freqMonths);
  return next;
}

/**
 * 🛠️ UTILITY: PARSE FREQUENCY TO MONTHS
 * Extracts the HIGHEST number found in a string (e.g., "12-24 months" -> 24).
 */
function parseFrequencyToMonths(freqRaw) {
  if (!freqRaw) return null;
  const text = freqRaw.toString().toLowerCase().trim();
  
  const matches = text.match(/(\d+(\.\d+)?)/g);
  if (!matches) return null;
  
  const num = parseFloat(matches[matches.length - 1]);
  
  return text.includes("year") ? num * 12 : num;
}