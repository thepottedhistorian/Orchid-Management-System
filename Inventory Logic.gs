/**
 * -----------------------------------------------------------------------------
 * ORCHID INVENTORY - PROVISIONING & DATA INTEGRITY (InventoryLogic.gs)
 * -----------------------------------------------------------------------------
 * Updated: 2026-05-09
 * Purpose: Manages two-stage orchid entry and database synchronization.
 * -----------------------------------------------------------------------------
 */

/**
 * 🌸 STAGE 1: PROVISION NEW ROW
 * Identifies a row to reuse (hidden) or the next available empty row.
 */
function provisionNewOrchid() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inv = ss.getSheetByName(INVENTORY_SHEET);
  const ui = SpreadsheetApp.getUi();
  
  if (!inv) {
    ui.alert("Error: 'Inventory' sheet not found.");
    return;
  }

  const lastRow = inv.getLastRow();
  let targetRow = -1;

  // 1. Check for hidden rows to recycle
  for (let i = 2; i <= lastRow; i++) {
    if (inv.isRowHiddenByUser(i)) {
      targetRow = i;
      break;
    }
  }

  // 2. Navigation and User Notification
  if (targetRow !== -1) {
    inv.showRows(targetRow);
    inv.getRange(targetRow, 1).activate(); 
    ui.alert(`Recycled row ${targetRow} found. Please fill out the botanical data (Columns B through V) and then run 'Finalize'.`);
  } else {
    targetRow = lastRow + 1;
    inv.getRange(targetRow, 1).activate();
    ui.alert(`No hidden rows found. Moving to next empty row: ${targetRow}. Please fill out details and then run 'Finalize'.`);
  }
}

/**
 * 🌸 STAGE 2: FINALIZE & LINK LEDGER
 * Creates the individual sheet from Template and maps data to specific cells.
 */
function finalizeNewOrchid() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inv = ss.getSheetByName(INVENTORY_SHEET);
  const template = ss.getSheetByName(TEMPLATE_SHEET);
  const ui = SpreadsheetApp.getUi();
  
  const activeRow = inv.getActiveCell().getRow();
  
  // Capture all 22 columns of data (A through V)
  const data = inv.getRange(activeRow, 1, 1, 22).getValues()[0]; 
  
  const orchidId = data[0];   // Col A: ID #
  const orchidName = data[1]; // Col B: Orchid Name
  
  if (!orchidId || orchidId === "") {
    ui.alert("Critical Error: No ID found in Column A. This is required for sheet generation.");
    return;
  }

  // 1. Create the new sheet from Template
  if (ss.getSheetByName(String(orchidId))) {
    ui.alert(`Error: A sheet for ID ${orchidId} already exists.`);
    return;
  }
  
  const newSheet = template.copyTo(ss).setName(String(orchidId));
  newSheet.showSheet();

  // 2. Map Inventory Data to Template Cells
  // Botanical Info Section
  newSheet.getRange("B2").setValue(data[2]);  // Genus (Col C)
  newSheet.getRange("B3").setValue(data[5]);  // Species / Hybrid (Col F)
  newSheet.getRange("B4").setValue(data[6]);  // Alliance (Col G)
  newSheet.getRange("B5").setValue(data[1]);  // Orchid Name (Col B)
  newSheet.getRange("B6").setValue(data[7]);  // Native Range (Col H)
  
  // Status and Maintenance
  newSheet.getRange("D5").setValue(data[10]); // Bloom Status (Col K)
  newSheet.getRange("D10").setValue(data[21]); // Next Repot Due (Col V)

  // Care Profile Section
  newSheet.getRange("B11").setValue(data[14]); // Light Needs (Col O)
  newSheet.getRange("B12").setValue(data[16]); // Watering Needs (Col Q)
  newSheet.getRange("B13").setValue(data[17]); // Humidity (Col R)
  newSheet.getRange("B14").setValue(data[18]); // Potting Media (Col S)
  newSheet.getRange("B15").setValue(data[19]); // Fertilizer Routine (Col T)
  newSheet.getRange("B16").setValue(data[20]); // Repotting Schedule (Col U)

  // 3. Create Hyperlink in Inventory Sheet (Col A)
  const sheetId = newSheet.getSheetId();
  const gidLink = `#gid=${sheetId}`;
  inv.getRange(activeRow, 1).setFormula(`=HYPERLINK("${gidLink}", "${orchidId}")`);

  ui.alert(`Success! Ledger created for ${orchidName} (ID: ${orchidId}).`);
}

/**
 * 🏥 SYSTEM HEALTH CHECK
 * Verifies ID synchronization, missing ledger sheets, and formatting.
 */
function runSystemHealthCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inv = ss.getSheetByName(INVENTORY_SHEET);
  const log = [];

  if (!inv) return;

  const data = inv.getDataRange().getValues();
  const seenIDs = new Set();

  for (let i = 1; i < data.length; i++) {
    const row = i + 1;
    const id = data[i][0]; // Column A: ID
    const name = data[i][1]; // Column B: Name
    
    // Skip hidden (archived) rows
    if (inv.isRowHiddenByUser(row)) continue;
    if (!id || !name) continue;

    // Check for Duplicate IDs
    if (seenIDs.has(id)) {
      log.push(`Row ${row}: Duplicate ID ${id}`);
    } else {
      seenIDs.add(id);
    }

    // Verify individual sheet existence
    const sheet = ss.getSheetByName(String(id));
    if (!sheet) {
      log.push(`ID ${id}: Missing individual sheet`);
    } else if (sheet.isSheetHidden()) {
      log.push(`ID ${id}: Sheet is hidden but orchid is active`);
    }

    // NEW: Verify Frequency format (e.g. "12 months")
    // Assuming Frequency is in Column U (Index 20)
    const freqRaw = data[i][20]; 
    if (freqRaw && !parseFrequencyToMonths(freqRaw)) {
      log.push(`ID ${id}: Invalid frequency format "${freqRaw}"`);
    }
  }

  const ui = SpreadsheetApp.getUi();
  if (log.length === 0) {
    ui.alert("✅ System Health Check: Database is synchronized!");
  } else {
    ui.alert("⚠️ Issues Found:\n\n" + log.join("\n"));
  }
}

/**
 * 🛠️ HELPER: Parses string frequency into a numeric month value.
 */
function parseFrequencyToMonths(input) {
  if (!input) return null;
  const s = input.toString().toLowerCase();
  const match = s.match(/\d+/);
  if (!match) return null;
  const val = parseInt(match[0]);
  if (s.includes("year")) return val * 12;
  if (s.includes("month")) return val;
  return null;
}