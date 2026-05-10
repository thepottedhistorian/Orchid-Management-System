/**
 * -----------------------------------------------------------------------------
 * ORCHID MANAGEMENT SYSTEM - ARCHIVE MANAGEMENT (Archive.gs)
 * -----------------------------------------------------------------------------
 * Purpose: Handles the decommissioning of orchids with Dashboard synchronization.
 * Logic:
 * 1. Historical Record: Moves data to the Archive sheet for long-term tracking.
 * 2. Dashboard Sync: Clears the ID/Name from Inventory to ensure the Web App 
 * ignores the row (solving the "Hidden Row" visibility bug).
 * 3. Sheet Recycling: Wipes the individual ledger and renames it to "TMP_[ID]" 
 * so it can be identified and reused by the Provisioning engine.
 * -----------------------------------------------------------------------------
 */

/**
 * SECTION I: INTERNAL ARCHIVE ENGINE
 * @param {string} orchidId - The ID of the specimen to decommission.
 * @param {string} batchStatus - (Optional) Status provided by bulk tool.
 * @param {string} batchNotes - (Optional) Notes provided by bulk tool.
 */
function archiveOrchidById_(orchidId, batchStatus = null, batchNotes = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ui;
  try { ui = SpreadsheetApp.getUi(); } catch (e) { console.log("UI not available"); }

  const inventory = ss.getSheetByName(INVENTORY_SHEET);
  if (!inventory) return false;
  
  // Ensure Archive sheet exists, create with headers if missing
  // Updated Header Row (K1) to include "Target FC"
  let archiveSheet = ss.getSheetByName(ARCHIVE_SHEET);
  if (!archiveSheet) {
    archiveSheet = ss.insertSheet(ARCHIVE_SHEET);
    archiveSheet.getRange("A1:K1").setValues([[
      "Orchid ID","Name","Genus","Species / Hybrid","Acquisition Date",
      "Target FC", "Archived Date","Final Status","Last Bloom Date","Last Repot Date","Notes"
    ]]);
    archiveSheet.setFrozenRows(1);
  }

  // Find the orchid row in Inventory
  const data = inventory.getDataRange().getValues();
  let rowIndex = -1;
  let rowValues = [];

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][ID_COL]) === String(orchidId)) {
      rowIndex = i + 1;
      rowValues = data[i];
      break;
    }
  }

  if (rowIndex === -1) {
    if (!batchStatus && ui) ui.alert("Orchid ID " + orchidId + " not found in Inventory.");
    return false;
  }

  // Handle Status and Notes via UI if not provided via Bulk
  let finalStatus = batchStatus;
  let notes = batchNotes;

  if (batchStatus === null && ui) {
    const statusResp = ui.prompt("Archive Orchid " + orchidId, "Enter final status (e.g., Died, Gifted, Rehomed):", ui.ButtonSet.OK_CANCEL);
    if (statusResp.getSelectedButton() !== ui.Button.OK) return false;
    finalStatus = statusResp.getResponseText().trim() || "Archived";

    const notesResp = ui.prompt("Archive Orchid " + orchidId, "Optional notes:", ui.ButtonSet.OK_CANCEL);
    if (notesResp.getSelectedButton() !== ui.Button.OK) return false;
    notes = notesResp.getResponseText().trim();
  }

  // Fetch dates from individual ledger before resetting
  const orchidSheet = ss.getSheetByName(String(orchidId));
  let lastBloomDate = null;
  let lastRepotDate = null;

  if (orchidSheet) {
    // These ranges (Row 19 and 37) refer to the individual ledger template
    lastBloomDate = getMostRecentDateFromRange_(orchidSheet.getRange(19, 1, 15, 1)); 
    lastRepotDate = getMostRecentDateFromRange_(orchidSheet.getRange(37, 1, 15, 1)); 
  }

  /**
   * SECTION II: DATA MIGRATION
   * Moving record to Archive and clearing Inventory for Dashboard Sync.
   * Updated to include LIGHT_FC_COL mapping.
   */
  archiveSheet.appendRow([
    rowValues[ID_COL],          // ID (Col A)
    rowValues[NAME_COL],        // Name (Col B)
    rowValues[GENUS_COL],       // Genus (Col C)
    rowValues[5],               // Species/Hybrid (Col F) - Unaffected by shift
    rowValues[ACQ_COL],         // Acquisition (Col I)
    rowValues[LIGHT_FC_COL],    // NEW: Target FC (Col O)
    new Date(),                 // Date Archived
    finalStatus, 
    lastBloomDate || "N/A", 
    lastRepotDate || "N/A", 
    notes || ""
  ]);

  // CRITICAL FIX: Clear the ID and Name from Inventory row.
  // This ensures the Standalone Dashboard ignores this row entirely.
  inventory.getRange(rowIndex, 1, 1, inventory.getLastColumn()).clearContent();
  inventory.hideRows(rowIndex);

  /**
   * SECTION III: LEDGER RECYCLING
   * Wiping the sheet and marking it for reuse by the Provisioning engine.
   */
  if (orchidSheet) {
    const templateSheet = ss.getSheetByName("template");
    if (templateSheet) {
      orchidSheet.clear(); 
      templateSheet.getDataRange().copyTo(orchidSheet.getRange(1, 1));
      
      // Rename to "TMP_" so Provisioning knows this is an available recycled sheet
      try {
        orchidSheet.setName("TMP_" + orchidId);
      } catch (e) {
        // Fallback if TMP_ID already exists (unlikely but safe)
        orchidSheet.setName("RECYCLED_" + new Date().getTime());
      }
      orchidSheet.hideSheet();
    }
  }

  return true;
}

/**
 * SECTION IV: HELPERS & UTILITIES
 */
function getMostRecentDateFromRange_(range) {
  const values = range.getValues();
  let latest = null;
  for (let r = 0; r < values.length; r++) {
    const v = values[r][0];
    if (v instanceof Date && !isNaN(v.getTime())) {
      if (!latest || v.getTime() > latest.getTime()) latest = v;
    }
  }
  return latest;
}

/**
 * SECTION V: MENU INTERFACES
 */
function archiveOrchid() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt("Archive Orchid", "Enter the Orchid ID to archive:", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const orchidId = response.getResponseText().trim();
  if (orchidId) {
    const success = archiveOrchidById_(orchidId);
    if (success) ui.alert(`ID ${orchidId} has been moved to Archive and the ledger is ready for reuse.`);
  }
}

function bulkArchiveOrchids() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt("Bulk Archive", "Enter Orchid IDs separated by commas:", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const ids = response.getResponseText().split(",").map(id => id.trim()).filter(id => id.length > 0);
  if (ids.length === 0) return;

  const batchStatus = ui.prompt("Batch Archive", "Final status for all (e.g., Gifted):", ui.ButtonSet.OK).getResponseText() || "Archived";
  const batchNotes = ui.prompt("Batch Archive", "Notes for all:", ui.ButtonSet.OK).getResponseText() || "";

  let count = 0;
  ids.forEach(id => {
    if (archiveOrchidById_(id, batchStatus, batchNotes)) count++;
  });

  ui.alert(`Successfully archived ${count} orchids.`);
}