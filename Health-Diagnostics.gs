/**
 * -----------------------------------------------------------------------------
 * ORCHID INVENTORY - FINAL SYSTEM DIAGNOSTICS (Health-Diagnostics.gs)
 * -----------------------------------------------------------------------------
 * Purpose: Verifies the cross-talk between all system modules.
 * Checks:
 * 1. Global Column Alignment (Focus on Col I and Col S).
 * 2. Archive Consistency (Ensures hidden rows match Archive sheet).
 * 3. Web App Accessibility (Verifies properties for index.html).
 * -----------------------------------------------------------------------------
 */

/**
 * 🕵️‍♂️ RUN FULL SYSTEM DIAGNOSTIC
 * Accessible via the script editor to verify the entire "Bound" logic.
 */
function runFullSystemDiagnostic() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inv = ss.getSheetByName(INVENTORY_SHEET);
  const archive = ss.getSheetByName(ARCHIVE_SHEET);
  const ui = SpreadsheetApp.getUi();
  let report = ["🔍 SYSTEM DIAGNOSTIC REPORT", "--------------------------"];

  if (!inv) {
    ui.alert("CRITICAL ERROR: Inventory Sheet not found. Check Globals.gs");
    return;
  }

// 1. Check Column Alignment
  if (headers[ACQ_COL] && headers[ACQ_COL].toString().toLowerCase().includes("acq")) {
    report.push("✅ Column I (Acquisition): Correctly Aligned.");
  } else {
    report.push(`⚠️ Column I Mismatch: Found "${headers[ACQ_COL]}" (Expected index ${ACQ_COL})`);
  }

  // Update the label here to Column T (Index 19)
  if (headers[FREQ_COL] && headers[FREQ_COL].toString().toLowerCase().includes("freq")) {
    report.push(`✅ Column ${String.fromCharCode(65 + FREQ_COL)} (Frequency): Correctly Aligned.`);
  } else {
    report.push(`⚠️ Column Mismatch: Found "${headers[FREQ_COL]}" at index ${FREQ_COL} instead of Frequency.`);
  }

  // OPTIONAL: Add a check for your new Foot Candle column
  if (headers[LIGHT_FC_COL] && headers[LIGHT_FC_COL].toString().toLowerCase().includes("fc")) {
    report.push("✅ Column O (Target FC): Correctly Aligned.");
  }
  
  // 2. Check Archive vs. Hidden Rows
  const lastRow = inv.getLastRow();
  let hiddenCount = 0;
  for (let i = 2; i <= lastRow; i++) {
    if (inv.isRowHiddenByUser(i)) hiddenCount++;
  }
  
  const archiveCount = archive ? archive.getLastRow() - 1 : 0;
  report.push(`📊 Hidden Rows in Inventory: ${hiddenCount}`);
  report.push(`📊 Entries in Archive Sheet: ${archiveCount}`);

  if (hiddenCount !== archiveCount) {
    report.push("ℹ️ Note: Archive count and Hidden count differ. (This is normal if you manually hid rows).");
  }

  // 3. Web App Data Simulation
  try {
    const testData = getDashboardData();
    if (testData.collection && testData.collection.length > 0) {
      report.push(`✅ Web App Logic: Successfully mapped ${testData.collection.length} orchids.`);
    } else {
      report.push("⚠️ Web App Logic: No active orchids found for dashboard.");
    }
  } catch (e) {
    report.push(`❌ Web App Logic Error: ${e.message}`);
  }

  ui.alert(report.join("\n"));
}