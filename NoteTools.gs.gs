/**
 * -----------------------------------------------------------------------------
 * ORCHID INVENTORY - NOTE CLONING UTILITIES (NoteTools.gs)
 * -----------------------------------------------------------------------------
 * Purpose: Copies a Cell Note (the black triangle comment) from a source 
 * orchid sheet to all other numbered ID sheets in the collection.
 * Logic:
 * 1. Grabs the 'Note' content from cell C57 of the source sheet.
 * 2. Iterates through all sheets in the Master Workbook.
 * 3. Targets only numeric sheet names (Orchid IDs).
 * 4. Pastes the note into the same cell (C57) on every target.
 * -----------------------------------------------------------------------------
 */

function copyNoteToNumberedSheets() {
  // Uses MASTER_ID from Globals.gs for a persistent connection
  const ss = SpreadsheetApp.openById(MASTER_ID);

  // The "Master" sheet where the original C57 note lives (Orchid ID 1)
  const sourceSheetName = "1";
  const sourceSheet = ss.getSheetByName(sourceSheetName);

  if (!sourceSheet) {
    throw new Error("Source sheet '" + sourceSheetName + "' not found. Check the tab name.");
  }

  // Get the technical 'Note' metadata (the hover-text), not the cell value
  const sourceNote = sourceSheet.getRange("C57").getNote();

  const sheets = ss.getSheets();
  let updateCount = 0;

  sheets.forEach(sheet => {
    const name = sheet.getName().trim();

    /**
     * REGEX: Only update sheets whose names are digits (2, 3, 4, ...)
     * Explicitly skip the source sheet to avoid redundant work.
     */
    if (/^\d+$/.test(name) && name !== sourceSheetName) {
      sheet.getRange("C57").setNote(sourceNote);
      updateCount++;
    }
  });

  // UI Notification for the administrator
  try {
    SpreadsheetApp.getUi().alert(`Success: Note from Sheet ${sourceSheetName} copied to ${updateCount} orchid sheets.`);
  } catch (e) {
    console.log(`Note sync complete. ${updateCount} sheets updated.`);
  }
}