/**
 * -----------------------------------------------------------------------------
 * ORCHID INVENTORY - BULK WATERING & FERTILIZER UPDATES (Watering.gs)
 * -----------------------------------------------------------------------------
 * Purpose: Appends specific nutritional or maintenance notes to the most 
 * recent entry in the watering logs across the entire collection.
 * -----------------------------------------------------------------------------
 */

/**
 * Appends a specific fertilizer note to the last active row.
 * FIXED: Now includes a check to prevent duplicate notes in the same cell.
 */
function appendFertilizerNotes() {
  const ss = SpreadsheetApp.openById(MASTER_ID);
  const sheets = ss.getSheets();
  
  // This is the exact string we are looking for to prevent duplicates
  const noteContent = "50/50 well water + city water + fertilizer";
  let updateCount = 0;

  sheets.forEach(sheet => {
    const sheetName = sheet.getName().trim();
    
    if (/^\d+$/.test(sheetName)) {
      const lastRow = sheet.getLastRow();
      if (lastRow < 56) return; 

      const columnA = sheet.getRange(56, 1, lastRow - 55, 1).getValues();
      let actualLastRow = 0;
      
      for (let i = columnA.length - 1; i >= 0; i--) {
        if (columnA[i][0] !== "" && columnA[i][0] !== null) {
          actualLastRow = i + 56;
          break;
        }
      }

      if (actualLastRow > 0) {
        const cell = sheet.getRange(actualLastRow, 3); // Column C (Notes)
        const existingNote = String(cell.getValue()); // Ensure it's treated as a string
        
        // --- THE FIX: PREVENT DUPLICATES ---
        // If the note already contains our fertilizer string, SKIP IT.
        if (existingNote.includes(noteContent)) {
          console.log(`Skipping Orchid #${sheetName}: Note already added.`);
          return; 
        }
        // ------------------------------------
        
        const newNote = (existingNote && existingNote !== "") 
                        ? `${existingNote}; ${noteContent}` 
                        : noteContent;
        
        cell.setValue(newNote);
        updateCount++;
        console.log(`Updated Orchid #${sheetName} at Row ${actualLastRow}`);
      }
    }
  });
  
  console.log(`Bulk update complete. ${updateCount} orchid sheets updated.`);
  
  try {
    SpreadsheetApp.getUi().alert(`Success: Fertilizer notes added to ${updateCount} orchid logs.`);
  } catch (e) {}
}