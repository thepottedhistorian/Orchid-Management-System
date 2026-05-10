/**
 * Pops a prompt to sync a specific date from the Maintenance Log to ID sheets.
 */
function syncMaintenanceByDatePrompt() {
  const ui = SpreadsheetApp.getUi();
  
  // 1. Ask for the date
  const response = ui.prompt(
    'Historical Sync',
    'Enter the date to sync (YYYY-MM-DD):',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;
  const targetDateStr = response.getResponseText().trim();

  // Validate format basic check
  if (!/^\d{4}-\d{2}-\d{2}$/.test(targetDateStr)) {
    ui.alert("Invalid format. Please use YYYY-MM-DD (e.g., 2026-04-27).");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("Maintenance Log");
  const logData = logSheet.getDataRange().getValues();
  let syncCount = 0;

  for (let i = 1; i < logData.length; i++) {
    const orchidID = logData[i][0]; 
    const method = logData[i][2];   
    const dateTime = logData[i][3]; 
    const notes = logData[i][4];    
    
    if (dateTime instanceof Date && orchidID) {
      const formattedDate = Utilities.formatDate(dateTime, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
      
      if (formattedDate === targetDateStr) {
        const orchidSheet = ss.getSheetByName(String(orchidID).trim());
        
        if (orchidSheet) {
          let targetRow = orchidSheet.getRange("A100").getNextDataCell(SpreadsheetApp.Direction.UP).getRow() + 1;
          if (targetRow < 56) targetRow = 56;

          // Duplicate check
          const lastEntry = orchidSheet.getRange(Math.max(1, targetRow - 1), 1).getValue();
          if (lastEntry.toString() === dateTime.toString()) continue; 

          orchidSheet.getRange(targetRow, 1).setValue(dateTime);
          orchidSheet.getRange(targetRow, 2).setValue(method);
          orchidSheet.getRange(targetRow, 3).setValue(notes || "");
          syncCount++;
        }
      }
    }
  }
  
  ui.alert("Sync Complete. Updated " + syncCount + " records for " + targetDateStr);
}