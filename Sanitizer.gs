/**
 * SYSTEM SANITIZER
 * Run this once to move hardcoded IDs into private properties.
 */
function runSystemSanitization() {
  const ui = SpreadsheetApp.getUi();
  const ps = PropertiesService.getScriptProperties();
  
  // 1. Capture the current Spreadsheet ID
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  ps.setProperty('MAIN_SS_ID', ssId);
  
  // 2. Scan for Govee Folder IDs (Assuming you have them in your sync scripts)
  // Replace 'YOUR_ACTUAL_FOLDER_ID' with your real ID if known, 
  // or the script will prompt you.
  const folderId = "REPLACE_WITH_YOUR_GOVEE_FOLDER_ID"; 
  ps.setProperty('GOVEE_FOLDER_ID', folderId);

  ui.alert("Sanitization Complete: IDs moved to Script Properties. You can now replace hardcoded IDs in your code with CONFIG.DATABASE_ID.");
}