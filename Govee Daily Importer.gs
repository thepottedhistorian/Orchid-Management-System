/**
 * -----------------------------------------------------------------------------
 * ORCHID MANAGEMENT SYSTEM - GOVEE CLIMATE IMPORTER (Govee Importer.gs)
 * -----------------------------------------------------------------------------
 * Purpose: Captures high-fidelity environmental samples (9AM/3PM/9PM).
 * Version: 5.1.0 - Sanitized for Public Repository
 * -----------------------------------------------------------------------------
 */

/**
 * SCRUBBED: Points to your private Govee Data folder via PropertiesService.
 * Ensure 'GOVEE_IMPORT_ID' is set in your Project Settings via Sanitizer.gs.
 */
const GOVEE_IMPORT_FOLDER_ID = PropertiesService.getScriptProperties().getProperty('GOVEE_IMPORT_ID');

/**
 * Main import trigger for daily climate parity.
 * Overhauled in v5.0 to capture high-fidelity samples at 9 AM, 3 PM, and 9 PM.
 */
function importGoveeDaily() {
  const ui = SpreadsheetApp.getUi();
  
  // Safety check for sanitized environments
  if (!GOVEE_IMPORT_FOLDER_ID) {
    ui.alert("Configuration Error: Govee Import Folder ID not found in Script Properties. Please run the Sanitizer or check Project Settings.");
    return;
  }

  try {
    const folder = DriveApp.getFolderById(GOVEE_IMPORT_FOLDER_ID);
    const files = folder.getFiles();
    
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      
      // logic for processing 30-day rolling samples...
      // (Rest of your specific processing logic remains unchanged)
    }
    
    console.log("Govee daily sync completed successfully.");
    
  } catch (e) {
    console.error("Govee Import Failed: " + e.toString());
    ui.alert("Govee Sync Error: " + e.message);
  }
}

/**
 * Refreshes cabinet tables and updates the greenhouse dashboard averages.
 */
function generateCabinetTables() {
  // Existing logic for table generation...
}