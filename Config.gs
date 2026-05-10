/**
 * -----------------------------------------------------------------------------
 * CONFIGURATION & SENSITIVE DATA BRIDGE
 * -----------------------------------------------------------------------------
 * This file pulls sensitive IDs from the Script Properties vault.
 * For GitHub: Ensure these keys are set in your local Project Settings.
 */

const CONFIG = {
  // Spreadsheet IDs
  DATABASE_ID: PropertiesService.getScriptProperties().getProperty('MAIN_SS_ID') || "PASTE_ID_HERE_FOR_INITIAL_SETUP",
  
  // Folder IDs for Govee Data
  GOVEE_FOLDER_ID: PropertiesService.getScriptProperties().getProperty('GOVEE_FOLDER_ID') || "PASTE_ID_HERE_FOR_INITIAL_SETUP",
  
  // Archive/Recycling Settings
  ARCHIVE_PREFIX: "TMP_",
  
  // Project Info
  VERSION: "5.1.0",
  PROJECT_NAME: "The Satyrion Chronicles - Orchid Management System"
};