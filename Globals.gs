/**
 * -----------------------------------------------------------------------------
 * ORCHID MANAGEMENT SYSTEM - GLOBAL CONFIGURATION (Globals.gs)
 * -----------------------------------------------------------------------------
 * Purpose: Centralized coordinate and naming definitions for the system.
 * Version: 5.1.0 - Sanitized for Public Repository
 * -----------------------------------------------------------------------------
 */

// This automatically detects the ID of the spreadsheet it is attached to (Inventory)
// This is safe to leave as-is for GitHub.
const MASTER_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

/**
 * SCRUBBED SENSITIVE IDs
 * These values are pulled from the Script Properties vault.
 * Ensure you have run your Sanitizer script to populate these keys.
 */
const KEIKI_ID = PropertiesService.getScriptProperties().getProperty('KEIKI_ID');
const GOVEE_ARCHIVE_FOLDER_ID = PropertiesService.getScriptProperties().getProperty('GOVEE_ARCHIVE_ID');

/**
 * 1. INVENTORY COLUMN INDICES (0-indexed)
 * Used for the "Inventory" and "Keiki Inventory" sheets.
 */
const ID_COL = 0;           // Column A
const NAME_COL = 1;          // Column B
const GENUS_COL = 2;         // Column C
const ACQ_COL = 8;           // Column I
const REPOT_NOTES_COL = 9;   // Column J
const BLOOM_STATUS_COL = 10; // Column K
const LAST_BLOOM_COL = 11;   // Column L
const LIGHT_FC_COL = 14;     // Column O
const FREQ_COL = 19;         // Column T
const NEXT_COL = 20;         // Column U

/**
 * 2. GREENHOUSE DASHBOARD MAPPING
 * Maps the summary data range (C4:C11) on the "Greenhouse Dashboard" sheet.
 */
const DASHBOARD_SHEET = "Greenhouse Dashboard";
const DASH_DATA_RANGE = "C4:C11"; 
const DASH_TOTAL_IDX = 0;    // Row 4: Total Collection
const DASH_BLOOM_IDX = 1;    // Row 5: Currently in Bloom
const DASH_WATER_DAY_IDX = 2;// Row 6: Last Watering Day
const DASH_DAYS_SINCE_IDX = 3;// Row 7: Days Since Watered
const DASH_RES_IDX = 6;      // Row 10: Research Progress %
const DASH_HOURS_IDX = 7;    // Row 11: Total Dedicated Project Time (Hrs)

/**
 * 3. SHEET NAMES (CORE ENGINE)
 */
const INVENTORY_SHEET = "Inventory";
const KEIKI_INVENTORY = "Keiki Inventory";
const ARCHIVE_SHEET = "Archived Orchids";
const MAINT_LOG_SHEET = "Maintenance Log";
const RESEARCH_SHEET_NAME = "Research Tracker";
const TEMPLATE_SHEET = "Template";

/**
 * 4. ENVIRONMENTAL DATA TABS
 */
const RUDSTA_DATA = "Office Cabinet Data";     // Now stores Rudsta + Terrarium
const MILSBO_DATA = "Living Room Cabinet Data";

const OFFICE_TAB = "Office Cabinet";
const LIVING_ROOM_TAB = "Living Room Cabinet";
const TERRARIUM_TAB = "Lizard Terrarium";

/**
 * 5. DISPLAY & CHART MAPPING
 */
const DISPLAY_NAMES = {
  "Rudsta top shelf": "Rudsta Top",
  "Rudsta 2 (mid)": "Rudsta Mid",
  "Rudsta 3 (bottom)": "Rudsta Bottom",
  "6 Milsbo upper": "Milsbo Upper",
  "7 Milsbo lower": "Milsbo Lower",
  "Terrarium": "Terrarium"
};