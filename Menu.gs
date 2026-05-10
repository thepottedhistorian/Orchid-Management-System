/**
 * -----------------------------------------------------------------------------
 * ORCHID MANAGEMENT SYSTEM - CONSOLIDATED MENU (Menu.gs)
 * -----------------------------------------------------------------------------
 * Version: 5.1.2
 * Updated: 2026-05-10
 * Project: The Satyrion Chronicles / Orchid Tracker
 * * CHANGE LOG:
 * - Sanitized all external URLs using ScriptProperties for GitHub security.
 * - Integrated automatic Sidebar launch on document open.
 * - Added visibility utilities for managing a large collection of ID sheets.
 * -----------------------------------------------------------------------------
 */

/**
 * Standard trigger that builds the menu and launches the sidebar.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("🌸 Orchid Tools")
    // Navigation & Layout
    .addItem("🧭 Open Navigation Sidebar", "showSidebar")
    .addItem("🌱 Open Keiki Workbook", "openKeikiWorkbookLink")
    .addSeparator()

    // STAGE 1 & 2: Provisioning Workflow
    .addItem("✨ 1. Provision New Row", "provisionNewOrchid")
    .addItem("🔗 2. Finalize & Link Ledger", "finalizeNewOrchid")
    .addSeparator()
    
    // Archiving Section
    .addItem("📦 Archive Orchid", "archiveOrchid")
    .addItem("📦 Bulk Archive Orchids", "bulkArchiveOrchids")
    .addSeparator()
    
    // Govee Data Section
    .addItem("📊 Sync Govee Data (Gmail/Drive)", "importGoveeDaily")
    .addItem("🔄 Refresh Cabinet Tables", "generateCabinetTables")
    .addSeparator()

    // Watering & Maintenance Log Tools
    .addSubMenu(
      ui.createMenu("🚿 Watering & Maintenance")
        .addItem("Log Bulk Watering (Selected Rows)", "showWateringModal")
        .addItem("🔄 Sync History by Date", "syncMaintenanceByDatePrompt")
        .addSeparator()
        .addItem("🌿 Bulk Add Fertilizer Notes", "appendFertilizerNotes")
        .addItem("🧹 Clean Duplicate Logs (Date Prompt)", "cleanDuplicateOrchidEntriesByDate")
    )
    .addSeparator()

    // Repotting Tools Submenu
    .addSubMenu(
      ui.createMenu("📅 Repotting Tools")
        .addItem("Step 1: Sync Log Column H (from A38)", "transferA38ToLogColumnH")
        .addItem("Step 2: Calculate Next Due", "calculateOnlyMaintenanceLogG")
        .addItem("Step 3: Push Final Dates to ID Sheets", "pushInventoryTToOrchidD10")
        .addSeparator()
        .addItem("👁️ Preview Repot Calculations", "previewRecalcRepotDates")
    )
    .addSeparator()

    // System Formatting Submenu
    .addSubMenu(
      ui.createMenu("🛠️ System Formatting")
        .addItem("Refresh All Bloom Icons", "refreshAllBloomStatusFormatting")
        .addItem("Refresh All Stoplight Dates", "applyAllSystemStoplights") 
        .addItem("Fix All Timestamps", "fixAllPhasesSequentially")
        .addSeparator()
        .addItem("🙈 Hide Current Sheet", "hideActiveSheet")
        .addItem("👁️ Show All Hidden Sheets", "showAllHiddenSheets")
    )
    .addSeparator()
    .addItem("🏥 Run System Health Check", "runSystemHealthCheck")
    .addToUi();

  // Auto-launch the sidebar on workbook open
  showSidebar();
}

/**
 * -----------------------------------------------------------------------------
 * NAVIGATION & SIDEBAR UTILITIES
 * -----------------------------------------------------------------------------
 */

/**
 * Renders the Sidebar.html file to the UI.
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Orchid Tracker Navigator')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Facilitates direct navigation to a tab by name or ID.
 * Called by Sidebar.html
 */
function setActiveSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    ss.setActiveSheet(sheet);
  } else {
    throw new Error('Sheet "' + sheetName + '" not found.');
  }
}

/**
 * Opens the Keiki Workbook Link (called from Menu)
 * Sanitized for GitHub using ScriptProperties.
 */
function openKeikiWorkbookLink() {
  const props = PropertiesService.getScriptProperties();
  const url = props.getProperty('KEIKI_WORKBOOK_URL');
  
  if (!url) {
    SpreadsheetApp.getUi().alert("Error: Keiki Workbook URL not found in Script Properties. Please add it to Project Settings.");
    return;
  }

  const html = HtmlService.createHtmlOutput(
    '<html><script>' +
    'var win = window.open("' + url + '", "_blank");' +
    'if(win){ google.script.host.close(); }' +
    'else { alert("Pop-up blocked! Please allow pop-ups for this sheet."); }' +
    '</script></html>'
  ).setWidth(10).setHeight(10);
  
  SpreadsheetApp.getUi().showModalDialog(html, "Opening Keiki Workbook...");
}

/**
 * Helper for Sidebar to get the sanitized URL.
 */
function getKeikiUrl() {
  return PropertiesService.getScriptProperties().getProperty('KEIKI_WORKBOOK_URL');
}

/**
 * -----------------------------------------------------------------------------
 * VISIBILITY UTILITIES
 * -----------------------------------------------------------------------------
 */

/**
 * Hides the currently focused sheet to keep the tab bar clean.
 */
function hideActiveSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  // Prevent hiding the last visible sheet
  const visibleSheets = ss.getSheets().filter(s => !s.isSheetHidden());
  if (visibleSheets.length > 1) {
    sheet.hideSheet();
  } else {
    SpreadsheetApp.getUi().alert("Digital Steward Alert: Cannot hide the only visible sheet.");
  }
}

/**
 * Reveals all previously hidden sheets (ID ledgers) in the workbook.
 */
function showAllHiddenSheets() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  sheets.forEach(sheet => sheet.showSheet());
}