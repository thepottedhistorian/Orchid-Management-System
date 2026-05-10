/**
 * -----------------------------------------------------------------------------
 * ORCHID MANAGEMENT SYSTEM - CONSOLIDATED MENU (Menu.gs)
 * -----------------------------------------------------------------------------
 * Version: 5.1.0
 * Updated: 2026-05-09
 * * CHANGELOG:
 * - ADDED: Navigation Sidebar (Sidebar.html) for inventory-focused management.
 * - ADDED: Auto-open trigger in onOpen to keep Sidebar persistent.
 * - ADDED: Sheet visibility utilities (hideActiveSheet / showAllHiddenSheets).
 * - UPDATED: Removed research-heavy links to focus on Active Inventory.
 * -----------------------------------------------------------------------------
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("🌸 Orchid Tools")
    // Navigation & Layout
    .addItem("🧭 Open Navigation Sidebar", "showSidebar")
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
 * Hides the currently focused sheet.
 */
function hideActiveSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  // Prevent hiding the last visible sheet
  const visibleSheets = ss.getSheets().filter(s => !s.isSheetHidden());
  if (visibleSheets.length > 1) {
    sheet.hideSheet();
  } else {
    SpreadsheetApp.getUi().alert("Cannot hide the only visible sheet.");
  }
}

/**
 * Reveals all previously hidden sheets in the workbook.
 */
function showAllHiddenSheets() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  sheets.forEach(sheet => sheet.showSheet());
}