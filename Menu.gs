/**
 * -----------------------------------------------------------------------------
 * ORCHID MANAGEMENT SYSTEM - CONSOLIDATED MENU (Menu.gs)
 * -----------------------------------------------------------------------------
 * Version: 6.1.0
 * Updated: 2026-08-08
 * Project: The Satyrion Chronicles / Orchid Tracker
 * CHANGE LOG:
 * - Sanitized all external URLs using ScriptProperties for GitHub security.
 * - Keeps the navigation sidebar available from the menu without auto-launching it.
 * - Added visibility utilities for managing a large collection of ID sheets.
 * - Integrated Export to XLSX utility into System Formatting submenu.
 * - Added Date Standardization functions to System Formatting submenu.
 * - Integrated System Manual & README viewer launcher (showSystemReadmeModal).
 * - Added searchAndNavigateSheetOrOrchid to query sheet names/numbers and cell B5.
 * -----------------------------------------------------------------------------
 */

/**
 * Standard trigger that builds the Orchid Tools menu.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("🌸 Orchid Tools")
    .addItem("🧭 Open Navigation Sidebar", "showSidebar")
    .addItem("🌱 Open Keiki Workbook", "openKeikiWorkbookLink")
    .addSeparator()

    .addItem("✨ 1. Provision New Row", "provisionNewOrchid")
    .addItem("🔗 2. Finalize & Link Ledger", "finalizeNewOrchid")
    .addSeparator()

    .addItem("📦 Archive Orchid", "archiveOrchid")
    .addItem("📦 Bulk Archive Orchids", "bulkArchiveOrchids")
    .addSeparator()

    .addItem("📊 Sync Govee Data (Gmail/Drive)", "importGoveeDaily")
    .addItem("🔄 Refresh Cabinet Tables", "generateCabinetTables")
    .addSeparator()

    .addSubMenu(
      ui.createMenu("🚿 Watering & Maintenance")
        .addItem("Log Bulk Watering (Selected Rows)", "showWateringModal")
        .addItem("🔄 Sync History by Date", "syncMaintenanceByDatePrompt")
        .addSeparator()
        .addItem("🌿 Bulk Add Fertilizer Notes", "appendFertilizerNotes")
        .addItem("🧹 Clean Duplicate Logs (Date Prompt)", "cleanDuplicateOrchidEntriesByDate")
    )
    .addSeparator()

    .addSubMenu(
      ui.createMenu("📅 Repotting Tools")
        .addItem("Step 1: Sync Log Column H (from A38)", "transferA38ToLogColumnH")
        .addItem("Step 2: Calculate Next Due", "calculateOnlyMaintenanceLogG")
        .addItem("Step 3: Push Final Dates to ID Sheets", "pushInventoryTToOrchidD10")
        .addSeparator()
        .addItem("👁️ Preview Repot Calculations", "showRepotPreviewModal")
    )
    .addSeparator()

    .addSubMenu(
      ui.createMenu("🛠️ System Formatting")
        .addItem("Refresh All Bloom Icons", "refreshAllBloomStatusFormatting")
        .addItem("Refresh All Stoplight Dates", "applyAllSystemStoplights")
        .addItem("Fix All Timestamps", "fixAllPhasesSequentially")
        .addSeparator()
        .addItem("📅 Standardize Dates (Selected Range)", "standardizeSelectedDates")
        .addItem("📅 Standardize Dates (Active Sheet)", "standardizeActiveSheetDates")
        .addItem("📅 Standardize All Inventory & ID Dates (1-65)", "standardizeAllOrchidInventoryDates")
        .addSeparator()
        .addItem("📊 Export Clean XLSX to Excel", "exportWorkbookToXLSX")
        .addSeparator()
        .addItem("🙈 Hide Current Sheet", "hideActiveSheet")
        .addItem("👁️ Show All Hidden Sheets", "showAllHiddenSheets")
        .addItem("🌸 Recalculate Bloom Durations (1-65)", "updateAllBloomDurations")
    )
    .addSeparator()
    .addItem("🏥 Run System Health Check", "runSystemHealthCheck")
    .addItem("📖 System Manual & README", "showSystemReadmeModal")
    .addToUi();
}

/**
 * Renders the Sidebar.html file to the UI.
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("Orchid Tracker Navigator")
    .setWidth(300);

  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Facilitates direct navigation to a tab by name or ID.
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
 * Searches sheets by sheet name/number or Orchid Name in cell B5.
 *
 * @param {string} query Sheet number, name, or orchid name.
 * @return {string} Activated sheet name.
 */
function searchAndNavigateSheetOrOrchid(query) {
  if (!query || typeof query !== "string") {
    throw new Error("Please enter a valid search term.");
  }

  const trimmedQuery = query.trim().toLowerCase();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let targetSheet = null;

  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    const sheetName = sheet.getName().toLowerCase();

    if (sheetName === trimmedQuery || sheetName.includes(trimmedQuery)) {
      targetSheet = sheet;
      break;
    }
  }

  if (!targetSheet) {
    for (let i = 0; i < sheets.length; i++) {
      const sheet = sheets[i];

      try {
        if (sheet.getLastRow() >= 5) {
          const b5Value = sheet.getRange("B5").getValue();

          if (
            b5Value &&
            typeof b5Value === "string" &&
            b5Value.toLowerCase().includes(trimmedQuery)
          ) {
            targetSheet = sheet;
            break;
          }
        }
      } catch (err) {
        console.warn("Could not read B5 for sheet: " + sheet.getName(), err);
      }
    }
  }

  if (targetSheet) {
    ss.setActiveSheet(targetSheet);
    return targetSheet.getName();
  }

  throw new Error('No matching sheet or orchid name found for: "' + query + '"');
}

/**
 * Opens the Keiki Workbook Link.
 */
function openKeikiWorkbookLink() {
  const props = PropertiesService.getScriptProperties();
  const url = props.getProperty("KEIKI_WORKBOOK_URL");

  if (!url) {
    SpreadsheetApp.getUi().alert(
      "Error: Keiki Workbook URL not found in Script Properties. Please add it to Project Settings."
    );
    return;
  }

  const html = HtmlService.createHtmlOutput(
    '<html><script>' +
      'var win = window.open("' + url + '", "_blank");' +
      'if(win){ google.script.host.close(); }' +
      'else { alert("Pop-up blocked! Please allow pop-ups for this sheet."); }' +
      "</script></html>"
  )
    .setWidth(10)
    .setHeight(10);

  SpreadsheetApp.getUi().showModalDialog(html, "Opening Keiki Workbook...");
}

/**
 * Returns the sanitized Keiki workbook URL for Sidebar.html.
 */
function getKeikiUrl() {
  return PropertiesService.getScriptProperties().getProperty("KEIKI_WORKBOOK_URL");
}

/**
 * Hides the currently focused sheet, unless it is the last visible sheet.
 */
function hideActiveSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const visibleSheets = ss.getSheets().filter(s => !s.isSheetHidden());

  if (visibleSheets.length > 1) {
    sheet.hideSheet();
  } else {
    SpreadsheetApp.getUi().alert(
      "Digital Steward Alert: Cannot hide the only visible sheet."
    );
  }
}

/**
 * Reveals all previously hidden sheets.
 */
function showAllHiddenSheets() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  sheets.forEach(sheet => sheet.showSheet());
}

/**
 * Launches the System README and documentation viewer.
 */
function showSystemReadmeModal() {
  const html = HtmlService.createHtmlOutputFromFile("HelpFile")
    .setWidth(950)
    .setHeight(750);

  SpreadsheetApp.getUi().showModalDialog(
    html,
    "Orchid Management System - Documentation & Manual"
  );
}