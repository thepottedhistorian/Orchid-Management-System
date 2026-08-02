/**
 * ============================================================================
 * ORCHID INVENTORY WORKBOOK EXPORT UTILITY (V6 - BULK OPTIMIZED)
 * Author: Digital Steward / System Automation
 * File: ExportToXLSX.gs
 * Description: High-speed flattened copy generator for Excel ingestion.
 *              Uses bulk range operations to prevent Apps Script timeouts.
 * ============================================================================
 */

function exportWorkbookToXLSX() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const originalName = ss.getName();
  
  // --------------------------------------------------------------------------
  // 1. CREATE TEMPORARY DRIVE COPY
  // --------------------------------------------------------------------------
  const tempFile = DriveApp.getFileById(ss.getId()).makeCopy(`${originalName}_Excel_Export`);
  const tempSs = SpreadsheetApp.openById(tempFile.getId());
  const sheets = tempSs.getSheets();

  // --------------------------------------------------------------------------
  // 2. BATCH CLEANING & FLATTENING LOOP
  // --------------------------------------------------------------------------
  sheets.forEach(sheet => {
    const maxRows = sheet.getMaxRows();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    // Clear conditional formatting rules (stoplights/colors)
    sheet.clearConditionalFormatRules();

    if (lastRow > 0 && lastCol > 0) {
      const range = sheet.getRange(1, 1, lastRow, lastCol);
      
      // Flatten dynamic formulas to static values
      range.setValues(range.getValues());
      
      // Clear Data Validation (checkboxes/dropdowns)
      range.clearDataValidations();

      // Safely unhide active data rows
      sheet.unhideRow(sheet.getRange(1, 1, lastRow, 1));
    }

    // FAST BULK OPERATION: Set row height for all rows in a single API call
    if (maxRows > 0) {
      sheet.setRowHeights(1, maxRows, 21);
    }

    // Remove drawing objects or floating overlays
    const drawings = sheet.getDrawings();
    drawings.forEach(drawing => drawing.remove());
  });

  SpreadsheetApp.flush();

  // --------------------------------------------------------------------------
  // 3. GENERATE DOWNLOAD DIALOG
  // --------------------------------------------------------------------------
  const url = `https://docs.google.com/spreadsheets/d/${tempFile.getId()}/export?format=xlsx`;
  
  const htmlOutput = HtmlService.createHtmlOutput(
    `<div style="font-family: Arial, sans-serif; padding: 15px;">
       <h3>Export Ready!</h3>
       <p>Your workbook has been flattened, unhidden, and scrubbed for Excel export.</p>
       <p><a href="${url}" target="_blank" style="background-color: #2e7d32; color: white; padding: 10px 15px; text-decoration: none; border-radius: 4px; display: inline-block;">Download XLSX File</a></p>
       <p style="font-size: 11px; color: #666;">Note: You can safely delete the temporary file from Google Drive after downloading.</p>
     </div>`
  ).setWidth(400).setHeight(200);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Export to Microsoft Excel');
}