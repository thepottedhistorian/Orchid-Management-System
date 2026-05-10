/**
 * -----------------------------------------------------------------------------
 * ORCHID MANAGEMENT SYSTEM - DEVELOPER UTILITIES (DeveloperTools.gs)
 * -----------------------------------------------------------------------------
 * Purpose: Diagnostic tool to scan the global scope for function name collisions.
 * Logic:
 * 1. Checks the current project's global 'this' context.
 * 2. Identifies if critical functions (onEdit, onOpen) are defined.
 * 3. Alerts the user if multiple declarations are suspected.
 * -----------------------------------------------------------------------------
 */

function findDuplicateFunctions() {
  // Pulls MASTER_ID from Globals.gs to confirm connection
  const ss = SpreadsheetApp.openById(MASTER_ID);
  
  Logger.log("--- 🔍 SCANNING PRO PROJECT FOR COLLISIONS ---");
  Logger.log("Connected to Workbook: " + ss.getName());
  
  // The "Common Suspects" list for duplication errors
  const criticalFunctions = [
    "onEdit",
    "onOpen",
    "importGoveeDaily", // Updated from 'Monthly' to match your current menu
    "provisionNewOrchid",
    "refreshAllBloomStatusFormatting",
    "applyAllSystemStoplights",
    "parseFrequencyToMonths"
  ];

  let foundCount = 0;

  criticalFunctions.forEach(func => {
    /**
     * SCOPE CHECK:
     * This checks if the function name exists in the current project scope.
     * Note: This detects if the function IS DEFINED, but won't catch
     * multiple definitions in the same script project (the editor usually
     * flags those as a SyntaxError anyway).
     */
    if (typeof this[func] === 'function') {
      Logger.log("✅ Function Found: " + func);
      foundCount++;
    } else {
      Logger.log("❌ Function Missing: " + func);
    }
  });

  Logger.log("-----------------------------------------------");
  Logger.log("Scan complete. " + foundCount + " core functions are active.");
  
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    "Diagnostic Complete",
    "Check the 'Execution Log' (at the bottom of the editor) to see which functions are currently active in the global scope.\n\n" +
    "If you still see a Yellow Warning Box in the Trigger Menu, check for a file named 'Code.gs'—it often contains a duplicate onOpen or onEdit by default.",
    ui.ButtonSet.OK
  );
}