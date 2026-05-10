/**
 * PROJECT: The Potted Historian - Greenhouse Status
 * PURPOSE: Data retrieval for the sidebar dashboard.
 */


function doGet(e) {
  // Ensure the HTML file name in your project is 'GreenhouseEmbed'
  return HtmlService.createTemplateFromFile('GreenhouseEmbed')
    .evaluate()
    .setTitle('Greenhouse Status')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/**
 * Pulls dynamic data from the 'Greenhouse Dashboard' sheet.
 */
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Greenhouse Dashboard");
 
  // Pulls cells C4 to C11 as an array
  const data = sheet.getRange("C4:C11").getValues();
 
  return {
    total: data[0][0],       // Cell C4: Total Collection
    bloom: data[1][0],       // Cell C5: Currently in Bloom
    
    // CHANGED TO data[2] to target Cell C6
    lastWater: data[2][0] instanceof Date ? 
               Utilities.formatDate(data[2][0], Session.getScriptTimeZone(), "MMMM dd, yyyy") : 
               data[2][0],   // Cell C6: Last Watering Day
               
    researchPct: (data[6][0] * 100).toFixed(2), // Cell C10: Research Progress %
    totalHours: data[7][0]   // Cell C11: Total Project Time
  };
}
