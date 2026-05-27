/**
 * -----------------------------------------------------------------------------
 * PROJECT: Greenhouse Inventory - webapp.gs (OFFICIAL)
 * -----------------------------------------------------------------------------
 * Section I: Web Service Handlers
 * Section II: Data Aggregation Engine (Master + Keiki)
 * Section III: Climate Analytics
 * Section IV: Changelog Retrieval
 * -----------------------------------------------------------------------------
 */

/**
 * SECTION I: STARTUP
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('inventory_index')
    .evaluate()
    .setTitle('The Potted Historian | Greenhouse Inventory')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * SECTION II: DATA AGGREGATION ENGINE
 */
function getDashboardData() {
  // Guard clause to ensure configuration exists
  if (typeof MASTER_SS_ID === 'undefined' || !MASTER_SS_ID) {
    throw new Error("Configuration Error: MASTER_SS_ID is not defined.");
  }

  const ss = SpreadsheetApp.openById(MASTER_SS_ID); 
  const masterInv = ss.getSheetByName(INVENTORY_SHEET);
  const resSheet = ss.getSheetByName(RESEARCH_SHEET_NAME);
  const dashSheet = ss.getSheetByName(DASH_SHEET_NAME);
  
  let collection = [];
  
  /**
   * Internal Helper: Processes rows from any workbook.
   * Uses displayValues for status, and getValues for IDs.
   */
  const processSheet = (sheet, sourceSS, isKeikiFile) => {
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return; 

    const data = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
    const displayValues = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getDisplayValues();
    
    for (let i = 1; i < data.length; i++) {
      const id = String(data[i][ID_COL]).trim();
      if (!id || isNaN(id) || !data[i][NAME_COL]) continue;
      
      const rawBloom = (displayValues[i][BLOOM_STATUS_COL] || "").trim();
      const isBlooming = /bloom|bud|spike/i.test(rawBloom) && !rawBloom.toLowerCase().includes("not in bloom");
      
      const nextRepotStr = displayValues[i][NEXT_COL];
      let repotColor = "#10b981"; 
      if (nextRepotStr) {
        const today = new Date();
        const nextDate = new Date(nextRepotStr);
        if (nextDate < today) repotColor = "#ef4444";
        else if ((nextDate - today) / (1000 * 60 * 60 * 24) < 30) repotColor = "#f59e0b";
      }

      let plantSheet = sourceSS.getSheetByName(id);
      let targetUrl = plantSheet ? `${sourceSS.getUrl()}#gid=${plantSheet.getSheetId()}` : sourceSS.getUrl();
      let bloomAvg = plantSheet ? plantSheet.getRange("B8").getDisplayValue() || "---" : "---";

      collection.push({
        id: parseInt(id),
        genus: data[i][GENUS_COL] || "---",
        name: data[i][NAME_COL] || "---",
        url: targetUrl,
        isBlooming: isBlooming,
        bloomDisplay: isBlooming ? rawBloom : "",
        acquisition: displayValues[i][ACQ_COL] || "---",
        lastBloom: displayValues[i][LAST_BLOOM_COL] || "---",
        bloomAvg: bloomAvg,
        repotColor: repotColor,
        isKeiki: isKeikiFile 
      });
    }
  };

  // Run the processing for Master
  processSheet(masterInv, ss, false);
  
  // Run the processing for Keiki
  try {
    const keikiSS = SpreadsheetApp.openById(KEIKI_SS_ID);
    processSheet(keikiSS.getSheetByName("Keiki Inventory"), keikiSS, true);
  } catch (e) { 
    console.warn("Keiki access skip: " + e.message); 
  }

  // Final sort: Keikis at bottom, sorted by ID
  collection.sort((a, b) => (a.isKeiki !== b.isKeiki) ? (a.isKeiki ? 1 : -1) : (a.id - b.id));

  // Climate and Stats mapping (Calls the external function)
  const climateStats = calculateClimateVitals(ss);
  
  // Return the complete object
  return {
    collection: collection,
    stats: {
      total: dashSheet.getRange("C4").getValue(),
      blooming: collection.filter(p => p.isBlooming).length,
      research: dashSheet.getRange("C10").getDisplayValue(),
      hours: Math.round(dashSheet.getRange("C11").getValue() || 0),
      temp: climateStats.avgTemp + "°F",
      hum: climateStats.avgHum,
      highLow: climateStats.highLowStr,
      vpd: climateStats.vpdStr
    },
    tasks: resSheet ? resSheet.getDataRange().getValues().filter(row => row[1] === true).map(row => row[2]) : []
  };
}
/**
 * SECTION III: CLIMATE ANALYTICS
 */
function calculateClimateVitals(ss) {
const sheets = ["Living Room Cabinet", "Office Cabinet"]; 
  
  let allTemps = [], allHums = [];
  const cutoff = new Date(new Date().getTime() - (24 * 60 * 60 * 1000));

  sheets.forEach(name => { // The code uses 'sheets' here
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    console.log("Checking sheet: " + name + " | Rows found: " + data.length); // ADD THIS
    data.slice(1).forEach(row => {
      if (row[0] instanceof Date && row[0] >= cutoff) {
        if (row[1] != null && row[1] !== "") allTemps.push(row[1]);
        if (row[2] != null && row[2] !== "") allHums.push(row[2]);
      }
    });
  });

  let acVpds = [];
  try {
    const climateSS = SpreadsheetApp.openById(CLIMATE_LOGS_SS_ID);
    const acSheet = climateSS.getSheetByName("ACInfinityLog");
    const acData = acSheet.getDataRange().getValues();
    acData.slice(1).forEach(row => {
      if (row[4] != null && row[4] !== "") acVpds.push(parseFloat(row[4]));
    });
  } catch (e) { console.warn("AC Infinity fetch failed: " + e.message); }

  if (allTemps.length === 0) return { avgTemp: "--", avgHum: "--", highLowStr: "Waiting for Sync...", vpdStr: "--" };

  const avgTemp = (allTemps.reduce((a, b) => a + b, 0) / allTemps.length).toFixed(1);
  const avgHum = (allHums.reduce((a, b) => parseFloat(b), 0) / allHums.length).toFixed(1);
  const hT = Math.max(...allTemps).toFixed(1);
  const lT = Math.min(...allTemps).toFixed(1);
  const hH = Math.max(...allHums).toFixed(0);
  const lH = Math.min(...allHums).toFixed(0);
  const avgVpd = acVpds.length > 0 ? (acVpds.reduce((a, b) => a + b, 0) / acVpds.length).toFixed(2) : "--";

  return { 
    avgTemp: avgTemp, 
    avgHum: avgHum, 
    highLowStr: `High: ${hT}°F / ${hH}% RH | Low: ${lT}°F / ${lH}% RH`,
    vpdStr: `Avg VPD: ${avgVpd} kPa`
  };
}

/**
 * SECTION IV: CHANGELOG RETRIEVAL
 */
function getChangelogFromServer() {
  return HtmlService.createHtmlOutputFromFile('ChangeLog').getContent();
}
