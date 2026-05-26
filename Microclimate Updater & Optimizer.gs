/**
 * ============================================================================
 * ORCHID MANAGEMENT SYSTEM - MICROCLIMATE UPDATER & OPTIMIZER
 * ============================================================================
 * Creative Studio: SCA Webministry / The Potted Historian
 * Repository: Orchid Management System (History-First Architecture)
 * File: Current Orchid Inventory Workbook
 * * Description:
 * Unifies historical microclimate feeds from the centralized "Climate Logs"
 * workbook into the active production sheets. Implements an automatic 90-day
 * seasonal lifecycle pruning sequence to maximize dashboard responsiveness.
 * ============================================================================
 */

// Centralized Configuration Constants
const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();
const CLIMATE_LOGS_SPREADSHEET_ID = SCRIPT_PROPERTIES.getProperty('CLIMATE_LOGS_ID');
const RETENTION_DAYS = 90;

function syncMultiSheetClimateLogs() {
  Logger.log('Initializing cross-workbook microclimate synchronization...');
  
  const localSs = SpreadsheetApp.getActiveSpreadsheet();
  const officeDest = localSs.getSheetByName('Office Cabinet Data');
  const livingDest = localSs.getSheetByName('Living Room Cabinet Data');
  
  // Establish connection to the remote tracking resource
  const remoteSs = SpreadsheetApp.openById(CLIMATE_LOGS_SPREADSHEET_ID);
  
  if (!officeDest || !livingDest) {
    throw new Error('Critical Error: Production target tracking sheets are missing from this workbook.');
  }
  
  // Fetch local tracking boundaries to enforce data integrity and block duplicate rows
  const lastOfficeTime = getLastTimestampMark(officeDest);
  const lastLivingTime = getLastTimestampMark(livingDest);
  
  Logger.log(`Sync Boundaries - Office Last Row: ${lastOfficeTime.toISOString()}, Living Room Last Row: ${lastLivingTime.toISOString()}`);
  
  let officeBuffer = [];
  let livingBuffer = [];
  
  // ==========================================================================
  // SEGMENT 1: PARSE SEPARATE GOVEE SHEETS (3-Column Layout)
  // ==========================================================================
  const goveeSources = [
    { name: 'Govee Milsbo',      destination: 'living', label: 'Milsbo Ambient' },
    { name: 'Govee Rudsta A',    destination: 'office', label: 'Rudsta A Ambient' },
    { name: 'Govee Rudsta B',    destination: 'office', label: 'Rudsta B Ambient' },
    { name: 'Govee Terrarium',   destination: 'office', label: 'Terrarium Ambient' }
  ];
  
  goveeSources.forEach(function(source) {
    const sheet = remoteSs.getSheetByName(source.name);
    if (!sheet) {
      Logger.log(`Status Check: Worksheet "${source.name}" has not logged data yet.`);
      return;
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return; // Worksheet contains headers only
    
    // Read lines chronologically
    for (let i = 1; i < data.length; i++) {
      let row = data[i];
      let timestamp = new Date(row[0]);
      let temp = row[1];
      let humidity = row[2];
      
      if (isNaN(timestamp.getTime())) continue; // Skip malformed date rows
      
      // Enforce data sync boundaries
      if (source.destination === 'office' && timestamp > lastOfficeTime) {
        officeBuffer.push([timestamp, temp, humidity, source.label]);
      } else if (source.destination === 'living' && timestamp > lastLivingTime) {
        livingBuffer.push([timestamp, temp, humidity, source.label]);
      }
    }
  });
  
  // ==========================================================================
  // SEGMENT 2: PARSE ORIGINAL AC INFINITY LOG (5-Column Layout)
  // ==========================================================================
  const acInfinitySheet = remoteSs.getSheetByName('ACInfinityLog');
  if (acInfinitySheet) {
    const acData = acInfinitySheet.getDataRange().getValues();
    if (acData.length > 1) {
      for (let i = 1; i < acData.length; i++) {
        let row = acData[i];
        let timestamp = new Date(row[0]);
        let cabinetTag = String(row[1]).toLowerCase();
        let temp = row[2];
        let humidity = row[3];
        
        if (isNaN(timestamp.getTime())) continue;
        
        // Sort AC Infinity entries to their matching spatial paths
        if ((cabinetTag.includes('rudsta') || cabinetTag.includes('terrarium')) && timestamp > lastOfficeTime) {
          officeBuffer.push([timestamp, temp, humidity, row[1]]);
        } else if (cabinetTag.includes('milsbo') && timestamp > lastLivingTime) {
          livingBuffer.push([timestamp, temp, humidity, row[1]]);
        }
      }
    }
  }
  
  // ==========================================================================
  // SEGMENT 3: SAFE BATCH WRITE EXECUTIONS
  // ==========================================================================
  if (officeBuffer.length > 0) {
    officeDest.getRange(officeDest.getLastRow() + 1, 1, officeBuffer.length, officeBuffer[0].length).setValues(officeBuffer);
    Logger.log(`Sync Engine Success: Appended ${officeBuffer.length} clean records to Office Cabinet Data.`);
  } else {
    Logger.log('Sync Engine Check: No new Office records detected.');
  }
  
  if (livingBuffer.length > 0) {
    livingDest.getRange(livingDest.getLastRow() + 1, 1, livingBuffer.length, livingBuffer[0].length).setValues(livingBuffer);
    Logger.log(`Sync Engine Success: Appended ${livingBuffer.length} clean records to Living Room Cabinet Data.`);
  } else {
    Logger.log('Sync Engine Check: No new Living Room records detected.');
  }

  // ==========================================================================
  // SEGMENT 4: LIFECYCLE MANAGEMENT ENGINE (90-Day Seasonal Pruning)
  // ==========================================================================
  pruneExpiredClimateLogs(localSs);
}

/**
 * Sweeps the active log worksheets and cleanses records exceeding the retention threshold.
 * Scans top-down to efficiently drop expired arrays in a single structural operation.
 * * @param {Spreadsheet} localSs The active spreadsheet instance to evaluate
 */
function pruneExpiredClimateLogs(localSs) {
  const targetSheets = ['Office Cabinet Data', 'Living Room Cabinet Data'];
  const now = new Date();
  const cutoffTimestamp = now.getTime() - (RETENTION_DAYS * 24 * 60 * 60 * 1000);
  
  Logger.log(`Executing architectural lifecycle purge. Cutoff threshold: ${new Date(cutoffTimestamp).toISOString()}`);
  
  targetSheets.forEach(function(sheetName) {
    const sheet = localSs.getSheetByName(sheetName);
    if (!sheet) return;
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return; // Headers only
    
    // Harvest the entire Column A index for rapid memory comparison
    const timestampRange = sheet.getRange(1, 1, lastRow, 1).getValues();
    let rowsToPurge = 0;
    
    for (let i = 1; i < timestampRange.length; i++) {
      let entryDate = new Date(timestampRange[i][0]);
      if (!isNaN(entryDate.getTime()) && entryDate.getTime() < cutoffTimestamp) {
        rowsToPurge++;
      } else {
        // Enforce an immediate exit loop the millisecond active logs are detected
        break; 
      }
    }
    
    if (rowsToPurge > 0) {
      sheet.deleteRows(2, rowsToPurge);
      Logger.log(`Optimization Engine: Dropped ${rowsToPurge} stale rows from "${sheetName}".`);
    } else {
      Logger.log(`Optimization Engine: "${sheetName}" database remains within standard boundaries.`);
    }
  });
}

/**
 * Helper: Grabs the absolute latest logged chronological entry point from a target tab
 * @param {Sheet} sheet The active sheet instance to analyze
 * @return {Date} The last logged timestamp, or epoch start if empty
 */
function getLastTimestampMark(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return new Date(0);
  const val = sheet.getRange(lastRow, 1).getValue();
  return isNaN(new Date(val).getTime()) ? new Date(0) : new Date(val);
}