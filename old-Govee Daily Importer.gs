/**
 * -----------------------------------------------------------------------------
 * ORCHID MANAGEMENT SYSTEM - INTEGRATED SYNC & LONG-TERM ARCHIVE
 * -----------------------------------------------------------------------------
 * Version: 9.1.0 - Clean Deployment
 * -----------------------------------------------------------------------------
 */

const SCRIPT_PROP = PropertiesService.getScriptProperties();
const FOLDER_ID = SCRIPT_PROP.getProperty('GOVEE_FOLDER_ID');
const ARCHIVE_ID = SCRIPT_PROP.getProperty('GOVEE_ARCHIVE_ID');
const LABEL_NAME = SCRIPT_PROP.getProperty('GMAIL_LABEL');

function importGoveeDaily() {
  let ui;
  try { ui = SpreadsheetApp.getUi(); } catch (e) { ui = null; }
  
  if (!FOLDER_ID || !LABEL_NAME) return;

  try {
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const label = GmailApp.getUserLabelByName(LABEL_NAME);
    if (!label) return;

    // --- PHASE 1: HARVEST ---
    const threads = label.getThreads();
    let newFilesCount = 0;

    threads.forEach(thread => {
      const messages = thread.getMessages();
      let threadHasData = false;
      messages.forEach(message => {
        if (message.isUnread()) {
          const attachments = message.getAttachments();
          attachments.forEach(attachment => {
            if (attachment.getName().endsWith('.csv')) {
              folder.createFile(attachment);
              newFilesCount++;
              threadHasData = true;
            }
          });
          message.markRead();
        }
      });
      if (threadHasData) thread.moveToArchive();
    });

    // --- PHASE 2: PROCESS ---
    processFolderContents(folder);

  } catch (e) {
    console.error("Govee Sync Failed: " + e.toString());
    if (ui) ui.alert("Govee Sync Error: " + e.message);
  }
}

function processFolderContents(folder) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const files = folder.getFiles();
  const archiveFolder = DriveApp.getFolderById(ARCHIVE_ID);

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    
    if (file.getParents().next().getId() === FOLDER_ID) {
      console.log(`Processing: ${fileName}`);
      const csvData = Utilities.parseCsv(file.getBlob().getDataAsString());
      let mapping = getMapping(fileName);
      
      if (mapping.sheetName) {
        const rollingSheet = ss.getSheetByName(mapping.sheetName);
        const dataLogSheet = ss.getSheetByName(mapping.sheetName + " Data");
        
        if (rollingSheet && dataLogSheet) {
          let rollingRows = [];
          let logRows = [];

          for (let i = 1; i < csvData.length; i++) {
            const row = csvData[i];
            if (!row[0]) continue;
            const timestamp = new Date(row[0]);
            const hour = timestamp.getHours();

            if ([9, 15, 21].includes(hour)) {
              rollingRows.push([timestamp, row[1], row[2]]);
              logRows.push([timestamp, row[1], row[2], mapping.sensorLabel]);
            }
          }

          if (rollingRows.length > 0) {
            let lastRowRolling = getLastRowInColumn(rollingSheet, mapping.col);
            rollingSheet.getRange(lastRowRolling + 1, mapping.col, rollingRows.length, 3).setValues(rollingRows);
          }

          if (logRows.length > 0) {
            dataLogSheet.getRange(dataLogSheet.getLastRow() + 1, 1, logRows.length, 4).setValues(logRows);
          }
        }
      }
      file.moveTo(archiveFolder);
    }
  }
}

function getMapping(fileName) {
  let result = { sheetName: "", col: 1, sensorLabel: "" };
  if (fileName.includes("Milsbo")) {
    result.sheetName = "Living Room Cabinet";
    result.sensorLabel = fileName.includes("Upper") ? "Milsbo Upper" : "Milsbo Lower";
    result.col = fileName.includes("Upper") ? 1 : 5;
  } else if (fileName.includes("Rudsta") || fileName.includes("Terrarium")) {
    result.sheetName = "Office Cabinet";
    if (fileName.includes("top")) { result.col = 1; result.sensorLabel = "Rudsta Top"; }
    else if (fileName.includes("mid")) { result.col = 5; result.sensorLabel = "Rudsta Mid"; }
    else if (fileName.includes("bottom")) { result.col = 9; result.sensorLabel = "Rudsta Bottom"; }
    else if (fileName.includes("Terrarium")) { result.col = 13; result.sensorLabel = "Terrarium"; }
  }
  return result;
}

function getLastRowInColumn(sheet, column) {
  const data = sheet.getRange(1, column, sheet.getMaxRows()).getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][0] !== "") return i + 1;
  }
  return 1;
}