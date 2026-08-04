/**
 * ==============================================================================
 * PROGRAM NAME   : Botanical Archivist (Orchid Collection Ledger)
 * MODULE NAME    : Full Collection Harvester & Staging Sync
 * AUTHOR         : The Potted Historian
 * SCRIPT VERSION : 9.0.0 (Production Release)
 * DESCRIPTION    : Extracts core metadata from Sheet39 into Import_Manifest,
 *                  parses sub-table logs (Bloom, Repotting, Watering, Notes) 
 *                  from ID sheets 1-65, and populates Maintenance_Log and 
 *                  Orchid_Specimen_Template with Orchid ID attribution.
 * ==============================================================================
 */

function runFullOrchidStagingSync() {
  const STAGING_SHEET_ID = "1C15_IpkCYYkN5N0NI8xS2HBfZm9PSo2nMh4pzEqBBeQ";
  const MANIFEST_TAB_NAME = "Import_Manifest";
  const MAINTENANCE_TAB_NAME = "Maintenance_Log";
  const TEMPLATE_TAB_NAME = "Orchid_Specimen_Template";

  const sourceSS = SpreadsheetApp.getActiveSpreadsheet();
  let targetSS;

  try {
    targetSS = SpreadsheetApp.openById(STAGING_SHEET_ID);
  } catch (err) {
    SpreadsheetApp.getUi().alert("Connection Error: Unable to access the Manifest Import workbook.");
    return;
  }

  const manifestSheet = targetSS.getSheetByName(MANIFEST_TAB_NAME);
  const maintenanceSheet = targetSS.getSheetByName(MAINTENANCE_TAB_NAME);
  const templateSheet = targetSS.getSheetByName(TEMPLATE_TAB_NAME);

  if (!manifestSheet) {
    SpreadsheetApp.getUi().alert("Schema Error: Target tab '" + MANIFEST_TAB_NAME + "' was not found.");
    return;
  }

  function formatDate(val) {
    if (!val || val === "—" || val === "-" || val === "") return "—";
    if (val instanceof Date) {
      return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
    const strVal = String(val).trim();
    if (strVal.match(/^\d{4}-\d{2}-\d{2}/)) return strVal.substring(0, 10);
    
    const parsed = new Date(strVal);
    if (!isNaN(parsed.getTime())) {
      return Utilities.formatDate(parsed, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
    return strVal;
  }

  // --- 1. INDEX INDIVIDUAL ID SHEETS (1-65) ---
  const allSheets = sourceSS.getSheets();
  const idSheetMap = {};

  allSheets.forEach(sheet => {
    const sName = sheet.getName().trim().toLowerCase();
    if (sName.includes("sheet")) return;
    
    const numMatch = sName.match(/\d+/);
    if (numMatch) {
      const idNum = parseInt(numMatch[0], 10);
      if (idNum >= 1 && idNum <= 65) {
        idSheetMap[idNum] = sheet;
      }
    }
  });

  const specLogsMap = {};
  const cleanMaintenanceRows = [];
  const cleanTemplateRows = [];

  for (let idNum = 1; idNum <= 65; idNum++) {
    const sheet = idSheetMap[idNum];
    if (!sheet) continue;

    const lastR = Math.min(sheet.getLastRow(), 100);
    const lastC = Math.min(sheet.getLastColumn(), 15);
    if (lastR < 2 || lastC < 2) continue;

    const data = sheet.getRange(1, 1, lastR, lastC).getValues();

    let bloomLogs = [];
    let currentSection = "";

    for (let r = 0; r < data.length; r++) {
      const rowStr = data[r].join(" ").trim();

      if (rowStr.includes("Bloom Log")) { currentSection = "BLOOM"; continue; }
      if (rowStr.includes("Observances") || (rowStr.includes("Notes") && !rowStr.includes("Care Profile"))) { 
        currentSection = "OBSERVATIONS"; 
        continue; 
      }
      if (rowStr.includes("Repotting Log")) { currentSection = "REPOT"; continue; }
      if (rowStr.includes("Watering Log")) { currentSection = "WATER"; continue; }

      if (rowStr.includes("In Bloom") && rowStr.includes("Out of Bloom")) continue;
      if (rowStr.includes("Date") && (rowStr.includes("Method") || rowStr.includes("Notes"))) continue;

      const nonCols = data[r].filter(cell => String(cell).trim() !== "");
      if (nonCols.length === 0) continue;

      // Category: Bloom Log
      if (currentSection === "BLOOM" && nonCols.length >= 2) {
        const inDate = formatDate(nonCols[0]);
        const outDate = formatDate(nonCols[1]);
        const dur = nonCols[2] ? String(nonCols[2]).trim() : "—";
        if (inDate && inDate !== "—") {
          bloomLogs.push({ inDate: inDate, outDate: outDate, duration: dur });
          
          cleanMaintenanceRows.push([idNum, inDate, "Bloom Event", "Status Update", `Out of Bloom: ${outDate} | Duration: ${dur}`, "System Import", "—"]);
          cleanTemplateRows.push([idNum, inDate, "Bloom Log", "Status Update", inDate, outDate, dur]);
        }
      }

      // Category: Observances / Notes
      if (currentSection === "OBSERVATIONS" && nonCols.length >= 2) {
        const dt = formatDate(nonCols[0]);
        const note = String(nonCols[1]).trim();
        if (dt && note && dt !== "—") {
          cleanMaintenanceRows.push([idNum, dt, "Observation", "General Note", note, "System Import", "—"]);
          cleanTemplateRows.push([idNum, dt, "Notes", "General Note", dt, "—", note]);
        }
      }

      // Category: Repotting Log
      if (currentSection === "REPOT" && nonCols.length >= 2) {
        const dt = formatDate(nonCols[0]);
        const note = String(nonCols[1]).trim();
        if (dt && dt !== "—") {
          cleanMaintenanceRows.push([idNum, dt, "Repotting", "Media Refresh", note, "System Import", "—"]);
          cleanTemplateRows.push([idNum, dt, "Repotting Log", "Media Refresh", dt, "—", note]);
        }
      }

      // Category: Watering Log
      if (currentSection === "WATER" && nonCols.length >= 2) {
        const dt = formatDate(nonCols[0]);
        const method = String(nonCols[1]).trim();
        const note = nonCols[2] ? String(nonCols[2]).trim() : "";
        if (dt && dt !== "—") {
          cleanMaintenanceRows.push([idNum, dt, "Watering", method, note || "—", "System Import", "—"]);
          cleanTemplateRows.push([idNum, dt, "Watering Log", method, dt, "—", note || "Standard Water Event"]);
        }
      }
    }

    specLogsMap[idNum] = { latestBloom: bloomLogs.length > 0 ? bloomLogs[0] : null };
  }

  // --- 2. HARVEST MASTER METADATA FROM SHEET39 ---
  const sheet39 = sourceSS.getSheetByName("Sheet39") || sourceSS.getSheetByName("Sheet 39");
  const cleanManifestRows = [];

  if (sheet39) {
    const s39LastR = Math.min(sheet39.getLastRow(), 200);
    const s39Data = sheet39.getRange(1, 1, s39LastR, 25).getValues();
    const headers = s39Data[0].map(h => String(h).trim().toLowerCase());

    function getColIdx(possibleNames) {
      return headers.findIndex(h => possibleNames.some(p => h.includes(p.toLowerCase())));
    }

    const idxId = getColIdx(["id"]);
    const idxName = getColIdx(["orchid name", "name"]);
    const idxGenus = getColIdx(["genus"]);
    const idxHabitat = getColIdx(["growth habitat", "habitat"]);
    const idxLineage = getColIdx(["species/hybrid lineage", "species/hybrid", "lineage"]);
    const idxEndemic = getColIdx(["endemic to", "native range", "range"]);
    const idxAcqDate = getColIdx(["acquisition date", "acq date"]);
    const idxSource = getColIdx(["source", "vendor"]);
    const idxBloomStatus = getColIdx(["in bloom", "bloom status"]);
    const idxLastBloom = getColIdx(["last bloom date", "last bloom"]);
    const idxLight = getColIdx(["light needs", "light"]);
    const idxWater = getColIdx(["watering needs", "water"]);
    const idxHumidity = getColIdx(["humidity"]);
    const idxMedia = getColIdx(["potting media", "pot type/media", "medium"]);
    const idxFert = getColIdx(["fertilizer routine", "fertilizer"]);
    const idxRepotFreq = getColIdx(["repot every", "repotting schedule"]);
    const idxNextRepot = getColIdx(["next repot due"]);

    for (let r = 1; r < s39Data.length; r++) {
      const row = s39Data[r];
      if (!row[idxName] && !row[idxId]) continue;

      const specId = idxId !== -1 && row[idxId] ? parseInt(row[idxId], 10) : r;
      if (isNaN(specId) || specId < 1 || specId > 65) continue;

      const genus = idxGenus !== -1 && row[idxGenus] ? String(row[idxGenus]).trim() : "Phalaenopsis";
      const name = idxName !== -1 && row[idxName] ? String(row[idxName]).trim() : `Orchid #${specId}`;
      const lineage = idxLineage !== -1 && row[idxLineage] ? String(row[idxLineage]).trim() : "complex hybrid";
      const habitat = idxHabitat !== -1 && row[idxHabitat] ? String(row[idxHabitat]).trim() : "Epiphyte";
      const range = idxEndemic !== -1 && row[idxEndemic] ? String(row[idxEndemic]).trim() : "Hybrid";
      const acqDate = idxAcqDate !== -1 ? formatDate(row[idxAcqDate]) : "—";
      const vendor = idxSource !== -1 && row[idxSource] ? String(row[idxSource]).trim() : "—";
      const cost = "—";
      const bloomStatus = idxBloomStatus !== -1 && row[idxBloomStatus] ? String(row[idxBloomStatus]).trim() : "Not in Bloom";
      const lastBloom = idxLastBloom !== -1 ? formatDate(row[idxLastBloom]) : "—";

      const light = idxLight !== -1 && row[idxLight] ? String(row[idxLight]).trim() : "Low (500–1,000 fc)";
      const water = idxWater !== -1 && row[idxWater] ? String(row[idxWater]).trim() : "Wet-Dry Cycle";
      const media = idxMedia !== -1 && row[idxMedia] ? String(row[idxMedia]).trim() : "Fir bark + perlite + charcoal";
      const humidity = idxHumidity !== -1 && row[idxHumidity] ? String(row[idxHumidity]).trim() : "50%–70% RH";
      const fert = idxFert !== -1 && row[idxFert] ? String(row[idxFert]).trim() : "1–2 weeks";
      const repotFreq = idxRepotFreq !== -1 && row[idxRepotFreq] ? String(row[idxRepotFreq]).trim() : "1–2 years";
      const nextRepot = idxNextRepot !== -1 ? formatDate(row[idxNextRepot]) : "—";

      const logData = specLogsMap[specId] || { latestBloom: null };

      cleanManifestRows.push([
        specId,
        genus,
        name,
        lineage,
        habitat,
        range,
        acqDate,
        vendor,
        cost,
        bloomStatus,
        lastBloom,
        logData.latestBloom ? logData.latestBloom.inDate : "—",
        logData.latestBloom ? logData.latestBloom.outDate : "—",
        logData.latestBloom ? logData.latestBloom.duration : "—",
        light,
        water,
        media,
        humidity,
        fert,
        repotFreq,
        nextRepot,
        "" // Clean empty placeholder for observation text
      ]);
    }
  }

  // --- 3. WRITE TO STAGING WORKBOOK TABS ---
  if (cleanManifestRows.length > 0) {
    const lastManifestRow = Math.max(manifestSheet.getLastRow(), 2);
    manifestSheet.getRange(2, 1, lastManifestRow, 22).clearContent();
    manifestSheet.getRange(2, 1, cleanManifestRows.length, 22).setValues(cleanManifestRows);

    if (maintenanceSheet && cleanMaintenanceRows.length > 0) {
      const lastMaintRow = Math.max(maintenanceSheet.getLastRow(), 2);
      maintenanceSheet.getRange(2, 1, lastMaintRow, 7).clearContent();
      maintenanceSheet.getRange(2, 1, cleanMaintenanceRows.length, 7).setValues(cleanMaintenanceRows);
    }

    if (templateSheet && cleanTemplateRows.length > 0) {
      const lastTmplRow = Math.max(templateSheet.getLastRow(), 2);
      templateSheet.getRange(2, 1, lastTmplRow, 7).clearContent();
      templateSheet.getRange(2, 1, cleanTemplateRows.length, 7).setValues(cleanTemplateRows);
    }

    SpreadsheetApp.getUi().alert(
      "Staging Sync Complete!\n\n" +
      "• Import_Manifest: " + cleanManifestRows.length + " clean profiles staged.\n" +
      "• Maintenance_Log: " + cleanMaintenanceRows.length + " activity events staged.\n" +
      "• Orchid_Specimen_Template: " + cleanTemplateRows.length + " categorized event rows staged with Orchid IDs."
    );
  } else {
    SpreadsheetApp.getUi().alert("Sync Warning: No master records found on Sheet39.");
  }
}