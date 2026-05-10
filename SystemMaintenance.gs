/**
 * -----------------------------------------------------------------------------
 * ORCHID INVENTORY - DRIVE ARCHIVE CLEANUP (SystemMaintenance.gs)
 * -----------------------------------------------------------------------------
 * Purpose: Deletes archived Govee CSV files from Google Drive that are 
 * older than 60 days to prevent folder clutter.
 * -----------------------------------------------------------------------------
 */

function cleanupGoveeArchiveFolder() {
  // Pulls the Folder ID from Globals.gs
  const archiveFolder = DriveApp.getFolderById(GOVEE_ARCHIVE_FOLDER_ID);
  const files = archiveFolder.getFiles();
  
  // Set the cutoff for 60 days ago
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - 60);
  
  let deletedCount = 0;

  while (files.hasNext()) {
    const file = files.next();
    
    /**
     * TRASH LOGIC:
     * If the file was created before the cutoff date, move it to the trash.
     * We use setTrashed(true) rather than permanent deletion to allow for
     * manual recovery if a sync error is discovered later.
     */
    if (file.getDateCreated() < cutoffDate) {
      file.setTrashed(true); 
      deletedCount++;
    }
  }

  // Logs the result to the Apps Script Execution Log
  console.log(`Maintenance Complete: ${deletedCount} old Govee CSVs moved to Trash.`);
}