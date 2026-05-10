/**
 * =============================================================================
 * ORCHID MANAGEMENT SYSTEM - TECHNICAL README
 * =============================================================================
 * PROJECT: The Satyrion Chronicles Analytical Engine
 * INCEPTION DATE: 2026-01-25
 * LAST AUDIT DATE: 2026-05-08
 * REPOSITORY: Google Apps Script (GAS) - Distributed Architecture
 * -----------------------------------------------------------------------------
 * * 🖋️ PROJECT PURPOSE
 * -----------------------------------------------------------------------------
 * A high-fidelity digital twin for professional orchid cultivation and 
 * pre-17th-century botanical research. This system acts as the primary data 
 * source for "The Satyrion Chronicles" historical and analytical series.
 *
 * CORE OBJECTIVES:
 * 1. AUTOMATED LIFECYCLE TRACKING: Dynamic repotting calculations based on 
 *    species-specific growth rates (Column S) and historical care data.
 *
 * 2. MULTI-ZONE CLIMATE INTEGRATION: Real-time telemetry for Milsbo (Living Room), 
 *    Rudsta (Office), and Lizard Terrarium environments via Govee Cloud API.
 *
 * 3. SHEET RECYCLING & PROVISIONING: Optimizes workbook performance by 
 *    reclaiming decommissioned ID sheets (Archive.gs) for new acquisitions.
 *
 * 4. SCHOLARLY DOCUMENTATION: Maintains a permanent archive of bloom cycles 
 *    and cultivation milestones for long-term botanical study.
 * -----------------------------------------------------------------------------
 * * 🛠️ SYSTEM ARCHITECTURE & FILE DIRECTORY
 * -----------------------------------------------------------------------------
 * 🟢 TIER 1: THE WORKBOOK ENGINE (Container-Bound)
 * - main-onEdit.gs: Primary event-driven sync; anchors logs to Row 56.
 * - globals.gs: Centralized ID management and API endpoint configuration.
 * - Archive.gs: "Sheet Recycling" logic; renames dead sheets to "TMP_[ID]".
 * - Govee daily importer.gs: Hourly fetch of 30-day rolling climate samples.
 * - Health-Diagnostics.gs: Validates column alignment (Col I/S) across 50+ tabs.
 * - Inventory Logic.gs: Core provisioning and target Foot Candle (Col O) mapping.
 *
 * 🔵 TIER 2: THE ANALYTICAL DASHBOARD (Standalone Web App)
 * - webapp.gs: Data aggregation; bridges Master and Keiki registries.
 * - inventory_index.html: Frontend UI; real-time search and milestone rotator.
 * -----------------------------------------------------------------------------
 * * 📊 DATA FLOW REFERENCE (COLUMN & CELL MAPPING)
 * -----------------------------------------------------------------------------
 * 1. WATERING & NUTRITION (Maintenance Log -> ID Sheets)
 * -> Edit Col C/E triggers Timestamp in Col D.
 * -> Appends nutritonal data (e.g., Well Water Mix) to ID Sheet [A56+].
 *
 * 2. REPOTTING (Maintenance Log -> Global Sync)
 * -> Edit Col H/I updates ID Sheet [A38/B38].
 * -> Calculates Next Due for: Log [Col G], Inventory [Col T], ID Sheet [D11].
 *
 * 3. BLOOM STATUS (Inventory -> ID Sheets)
 * -> Edit Col K updates ID Sheet [D5] and appends to lifetime log [A20:A].
 *
 * 4. CLIMATE (Govee API -> Dashboard)
 * -> Aggregates 9AM, 3PM, and 9PM samples from Living Room, Office, & Terrarium.
 * -> Feeds rolling 48hr averages to the web-facing dashboard.
 * -----------------------------------------------------------------------------
 * * ⚠️ DEVELOPER & MAINTENANCE NOTES
 * -----------------------------------------------------------------------------
 * - RECYCLING PROTOCOL: Never delete ID sheets manually. Use the Archive 
 *   utility to mark them as "TMP_". The Provisioning engine checks for these 
 *   before creating new sheets to prevent Google Sheets tab limits.
 * - FOOT CANDLE TARGETS: Column O in the Master Inventory is the primary 
 *   lookup for light intensity requirements during new specimen setup.
 * - SYNC PARITY: The Dashboard utilizes a Time-Driven Trigger for Keiki 
 *   registry parity. Manual runs of 'getDashboardData' resolve sync hangs.
 * * =============================================================================
 */

/**
 * -----------------------------------------------------------------------------
 * * 🆘 TROUBLESHOOTING & CONNECTION RECOVERY
 * -----------------------------------------------------------------------------
 * 1. "COLUMN MISALIGNMENT" ERROR:
 * -> Run 'runFullSystemDiagnostic' from Health-Diagnostics.gs to verify that 
 *    Log Column I (Maintenance) and Inventory Column S (Frequency) match.
 *
 * 2. BLANK CLIMATE READOUTS:
 * -> Verify Govee API Key in globals.gs. Ensure "Lizard Terrarium" and 
 *    "Office Cabinet" tabs haven't been renamed, as this breaks the importer.
 *
 * 3. NEW SPECIMEN FAILURES:
 * -> Check for available "TMP_" sheets. If none exist, ensure 'provisionNewOrchid' 
 *    has permission to create new tabs based on the "Template" sheet.
 * -----------------------------------------------------------------------------
 * * 📂 SYSTEM DIRECTORY (DRIVE & FOLDER IDS)
 * -----------------------------------------------------------------------------
 * - MASTER WORKBOOK ID: 1UrcoDPCyKMuMPp9-sYqYoSnZbNQZvxK09ISzvIGaSHo
 * - KEIKI REGISTRY ID: 1MDTYm3CUR5rDS--A0HWD_Q0VXpC6ExkaPfV73ypmhoI
 * - GOVEE ARCHIVE FOLDER: 1FKpXI0UavjJcwC3jYiqgeBV94MZjLwbT
 * - SYSTEM BACKUP FOLDER: [User to insert updated Backup Folder ID]
 * -----------------------------------------------------------------------------
 * * 📜 MISSION STATEMENT
 * -----------------------------------------------------------------------------
 * This system serves as the definitive analytical record for the collection. 
 * Its purpose is to bridge the gap between daily greenhouse maintenance and 
 * long-term botanical documentation. Through automated tracking of repotting 
 * cycles, bloom history, and climate data, the system maintains a high-fidelity 
 * archive of plant health to support ongoing research and conservation efforts.
 * =============================================================================
 */