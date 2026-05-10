# Orchid Management System: The Satyrion Chronicles
**Version:** 5.1.0  
**Archivist:** R. Short (Ailis inghean Ui Riagain)  
**Affiliations:** The Potted Historian | Mystic Oaks Trading Co. | SCA Barony of the Sacred Stone

---

## 📜 Project Overview
This repository contains the core automation engine for a technical **"Digital Twin"** architecture. It is designed to manage a living collection of over 50 orchid specimens while supporting pre-17th-century botanical research. 

By integrating modern environmental sensor data with historical nomenclature, this system ensures that each specimen—from common hybrids to rare historical species—is tracked with archival precision.

## 🚀 Technical Features

### 1. High-Fidelity Climate Monitoring
- **Govee Cloud Integration:** Daily sampling at **09:00, 15:00, and 21:00** to capture diurnal temperature and humidity shifts.
- **Multi-Cabinet Support:** Independent tracking for IKEA Milsbo and Rudsta grow cabinets, as well as a specialized lizard terrarium.
- **30-Day Rolling Parity:** Automated data rotation to maintain performance while preserving historical trends.

### 2. Archival Inventory Logic
- **Automated Provisioning:** Generates unique Serial IDs and dedicated tracking tabs for new acquisitions.
- **Maintenance Lifecycle:** Standardized logging for watering, repotting, and "bloom status" cycles.
- **Sheet Recycling Protocol:** Logic to stay within Google Sheets limitations by decommissioning and repurposing temporary "TMP_" tabs.

### 3. Custom Navigator Sidebar
- An HTML5/CSS3 integrated sidebar that allows for rapid navigation between "Inventory," "Maintenance," and "Research" views without manual tab scrolling.

## 🛠️ Security & Architecture
- **Zero-Footprint Sanitization:** This codebase utilizes `PropertiesService` for all sensitive Folder and Spreadsheet IDs.
- **Environment:** Google Apps Script (V8 Engine).
- **Relational Mapping:** 0-indexed column mapping for high-speed data retrieval across the "Inventory" and "Keiki Registry" workbooks.

## 🌿 Scholarly Intent
The data collected by this system directly informs the research for **The Satyrion Chronicles**, exploring the intersection of medieval botanical folklore and modern orchid care. 


#### *“As the bee tends the hive, so the steward tends the scroll.”* 



**© 2026 The Potted Historian. All Rights Reserved.**

*This code is provided for educational and archival purposes. Unauthorized redistribution or commercial use is prohibited.*
