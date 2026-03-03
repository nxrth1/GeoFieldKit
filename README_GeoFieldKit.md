# 🌍 GeoFieldKit

> **A Python pipeline for processing, validating, and reporting on ODK geospatial field data collections.**

GeoFieldKit automates the full workflow from raw ODK/KoboToolbox exports to clean reports — tracking worker attendance, GPS data quality, payroll calculations, and interactive maps. Built for field survey teams collecting geospatial points using mobile data collection tools.

---

## 📌 Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Project Structure](#project-structure)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [Scripts](#scripts)
- [Configuration](#configuration)
- [Outputs](#outputs)
- [Use Cases](#use-cases)
- [Tech Stack](#tech-stack)
- [Contributing](#contributing)
- [License](#license)

---

## Overview

Field teams using ODK Collect or KoboToolbox generate large volumes of raw submission data — messy column names, duplicate GPS fields, missing values, and no easy way to track who worked, when, and how accurately. **GeoFieldKit** solves this with 7 focused Python scripts that take you from raw export to professional reports in minutes.

Designed for teams collecting **10 GPS points per worker per day** across field surveys, land use mapping, environmental monitoring, and similar data collection programs.

---

## ✨ Features

- 🧹 **Smart ODK Cleaning** — strips redundant system columns, renames fields, fixes types
- 📅 **Attendance Tracking** — clock-in/out analysis, early/late/absent flags, attendance matrix
- 💰 **Payroll Calculation** — daily pay with overtime rules, monthly summaries
- 📍 **GPS Quality Control** — accuracy checks, boundary validation, duplicate detection, PASS/REVIEW/FAIL ratings
- 📊 **Automated Reporting** — executive summary with embedded charts (hours, validity, late arrivals, weekly trend)
- 🗺️ **Interactive Maps** — colour-coded Leaflet maps, per-worker layers, heatmaps, GeoJSON/QGIS export
- ⚙️ **Full Automation** — folder watcher + email delivery when new ODK export lands

---

## 📁 Project Structure

```
GeoFieldKit/
│
├── script_1_cleaning.py        # ODK data cleaning & standardization
├── script_2_attendance.py      # Attendance & clock-in analysis
├── script_3_payroll.py         # Payroll calculation with overtime
├── script_4_gps_qc.py          # GPS & data quality control
├── script_5_reporting.py       # Excel reports with embedded charts
├── script_6_geospatial.py      # Interactive maps & GeoJSON export
├── script_7_automation.py      # Full pipeline automation + email
│
├── ODK_Raw_Export.xlsx          # ← Drop your ODK export here
│
├── outputs/                     # All generated reports go here
│   ├── cleaned_data.xlsx
│   ├── attendance_report.xlsx
│   ├── payroll_report.xlsx
│   ├── quality_control_report.xlsx
│   ├── full_report.xlsx
│   ├── map_all_points.html
│   ├── map_by_worker.html
│   ├── heatmap.html
│   └── points_export.geojson
│
├── incoming/                    # Drop new exports here for auto-processing
├── pipeline.log                 # Automatic run logs
└── README.md
```

---

## 🛠️ Installation

**Requirements:** Python 3.8+

```bash
# Clone the repository
git clone https://github.com/nxrth1/GeoFieldKit.git
cd GeoFieldKit

# Install dependencies
pip install pandas openpyxl matplotlib folium watchdog
```

---

## 🚀 Quick Start

1. Place your ODK/KoboToolbox `.xlsx` export in the project folder and rename it `ODK_Raw_Export.xlsx`

2. Run the full pipeline at once:

```bash
python script_7_automation.py --file ODK_Raw_Export.xlsx --no-email
```

3. Or run scripts individually in order:

```bash
python script_1_cleaning.py       # Always run first
python script_2_attendance.py
python script_3_payroll.py
python script_4_gps_qc.py
python script_5_reporting.py
python script_6_geospatial.py
```

4. Find all outputs in the `outputs/` folder.

---

## 📜 Scripts

### Script 1 — Data Cleaning
Loads the raw ODK export and produces a clean, analysis-ready dataset.
- Drops redundant ODK system columns (`subscriberid`, `simserial`, hidden `_gps_*` duplicates)
- Renames ODK-style columns (`gps_location-Latitude` → `latitude`)
- Fixes data types — dates, times, numerics, categoricals
- Flags missing GPS values
- Adds computed `hours_worked` and `clock_in_status` columns

```bash
python script_1_cleaning.py
# Input:  ODK_Raw_Export.xlsx
# Output: cleaned_data.xlsx, cleaned_data.csv
```

---

### Script 2 — Attendance Analysis
Builds a full attendance register from the cleaned data.
- Clock-in status per worker per day: Early / On Time / Late (with exact minutes)
- Detects absent workers (expected but no submission)
- Calculates total hours, average hours, attendance rate %
- Flags chronic latecomers (configurable threshold)
- Outputs a calendar-grid attendance matrix

```bash
python script_2_attendance.py
# Input:  cleaned_data.csv
# Output: attendance_report.xlsx (3 sheets)
```

---

### Script 3 — Payroll Calculation
Calculates pay for each worker based on hours worked.
- Configurable hourly rates per worker
- Overtime rules: hours > 8 billed at 1.5× rate
- Classifies days: Full Day / Half Day / Overtime / Absent
- Monthly breakdown per worker
- Grand total payroll summary

```bash
python script_3_payroll.py
# Input:  cleaned_data.csv
# Output: payroll_report.xlsx (3 sheets)
```

---

### Script 4 — GPS Quality Control
Validates every GPS point against a set of quality rules.
- Accuracy threshold check (default: < 10 metres)
- Boundary box validation (configurable to your study area)
- Duplicate coordinate detection per worker per day
- Timestamp gap checks (flags suspiciously fast collections)
- Distance between consecutive points
- Per-worker QC rating: **PASS** / **REVIEW** / **FAIL**

```bash
python script_4_gps_qc.py
# Input:  cleaned_data.csv
# Output: quality_control_report.xlsx (3 sheets)
```

---

### Script 5 — Reporting & Charts
Generates a polished Excel report with KPI cards and embedded charts.
- Executive summary sheet with 8 KPI metrics
- Bar chart: Total hours per worker
- Bar chart: GPS validity % per worker (colour-coded PASS/REVIEW/FAIL)
- Bar chart: Late clock-ins per worker
- Line chart: Weekly validity trend

```bash
python script_5_reporting.py
# Input:  cleaned_data.csv
# Output: full_report.xlsx
```

---

### Script 6 — Geospatial Maps
Produces interactive HTML maps and GIS-ready exports.
- `map_all_points.html` — all points colour-coded (valid=green, bad accuracy=red, out of bounds=orange)
- `map_by_worker.html` — one layer per worker, toggle on/off
- `heatmap.html` — collection density heatmap
- `points_export.geojson` — import directly into QGIS, ArcGIS, or Google Earth

```bash
python script_6_geospatial.py
# Input:  cleaned_data.csv
# Output: map_*.html, points_export.geojson
```

---

### Script 7 — Automation
Orchestrates all 6 scripts and handles delivery.
- Run the full pipeline on a single file
- Watch a folder and auto-trigger when a new export is dropped in
- Email the completed reports to a list of recipients
- Logs every run to `pipeline.log`

```bash
# Run once on a file
python script_7_automation.py --file ODK_Raw_Export.xlsx

# Watch a folder for new exports
python script_7_automation.py --watch ./incoming

# Run without sending email
python script_7_automation.py --file ODK_Raw_Export.xlsx --no-email
```

---

## ⚙️ Configuration

Each script has a `CONFIG` section at the top. Key settings to update for your project:

| Setting | Script | Default | Description |
|---|---|---|---|
| `HOURLY_RATES` | Script 3 | 18.00 | Pay rate per worker ID |
| `STANDARD_HOURS` | Script 3 | 8.0 | Hours before overtime kicks in |
| `OVERTIME_MULT` | Script 3 | 1.5 | Overtime pay multiplier |
| `ACCURACY_THRESHOLD` | Scripts 4 & 6 | 10.0 m | Max acceptable GPS accuracy |
| `BOUNDARY` | Scripts 4 & 6 | Nairobi area | Lat/lon bounding box for your study area |
| `POINTS_REQUIRED` | Scripts 4 & 5 | 10 | Expected points per worker per day |
| `REQUIRED_START` | Scripts 1 & 2 | 08:00 | Required clock-in time |
| `LATE_THRESHOLD` | Script 2 | 5 min | Grace period before marked Late |
| `CHRONIC_LATE` | Script 2 | 3 days | Days late before flagged chronic |
| `EMAIL_CONFIG` | Script 7 | Placeholder | SMTP settings and recipient list |

---

## 📤 Outputs

| File | Description |
|---|---|
| `cleaned_data.xlsx/.csv` | Standardized ODK data ready for analysis |
| `attendance_report.xlsx` | Daily register, worker summary, attendance matrix |
| `payroll_report.xlsx` | Daily register, worker payroll summary, monthly totals |
| `quality_control_report.xlsx` | Point-level flags, daily QC, worker PASS/REVIEW/FAIL |
| `full_report.xlsx` | Executive summary with KPI cards and 4 charts |
| `map_all_points.html` | Interactive map of all points (open in browser) |
| `map_by_worker.html` | Per-worker toggle map |
| `heatmap.html` | Collection density heatmap |
| `points_export.geojson` | GIS-ready export for QGIS / ArcGIS |

---

## 🎯 Use Cases

- **Land use mapping** — validate and report on field teams collecting land cover data
- **Environmental surveys** — track GPS collection quality across a study area
- **Agricultural assessments** — monitor field agent attendance and productivity
- **Infrastructure surveys** — QC GPS points for utility or road mapping
- Any ODK/KoboToolbox project with **daily point collection targets and field workers to track**

---

## 🧰 Tech Stack

| Library | Purpose |
|---|---|
| `pandas` | Data cleaning, analysis, grouping |
| `openpyxl` | Excel file creation and formatting |
| `matplotlib` | Chart generation |
| `folium` | Interactive Leaflet maps |
| `watchdog` | Folder monitoring for automation |
| `smtplib` | Email delivery |

---

## 🤝 Contributing

Contributions are welcome! To contribute:

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/your-feature-name
`
3. Commit your changes: `git commit -m "Add your feature"`
4. Push to your branch: `git push origin feature/your-feature-name`
5. Open a Pull Request

Please keep each script self-contained and include comments for any new config options.

---

## 📄 License

This project is licensed under the MIT License — see the [LICENSE](LICENSE) file for details.

---

## 👤 Author

**Mark Mwari**
GitHub: [@nxrth1](https://github.com/nxrth1)

---

*Built for field data teams working with ODK Collect and KoboToolbox.*
