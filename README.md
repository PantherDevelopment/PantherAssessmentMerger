# Panther Assessment Merger

A desktop application for merging faculty course assessment data into departmental master spreadsheets and combining semester files into yearly reports.

![Platform](https://img.shields.io/badge/platform-macOS-lightgrey.svg)

## Overview

Panther Assessment Merger streamlines two common assessment coordinator tasks:

1. **Merging faculty data** — aggregate student assessment scores from multiple faculty data files into the departmental master spreadsheet, preserving all Excel formatting
2. **Combining semesters** — merge semester master files into a single yearly file, with correct handling of students who retook courses

All processing happens locally on your machine — no data is ever uploaded or shared. FERPA compliant.

---

## Installation

### Download

[📥 Download Latest Release](https://github.com/DarbyP/panther-assessment-merger/releases/latest)

- **macOS**: Download `Panther-Assessment-Merger-macOS.dmg` (Apple Silicon — M1/M2/M3/M4)
- **Windows**: Download `Panther-Assessment-Merger-Windows.zip`

### macOS
1. Open the downloaded DMG file
2. Drag "Panther Assessment Merger" to your Applications folder
3. Right-click the app → "Open" (first time only, to bypass Gatekeeper)

### Windows
1. Extract the downloaded ZIP file
2. Double-click `Panther Assessment Merger.exe`
3. Windows will show an "Unknown Publisher" security warning — this is expected. Click "More info" → "Run anyway" to proceed. The warning appears because the Windows version is not commercially code-signed, but the file is safe.

---

## Features

- 🎯 **Auto-detects matching columns** between your data and the master spreadsheet
- 📊 **Multi-file support** — load multiple faculty files at once (any mix of courses)
- ✅ **Column match preview** — color-coded display shows exactly what will merge before you commit
- 🎨 **Preserves Excel formatting** — colors, styles, formulas, and hidden columns are untouched
- 🔍 **Reports missing students** — lists any student IDs in your data not found in master
- 📅 **Semester combining** — merges Fall/Spring/Summer files into a yearly file with correct priority ordering
- 🔔 **Auto update notifications** — checks GitHub for new versions on startup
- 🖥️ **Universal Mac build** — single DMG runs natively on both Intel and Apple Silicon
- 🔒 **FERPA compliant** — all data stays local, no web interface

---

## Usage

### Tab 1: Merge Data into Master

Use this tab to add faculty assessment data into the departmental master spreadsheet.

#### Step 1 — Load Master File
Click **"Load Master File"** and select the departmental master spreadsheet (`.xlsx`). The file must contain a `STUDENT_ID` column.

#### Step 2 — Load Your Data File(s)
Click **"Add Data File(s)"** and select one or more faculty data files. You can select multiple files at once using Cmd+Click (Mac) or Ctrl+Click (Windows). Each file must contain a `STUDENT_ID` column.

Files can contain data for any course or mix of courses — the app figures out what matches automatically.

#### Step 3 — Review Column Match Preview
The preview tree shows what will happen for each file:

| Color | Meaning |
|-------|---------|
| 🟢 Green | Columns that will be merged |
| 🟠 Orange | Some columns match, some don't (unmatched columns ignored with warning) |
| 🔴 Red | No matching columns — file will be skipped |

#### Step 4 — Merge & Save
Click **"Merge & Save"**. The master file is updated in place — all formatting, colors, and structure are preserved. Only cells with data in your files are written; blank cells never overwrite existing master data.

The Status Log reports:
- Number of students updated
- Student IDs found in your data but not in the master (worth investigating)

---

### Tab 2: Combine Semesters into Yearly

Use this tab to combine semester master files into a single yearly report.

#### Step 1 — Load Files
Click **"Add File(s)"** and select the semester master files to combine.

#### Step 2 — Assign Type to Each File
Use the dropdown to tag each file:

| Type | Description |
|------|-------------|
| Fall | Fall semester master file |
| Spring | Spring semester master file |
| Summer | Summer semester master file |
| Already Combined | A file already containing merged Fall + Spring data. Use only when adding Summer to an existing combined file. Only one allowed per operation. |

**Priority order (lowest → highest):** Fall → Spring → Summer → Already Combined

For students appearing in multiple files, data from the higher-priority file is used, column by column. Blank cells never overwrite existing data.

#### Step 3 — Save Yearly File
Click **"Combine & Save Yearly File"** and choose a save location. All students from all semester files are included in the output. Original files are not modified.

---

## Data Format

Your faculty data files must include:
- A `STUDENT_ID` column matching the IDs in the master spreadsheet
- Assessment columns named to match the master spreadsheet exactly

Example:
```
STUDENT_ID | PSY3421_ExamQ_1.1.2 | PSY3421_ExamQ_1.2.1 | PSY3421_BMP_Methods_1.2.2
123456     | 85                  | 90                  | 78
234567     | 92                  | 88                  | 95
```

Column name format: `COURSE_AssessmentType_OutcomeID`
- `PSY3421_ExamQ_1.1.2` → PSY3421 course, exam question, outcome 1.1.2
- `PSY4514_Methods_3.1.1` → PSY4514 course, methods assessment, outcome 3.1.1

---

## Updates

The app checks for new versions automatically on startup. If an update is available, you will be prompted with a link to download it. You can also check manually via **Help → Check for Updates**.

---

## Running from Source

If you prefer to run from source rather than the standalone app:

```bash
# Install dependencies
pip3 install PyQt6 pandas openpyxl

# Run
python3 assessment_merger.py
```

---

## Requirements

| Platform | Requirement |
|----------|-------------|
| macOS | Apple Silicon (M1 or later) |
| Windows | Run from source only (Python 3.8+, PyQt6, pandas, openpyxl) |

---

## Support

- **Contact**: Darby Proctor, Ph.D. — Florida Institute of Technology
- **Issues**: [GitHub Issues](https://github.com/DarbyP/panther-assessment-merger/issues)
- **In-app guide**: Help → User Guide

---

## Tips

- Always keep a backup of the master file before merging
- Student IDs must match exactly (same format, no extra spaces)
- The app works for any Canvas-using institution — column naming just needs to follow the `COURSE_Type_Outcome` format used by your department

---

Developed by Darby Proctor, Ph.D. — Florida Institute of Technology  
Built with PyQt6, pandas, and openpyxl
