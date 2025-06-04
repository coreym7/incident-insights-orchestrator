# Incident Insights Orchestrator

This project automates the analysis and reporting of incident logs across multiple dimensions—including time, grade level, location, and staff member. Built using modular Python functions, it transforms raw log data into structured Excel workbooks with summary tables, per-building breakdowns, and multi-tab outputs designed for operational decision-making.

---

## 🔧 Features

- **ETL Pipeline**: Ingests raw CSV data and maps columns to a standardized schema
- **Multi-Dimensional Analysis**:
  - Incidents by hour (with AM/PM conversion)
  - Location breakdowns
  - Grade-level distribution
  - Top incident authors and students
  - Date-based analysis with weekday averages
- **Excel Output**:
  - Styled summary reports with grouped sections and borders
  - Separate sheets for raw logs and detailed breakdowns
  - Per-building output folders with sanitized file paths
- **Extensible Design**:
  - Easily add new metric types or change output format
  - Built for real-world district use with ambiguous, messy inputs

---

## 🚀 Technologies Used

- Python 3
- pandas
- openpyxl
- matplotlib (optional for future graphing)
- re / datetime / collections

---

## 📁 Sample Input

The system expects a CSV file with the following columns (customizable via `column_mapping`):

- `Student Name`
- `Grade Level`
- `Incident Date`
- `Incident Time`
- `Incident Location`
- `Subtype Name`
- `Entry Author`
- `Student School`

---

## 🗂️ Output Structure

For each school or site, the system generates:

```
/output/
└── Building_Name/
    ├── Building_Name_Report_YYYY-MM-DD.xlsx
    └── (optional) CSV backups
```

Each report includes:
- A summary sheet with tabular insights
- A sheet of detailed log entries
- Optional hourly-location cross breakdown

---

## ⚙️ Execution

To run the full system:

```bash
python main.py
```

Make sure to adjust:

- `input_file = "your_file.csv"`
- `column_mapping = {...}` (map raw column names to normalized keys)
- `output_folder = "./output"`
- `workbook_name = "Your_Report.xlsx"`

---

## 🛡️ Disclosure

This repo contains anonymized and domain-agnostic logic. No sensitive data, credentials, or private dependencies are included. Originally designed for K–12 compliance reporting, but fully adaptable to other domains (HR, safety, healthcare, etc.).
