# Timesheet Governance

## Overview

Timesheet Governance is an AI-powered solution that reconciles employee timesheets between the **Client Timesheet** and the **Parent Company Timesheet**.

The solution automatically understands different Excel formats, maps employees across systems, compares weekly working hours, and updates the client timesheet.

---

# Architecture

```text
                Client Timesheet
                       │
                       │
                Parent Timesheet
                       │
                       ▼
              Schema Analyzer (Gemini)
                       │
        Dynamically identifies:
        • Header row
        • Sheet type
        • Employee column
        • Date columns
        • Weekly hours column
                       │
                       ▼
          Employee Mapper
              (Gemini AI)
                       │
        Maps employees even when
        names differ across systems
                       │
                       ▼
            Hours Comparator
        • Compare Monday–Friday
        • Ignore Saturday & Sunday
        • Validate weekly hours
                       │
                       ▼
          Excel Update Engine
        • Saturday → ASPIRE
        • Sunday → FG
        • Highlight Employee Name
        • Highlight ASPIRE & FG cells
                       │
                       ▼
          Updated Client Timesheet
```

---

# Current Project Structure

```text
Timesheet Governance/
│
├── config.py
├── schema_analyzer.py
├── employeer_mapper.py
├── dummy_main.py
├── employee_mapping.json
│
├── input/
│   ├── Client_Timesheet.xlsx
│   └── Parent_Timesheet.xlsx
│
└── output/
```

---

# Tech Stack

* Python
* Pandas
* OpenPyXL
* RapidFuzz
* Google Gemini (`google-genai`)

---

# Current Status

* ✅ Schema Analyzer
* 🚧 Employee Mapper
* ⏳ Hours Comparator
* ⏳ Excel Update Engine
* ⏳ Main Workflow
