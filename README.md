# Data Cleanup & Excel Reporting Tool 🧹📊

A Python automation tool that takes messy, inconsistent data, cleans it automatically, and generates a professionally formatted Excel report.

## Features
- Detects and removes duplicate rows
- Standardises text fields (names, emails, country, status)
- Validates and fixes date formats (multiple formats supported)
- Removes invalid entries (negative amounts, bad emails, blank fields)
- Generates a color-coded Excel report with Summary and Clean Data sheets

## Tech Stack
Python, pandas, openpyxl, Faker

## How to Run
1. Install dependencies: `pip install pandas openpyxl faker`
2. Run: `python main.py`
3. Check `clean_report.xlsx` for the generated report

## Output
- `messy_data.csv` — the raw generated input data
- `clean_report.xlsx` — formatted Excel report with summary statistics and clean data
