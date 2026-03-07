# Automated Marketing Report Generator 📊

## What This Project Does
This Python automation tool takes raw Google Ads campaign data,
calculates key marketing KPIs (Key Performance Indicators) automatically,
and generates a professionally formatted Excel report — in seconds.

A task that normally takes a Marketing Analyst 2 hours every Monday
now runs in under 5 seconds.

---

## Business Problem It Solves
Every marketing team receives raw campaign data weekly.
Someone has to manually:
- Open the CSV file
- Calculate CTR, CPC and ROAS using Excel formulas
- Format the report with colors and headers
- Send it to the manager

This project automates that entire workflow end to end.

---

## KPIs (Key Performance Indicators) Calculated Automatically

| KPI | Full Form | What It Means |
|-----|-----------|---------------|
| CTR | Click Through Rate | Out of everyone who saw the ad, how many clicked? |
| CPC | Cost Per Click | How much did each click cost in INR? |
| ROAS | Return on Ad Spend | For every 1 INR spent, how much value came back? |

---

## What the Output Looks Like
- Professional Excel report with styled headers
- Green rows = Strong performing campaigns (ROAS >= 150)
- Red rows = Poor performing campaigns (ROAS < 80)
- Summary section at the bottom showing total spend, total conversions and best campaign

---

## Project Structure
```
marketing_report_automation/
│
├── create_data.py        # Generates mock Google Ads campaign data as CSV
├── analyze_data.py       # Loads data and calculates KPIs using Pandas
├── generate_report.py    # Exports formatted Excel report using OpenPyXL
└── .gitignore            # Excludes generated files from GitHub
```

---

## Tools and Libraries Used
- **Python** — Core programming language
- **Pandas** — Data loading, cleaning and KPI calculation
- **OpenPyXL** — Excel report generation and formatting
- **CSV** — Raw data storage

---

## How to Run This Project

**Step 1:** Clone the repository
```
git clone https://github.com/Margos-s/marketing-report-automation.git
```

**Step 2:** Install required libraries
```
pip install pandas openpyxl
```

**Step 3:** Generate the data
```
python create_data.py
```

**Step 4:** Analyze the data
```
python analyze_data.py
```

**Step 5:** Generate the Excel report
```
python generate_report.py
```

**Step 6:** Open `campaign_report.xlsx` in your project folder

---

## Built By
**Margos** — Aspiring Data & Marketing Analyst
actively building skills in Python, Pandas, SQL and Tableau

---

*This project is part of my analytics portfolio demonstrating
Python automation, data analysis and business thinking.*
