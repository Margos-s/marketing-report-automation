# ================================================
# Automated Marketing Report Generator
# Built by: Margos
# Purpose: Automatically analyze Google Ads campaign
#          data and generate a formatted Excel report
# ================================================

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ------------------------------------------------
# STEP 1 - Load raw campaign data from CSV file
# ------------------------------------------------
df = pd.read_csv("campaign_data.csv")

# ------------------------------------------------
# STEP 2 - Calculate KPIs (Key Performance Indicators)
# ------------------------------------------------

# Click Through Rate (CTR) - Out of everyone who saw the ad, how many clicked?
df["CTR (%)"] = round((df["Clicks"] / df["Impressions"]) * 100, 2)

# Cost Per Click (CPC) - How much did each click cost in INR?
df["CPC (INR)"] = round(df["Spend"] / df["Clicks"], 2)

# Return on Ad Spend (ROAS) - For every 1 INR spent, how much value came back?
df["ROAS"] = round((df["Conversions"] * 800) / df["Spend"], 2)

# ------------------------------------------------
# STEP 3 - Create a new Excel Workbook
# ------------------------------------------------
wb = Workbook()
ws = wb.active
ws.title = "Campaign Report"

# ------------------------------------------------
# STEP 4 - Write the report title
# ------------------------------------------------
ws["A1"] = "Google Ads Campaign Performance Report"
ws["A1"].font = Font(bold=True, size=14)
ws["A2"] = "Generated automatically using Python"
ws["A2"].font = Font(italic=True, size=10)

# ------------------------------------------------
# STEP 5 - Write column headers with dark blue styling
# ------------------------------------------------
headers = list(df.columns)
for col_num, header in enumerate(headers, 1):
    cell = ws.cell(row=4, column=col_num, value=header)
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="1F4E79")
    cell.alignment = Alignment(horizontal="center")

# ------------------------------------------------
# STEP 6 - Write data rows with color coding
# Green = Strong Return on Ad Spend (ROAS >= 150)
# Red   = Poor Return on Ad Spend (ROAS < 80)
# White = Average performance
# ------------------------------------------------

# Define row colors
green    = PatternFill("solid", fgColor="C6EFCE")
red      = PatternFill("solid", fgColor="FFC7CE")
no_color = PatternFill(fill_type=None)

for row_num, row_data in enumerate(df.values, 5):
    roas_value = row_data[-1]

    # Assign color based on Return on Ad Spend (ROAS) performance
    if roas_value >= 150:
        row_color = green
    elif roas_value < 80:
        row_color = red
    else:
        row_color = no_color

    # Write each cell and apply the row color
    for col_num, value in enumerate(row_data, 1):
        cell = ws.cell(row=row_num, column=col_num, value=value)
        cell.alignment = Alignment(horizontal="center")
        cell.fill = row_color

# ------------------------------------------------
# STEP 7 - Add summary section at the bottom
# ------------------------------------------------
summary_row = len(df) + 7

ws.cell(row=summary_row, column=1, value="SUMMARY").font = Font(bold=True, size=12)

ws.cell(row=summary_row + 1, column=1, value="Total Spend (INR):")
ws.cell(row=summary_row + 1, column=2, value=round(df["Spend"].sum(), 2))

ws.cell(row=summary_row + 2, column=1, value="Total Conversions:")
ws.cell(row=summary_row + 2, column=2, value=int(df["Conversions"].sum()))

best_campaign = df.loc[df["ROAS"].idxmax(), "Campaign"]
ws.cell(row=summary_row + 3, column=1, value="Best Campaign (ROAS):")
ws.cell(row=summary_row + 3, column=2, value=best_campaign)

# ------------------------------------------------
# STEP 8 - Auto adjust all column widths
# ------------------------------------------------
for col_num in range(1, len(headers) + 1):
    col_letter = get_column_letter(col_num)
    ws.column_dimensions[col_letter].width = 20

# ------------------------------------------------
# STEP 9 - Save the final Excel report
# ------------------------------------------------
wb.save("campaign_report.xlsx")
print("Campaign Report generated successfully!")
print("Check your project folder for campaign_report.xlsx")
print(f"Best Campaign: {best_campaign}")