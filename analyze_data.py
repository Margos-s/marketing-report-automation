# ================================================
# Campaign Data Analyzer
# Purpose: Loads raw campaign data and calculates
#          Key Performance Indicators (KPIs)
# ================================================

import pandas as pd

# Load raw campaign data
df = pd.read_csv("campaign_data.csv")

print("--- Raw Data ---")
print(df)
print()

# Click Through Rate (CTR) - Out of everyone who saw the ad, how many clicked?
df["CTR (%)"] = round((df["Clicks"] / df["Impressions"]) * 100, 2)

# Cost Per Click (CPC) - How much did each click cost in INR?
df["CPC (INR)"] = round(df["Spend"] / df["Clicks"], 2)

# Return on Ad Spend (ROAS) - For every 1 INR spent, how much value came back?
df["ROAS"] = round((df["Conversions"] * 800) / df["Spend"], 2)

print("--- Data with KPIs ---")
print(df)