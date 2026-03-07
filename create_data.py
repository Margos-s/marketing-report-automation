# ================================================
# Campaign Data Generator
# Purpose: Creates mock Google Ads campaign data
#          and saves it as a CSV file
# ================================================

import csv

# Raw Google Ads campaign data
campaigns = [
    ["Campaign", "Impressions", "Clicks", "Spend", "Conversions"],
    ["Brand Awareness",      120000, 3600, 1800.00, 180],
    ["Product Launch",        95000, 4750, 2375.00, 285],
    ["Retargeting",           60000, 3000,  900.00, 270],
    ["Competitor Keywords",   80000, 2400, 1600.00,  96],
    ["Seasonal Promo",       110000, 5500, 2200.00, 385],
]

# Save data to CSV file
with open("campaign_data.csv", "w", newline="") as file:
    writer = csv.writer(file)
    writer.writerows(campaigns)

print("campaign_data.csv created successfully!")