import pandas as pd
import numpy as np

file_path = r"C:\Users\DenisG\Documents\Workplace\Data\Prices\Line up monitoring_filled.xlsx"
df = pd.read_excel(file_path)

# Ensure numeric format for Price column
df['Unnamed: 2'] = pd.to_numeric(df['Unnamed: 2'], errors='coerce')

# Forward fill brand if needed to group, though we might not even need it if we group by characteristics or just pair them
# Characteristics are in columns 0, 4, 5, 6, 7, 8, 9
char_cols = ['Unnamed: 0', 'Unnamed: 4', 'Unnamed: 5', 'Unnamed: 6', 'Unnamed: 7', 'Unnamed: 8', 'Unnamed: 9']

# Forward-fill Brand (Unnamed: 0) to ensure 'Lowest Price LFR' has the brand of the 'Lowest Price' above it
df['Unnamed: 0'] = df['Unnamed: 0'].ffill()

errors = []

# 1. Check for prices <= 0
for idx, row in df.iterrows():
    if idx == 0: continue # header
    price = row['Unnamed: 2']
    if pd.notna(price) and price <= 0:
        brand = row['Unnamed: 0']
        model = row['Unnamed: 3']
        price_type = row['Unnamed: 1']
        errors.append(f"Row {idx+2}: Price is {price} (<= 0) for Brand '{brand}', Model '{model}', Type '{price_type}'.")

# 2. Check Lowest Price LFR >= Lowest Price for the same characteristics
# Group by the characteristic columns
for name, group in df.iloc[1:].groupby(char_cols, dropna=False):
    # Filter 'Lowest Price' and 'Lowest Price LFR'
    lp_rows = group[group['Unnamed: 1'] == 'Lowest Price']
    lfr_rows = group[group['Unnamed: 1'] == 'Lowest Price LFR']
    
    # Normally there should be 1 of each per group
    if not lp_rows.empty and not lfr_rows.empty:
        lp_val = lp_rows['Unnamed: 2'].values[0]
        lfr_val = lfr_rows['Unnamed: 2'].values[0]
        
        if pd.notna(lp_val) and pd.notna(lfr_val):
            if lfr_val < lp_val:
                # Get some info for the error message
                brand = lp_rows['Unnamed: 0'].values[0]
                model = lp_rows['Unnamed: 3'].values[0]
                errors.append(f"Lowest Price LFR ({lfr_val}) < Lowest Price ({lp_val}) for Brand '{brand}' (Model: '{model}').")

if errors:
    print("STATUS: REJECTED")
    print("Errors found:")
    for e in errors:
        print("-", e)
else:
    print("STATUS: APPROVED")
    print("No errors found.")
