import pandas as pd
import numpy as np

print("Starting the script...")

# Paths
main_file = r"C:\Users\DenisG\Documents\Workplace\Data\GFK\GFK data TTL.xlsx"
dict1_file = r"C:\Users\DenisG\Documents\Workplace\Data\GFK\Книга1.xlsx"
dict2_file = r"C:\Users\DenisG\Documents\Workplace\Data\GFK\Книга2.xlsx"

# 1. Load data
print("Loading dictionaries...")
dict1 = pd.read_excel(dict1_file, usecols=['Model Name', 'Brand', 'SEGMENT'])
dict2 = pd.read_excel(dict2_file, usecols=['Model Name', 'Brand', 'SEGMENT', 'SEGMENT2'])

# Drop duplicates in dictionaries to avoid cartesian products
dict1 = dict1.drop_duplicates(subset=['Model Name', 'Brand'], keep='first')
dict2 = dict2.drop_duplicates(subset=['Model Name', 'Brand'], keep='first')

print("Loading main data...")
df = pd.read_excel(main_file, sheet_name='Data')

# 2. Cleanup columns
# Drop 'Segments' if it exists
if 'Segments' in df.columns:
    df.drop(columns=['Segments'], inplace=True)
# Keep original column order (without 'Segments') but with 'Segment' and 'Sub_Segment'
# If they don't exist, we will add them later.

# 3. Process data
mask_lt_2025 = df['Year'] < 2025
mask_ge_2025 = df['Year'] >= 2025

# < 2025
df_lt = df[mask_lt_2025].copy()
if 'Segment' in df_lt.columns:
    df_lt.drop(columns=['Segment', 'Sub_Segment'], inplace=True, errors='ignore')
df_lt = df_lt.merge(dict1, on=['Model Name', 'Brand'], how='left')
df_lt.rename(columns={'SEGMENT': 'Segment'}, inplace=True)
df_lt['Sub_Segment'] = np.nan

# >= 2025
df_ge = df[mask_ge_2025].copy()
if 'Segment' in df_ge.columns:
    df_ge.drop(columns=['Segment', 'Sub_Segment'], inplace=True, errors='ignore')
df_ge = df_ge.merge(dict2, on=['Model Name', 'Brand'], how='left')
df_ge.rename(columns={'SEGMENT': 'Segment', 'SEGMENT2': 'Sub_Segment'}, inplace=True)

# Combine back
df_final = pd.concat([df_lt, df_ge], ignore_index=True)

# Restore column order
cols_to_front = ['Model Name', 'Brand', 'Segment', 'Sub_Segment']
other_cols = [c for c in df_final.columns if c not in cols_to_front]
df_final = df_final[cols_to_front + other_cols]

# 4. Reporting
total_rows = len(df_final)
mapped_segment = df_final['Segment'].notna().sum()
mapped_sub_segment = df_final['Sub_Segment'].notna().sum()

print(f"Total rows: {total_rows}")
print(f"Rows with mapped Segment: {mapped_segment}")
print(f"Rows with mapped Sub_Segment: {mapped_sub_segment}")

# 5. Save back to Excel
print("Saving back to Excel...")
with pd.ExcelWriter(main_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df_final.to_excel(writer, sheet_name='Data', index=False)

print("Done!")
