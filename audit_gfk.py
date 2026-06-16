import pandas as pd
import os
import glob
import re
import sys

gfk_dir = r'/app/Workplace/Data/GFK/Monthly reports'
ttl_file = r'/app/Workplace/Data/GFK/GFK data TTL.xlsx'

try:
    df_ttl = pd.read_excel(ttl_file, sheet_name='Data')
except Exception as e:
    print(f"FAILED: Cannot read {ttl_file}: {e}")
    sys.exit(1)

all_files = glob.glob(os.path.join(gfk_dir, '*.xlsx'))
files_to_check = [f for f in all_files if not re.search(r'\(\d+\)\.xlsx$', f)]

if not files_to_check:
    print("FAILED: No raw files found.")
    sys.exit(1)

errors = []

for file_path in files_to_check:
    filename = os.path.basename(file_path)
    match = re.search(r'-([A-Z][a-z]{2})\s+(\d{4})', filename)
    if not match:
        continue
    month, year = match.groups()
    year = int(year)
    
    # Calculate total in TTL
    ttl_subset = df_ttl[(df_ttl['Year'] == year) & (df_ttl['Month'] == month)]
    ttl_units = pd.to_numeric(ttl_subset['Units'], errors='coerce').sum()
    
    if ttl_units == 0:
        continue # Skip if empty
        
    # Read raw
    try:
        temp_df = pd.read_excel(file_path, header=None, nrows=20)
        idx_list = temp_df[temp_df.iloc[:, 0] == 'Model Name'].index
        if len(idx_list) == 0:
            continue
        idx = idx_list[0]
        
        df_raw = pd.read_excel(file_path, header=idx)
        df_raw = df_raw[df_raw['Model Name'].notna()]
        df_raw = df_raw[df_raw['Model Name'] != 'Market']
        df_raw = df_raw[df_raw['Brand'] != '-']
        raw_units = pd.to_numeric(df_raw['Units'], errors='coerce').sum()
        
        diff = abs(ttl_units - raw_units)
        if diff > 0.1:
            errors.append(f"{month} {year}: Raw={raw_units:.2f}, TTL={ttl_units:.2f}, Diff={diff:.2f}")
    except Exception as e:
        print(f"Error checking {filename}: {e}")

if errors:
    print("FAILED: Data mismatch found!")
    for e in errors:
        print(" - " + e)
    sys.exit(1)
else:
    print("SUCCESS: All months match perfectly between raw files and TTL database.")
    sys.exit(0)
