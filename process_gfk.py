import pandas as pd
import os
import glob
import re
from collections import defaultdict

gfk_dir = r'C:\Users\DenisG\Documents\Workplace\Data\GFK\Monthly reports'
output_file = r'C:\Users\DenisG\Documents\Workplace\Data\GFK\GFK data TTL.xlsx'

all_files = glob.glob(os.path.join(gfk_dir, '*.xlsx'))

groups = defaultdict(list)
for f in all_files:
    filename = os.path.basename(f)
    match = re.search(r'-([A-Z][a-z]{2})\s+(\d{4})', filename)
    if match:
        month, year = match.groups()
        groups[(month, int(year))].append(f)

all_data = []

sort_keys = ['Model Name', 'Brand', 'Display Size', 'Gaming PCs', 'GPU Model', 
             'Height (MM)', 'Operating Syst.', 'Processor', 'Processor Numbe', 
             'RAM IN GB', 'Segments', 'Storage IN GB', 'V-RAM (GB)', 'Units']

for (month, year), files in groups.items():
    if len(files) == 0:
        continue
        
    dfs = []
    for f in files:
        try:
            temp = pd.read_excel(f, header=None, nrows=20)
            idx_list = temp[temp.iloc[:, 0] == 'Model Name'].index
            if len(idx_list) == 0:
                continue
            idx = idx_list[0]
            
            df = pd.read_excel(f, header=idx)
            df = df[df['Model Name'].notna()]
            df = df[df['Model Name'] != 'Market']
            df = df[df['Brand'] != '-']
            df = df.loc[:, ~df.columns.astype(str).str.contains('^Unnamed')]
            
            avg = pd.to_numeric(df['Average Price (Units)'], errors='coerce').mean()
            currency = 'USD' if avg < 10000 else 'KZT'
            df['currency_type'] = currency
            dfs.append(df)
        except Exception as e:
            pass
            
    if not dfs:
        continue
        
    df_kzt = next((d for d in dfs if d['currency_type'].iloc[0] == 'KZT'), None)
    df_usd = next((d for d in dfs if d['currency_type'].iloc[0] == 'USD'), None)
    
    if df_kzt is not None and df_usd is not None:
        df_kzt = df_kzt.rename(columns={'Revenue': 'Revenue (KZT)', 'Average Price (Units)': 'Average Price (KZT)'})
        df_usd = df_usd.rename(columns={'Revenue': 'Revenue (USD)', 'Average Price (Units)': 'Average Price (USD)'})
        
        # Sort both by identical keys to ensure perfect alignment
        keys_present = [k for k in sort_keys if k in df_kzt.columns]
        df_kzt = df_kzt.sort_values(by=keys_present).reset_index(drop=True)
        df_usd = df_usd.sort_values(by=keys_present).reset_index(drop=True)
        
        merged = df_kzt.copy()
        if 'Revenue (USD)' in df_usd.columns:
            merged['Revenue (USD)'] = df_usd['Revenue (USD)']
            merged['Average Price (USD)'] = df_usd['Average Price (USD)']
    else:
        merged = dfs[0]
        curr = merged['currency_type'].iloc[0]
        if curr == 'KZT':
            merged = merged.rename(columns={'Revenue': 'Revenue (KZT)', 'Average Price (Units)': 'Average Price (KZT)'})
            merged['Revenue (USD)'] = None
            merged['Average Price (USD)'] = None
        else:
            merged = merged.rename(columns={'Revenue': 'Revenue (USD)', 'Average Price (Units)': 'Average Price (USD)'})
            merged['Revenue (KZT)'] = None
            merged['Average Price (KZT)'] = None

    merged['Year'] = year
    merged['Month'] = month
    merged = merged.drop(columns=['currency_type'], errors='ignore')
    all_data.append(merged)
    print(f"Processed {month} {year} - Rows: {len(merged)}")

if all_data:
    final_df = pd.concat(all_data, ignore_index=True)
    
    col_order = ['Model Name', 'Brand', 'Display Size', 'Gaming PCs', 'GPU Model', 
                 'Height (MM)', 'Operating Syst.', 'Processor', 'Processor Numbe', 
                 'RAM IN GB', 'Segments', 'Storage IN GB', 'V-RAM (GB)', 'Units', 
                 'Year', 'Month', 'Revenue (KZT)', 'Average Price (KZT)', 
                 'Revenue (USD)', 'Average Price (USD)']
    
    for col in col_order:
        if col not in final_df.columns:
            final_df[col] = None
            
    final_df = final_df[col_order]
    
    for col in ['Units', 'Revenue (KZT)', 'Revenue (USD)', 'Average Price (KZT)', 'Average Price (USD)']:
        final_df[col] = pd.to_numeric(final_df[col], errors='coerce')
        
    with pd.ExcelWriter(output_file, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        final_df.to_excel(writer, sheet_name='Data', index=False)
    print(f"\nSuccessfully updated {output_file}. Total rows: {len(final_df)}")
else:
    print("No data processed.")
