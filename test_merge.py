import pandas as pd
import os

source_dir = r"C:\Users\DenisG\Documents\Workplace\Data\GFK\Monthly reports"
files = ['Mobile Computing in Kazakhstan-products-Apr 2023.xlsx', 'Mobile Computing in Kazakhstan-products-Apr 2023 (1).xlsx']

def process_file(file_path):
    df_raw = pd.read_excel(file_path, header=None)
    
    ts_val = None
    for i in range(min(20, len(df_raw))):
        if str(df_raw.iloc[i, 0]).strip().lower() == 'timestamp':
            ts_val = str(df_raw.iloc[i, 1]).strip()
            break
            
    month, year = ts_val.split()
    
    header_row = -1
    for i in range(min(30, len(df_raw))):
        if str(df_raw.iloc[i, 0]).strip().lower() == 'model name':
            header_row = i
            break
            
    df = pd.read_excel(file_path, header=header_row)
    df = df.dropna(subset=['Model Name'])
    df = df[df['Model Name'] != 'Market']
    
    rev_col = [c for c in df.columns if 'Revenue' in str(c)][0]
    unit_col = [c for c in df.columns if 'Units' in str(c)][0]
    avg_col = [c for c in df.columns if 'Average' in str(c)][0]
    
    df[rev_col] = pd.to_numeric(df[rev_col], errors='coerce')
    df[unit_col] = pd.to_numeric(df[unit_col], errors='coerce')
    df[avg_col] = pd.to_numeric(df[avg_col], errors='coerce')
    
    valid = (df[unit_col] > 0)
    avg_price = (df.loc[valid, rev_col] / df.loc[valid, unit_col]).mean()
    
    currency = 'KZT' if avg_price > 100000 else 'USD' if avg_price < 10000 else 'UNKNOWN'
    
    model_cols = [c for c in df.columns if c not in [rev_col, unit_col, avg_col]]
    
    return {
        'month': month, 'year': year, 'currency': currency,
        'df': df, 'model_cols': model_cols,
        'rev_col': rev_col, 'unit_col': unit_col, 'avg_col': avg_col
    }

data = []
for f in files:
    data.append(process_file(os.path.join(source_dir, f)))

print(f"File 1: {data[0]['currency']} | File 2: {data[1]['currency']}")
