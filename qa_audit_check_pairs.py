import pandas as pd

file_path = r"C:\Users\DenisG\Documents\Workplace\Data\Prices\Line up monitoring_filled.xlsx"
df = pd.read_excel(file_path)
df['Unnamed: 2'] = pd.to_numeric(df['Unnamed: 2'], errors='coerce')
df['Unnamed: 0'] = df['Unnamed: 0'].ffill()

char_cols = ['Unnamed: 0', 'Unnamed: 4', 'Unnamed: 5', 'Unnamed: 6', 'Unnamed: 7', 'Unnamed: 8', 'Unnamed: 9']

for name, group in df.iloc[1:].groupby(char_cols, dropna=False):
    lp_rows = group[group['Unnamed: 1'] == 'Lowest Price']
    lfr_rows = group[group['Unnamed: 1'] == 'Lowest Price LFR']
    
    if not lp_rows.empty and not lfr_rows.empty:
        lp_val = lp_rows['Unnamed: 2'].values[0]
        lfr_val = lfr_rows['Unnamed: 2'].values[0]
        if pd.notna(lp_val) and pd.notna(lfr_val):
            print(f"Group {name[:2]}... -> LP: {lp_val}, LFR: {lfr_val}")
