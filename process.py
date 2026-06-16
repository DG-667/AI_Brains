import pandas as pd
import numpy as np
import os

prices_dir = r"C:\Users\DenisG\Documents\Workplace\Data\Prices"
matritsa_file = [f for f in os.listdir(prices_dir) if f != "Line up monitoring.xlsx" and f.endswith(".xlsx")][0]
matritsa_path = os.path.join(prices_dir, matritsa_file)
lineup_path = os.path.join(prices_dir, "Line up monitoring.xlsx")
wp_path = r"C:\Users\DenisG\Documents\Workplace\Data\WorkPlace.xlsx"

# 1. Read files
df_mat = pd.read_excel(matritsa_path)
df_mat = df_mat[df_mat.iloc[:, 0] == 'Ноутбуки/Ультрабуки']

df_lineup = pd.read_excel(lineup_path)
df_wp = pd.read_excel(wp_path, sheet_name="Data")

# Ensure WP stock is numeric
df_wp['Stock'] = pd.to_numeric(df_wp['Stock'], errors='coerce').fillna(0)
stock_map = df_wp.groupby('Part number')['Stock'].sum().to_dict()

# 2. Helpers
lfr_columns = [
    'Сулпак (Алматы)',
    'Технодом (Алматы)',
    'Мечта (Астана)',
    'Эврика (Шымкент)',
    'DNS (Нур-Султан)',
    'Белый ветер (Караганда)',
    '(Каспи алматы) Белый Ветер'
]

def is_match(lineup_val, mat_val):
    if pd.isna(lineup_val) or str(lineup_val).strip() == "":
        return True
    l_str = str(lineup_val).replace(' ', '').replace('GB', 'G').replace('Gb', 'G').lower()
    m_str = str(mat_val).replace(' ', '').replace('GB', 'G').replace('Gb', 'G').lower()
    return l_str in m_str

# Identify price columns in Matritsa
# "Минимальная цена" is column 12. "RRP" is 13.
# We'll take all columns from RRP onwards.
start_idx = list(df_mat.columns).index('RRP')
all_price_cols = df_mat.columns[start_idx:]
available_lfr_cols = [c for c in lfr_columns if c in df_mat.columns]

# 3. Process Line Up
current_brand = None

# Create a copy to edit
df_out = df_lineup.copy()

for i, row in df_out.iterrows():
    if not pd.isna(row.iloc[0]) and str(row.iloc[0]).strip() != "":
        current_brand = str(row.iloc[0]).strip()
    
    target = row.iloc[1]
    if pd.isna(target) or target not in ['Lowest Price', 'Lowest Price LFR']:
        continue
        
    display = row.iloc[4]
    cpu = row.iloc[5]
    ram = row.iloc[6]
    storage = row.iloc[7]
    video = row.iloc[8]
    os_ = row.iloc[9]
    
    if pd.isna(cpu) and pd.isna(ram) and pd.isna(storage):
        continue # skip rows without characteristics
        
    # Filter Matritsa
    mask_brand = df_mat['Бренд'].apply(lambda x: is_match(current_brand, x))
    mask_disp = df_mat['Display'].apply(lambda x: is_match(display, x))
    mask_cpu = df_mat['CPU'].apply(lambda x: is_match(cpu, x))
    mask_ram = df_mat['RAM'].apply(lambda x: is_match(ram, x))
    mask_hdd = df_mat['HDD'].apply(lambda x: is_match(storage, x))
    mask_video = df_mat['Video'].apply(lambda x: is_match(video, x))
    mask_os = df_mat['OS'].apply(lambda x: is_match(os_, x))
    
    matched = df_mat[mask_brand & mask_disp & mask_cpu & mask_ram & mask_hdd & mask_video & mask_os].copy()
    
    if matched.empty:
        df_out.iat[i, 2] = 'N/A'
        df_out.iat[i, 3] = 'N/A'
        continue
        
    if current_brand.lower() == 'lenovo':
        # Apply Lenovo rule: get stock for each part number
        matched['stock_val'] = matched['Артикул'].apply(lambda x: stock_map.get(x, 0))
        max_stock = matched['stock_val'].max()
        matched = matched[matched['stock_val'] == max_stock]
        # if multiple have the same max stock, we just keep them all and find minimum price among them
        
    # Calculate price
    best_price = np.inf
    best_model = 'N/A'
    
    price_cols_to_check = all_price_cols if target == 'Lowest Price' else available_lfr_cols
    
    for _, m_row in matched.iterrows():
        prices = pd.to_numeric(m_row[price_cols_to_check], errors='coerce')
        min_p = prices.min()
        if pd.notna(min_p) and min_p < best_price:
            best_price = min_p
            best_model = m_row['Наименование товара']
            
    if best_price == np.inf:
        df_out.iat[i, 2] = 'N/A'
        df_out.iat[i, 3] = 'N/A'
    else:
        df_out.iat[i, 2] = best_price
        df_out.iat[i, 3] = best_model

out_path = os.path.join(prices_dir, "Line up monitoring_filled.xlsx")
df_out.to_excel(out_path, index=False)
print(f"Saved to {out_path}")
