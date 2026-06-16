import pandas as pd
import numpy as np
import os

prices_dir = r"C:\Users\DenisG\Documents\Workplace\Data\Prices"
matritsa_file = [f for f in os.listdir(prices_dir) if "Line up" not in f and f.endswith(".xlsx") and not f.startswith("~$")][0]
matritsa_path = os.path.join(prices_dir, matritsa_file)
lineup_path = os.path.join(prices_dir, "Line up monitoring.xlsx")

df_mat = pd.read_excel(matritsa_path)
df_mat = df_mat[df_mat.iloc[:, 0] == 'Ноутбуки/Ультрабуки']

df_lineup = pd.read_excel(lineup_path)
df_out = df_lineup.copy()

current_brand = None

for i, row in df_out.iterrows():
    if not pd.isna(row.iloc[0]) and str(row.iloc[0]).strip() != "":
        current_brand = str(row.iloc[0]).strip()
    
    target = str(row.iloc[1]).strip()
    if target not in ['Lowest Price', 'Lowest Price LFR']:
        continue
        
    print(f"Row {i}: Brand={current_brand}, Target={target}, before update Price={df_out.iat[i, 2]}")
    # Force update
    df_out.iat[i, 2] = "TEST"
    df_out.iat[i, 3] = "MODEL"
    
df_out.to_excel(os.path.join(prices_dir, "test_write.xlsx"), index=False)
print("Wrote test_write.xlsx")
