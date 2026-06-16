import pandas as pd
import numpy as np
import os

prices_dir = r"C:\Users\DenisG\Documents\Workplace\Data\Prices"
matritsa_file = [f for f in os.listdir(prices_dir) if "Line up" not in f and not f.startswith("~$") and f.endswith(".xlsx") and "test" not in f][0]
matritsa_path = os.path.join(prices_dir, matritsa_file)

df_mat = pd.read_excel(matritsa_path)
df_mat = df_mat[df_mat.iloc[:, 0] == 'Ноутбуки/Ультрабуки']

def is_match(lineup_val, mat_val):
    if pd.isna(lineup_val) or str(lineup_val).strip() == "":
        return True
    l_str = str(lineup_val).replace(' ', '').replace('GB', 'G').replace('Gb', 'G').lower()
    m_str = str(mat_val).replace(' ', '').replace('GB', 'G').replace('Gb', 'G').lower()
    return l_str in m_str

print(f"Brand: {df_mat['Бренд'].apply(lambda x: is_match('Lenovo', x)).sum()}")
print(f"CPU: {df_mat['CPU'].apply(lambda x: is_match('Core i3', x)).sum()}")
print(f"RAM: {df_mat['RAM'].apply(lambda x: is_match('8G', x)).sum()}")
print(f"HDD: {df_mat['HDD'].apply(lambda x: is_match('256G', x)).sum()}")
print(f"Video: {df_mat['Video'].apply(lambda x: is_match('Int', x)).sum()}")
print(f"OS: {df_mat['OS'].apply(lambda x: is_match('Dos', x)).sum()}")
