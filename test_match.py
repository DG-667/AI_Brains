import pandas as pd
import json

prices_dir = r"C:\Users\DenisG\Documents\Workplace\Data\Prices"
matritsa_file = " 業 ७⮢ -  2.xlsx.xlsx" # We use os.listdir to find it safely
import os
matritsa_file = [f for f in os.listdir(prices_dir) if f != "Line up monitoring.xlsx" and f.endswith(".xlsx")][0]
matritsa_path = os.path.join(prices_dir, matritsa_file)

df_mat = pd.read_excel(matritsa_path)
df_mat = df_mat[df_mat.iloc[:, 0] == 'Ноутбуки/Ультрабуки']

def norm(x):
    if pd.isna(x): return ""
    return str(x).strip().lower()

sample_matches = df_mat[
    (df_mat['Бренд'].apply(norm) == 'lenovo') &
    (df_mat['CPU'].apply(norm) == 'core i3') &
    (df_mat['RAM'].apply(norm) == '8g')
]

out = []
for _, row in sample_matches.head(5).iterrows():
    out.append({
        "Brand": row["Бренд"],
        "Display": row["Display"],
        "CPU": row["CPU"],
        "RAM": row["RAM"],
        "HDD": row["HDD"],
        "Video": row["Video"],
        "OS": row["OS"],
        "Price RRP": row["RRP"]
    })

with open("C:/Users/DenisG/Documents/AI files/test_match.json", "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, indent=2)
