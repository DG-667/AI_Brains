import pandas as pd
import json
import os

prices_dir = r"C:\Users\DenisG\Documents\Workplace\Data\Prices"
matritsa_file = [f for f in os.listdir(prices_dir) if f != "Line up monitoring.xlsx" and f.endswith(".xlsx")][0]
matritsa_path = os.path.join(prices_dir, matritsa_file)

df_mat = pd.read_excel(matritsa_path)
df_mat = df_mat[df_mat.iloc[:, 0] == 'Ноутбуки/Ультрабуки']

out = {
    "Videos": df_mat['Video'].dropna().unique().tolist(),
    "OSs": df_mat['OS'].dropna().unique().tolist()
}

with open("C:/Users/DenisG/Documents/AI files/matritsa_video_os.json", "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, indent=2)
