import pandas as pd
import json
import os

prices_dir = r"C:\Users\DenisG\Documents\Workplace\Data\Prices"
out_path = os.path.join(prices_dir, "Line up monitoring_filled.xlsx")

df_out = pd.read_excel(out_path)
out_sample = df_out.head(10).fillna("").to_dict(orient="records")

with open("C:/Users/DenisG/Documents/AI files/out_check.json", "w", encoding="utf-8") as f:
    json.dump(out_sample, f, ensure_ascii=False, indent=2)
