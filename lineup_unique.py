import pandas as pd
import json

lineup_path = r"C:\Users\DenisG\Documents\Workplace\Data\Prices\Line up monitoring.xlsx"
df_lineup = pd.read_excel(lineup_path)

out = {
    "Brands": df_lineup.iloc[:, 0].dropna().unique().tolist(),
    "Targets": df_lineup.iloc[:, 1].dropna().unique().tolist(),
    "Displays": df_lineup.iloc[:, 4].dropna().unique().tolist(),
    "CPUs": df_lineup.iloc[:, 5].dropna().unique().tolist(),
    "RAMs": df_lineup.iloc[:, 6].dropna().unique().tolist(),
    "Storages": df_lineup.iloc[:, 7].dropna().unique().tolist(),
    "Videos": df_lineup.iloc[:, 8].dropna().unique().tolist(),
    "OSs": df_lineup.iloc[:, 9].dropna().unique().tolist()
}

with open("C:/Users/DenisG/Documents/AI files/lineup_unique.json", "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, indent=2)
