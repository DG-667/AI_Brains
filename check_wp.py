import pandas as pd
import json
wp_path = r"C:\Users\DenisG\Documents\Workplace\Data\WorkPlace.xlsx"
df_wp_data = pd.read_excel(wp_path, sheet_name="Data", nrows=5)
with open("C:/Users/DenisG/Documents/AI files/wp_info.json", "w", encoding="utf-8") as f:
    json.dump({
        "WP Data Columns": df_wp_data.columns.tolist(),
        "WP Data Sample": df_wp_data.head(2).fillna("").to_dict(orient="records")
    }, f, ensure_ascii=False, indent=2)
