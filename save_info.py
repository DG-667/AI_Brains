import pandas as pd
import os
import json

prices_dir = r"C:\Users\DenisG\Documents\Workplace\Data\Prices"
matritsa_file = [f for f in os.listdir(prices_dir) if f != "Line up monitoring.xlsx" and f.endswith(".xlsx")][0]
matritsa_path = os.path.join(prices_dir, matritsa_file)
lineup_path = os.path.join(prices_dir, "Line up monitoring.xlsx")
wp_path = r"C:\Users\DenisG\Documents\Workplace\Data\WorkPlace.xlsx"

out_dict = {}

df_mat = pd.read_excel(matritsa_path, nrows=5)
out_dict["Matritsa Columns"] = df_mat.columns.tolist()
# Let's get unique categories
out_dict["Categories"] = df_mat.iloc[:, 0].unique().tolist() # assuming first col is category or so
out_dict["Brands"] = df_mat.iloc[:, 1].unique().tolist() # assuming 2nd col is brand

df_lineup = pd.read_excel(lineup_path, nrows=20)
# To see what the lineup structure actually is
out_dict["Lineup Sample"] = df_lineup.head(20).fillna("").to_dict(orient="records")

df_wp = pd.read_excel(wp_path, sheet_name=None, nrows=10) # check all sheets
out_dict["WP Sheets"] = list(df_wp.keys())
out_dict["WP Columns Sheet 1"] = df_wp[list(df_wp.keys())[0]].columns.tolist()

with open("C:/Users/DenisG/Documents/AI files/info.json", "w", encoding="utf-8") as f:
    json.dump(out_dict, f, ensure_ascii=False, indent=2)
