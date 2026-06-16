import pandas as pd
import os

prices_dir = r"C:\Users\DenisG\Documents\Workplace\Data\Prices"
matritsa_file = [f for f in os.listdir(prices_dir) if f != "Line up monitoring.xlsx" and f.endswith(".xlsx")][0]
matritsa_path = os.path.join(prices_dir, matritsa_file)
lineup_path = os.path.join(prices_dir, "Line up monitoring.xlsx")
wp_path = r"C:\Users\DenisG\Documents\Workplace\Data\WorkPlace.xlsx"

print("Matritsa Path:", matritsa_path)
df_mat = pd.read_excel(matritsa_path, nrows=5)
print("Matritsa Columns:", df_mat.columns.tolist())

df_lineup = pd.read_excel(lineup_path, nrows=10)
print("Line Up Columns:", df_lineup.columns.tolist())

df_wp = pd.read_excel(wp_path, nrows=5)
print("WorkPlace Columns:", df_wp.columns.tolist())
