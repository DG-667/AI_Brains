import pandas as pd
import os

prices_dir = r"C:\Users\DenisG\Documents\Workplace\Data\Prices"
matritsa_file = [f for f in os.listdir(prices_dir) if "Line up" not in f and not f.startswith("~$") and f.endswith(".xlsx") and "test" not in f][0]
matritsa_path = os.path.join(prices_dir, matritsa_file)

df_mat = pd.read_excel(matritsa_path)
print(df_mat['HDD'].dropna().unique()[:10])
