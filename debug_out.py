import pandas as pd
import numpy as np
import os

prices_dir = r"C:\Users\DenisG\Documents\Workplace\Data\Prices"
out_path = os.path.join(prices_dir, "Line up monitoring_filled.xlsx")

df_out = pd.read_excel(out_path)
print(df_out[['Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3']].head(10))
