import pandas as pd

file_path = r"C:\Users\DenisG\Documents\Workplace\Data\Prices\Line up monitoring_filled.xlsx"
df = pd.read_excel(file_path)

print("Columns:", df.columns.tolist())
print("First 15 rows:")
print(df.head(15).to_string())

# Also let's print unique values in Unnamed: 1
if 'Unnamed: 1' in df.columns:
    print("\nUnique values in Unnamed: 1:", df['Unnamed: 1'].unique())
