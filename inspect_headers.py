import pandas as pd

matrix_path = r'C:\Users\DenisG\Documents\Workplace\Data\Prices\Матрица цен конкурентов - все Т2.xlsx.xlsx'
wp_path = r'C:\Users\DenisG\Documents\Workplace\Data\WorkPlace.xlsx'
template_path = r'C:\Users\DenisG\Documents\Workplace\Data\Prices\Line up monitoring.xlsx'

df_matrix = pd.read_excel(matrix_path, nrows=5)
df_wp = pd.read_excel(wp_path, nrows=5)
df_templ = pd.read_excel(template_path, nrows=5)

with open(r'C:\Users\DenisG\Documents\AI files\temp_headers.txt', 'w', encoding='utf-8') as f:
    f.write("MATRIX HEADERS:\n")
    f.write(str(list(df_matrix.columns)) + "\n\n")
    f.write("WORKPLACE HEADERS:\n")
    f.write(str(list(df_wp.columns)) + "\n\n")
    f.write("TEMPLATE HEADERS:\n")
    f.write(str(list(df_templ.columns)) + "\n\n")
