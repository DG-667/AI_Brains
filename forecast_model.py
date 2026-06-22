import pandas as pd
import numpy as np
import json
import warnings
warnings.filterwarnings('ignore')

# 1. Сгенерируем исторические данные (2024 - mid 2026)
dates = pd.date_range(start='2024-01-01', end='2026-12-01', freq='MS')
df = pd.DataFrame({'Date': dates})
df['Year'] = df['Date'].dt.year
df['Month'] = df['Date'].dt.month

usdkzt = []
for d in dates:
    if d.year == 2024:
        usdkzt.append(450 + (d.month / 12) * 73)
    elif d.year == 2025:
        usdkzt.append(523 + (d.month / 12) * 42)
    else:
        usdkzt.append(565 - (d.month / 12) * 15)
df['USD_KZT'] = usdkzt

inflation = []
for d in dates:
    if d.year == 2024:
        inflation.append(8.6 + np.sin(d.month) * 0.5)
    elif d.year == 2025:
        inflation.append(10.0 + (d.month / 12) * 2.3)
    else:
        inflation.append(12.3 - (d.month / 12) * 2.5)
df['Inflation'] = inflation

base_sales = 100000
seasonal = [0.8, 0.85, 1.1, 0.9, 0.95, 0.9, 0.9, 1.0, 1.05, 1.0, 1.3, 1.5]
sales = []
for i, row in df.iterrows():
    if row['Date'] < pd.to_datetime('2026-07-01'):
        fx_effect = (row['USD_KZT'] / 450) ** -0.8
        inf_effect = (row['Inflation'] / 8.6) ** -0.3
        noise = np.random.normal(1, 0.05)
        s = base_sales * seasonal[int(row['Month'])-1] * fx_effect * inf_effect * noise
        sales.append(int(s))
    else:
        sales.append(np.nan)
df['Sales_Units'] = sales

# 2. Моделирование
train = df[df['Date'] < pd.to_datetime('2026-07-01')].copy()
train['log_sales'] = np.log(train['Sales_Units'])
train['log_usd'] = np.log(train['USD_KZT'])
train['log_inf'] = np.log(train['Inflation'])

# Создаем матрицу признаков (с дамми-переменными месяцев)
X_cols = ['log_usd', 'log_inf']
for m in range(2, 13):
    col = f'M_{m}'
    train[col] = (train['Month'] == m).astype(int)
    X_cols.append(col)

# Добавляем константу
train['const'] = 1
X_cols.append('const')

X = train[X_cols].values
y = train['log_sales'].values

# OLS via numpy
coefs, residuals, rank, s = np.linalg.lstsq(X, y, rcond=None)

# Коэффициенты
elasticity_usd = coefs[0]
elasticity_inf = coefs[1]

# 3. Прогнозирование на 2026 (июль - декабрь)
test = df[df['Date'] >= pd.to_datetime('2026-07-01')].copy()
test['log_usd'] = np.log(test['USD_KZT'])
test['log_inf'] = np.log(test['Inflation'])

for m in range(2, 13):
    test[f'M_{m}'] = (test['Month'] == m).astype(int)
test['const'] = 1

X_test = test[X_cols].values
log_pred = X_test @ coefs
test['Forecast_Units'] = np.exp(log_pred)

# 4. Применение бизнес-ограничений (Capping)
df_2025 = df[(df['Year'] == 2025) & (df['Month'] >= 7)].set_index('Month')

capped_forecast = []
for i, row in test.iterrows():
    m = row['Month']
    sales_last_year = df_2025.loc[m, 'Sales_Units']
    predicted = row['Forecast_Units']
    
    yoy_drop = (predicted - sales_last_year) / sales_last_year
    if yoy_drop < -0.20:
        capped_val = sales_last_year * 0.80
        capped_forecast.append(int(capped_val))
    else:
        capped_forecast.append(int(predicted))

test['Capped_Forecast'] = capped_forecast
test['Average_Ticket'] = 150000 * (test['Inflation'] / 8.6)
test['Forecast_Revenue_M_KZT'] = (test['Capped_Forecast'] * test['Average_Ticket']) / 1e6

results = {
    'elasticity_usd': float(elasticity_usd),
    'elasticity_inf': float(elasticity_inf),
    'forecast_table': []
}

for i, row in test.iterrows():
    m = int(row['Month'])
    yoy_units = (row['Capped_Forecast'] / df_2025.loc[m, 'Sales_Units']) - 1
    
    results['forecast_table'].append({
        'Month': m,
        'USD_KZT': float(row['USD_KZT']),
        'Inflation': float(row['Inflation']),
        'Units': int(row['Capped_Forecast']),
        'YoY_Units': float(yoy_units),
        'Revenue_M_KZT': float(row['Forecast_Revenue_M_KZT'])
    })

with open('forecast_results_np.json', 'w', encoding='utf-8') as f:
    json.dump(results, f, indent=4)

print("Forecast completed and saved to forecast_results_np.json")
