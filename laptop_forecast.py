import pandas as pd
import numpy as np
import json
import warnings
warnings.filterwarnings('ignore')

np.random.seed(42)

dates = pd.date_range(start='2022-01-01', end='2026-12-01', freq='MS')
df = pd.DataFrame({'Date': dates})
df['Year'] = df['Date'].dt.year
df['Month'] = df['Date'].dt.month

def get_macro(d):
    y = d.year
    m = d.month
    if y == 2022: return 460 + np.random.normal(0, 10), 20.3 - (12-m)*0.5
    elif y == 2023: return 456 + np.random.normal(0, 5), 9.7 + (12-m)*0.2
    elif y == 2024: return 456 + (m/12)*69, 8.6 + np.random.normal(0, 0.5)
    elif y == 2025: return 525 + np.random.normal(0, 5), 8.6 + (m/12)*(12.3-8.6)
    else: return 530 + (m/12)*10, 12.3 - (m/12)*(12.3-10.4)

df[['USD_KZT', 'Inflation']] = df['Date'].apply(lambda x: pd.Series(get_macro(x)))

base_sales = 60000
seasonal = {1:0.8, 2:0.85, 3:0.9, 4:0.85, 5:0.85, 6:0.9, 7:0.95, 8:1.3, 9:1.4, 10:1.0, 11:1.2, 12:1.5}
df['Seasonality'] = df['Month'].map(seasonal)

df['USD_KZT_lag2'] = df['USD_KZT'].shift(2).bfill()
df['Inflation_lag1'] = df['Inflation'].shift(1).bfill()

df['Structural_Shock'] = np.where((df['Year'] == 2022) & (df['Month'].isin([3,4,5])), 0.8, 1.0)

elast_usd = -0.7
elast_inf = -0.3

sales = []
for i, row in df.iterrows():
    if row['Date'] <= pd.to_datetime('2026-05-01'):
        fx_idx = (row['USD_KZT_lag2'] / 450)
        inf_idx = (row['Inflation_lag1'] / 8.0)
        fx_eff = fx_idx ** elast_usd
        inf_eff = inf_idx ** elast_inf
        s = base_sales * row['Seasonality'] * fx_eff * inf_eff * row['Structural_Shock']
        s *= np.random.normal(1, 0.05)
        sales.append(int(s))
    else:
        sales.append(np.nan)
        
df['Sales_Units'] = sales

forecast_results = []
for scen in ['Base', 'Optimistic', 'Pessimistic']:
    df_scen = df.copy()
    for i, row in df_scen.iterrows():
        if row['Date'] > pd.to_datetime('2026-05-01'):
            if scen == 'Base':
                u, inf = 540, 10.4
            elif scen == 'Optimistic':
                u, inf = 520, 9.0
            else:
                u, inf = 570, 12.0
            
            df_scen.at[i, 'USD_KZT'] = u
            df_scen.at[i, 'Inflation'] = inf
            
    df_scen['USD_KZT_lag2'] = df_scen['USD_KZT'].shift(2).bfill()
    df_scen['Inflation_lag1'] = df_scen['Inflation'].shift(1).bfill()
    
    for i, row in df_scen.iterrows():
        if row['Date'] > pd.to_datetime('2026-05-01'):
            fx_idx = (row['USD_KZT_lag2'] / 450)
            inf_idx = (row['Inflation_lag1'] / 8.0)
            fx_eff = fx_idx ** elast_usd
            inf_eff = inf_idx ** elast_inf
            s = base_sales * row['Seasonality'] * fx_eff * inf_eff
            df_scen.at[i, 'Sales_Units'] = int(s)
            
            last_year_sales = df_scen[(df_scen['Year'] == 2025) & (df_scen['Month'] == row['Month'])]['Sales_Units'].values[0]
            yoy = (s / last_year_sales) - 1
            
            forecast_results.append({
                'Scenario': scen,
                'Month': row['Date'].strftime('%Y-%m'),
                'Forecast_Units': int(s),
                'YoY': float(yoy),
                'USD_KZT': u,
                'Inflation': inf
            })

with open("laptop_forecast_scenarios.json", "w") as f:
    json.dump(forecast_results, f, indent=4)
