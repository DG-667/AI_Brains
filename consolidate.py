import pandas as pd
import numpy as np
import os

def run_consolidation():
    dir_path = r"C:\Users\DenisG\Documents\Workplace\Data\GFK"
    file_kzt = os.path.join(dir_path, "Mobile Computing in Kazakhstan-products-May 2022.xlsx")
    file_usd = os.path.join(dir_path, "Mobile Computing in Kazakhstan-products-May 2022 (1).xlsx")
    target_file = os.path.join(dir_path, "GFK data TTL.xlsx")

    # 1. Read metadata to extract month and year
    meta_df = pd.read_excel(file_kzt, header=None, nrows=10)
    # 'Timestamp' is at row 3 (index 3), 'May 2022' is in column 1.
    date_str = meta_df.iloc[3, 1]
    month, year = str(date_str).split()

    # 2. Read actual data
    df_kzt = pd.read_excel(file_kzt, skiprows=8)
    df_usd = pd.read_excel(file_usd, skiprows=8)

    # Drop the first row which is the 'May 2022' string in data columns
    df_kzt = df_kzt.iloc[1:].copy()
    df_usd = df_usd.iloc[1:].copy()

    # Filter out 'Market' total or NaN models
    df_kzt = df_kzt[df_kzt['Model Name'].notna() & (df_kzt['Model Name'] != 'Market') & (df_kzt['Model Name'] != '-')]
    df_usd = df_usd[df_usd['Model Name'].notna() & (df_usd['Model Name'] != 'Market') & (df_usd['Model Name'] != '-')]

    # 3. Rename columns
    # KZT columns
    kzt_rename = {}
    if 'Revenue' in df_kzt.columns:
        kzt_rename['Revenue'] = 'Revenue (KZT)'
    if 'Sales Value' in df_kzt.columns:
        kzt_rename['Sales Value'] = 'Revenue (KZT)'
    if 'Average Price (Units)' in df_kzt.columns:
        kzt_rename['Average Price (Units)'] = 'Average Price (KZT)'
    elif 'Average Price' in df_kzt.columns:
        kzt_rename['Average Price'] = 'Average Price (KZT)'

    df_kzt.rename(columns=kzt_rename, inplace=True)

    # USD columns
    usd_rename = {}
    if 'Revenue' in df_usd.columns:
        usd_rename['Revenue'] = 'Revenue (USD)'
    if 'Sales Value' in df_usd.columns:
        usd_rename['Sales Value'] = 'Revenue (USD)'
    if 'Average Price (Units)' in df_usd.columns:
        usd_rename['Average Price (Units)'] = 'Average Price (USD)'
    elif 'Average Price' in df_usd.columns:
        usd_rename['Average Price'] = 'Average Price (USD)'

    df_usd.rename(columns=usd_rename, inplace=True)

    # 4. Merge
    # To avoid floating point issues on 'Units', we use dimensions for merge
    dim_cols = [c for c in df_kzt.columns if c not in ['Revenue (KZT)', 'Average Price (KZT)', 'Units']]
    dim_cols = [c for c in dim_cols if c in df_usd.columns]

    # Fill NaN in dimensions to ensure merge works properly (NaN != NaN in pandas merge unless specified)
    # Actually, pandas merge handles NaN as equal if we just merge, but to be safe:
    # df_kzt[dim_cols] = df_kzt[dim_cols].fillna('N/A')
    # df_usd[dim_cols] = df_usd[dim_cols].fillna('N/A')

    merged_df = pd.merge(df_kzt, df_usd[dim_cols + ['Revenue (USD)', 'Average Price (USD)']], on=dim_cols, how='left')

    # 5. Add Month and Year
    merged_df.insert(0, 'Year', year)
    merged_df.insert(0, 'Month', month)

    # 6. Deduplication logic
    final_df = merged_df
    if os.path.exists(target_file):
        try:
            ttl_df = pd.read_excel(target_file)
            if not ttl_df.empty:
                # filter out the current month/year
                ttl_df = ttl_df[~((ttl_df['Month'].astype(str) == month) & (ttl_df['Year'].astype(str) == year))]
                final_df = pd.concat([ttl_df, merged_df], ignore_index=True)
        except Exception as e:
            pass # if empty or corrupted, just overwrite

    # 7. Save to Excel
    with pd.ExcelWriter(target_file, engine='openpyxl') as writer:
        final_df.to_excel(writer, index=False, sheet_name='Data')

    rows_added = len(merged_df)
    print(f"<GFK_CONSOLIDATION_RESULT><STATUS>SUCCESS</STATUS><ROWS_ADDED>{rows_added}</ROWS_ADDED></GFK_CONSOLIDATION_RESULT>")

if __name__ == "__main__":
    try:
        run_consolidation()
    except Exception as e:
        print(f"<ERROR_USER_ACTION_REQUIRED> {str(e)}")
