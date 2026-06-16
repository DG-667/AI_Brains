import os
import sys
import pandas as pd
import json

def main():
    source_dir = r"C:\Users\DenisG\Documents\Workplace\Data\GFK\Monthly reports"
    target_file = r"C:\Users\DenisG\Documents\Workplace\Data\GFK\GFK data TTL.xlsx"
    
    # 1. Read files and group by (Month, Year)
    files = [f for f in os.listdir(source_dir) if f.endswith('.xlsx')]
    
    # Only process files for 2023 to save time as requested by user
    # "Ты можешь отфильтровать файлы, в названии которых есть '2023'"
    # files = [f for f in files if '2023' in f] # We will just process ALL files to be sure
    
    # Let's process 2023 only to speed up as per user prompt suggestion
    files = [f for f in files if '2023' in f]
    
    grouped = {}
    errors = []
    
    for f in files:
        file_path = os.path.join(source_dir, f)
        try:
            df_raw = pd.read_excel(file_path, header=None)
            
            # Find timestamp
            ts_val = None
            for i in range(min(20, len(df_raw))):
                if str(df_raw.iloc[i, 0]).strip().lower() == 'timestamp':
                    ts_val = str(df_raw.iloc[i, 1]).strip()
                    break
                    
            if not ts_val:
                errors.append(f"No timestamp found in {f}")
                continue
                
            parts = ts_val.split()
            if len(parts) < 2:
                errors.append(f"Invalid timestamp '{ts_val}' in {f}")
                continue
            month, year = parts[0], parts[1]
            
            # find header row
            header_row = -1
            for i in range(min(30, len(df_raw))):
                if str(df_raw.iloc[i, 0]).strip().lower() == 'model name':
                    header_row = i
                    break
                    
            if header_row == -1:
                errors.append(f"No header found in {f}")
                continue
                
            df = pd.read_excel(file_path, header=header_row)
            df = df.dropna(subset=['Model Name'])
            df = df[df['Model Name'] != 'Market']
            
            rev_cols = [c for c in df.columns if 'Revenue' in str(c)]
            unit_cols = [c for c in df.columns if 'Units' in str(c)]
            avg_cols = [c for c in df.columns if 'Average' in str(c)]
            
            if not rev_cols or not unit_cols or not avg_cols:
                errors.append(f"Missing numeric columns in {f}")
                continue
                
            rev_col = rev_cols[0]
            unit_col = unit_cols[0]
            avg_col = avg_cols[0]
            
            df[rev_col] = pd.to_numeric(df[rev_col], errors='coerce')
            df[unit_col] = pd.to_numeric(df[unit_col], errors='coerce')
            df[avg_col] = pd.to_numeric(df[avg_col], errors='coerce')
            
            valid = (df[unit_col] > 0)
            if valid.sum() == 0:
                errors.append(f"No valid units in {f}")
                continue
                
            avg_price = (df.loc[valid, rev_col] / df.loc[valid, unit_col]).mean()
            
            if avg_price > 100000:
                currency = 'KZT'
            elif avg_price < 10000:
                currency = 'USD'
            else:
                errors.append(f"Unrecognized currency. Average price: {avg_price} in {f}")
                continue
                
            model_cols = [c for c in df.columns if c not in [rev_col, unit_col, avg_col]]
            
            key = (month, year)
            if key not in grouped:
                grouped[key] = {}
                
            grouped[key][currency] = {
                'df': df,
                'model_cols': model_cols,
                'rev_col': rev_col,
                'unit_col': unit_col,
                'avg_col': avg_col,
                'file': f
            }
            
        except Exception as e:
            errors.append(f"Error reading {f}: {str(e)}")
            continue

    processed_pairs_count = 0
    new_rows = []
    
    target_columns = ['Month', 'Year', 'Model Name', 'Brand', 'Display Size', 'Gaming PCs', 'GPU Model', 'Height (MM)', 'Operating Syst.', 'Processor', 'Processor Numbe', 'RAM IN GB', 'Segments', 'Storage IN GB', 'V-RAM (GB)', 'Units', 'Revenue (KZT)', 'Average Price (KZT)', 'Revenue (USD)', 'Average Price (USD)']

    for (month, year), pair in grouped.items():
        if 'KZT' not in pair or 'USD' not in pair:
            errors.append(f"Missing pair for {month} {year}")
            continue
            
        kzt = pair['KZT']
        usd = pair['USD']
        
        kzt_df = kzt['df'].copy()
        usd_df = usd['df'].copy()
        
        merge_cols = kzt['model_cols']
        
        for c in merge_cols:
            kzt_df[c] = kzt_df[c].fillna('__NAN__')
            usd_df[c] = usd_df[c].fillna('__NAN__')
            
        usd_subset = usd_df[merge_cols + [usd['rev_col'], usd['avg_col']]].rename(columns={
            usd['rev_col']: 'Revenue (USD)',
            usd['avg_col']: 'Average Price (USD)'
        })
        
        kzt_df = kzt_df.rename(columns={
            kzt['rev_col']: 'Revenue (KZT)',
            kzt['avg_col']: 'Average Price (KZT)',
            kzt['unit_col']: 'Units'
        })
        
        merged = pd.merge(kzt_df, usd_subset, on=merge_cols, how='outer')
        
        for c in merge_cols:
            merged[c] = merged[c].replace('__NAN__', pd.NA)
            
        merged['Month'] = month
        merged['Year'] = year
        
        for col in target_columns:
            if col not in merged.columns:
                merged[col] = pd.NA
                
        merged = merged[target_columns]
        new_rows.append(merged)
        processed_pairs_count += 1

    if not new_rows:
        result = {"status": "error" if errors else "success", "processed_pairs_count": 0, "errors": errors}
        print(json.dumps(result))
        return

    new_data = pd.concat(new_rows, ignore_index=True)
    
    # Load target file and deduplicate
    try:
        if os.path.exists(target_file):
            ttl_df = pd.read_excel(target_file)
            
            # drop existing rows for the months and years we are updating
            updated_periods = set(grouped.keys())
            
            def is_updated(row):
                return (str(row['Month']), str(row['Year'])) in updated_periods
                
            mask = ttl_df.apply(is_updated, axis=1)
            ttl_df = ttl_df[~mask]
            
            final_df = pd.concat([ttl_df, new_data], ignore_index=True)
        else:
            final_df = new_data
            
        # force column order
        final_df = final_df[target_columns]
        
        # save
        final_df.to_excel(target_file, index=False)
        
        status = "success" if not errors else "partial"
        result = {"status": status, "processed_pairs_count": processed_pairs_count, "errors": errors}
        print(json.dumps(result))
        
    except PermissionError:
        result = {"status": "error", "processed_pairs_count": 0, "errors": ["PermissionError: Target file is locked. Close it and try again."]}
        print(json.dumps(result))
    except Exception as e:
        result = {"status": "error", "processed_pairs_count": 0, "errors": [str(e)]}
        print(json.dumps(result))

if __name__ == "__main__":
    main()
