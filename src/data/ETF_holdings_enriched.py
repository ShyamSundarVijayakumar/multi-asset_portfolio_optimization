import time
import requests
import yfinance as yf
import pandas as pd
from pathlib import Path
from datetime import datetime

def process_universal_etf(input_path, output_path, config, map_path = None):
    """
    Processes both CSV and Excel files, dynamically finds headers, 
    and standardizes the output format.
    """
    # 1. Load Manual Map safely
    manual_map = {}
    if map_path is not None and Path(map_path).exists():
        map_df = pd.read_csv(map_path, encoding='latin-1')
        map_df.columns = map_df.columns.str.lower().str.strip()
        for _, row in map_df.iterrows():
            name = str(row.get('component name', '')).strip()
            if name and name != 'nan':
                manual_map[name] = {
                    "Symbol": str(row.get('symbol', '')).strip(),
                    "ISIN": str(row.get('isin', '')).strip()
                }

    # 2. Read Input File & Dynamically Skip Top Text
    input_file = Path(input_path)
    
    if input_file.suffix in ['.xlsx', '.xls']:
        temp_df = pd.read_excel(input_path, header=None)
        # Find header based on ISIN column name
        header_idx = temp_df[temp_df.apply(lambda r: r.astype(str).str.strip().eq(config['col_isin']).any(), axis=1)].index
        
        if not header_idx.empty:
            df = pd.read_excel(input_path, header=header_idx[0])
        else:
            df = pd.read_excel(input_path)
    else:
        df = pd.read_csv(input_path, encoding='latin-1')

    # --- WEIGHT NORMALIZATION & FILTERING ---
    # 1. Strip '%' and convert to numeric so we can calculate sums
    df['temp_weight'] = df[config['col_weight']].astype(str).str.replace('%', '', regex=False)
    df['numeric_weight'] = pd.to_numeric(df['temp_weight'], errors='coerce')
    
    # 2. Filter: Drop the disclaimers and empty text by keeping only valid numbers
    df = df[df['numeric_weight'].notna()].copy()
    
    # 3. Detect Native Excel Decimal Conversion (The 0.1112 issue)
    # If the sum of the weight column is <= 2.0 (meaning total weight is ~1.0 / 100%),
    # Pandas read the underlying decimals. We must multiply by 100 to restore 11.12.
    if df['numeric_weight'].sum() <= 2.0:
        df['numeric_weight'] = df['numeric_weight'] * 100
        
    # 4. Clean up temporary column
    df = df.drop(columns=['temp_weight'])

    # 3. Process Data
    output_data = []
    today_date = datetime.now().strftime("%Y-%m-%d")
    
    print(f"Starting to process {len(df)} stocks from {input_file.name}...")
    
    for i, row in df.iterrows():
        comp_name = str(row.get(config['col_name'], '')).strip()
        file_isin = str(row.get(config['col_isin'], '')).strip()
        
        # Skip empty rows
        if not comp_name or comp_name == 'nan': continue

        print(f"[{i+1}/{len(df)}] Processing: {comp_name}...", end="", flush=True)

        # A. Check Manual Map
        manual = manual_map.get(comp_name, {})
        symbol = manual.get("Symbol")
        isin = manual.get("ISIN") or file_isin

        # B. Fallback to API: Search Yahoo by ISIN first (Highly Efficient)
        if not symbol and isin and isin != 'nan':
            try:
                url = f"https://query2.finance.yahoo.com/v1/finance/search?q={isin}"
                data = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=5).json()
                symbol = next((q['symbol'] for q in data.get("quotes", []) if q.get("quoteType") == "EQUITY"), None)
            except: pass
            
        # C. Fallback to API: Search Yahoo by Company Name
        if not symbol:
            try:
                url = f"https://query2.finance.yahoo.com/v1/finance/search?q={requests.utils.quote(comp_name)}"
                data = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=5).json()
                symbol = next((q['symbol'] for q in data.get("quotes", []) if q.get("quoteType") == "EQUITY"), None)
            except: pass

        # D. Fetch Sector/Industry using the Symbol
        info = yf.Ticker(symbol).info if symbol else {}

        # E. Extract the scaled weight
        val = row['numeric_weight']
        
        # rounds up to 4 decimal places to prevent Python floating point artifacts 
        # (e.g. 11.12000001)
        clean_weight_str = f"{round(val, 4):g}"

        # 4. Standardize Output Format
        out_row = {
            "Date": today_date,
            "Index Name": config['index_name'],
            "Index Ticker Symbol": config['index_ticker'],
            "Component Name": comp_name,
            "Weight %": clean_weight_str,
            "Symbol": symbol or "",
            "ISIN": isin or info.get("isin", ""),
            "Sector": info.get("sector") or "",
            "Industry": info.get("industry") or "",
            "Country": row.get(config['col_country']) or info.get("country") or ""
        }
        output_data.append(out_row)
        print(" Done.")
        time.sleep(0.5)

    # 5. Write Standardized Output
    out_df = pd.DataFrame(output_data)
    columns_order = ["Date", "Index Name", "Index Ticker Symbol", "Component Name", "Weight %", "Symbol", "ISIN", "Sector", "Industry", "Country"]
    out_df = out_df[columns_order]
    
    # Use utf-8-sig to prevent UnicodeEncodeError
    out_df.to_csv(output_path, index=False, encoding='utf-8-sig')
    print(f"\nSuccessfully saved standardized file to: {output_path}")