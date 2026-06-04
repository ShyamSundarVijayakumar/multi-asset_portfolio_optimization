import os
import requests
import difflib
import pandas as pd
import numpy as np
import yfinance as yf
from pathlib import Path
from datetime import datetime

# ==========================================
# CONFIGURATION
# ==========================================
 
base_dir = Path(os.getenv("data_dir", "."))
config_dir = Path(os.getenv("config_dir", "."))
OUTPUT_DIR = os.path.join(base_dir, "processed")
INPUT_FILES = [
    Path(OUTPUT_DIR) / "portfolio_platform1_input.csv",
    Path(OUTPUT_DIR) / "portfolio_platform2_input.csv",
    Path(OUTPUT_DIR) / "portfolio_platform3_input.csv",
    Path(OUTPUT_DIR) / "portfolio_platform4_input.csv",
    Path(OUTPUT_DIR) / "portfolio_platform5_input.csv"
]
CORRECTIONS_FILE = Path(config_dir) / "portfolio_corrections.csv"
Industry_Sector_Country_Fix_FILE = Path(config_dir) / "Industry_Sector_Country_Fix.csv"
FILE_CSV_OUT = Path(OUTPUT_DIR) / "Consolidated_Portfolio_Positions.csv"
FILE_XLSX_OUT = Path(OUTPUT_DIR) / "Overall_PnL_and_Tax_Summary.xlsx"

FX_CACHE = {}

# ==========================================
# CORE UTILITIES
# ==========================================
def clean_name(name):
    if pd.isna(name): return "Unknown"
    name = str(name).lower().replace("_", " ").strip()
    for suffix in [" ag", " gmbh", " plc", " inc", " corp", " se", "(acc)", "(dist)", " class a", " class b"]:
        name = name.replace(suffix, "")
    return name.strip().title()

def robust_date_parser(date_val):
    date_str = str(date_val).strip()
    for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%d.%m.%Y', '%Y/%m/%d'):
        try: return datetime.strptime(date_str, fmt)
        except ValueError: continue
    return pd.NaT

def get_live_fx_rate(from_curr, to_curr="EUR"):
    from_curr = from_curr.upper().strip()
    to_curr = to_curr.upper().strip()
    if from_curr in ["GBP", "GBX"]: from_curr = "GBP"
    if from_curr == to_curr: return 1.0

    pair = f"{from_curr}{to_curr}=X"
    if pair in FX_CACHE: return FX_CACHE[pair]

    try:
        ticker = yf.Ticker(pair)
        rate = ticker.info.get('currentPrice') or ticker.info.get('previousClose')
        if not rate:
            hist = ticker.history(period="1d")
            if not hist.empty: rate = hist['Close'].iloc[-1]
        if rate:
            FX_CACHE[pair] = float(rate)
            return float(rate)
    except Exception: pass
    
    if to_curr != "USD" and from_curr != "USD":
        try:
            rate1 = get_live_fx_rate(from_curr, "USD")
            rate2 = get_live_fx_rate("USD", to_curr)
            return rate1 * rate2
        except: pass
        
    return 1.0

def get_crypto_info_if_applicable(isin: str, name: str):
    if not (isin or "").strip().lower().startswith("crypto currency"): return None
    symbol = (name or "").strip().upper()
    for suf in ("-EUR", "-USD"):
        ticker_sym = f"{symbol}{suf}"
        try:
            t = yf.Ticker(ticker_sym)
            hist = t.history(period="5d")
            if hist.empty: continue
            price = float(hist["Close"].dropna().iloc[-1])
            if price <= 0: continue

            if suf == "-EUR": price_eur = price
            else: price_eur = price * float(get_live_fx_rate("USD", "EUR"))

            return {
                "price": price_eur, "currency": "EUR",
                "sector": "Cryptocurrency", "industry": "Digital Asset", "country": "Decentralized"
            }
        except Exception: continue
    return None

def get_ticker_info(isin, name):
    if isin == "Unknown" and name == "Unknown": return None

    crypto_res = get_crypto_info_if_applicable(isin, name)
    if crypto_res is not None: return crypto_res

    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
    search_queries = []

    if isin != "Unknown" and not isin.lower().startswith("crypto currency"): search_queries.append(isin)
    if name != "Unknown": search_queries.append(name)
        
    for query in search_queries:
        url = f"https://query2.finance.yahoo.com/v1/finance/search?q={query}"
        try:
            res = requests.get(url, headers=headers, timeout=7).json()
            quotes = res.get('quotes', [])
            if not quotes: continue
            
            symbol = quotes[0]['symbol']
            ticker = yf.Ticker(symbol)
            info = ticker.info
            
            raw_price = info.get('currentPrice') or info.get('previousClose') or info.get('regularMarketPrice')
            
            if not raw_price or raw_price == 0.0:
                hist = ticker.history(period="5d")
                if not hist.empty: raw_price = hist['Close'].dropna().iloc[-1]
            
            if not raw_price or raw_price == 0.0: continue
            
            currency = info.get('currency', 'EUR')
            fx_rate = get_live_fx_rate(currency, "EUR")
            
            price_in_eur = (raw_price / 100.0) * fx_rate if currency.upper() in ["GBP", "GBX"] else raw_price * fx_rate
            
            return {
                "price": price_in_eur, "currency": currency,
                "sector": info.get('sector', 'Unknown'), "industry": info.get('industry', 'Unknown'), "country": info.get('country', 'Unknown')
            }
        except Exception: continue
        
    return None

def truncate_4_decimals(val):
    try:
        if pd.isna(val): return 0.0
        return np.trunc(float(val) * 10000) / 10000
    except:
        return val

# ==========================================
# ENGINE PIPELINE EXECUTION
# ==========================================
def run_pipeline():
    df_corr = pd.DataFrame()
    if CORRECTIONS_FILE.exists():
        df_corr = pd.read_csv(CORRECTIONS_FILE)
        for c in ['wrong_isin', 'correct_isin', 'platform', 'manual_type', 'manual_quantity', 'manual_price', 'manual_date', 'security_name']:
            if c in df_corr.columns:
                df_corr[c] = df_corr[c].astype(str).str.strip().replace(['nan', 'NaN', 'None'], 'Unknown')
            else: df_corr[c] = 'Unknown'

    df_list = []
    for f in INPUT_FILES:
        if f.exists():
            temp_df = pd.read_csv(f)
            platform_name = f.stem.replace("portfolio_", "").replace("_input", "").replace("_", " ").title()
            temp_df['broker'] = platform_name
            
            if not df_corr.empty:
                platform_mask = df_corr['platform'].str.lower() == platform_name.lower()
                active_corrections = df_corr[platform_mask & (df_corr['wrong_isin'] != 'Unknown') & (df_corr['correct_isin'] != 'Unknown')]
                for _, corr in active_corrections.iterrows():
                    temp_df.loc[temp_df['isin'] == corr['wrong_isin'], 'isin'] = corr['correct_isin']
                    
            df_list.append(temp_df)
            
    if not df_list: raise FileNotFoundError("CRITICAL: No input CSV files found.")
    df = pd.concat(df_list, ignore_index=True)

    if not df_corr.empty:
        manual_entries = df_corr[(df_corr['manual_type'] != 'Unknown') & (df_corr['manual_quantity'] != 'Unknown')]
        if not manual_entries.empty:
            manual_records = []
            for _, row in manual_entries.iterrows():
                manual_records.append({
                    'security_name': row['security_name'],
                    'isin': row['correct_isin'],
                    'type': row['manual_type'].lower().strip(),
                    'quantity': row['manual_quantity'],
                    'price': row['manual_price'],
                    'date': row['manual_date'],
                    'broker': row['platform'] if row['platform'] != 'Unknown' else 'Manual_Adjustment',
                    'tax_withheld': row.get('manual_tax_withheld', 0.0), 
                    'broker_fee': row.get('manual_broker_fee', 0.0), 
                    'dividend_after_taxes': row.get('manual_dividend_after_taxes', 0.0)
                })
            df = pd.concat([df, pd.DataFrame(manual_records)], ignore_index=True)

    df['date'] = df['date'].apply(robust_date_parser)
    df = df.dropna(subset=['date']).copy()
    df['type'] = df['type'].str.strip().str.lower()
    df['clean_name'] = df['security_name'].apply(clean_name)
    df['year'] = df['date'].dt.year

    for col in ['quantity', 'price', 'tax_withheld', 'broker_fee', 'dividend_after_taxes']:
        if col in df.columns: 
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)

    df = df.sort_values(by=['date', 'type'], ascending=[True, False])

    name_to_isin = df[df['isin'] != 'Unknown'].set_index('clean_name')['isin'].to_dict()
    if not df_corr.empty:
        for _, corr in df_corr[(df_corr['wrong_isin'] != 'Unknown') & (df_corr['correct_isin'] != 'Unknown')].iterrows():
            cleaned_corr_name = clean_name(corr['security_name'])
            if cleaned_corr_name != 'Unknown': name_to_isin[cleaned_corr_name] = corr['correct_isin']

    for name in list(df['clean_name'].unique()):
        if name not in name_to_isin:
            matches = difflib.get_close_matches(name, name_to_isin.keys(), n=1, cutoff=0.8)
            if matches: name_to_isin[name] = name_to_isin[matches[0]]

    df['isin'] = df.apply(lambda r: name_to_isin.get(r['clean_name'], r['isin']), axis=1)
    df['asset_id'] = np.where(df['isin'].str.lower() != 'crypto currency', df['isin'], df['clean_name'])
    df['group_id'] = df['asset_id'] + "_" + df['broker'].astype(str)

    raw_active_portfolio, closed_portfolio_log = [], []
    dividend_logs, tax_logs = [], []
    pnl_summary = {k: 0.0 for k in ['Realized Gains (EUR)', 'Realized Losses / Harvestable (EUR)', 'Total Broker Fees Paid (EUR)']}

    for group_id, group in df.groupby('group_id'):
        fifo_queue = []
        broker_name = group['broker'].iloc[0]
        isin = group['isin'].iloc[0]
        display_name = group['security_name'].iloc[0].replace("_", " ")
        last_buy_date = pd.NaT
        
        for _, row in group.iterrows():
            qty, price = row['quantity'], row['price']
            pnl_summary['Total Broker Fees Paid (EUR)'] += row['broker_fee']
            if row['tax_withheld'] > 0: tax_logs.append({'Year': row['year'], 'Tax Amount (EUR)': row['tax_withheld']})
            
            if row['type'] == 'buy':
                fifo_queue.append({'qty': qty, 'price': price, 'date': row['date']})
                last_buy_date = row['date']
            elif row['type'] == 'sell':
                qty_to_sell, total_cost_basis_sold = qty, 0.0
                while qty_to_sell > 0 and fifo_queue:
                    oldest_buy = fifo_queue[0]
                    if oldest_buy['qty'] < qty_to_sell or np.isclose(oldest_buy['qty'], qty_to_sell, atol=1e-8):
                        matched_qty = oldest_buy['qty']
                        total_cost_basis_sold += matched_qty * oldest_buy['price']
                        qty_to_sell -= matched_qty
                        fifo_queue.pop(0)
                    else:
                        matched_qty = qty_to_sell
                        total_cost_basis_sold += matched_qty * oldest_buy['price']
                        oldest_buy['qty'] -= matched_qty
                        qty_to_sell = 0
                
                actual_qty_sold = qty - qty_to_sell
                if actual_qty_sold > 0:
                    avg_buy_price = total_cost_basis_sold / actual_qty_sold
                    realized_pnl = (price - avg_buy_price) * actual_qty_sold
                    pnl_summary['Realized Gains (EUR)' if realized_pnl > 0 else 'Realized Losses / Harvestable (EUR)'] += realized_pnl
                    closed_portfolio_log.append({
                        'ISIN': isin, 'Security Name': display_name, 'Broker': broker_name, 
                        'Sell Date': row['date'].strftime('%Y-%m-%d'), 'Quantity Sold': actual_qty_sold, 
                        'Sell Price (EUR)': price, 'Average Buy Price (EUR)': avg_buy_price, 'Total Profit or Loss (EUR)': realized_pnl
                    })
            elif row['type'] == 'dividend':
                if row['dividend_after_taxes'] > 0:
                    dividend_logs.append({'Year': row['year'], 'Dividend Amount (EUR)': row['dividend_after_taxes']})
                
        remaining_qty = sum(item['qty'] for item in fifo_queue)
        
        # Increased precision zero-cutoff here as well for crypto
        if remaining_qty < 1e-8: remaining_qty = 0.0
        
        avg_buy_price_remaining = (sum(item['qty'] * item['price'] for item in fifo_queue) / remaining_qty) if remaining_qty > 0 else 0.0
        
        if remaining_qty > 0:
            raw_active_portfolio.append({
                'ISIN': isin, 'Security Name': display_name, 'Asset_ID': group['asset_id'].iloc[0],
                'Last Bought Date': last_buy_date, 'Current Quantity': remaining_qty, 
                'Total Cost Basis Value': sum(item['qty'] * item['price'] for item in fifo_queue)
            })

    consolidated_active = []
    if raw_active_portfolio:
        df_raw_active = pd.DataFrame(raw_active_portfolio)
        for asset_id, group in df_raw_active.groupby('Asset_ID'):
            total_qty = group['Current Quantity'].sum()
            total_cost = group['Total Cost Basis Value'].sum()
            avg_price = total_cost / total_qty if total_qty > 0 else 0.0
            max_date = group['Last Bought Date'].max()
            consolidated_active.append({
                'ISIN': group['ISIN'].iloc[0],
                'Security Name': group['Security Name'].iloc[0],
                'Last Bought Date': max_date.strftime('%Y-%m-%d') if pd.notnull(max_date) else 'Unknown',
                'Current Quantity': total_qty,
                'Average Buy Price (EUR)': avg_price
            })

    print("Connecting to live metric streams...")
    for asset in consolidated_active:
        name_lower = asset['Security Name'].lower()
        isin_val = str(asset['ISIN']).upper().strip()

        if "real estate" in name_lower or isin_val == "REAL ESTATE":
            eur_inr = get_live_fx_rate("EUR", "INR")
            inr_eur = get_live_fx_rate("INR", "EUR")
            buy_date_str = asset['Last Bought Date']
            years_held = 0
            if buy_date_str != 'Unknown':
                buy_date = datetime.strptime(buy_date_str, '%Y-%m-%d')
                years_held = (datetime.now() - buy_date).days / 365.25

            initial_inr_price = asset['Average Buy Price (EUR)'] * 90.3753
            current_inr_price = initial_inr_price * ((1 + 0.135) ** years_held)
            current_eur_price = current_inr_price * inr_eur

            asset.update({
                'Current Price (EUR)': current_eur_price, 'Native Currency': 'EUR',
                'Sector': 'Real Estate', 'Industry': 'Property', 'Country': 'India',
                'Unrealized PnL (EUR)': (current_eur_price - asset['Average Buy Price (EUR)']) * asset['Current Quantity']
            })
            continue

        if isin_val in ["GC=F", "SI=F"] or "physical gold" in name_lower or "physical silver" in name_lower:
            target_ticker = "GC=F" if ("gold" in name_lower or isin_val == "GC=F") else "SI=F"
            try:
                t = yf.Ticker(target_ticker)
                
                hist = t.history(period="5d")
                price_usd_per_oz = float(hist["Close"].dropna().iloc[-1]) if not hist.empty else 0.0

                if price_usd_per_oz > 0:
                    # Convert global Troy Ounce standard to strictly Grams
                    price_usd_per_gram = price_usd_per_oz / 31.1034768

                    if target_ticker == "GC=F":
                        price_usd_per_gram = price_usd_per_gram * (22.0 / 24.0)

                    usd_inr = get_live_fx_rate("USD", "INR")
                    inr_eur = get_live_fx_rate("INR", "EUR")

                    price_inr_taxed = (price_usd_per_gram * usd_inr) * 1.185
                    current_eur_price = price_inr_taxed * inr_eur

                    asset.update({
                        'Current Price (EUR)': current_eur_price, 'Native Currency': 'USD',
                        'Sector': 'Commodities', 'Industry': 'Precious Metals', 'Country': 'Global',
                        'Unrealized PnL (EUR)': (current_eur_price - asset['Average Buy Price (EUR)']) * asset['Current Quantity']
                    })
                    continue
            except Exception: pass

        info = get_ticker_info(asset['ISIN'], asset['Security Name'])
        if info:
            asset.update({
                'Current Price (EUR)': info['price'], 'Native Currency': info['currency'],
                'Sector': info['sector'], 'Industry': info['industry'], 'Country': info['country'],
                'Unrealized PnL (EUR)': (info['price'] - asset['Average Buy Price (EUR)']) * asset['Current Quantity']
            })
        else:
            asset.update({
                'Current Price (EUR)': 0.0, 'Native Currency': 'Unknown', 
                'Sector': 'Unknown', 'Industry': 'Unknown', 'Country': 'Unknown', 'Unrealized PnL (EUR)': 0.0
            })

    df_div = pd.DataFrame(dividend_logs)
    if not df_div.empty:
        df_div_yearly = df_div.groupby('Year')['Dividend Amount (EUR)'].sum().reset_index()
        df_div_yearly.columns = ['Performance Metric', 'Value']
        df_div_yearly['Performance Metric'] = 'Dividends Collected in ' + df_div_yearly['Performance Metric'].astype(str) + ' (EUR)'
    else: df_div_yearly = pd.DataFrame(columns=['Performance Metric', 'Value'])

    df_tax = pd.DataFrame(tax_logs)
    if not df_tax.empty:
        df_tax_yearly = df_tax.groupby('Year')['Tax Amount (EUR)'].sum().reset_index()
        df_tax_yearly.columns = ['Performance Metric', 'Value']
        df_tax_yearly['Performance Metric'] = 'Total Taxes Withheld in ' + df_tax_yearly['Performance Metric'].astype(str) + ' (EUR)'
    else: df_tax_yearly = pd.DataFrame(columns=['Performance Metric', 'Value'])

    df_active_out = pd.DataFrame(consolidated_active)
    df_closed_out = pd.DataFrame(closed_portfolio_log)
    
    total_unrealized = df_active_out['Unrealized PnL (EUR)'].sum() if not df_active_out.empty else 0.0
    pnl_summary['Unrealized PnL (EUR)'] = total_unrealized
    pnl_summary['Net Realized PnL (EUR)'] = pnl_summary['Realized Gains (EUR)'] + pnl_summary['Realized Losses / Harvestable (EUR)']
    
    df_base_summary = pd.DataFrame(list(pnl_summary.items()), columns=['Performance Metric', 'Value'])
    df_summary_out = pd.concat([df_base_summary, df_div_yearly, df_tax_yearly], ignore_index=True)

    cols_to_round = ['Current Quantity', 'Average Buy Price (EUR)', 'Current Price (EUR)', 'Unrealized PnL (EUR)']

    # --- Apply Manual Industry/Sector/Country Fixes ---
    if Industry_Sector_Country_Fix_FILE.exists():
        df_fix = pd.read_csv(Industry_Sector_Country_Fix_FILE)
        # Ensure ISINs are strings for reliable merging
        df_fix['isin'] = df_fix['isin'].astype(str).str.strip().str.upper()
        
        # Merge the manual data into the active portfolio dataframe
        # We use 'left' join to keep existing data and fill missing gaps
        df_active_out['ISIN_UPPER'] = df_active_out['ISIN'].astype(str).str.strip().str.upper()
        df_active_out = df_active_out.merge(
            df_fix[['isin', 'sector', 'industry', 'country']], 
            left_on='ISIN_UPPER', 
            right_on='isin', 
            how='left'
        )
        
        # Fill original columns with manual data if YFinance returned 'Unknown'
        for col in ['sector', 'industry', 'country']:
            if col in df_active_out.columns:
                manual_col = col if col in df_fix.columns else None
                if manual_col:
                    df_active_out[col.capitalize()] = df_active_out[col.capitalize()].mask(
                        (df_active_out[col.capitalize()].isna()) | (df_active_out[col.capitalize()] == 'Unknown'), 
                        df_active_out[manual_col]
                    )
        
        # Cleanup temp columns
        df_active_out.drop(columns=['isin', 'ISIN_UPPER', 'sector', 'industry', 'country'], errors='ignore', inplace=True)
    
    if not df_active_out.empty:
        for col in cols_to_round:
            if col in df_active_out.columns:
                df_active_out[col] = df_active_out[col].apply(truncate_4_decimals)
        df_active_out.to_csv(FILE_CSV_OUT, index=False)
    else:
        pd.DataFrame(columns=['ISIN', 'Security Name', 'Last Bought Date', 'Current Quantity', 'Average Buy Price (EUR)']).to_csv(FILE_CSV_OUT, index=False)

    if not df_closed_out.empty:
        closed_cols = ['Quantity Sold', 'Sell Price (EUR)', 'Average Buy Price (EUR)', 'Total Profit or Loss (EUR)']
        for col in closed_cols:
            if col in df_closed_out.columns:
                df_closed_out[col] = df_closed_out[col].apply(truncate_4_decimals)

    df_summary_out['Value'] = df_summary_out['Value'].apply(truncate_4_decimals)

    with pd.ExcelWriter(FILE_XLSX_OUT, engine='openpyxl') as writer:
        df_summary_out.to_excel(writer, sheet_name='Performance Metrics', index=False)
        df_closed_out.to_excel(writer, sheet_name='Closed Positions Log', index=False)

    print("Pipeline completed execution successfully.")

if __name__ == "__main__":
    run_pipeline()