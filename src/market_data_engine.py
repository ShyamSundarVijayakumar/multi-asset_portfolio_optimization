import time
import pandas as pd
import yfinance as yf
from pathlib import Path

def update_central_fx_file(market_data_dir: Path) -> pd.Series:
    """
    Updates the central EURUSD.csv file with the latest yfinance data.
    Returns a Pandas Series of the FX rates indexed by Date for fast lookups.
    """
    fx_file = market_data_dir / "EURUSD.csv"
    
    print("--- Phase 1: Synchronizing Central FX Data ---")
    if not fx_file.exists():
        raise FileNotFoundError(f"CRITICAL: Central FX file not found at {fx_file}")
        
    # Load existing FX data
    df_fx = pd.read_csv(fx_file)
    df_fx['date'] = pd.to_datetime(df_fx['date'], format='mixed', dayfirst=True)
    max_date = df_fx['date'].max()
    
    print(f"Updating EURUSD=X from last known date: {max_date.strftime('%Y-%m-%d')}...")
    
    # Fetch missing dates
    t = yf.Ticker("EURUSD=X")
    new_fx = t.history(start=max_date.strftime('%Y-%m-%d'))
    
    if not new_fx.empty:
        new_fx.index = new_fx.index.tz_localize(None).normalize()
        
        new_df = pd.DataFrame({
            'date': new_fx.index,
            'rate': new_fx['Close']
        })
        
        combined_fx = pd.concat([df_fx, new_df], ignore_index=True)
        combined_fx = combined_fx.drop_duplicates(subset=['date'], keep='last').sort_values('date')
        
        save_df = combined_fx.copy()
        save_df['date'] = save_df['date'].dt.strftime('%d/%m/%Y')
        save_df.to_csv(fx_file, index=False)
        print("Central FX file successfully updated.\n")
    else:
        combined_fx = df_fx
        print("Central FX file is already up to date.\n")
        
    return combined_fx.set_index('date')['rate']

def update_and_fetch_market_data():
    """
    Fetches and updates stock data, strictly enforcing currency conversion via the central FX file.
    """
    config_dir = Path("../config") 
    market_data_dir = Path("../data/raw/market_data_for_risk_analysis")
    mapper_file = config_dir / "yfinance_ticker_mapper.csv"
    
    # Phase 1: Get the reliable local FX dictionary/series
    fx_series = update_central_fx_file(market_data_dir)
    
    print("--- Phase 2: Updating Market Assets ---")
    df_mapper = pd.read_csv(mapper_file)
    tickers_to_process = df_mapper['yfinance_tickters'].dropna().unique()
    
    today = pd.Timestamp.today().normalize()
    ten_years_ago = today - pd.DateOffset(years=10)

    for idx, ticker in enumerate(tickers_to_process, 1):
        print(f"[{idx}/{len(tickers_to_process)}] Processing {ticker}...")
        
        file_path = market_data_dir / f"{ticker}.csv"
        existing_df = pd.DataFrame()
        
        if file_path.exists():
            existing_df = pd.read_csv(file_path)
            existing_df['Date'] = pd.to_datetime(existing_df['Date'], format='mixed').dt.normalize()
            start_date = existing_df['Date'].max().strftime('%Y-%m-%d')
            print(f"      File exists. Updating from {start_date}...")
        else:
            start_date = ten_years_ago.strftime('%Y-%m-%d')
            print(f"      New file. Fetching 10 years of data from {start_date}...")

        try:
            t = yf.Ticker(ticker)
            hist = t.history(start=start_date, auto_adjust=True)
            
            if hist.empty:
                print(f"      [Warning] No data found for {ticker}.")
                time.sleep(2)
                continue
                
            hist.index = hist.index.tz_localize(None).normalize()
            currency = str(t.info.get('currency', 'EUR')).upper()
            
            if currency == 'USD':
                print("      Converting USD to EUR using central FX file...")
                fx_aligned = fx_series.reindex(hist.index).ffill().bfill()
                
                if fx_aligned.isna().any():
                    print(f"      [CRITICAL ERROR] Missing local FX data for {ticker}'s dates. Skipping save to protect data integrity.")
                    time.sleep(2)
                    continue
                
                hist['Close'] = hist['Close'] / fx_aligned
                
            elif currency != 'EUR':
                print(f"      [CRITICAL ERROR] Unsupported currency '{currency}'. Only EUR and USD are supported via local FX files. Skipping.")
                time.sleep(2)
                continue
            
            new_data = hist[['Close']].copy().reset_index()
            new_data.rename(columns={'Date': 'Date', 'Close': 'Adjusted close price'}, inplace=True)
            
            if not existing_df.empty:
                combined_df = pd.concat([existing_df, new_data], ignore_index=True)
                combined_df = combined_df.drop_duplicates(subset=['Date'], keep='last')
            else:
                combined_df = new_data
                
            combined_df = combined_df.sort_values('Date')
            combined_df['Date'] = combined_df['Date'].dt.strftime('%Y-%m-%d')
            
            combined_df.to_csv(file_path, index=False)
            print(f"      Saved safely to {file_path.name}")
            
        except Exception as e:
            print(f"      [ERROR] Network or processing failure for {ticker}: {e}")
            
        time.sleep(2.5)