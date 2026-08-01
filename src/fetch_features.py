"""
Downloads Yahoo Finance features for all tradable assets.
Converts all USD financial values to EUR automatically.
ALIGNS TO DAILY BUSINESS CALENDAR: Uses Stocks as a reference calendar.
Smashes weekend/holiday crypto data into the next valid business day (Monday).
"""

from __future__ import annotations

import time
import logging
from pathlib import Path
import numpy as np
import pandas as pd
import yfinance as yf

from src.config import (FILES, FEATURE_DIR, SYNTHETIC_ETFS, YFINANCE_SLEEP, YF_RAW_FILE)

# =============================================================================
# Directories & Logging
# =============================================================================
PROJECT_ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = PROJECT_ROOT / "data"
FX_DIR = DATA_DIR / "market_data_for_risk_analysis"
FX_DIR.mkdir(parents=True, exist_ok=True)
EURUSD_FILE = FX_DIR / "EURUSD.csv"
FEATURE_FILE = YF_RAW_FILE

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)
logger = logging.getLogger(__name__)

ASSET_CLASS_FILES = {
    "Stocks": FILES["stocks"],
    "ETF": FILES["etfs"],
    "Crypto": FILES["crypto"],
    "Commodity": FILES["commodities"],
}

# =============================================================================
# Clean Date Helper
# =============================================================================
def clean_date_series(series: pd.Series) -> pd.Series:
    """Forces dates to be purely Year-Month-Day at 00:00:00 with no timezones."""
    return pd.to_datetime(series, utc=True).dt.tz_convert(None).dt.normalize()

# =============================================================================
# Incremental & FX Logic
# =============================================================================
def update_and_load_eurusd() -> pd.Series:
    today = pd.Timestamp.today().normalize()
    
    if EURUSD_FILE.exists():
        fx_df = pd.read_csv(EURUSD_FILE, parse_dates=["Date"])
        last_date = fx_df["Date"].max()
    else:
        fx_df = pd.DataFrame(columns=["Date", "EURUSD"])
        last_date = today - pd.DateOffset(years=5)

    if last_date < today:
        fetch_start = last_date + pd.Timedelta(days=1)
        logger.info(f"Updating EURUSD FX rates from {fetch_start.date()} to {today.date()}")
        
        new_fx = yf.download("EURUSD=X", start=fetch_start, end=today + pd.Timedelta(days=1), progress=False)
        if not new_fx.empty:
            new_fx = new_fx.reset_index()
            if isinstance(new_fx.columns, pd.MultiIndex):
                new_fx.columns = [c[0] for c in new_fx.columns]
            
            new_fx = new_fx[["Date", "Close"]].rename(columns={"Close": "EURUSD"})
            fx_df = pd.concat([fx_df, new_fx], ignore_index=True)
            fx_df.drop_duplicates(subset=["Date"], keep="last", inplace=True)
            fx_df.to_csv(EURUSD_FILE, index=False)

    fx_df["Date"] = clean_date_series(fx_df["Date"])
    fx_df.set_index("Date", inplace=True)
    return fx_df["EURUSD"]

def get_incremental_dates() -> tuple[pd.Timestamp, pd.Timestamp, pd.DataFrame]:
    today = pd.Timestamp.today().normalize()
    
    if FEATURE_FILE.exists():
        existing_df = pd.read_parquet(FEATURE_FILE)
        existing_df["Date"] = clean_date_series(existing_df["Date"])
        last_date = existing_df["Date"].max()
        # Roll back 14 days to catch any weekend data that needs to merge into Monday
        start_date = last_date - pd.Timedelta(days=14)
        return start_date, today, existing_df
    else:
        starts = []
        for file in ASSET_CLASS_FILES.values():
            df = pd.read_csv(file)
            starts.append(pd.to_datetime(df.iloc[:, 0]).min())
        return min(starts), today, pd.DataFrame()

# =============================================================================
# Download & Cleaning Helpers
# =============================================================================
def get_tickers() -> pd.DataFrame:
    records = []
    for asset_class, file in ASSET_CLASS_FILES.items():
        df = pd.read_csv(file, nrows=1)
        tickers = list(df.columns[1:])
        for ticker in tickers:
            records.append({"Ticker": ticker, "Asset_Class": asset_class})
    return pd.DataFrame(records).drop_duplicates().sort_values("Ticker").reset_index(drop=True)

def flatten_columns(df: pd.DataFrame) -> pd.DataFrame:
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.get_level_values(0)
    return df

def clean_and_convert(df: pd.DataFrame, ticker: str, asset_class: str, fx_series: pd.Series) -> pd.DataFrame:
    if df.empty: return df

    df["Date"] = clean_date_series(df["Date"])
    keep_cols = ["Date", "Open", "High", "Low", "Close", "Volume", "Dividends", "Stock Splits"]
    df = df[[c for c in keep_cols if c in df.columns]].copy()
    
    try:
        currency = yf.Ticker(ticker).fast_info.get("currency", "EUR")
    except Exception:
        currency = "USD"
        
    if currency == "USD":
        df = df.merge(fx_series.reset_index(), on="Date", how="left")
        df["EURUSD"] = df["EURUSD"].ffill().bfill() 
        
        for col in ["Open", "High", "Low", "Close", "Dividends"]:
            if col in df.columns:
                df[col] = df[col] / df["EURUSD"]
        df.drop(columns=["EURUSD"], inplace=True)
    
    df["Ticker"] = ticker
    df["Asset_Class"] = asset_class
    df["Data_Source"] = "YahooFinance"
    return df

def download_all_features(start_date: pd.Timestamp, end_date: pd.Timestamp) -> pd.DataFrame:
    if start_date >= end_date: return pd.DataFrame()

    fx_series = update_and_load_eurusd()
    tickers_df = get_tickers()
    all_data = []

    logger.info(f"Downloading incrementally from {start_date.date()} to {end_date.date()}")

    for _, row in tickers_df.iterrows():
        ticker = row["Ticker"]
        asset_class = row["Asset_Class"]

        if asset_class == "ETF" and ticker in SYNTHETIC_ETFS: continue

        data = yf.download(ticker, start=start_date.strftime("%Y-%m-%d"), end=(end_date + pd.Timedelta(days=1)).strftime("%Y-%m-%d"), progress=False, actions=True)
        if not data.empty:
            data.reset_index(inplace=True)
            df = flatten_columns(data)
            df = clean_and_convert(df, ticker, asset_class, fx_series)
            all_data.append(df)
        time.sleep(YFINANCE_SLEEP)

    return pd.concat(all_data, ignore_index=True) if all_data else pd.DataFrame()

def append_synthetic_etfs(start_date: pd.Timestamp) -> pd.DataFrame:
    if not SYNTHETIC_ETFS: 
        return pd.DataFrame()

    tickers_to_process = [SYNTHETIC_ETFS] if isinstance(SYNTHETIC_ETFS, str) else SYNTHETIC_ETFS
    
    # Check for existing feature file and load it once
    existing_df = None
    if FEATURE_FILE.exists():
        existing_df = pd.read_parquet(FEATURE_FILE)
        if "Date" in existing_df.columns:
            existing_df["Date"] = pd.to_datetime(existing_df["Date"])
            
    # load the CSV only if needed for tickers without existing history
    etf_prices = None
    synthetic_frames = []

    for ticker in tickers_to_process:
        # Check if we have existing data for this specific ticker
        has_existing_data = False
        if existing_df is not None and "Ticker" in existing_df.columns:
            ticker_mask = existing_df["Ticker"] == ticker
            if ticker_mask.any():
                has_existing_data = True

        if has_existing_data:
            # --- CASE 1: Ticker exists in feature file. Fetch new data from yfinance ---
            last_date = existing_df.loc[ticker_mask, "Date"].max()
            fetch_start = (last_date + pd.Timedelta(days=1)).strftime('%Y-%m-%d')
            
            ticker_obj = yf.Ticker(ticker)
            yf_data = ticker_obj.history(start=fetch_start, auto_adjust=True)
            
            if not yf_data.empty:
                yf_data.reset_index(inplace=True)
                
                temp = pd.DataFrame({
                    "Date": clean_date_series(yf_data["Date"]),
                    "Open": np.nan, "High": np.nan, 
                    "Low": np.nan, "Close": yf_data["Close"],
                    "Volume": np.nan, "Dividends": np.nan, "Stock Splits": np.nan,
                    "Ticker": ticker, "Asset_Class": "ETF", "Data_Source": "Synthetic"
                })
                
                temp.dropna(subset=["Close"], inplace=True)
                if not temp.empty:
                    synthetic_frames.append(temp)
                    
        else:
            # --- CASE 2: Ticker NOT in feature file. Use original CSV fallback logic ---
            if etf_prices is None:
                etf_prices = pd.read_csv(FILES["etfs"])
                etf_prices["Date"] = clean_date_series(etf_prices["Date"])

            if ticker not in etf_prices.columns:
                continue
            
            temp = pd.DataFrame({
                "Date": etf_prices["Date"],
                "Open": np.nan, "High": np.nan, 
                "Low": np.nan, "Close": etf_prices[ticker],
                "Volume": np.nan, "Dividends": np.nan, "Stock Splits": np.nan,
                "Ticker": ticker, "Asset_Class": "ETF", "Data_Source": "Synthetic"
            })
            
            temp.dropna(subset=["Close"], inplace=True)
            if not temp.empty:
                synthetic_frames.append(temp)

    return pd.concat(synthetic_frames, ignore_index=True) if synthetic_frames else pd.DataFrame()

# =============================================================================
# Core Logic: Master Calendar Alignment (Smashes Weekends)
# =============================================================================
def align_to_business_days(df: pd.DataFrame) -> pd.DataFrame:
    """
    Uses Stocks as the master calendar. Any crypto date (weekend/holiday) 
    is rolled forward to the NEXT valid Stock trading day, then grouped.
    This safely deletes weekend dates from the final output.
    """
    if df.empty: return df

    df["Date"] = clean_date_series(df["Date"])

    # 1. Establish the "Master Calendar" using Stock Market dates
    stock_data = df[df["Asset_Class"] == "Stocks"]
    if not stock_data.empty:
        valid_dates = pd.Series(stock_data["Date"].unique()).sort_values()
    else:
        # Fallback if no stocks exist in the chunk
        valid_dates = pd.Series(pd.bdate_range(df["Date"].min(), df["Date"].max() + pd.Timedelta(days=5))).dt.normalize()

    # 2. Map every single date (including Sat/Sun) to the NEXT valid stock date
    master_dates_df = pd.DataFrame({"Business_Date": valid_dates, "Date": valid_dates})
    all_dates = pd.DataFrame({"Date": df["Date"].drop_duplicates().sort_values()})

    # Forward direction means Friday -> Friday, Saturday -> Monday, Sunday -> Monday
    date_mapping = pd.merge_asof(all_dates, master_dates_df, on="Date", direction="forward")
    df = df.merge(date_mapping, on="Date", how="left")
    
    # Drop rows at the end of the dataset if Monday hasn't happened yet (e.g., if you run this on Sunday)
    df.dropna(subset=["Business_Date"], inplace=True)

    # 3. Group by Ticker and the new Business_Date (Automatically drops weekend rows)
    agg_rules = {
        "Open": "first",     
        "High": "max",       
        "Low": "min",        
        "Close": "last",     
        "Volume": lambda x: x.sum(min_count=1),     
        "Dividends": lambda x: x.sum(min_count=1),
        "Stock Splits": lambda x: x.sum(min_count=1),
        "Asset_Class": "first",
        "Data_Source": "first",
    }
    agg_rules = {k: v for k, v in agg_rules.items() if k in df.columns}

    aligned_df = (
        df.sort_values(["Ticker", "Date"])
          .groupby(["Ticker", "Business_Date"])
          .agg(agg_rules)
          .reset_index()
    )

    # Rename the mapped column back to 'Date'
    aligned_df.rename(columns={"Business_Date": "Date"}, inplace=True)
    return aligned_df

# =============================================================================
# Pipeline Trigger
# =============================================================================
def fetch_yfinance_features() -> pd.DataFrame:
    logger.info("=" * 70)
    logger.info("Starting Download & Business Calendar Alignment (Weekend Smashing)")
    logger.info("=" * 70)
    
    start_date, end_date, existing_df = get_incremental_dates()
    
    # Download raw unaligned data
    new_data = download_all_features(start_date, end_date)
    new_data.dropna(subset=["Open", "High", "Low", "Close"], how="all", inplace=True)
    synth_data = append_synthetic_etfs(start_date)
    new_data = pd.concat([new_data, synth_data], ignore_index=True) if not synth_data.empty else new_data

    if not new_data.empty:
        # Combine raw new data with existing history
        if not existing_df.empty:
            final_df = pd.concat([existing_df, new_data], ignore_index=True)
            final_df["Date"] = clean_date_series(final_df["Date"])
            final_df.sort_values(["Ticker", "Date"], inplace=True)
            final_df.drop_duplicates(subset=["Ticker", "Date"], keep="last", inplace=True)
        else:
            final_df = new_data
            
        # PERFORM MASTER ALIGNMENT ON ENTIRE HISTORY
        # This scrubs the old file clean of any lingering weekends and safely maps the new ones
        final_df = align_to_business_days(final_df)

        # Set Dividends and Stock Splits to pd.NA for Crypto and Commodity assets
        if "Asset_Class" in final_df.columns:
            non_corporate_mask = final_df["Asset_Class"].isin(["Crypto", "Commodity"])
            for col in ["Dividends", "Stock Splits"]:
                if col in final_df.columns:
                    final_df.loc[non_corporate_mask, col] = pd.NA

        final_df.to_parquet(FEATURE_FILE, index=False)
      #  final_df.to_csv(FEATURE_DIR / "yfinance_raw_features.csv", index=False)
        logger.info(f"Saved {len(final_df):,} cleanly aligned daily records to : {FEATURE_FILE}")
    else:
        final_df = existing_df
        logger.info("No new data downloaded. Existing file is up to date.")

    logger.info("=" * 70)
    return final_df

if __name__ == "__main__":
    df = fetch_yfinance_features()