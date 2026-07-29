"""
Downloads Yahoo Finance features for all tradable assets.
Downloads: Open, High, Low, Volume, Dividends, Stock Splits.
Drops: Close, Adj Close.
Converts all USD financial values to EUR automatically.
Handles incremental updates efficiently.

Output: Long-format dataframe
"""

from __future__ import annotations

import time
import logging
from pathlib import Path
import numpy as np
import pandas as pd
import yfinance as yf

from src.config import (FILES, FEATURE_DIR, SYNTHETIC_ETFS, YFINANCE_SLEEP)

# =============================================================================
# Directories & Logging
# =============================================================================
PROJECT_ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = PROJECT_ROOT / "data"
FX_DIR = DATA_DIR / "market_data_for_risk_analysis"
FX_DIR.mkdir(parents=True, exist_ok=True)
EURUSD_FILE = FX_DIR / "EURUSD.csv"
FEATURE_FILE = FEATURE_DIR / "yfinance_raw_features.parquet"

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
# Incremental & FX Logic
# =============================================================================

def update_and_load_eurusd() -> pd.Series:
    """
    Updates the EURUSD.csv file incrementally and returns a Date-indexed Series 
    of the exchange rate for vectorized conversions.
    """
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

    fx_df.set_index("Date", inplace=True)
    return fx_df["EURUSD"]


def get_incremental_dates() -> tuple[pd.Timestamp, pd.Timestamp, pd.DataFrame]:
    """
    Determines start date based on existing parquet file to avoid re-downloading.
    Returns (start_date, end_date, existing_dataframe).
    """
    today = pd.Timestamp.today().normalize()
    
    if FEATURE_FILE.exists():
        existing_df = pd.read_parquet(FEATURE_FILE)
        last_date = pd.to_datetime(existing_df["Date"]).max()
        start_date = last_date + pd.Timedelta(days=1)
        return start_date, today, existing_df
    else:
        # Fallback to the earliest date from processed files if no parquet exists
        starts = []
        for file in ASSET_CLASS_FILES.values():
            df = pd.read_csv(file)
            starts.append(pd.to_datetime(df.iloc[:, 0]).min())
        return min(starts), today, pd.DataFrame()

# =============================================================================
# Download & Processing Helpers
# =============================================================================

def get_tickers() -> pd.DataFrame:
    records = []
    for asset_class, file in ASSET_CLASS_FILES.items():
        df = pd.read_csv(file, nrows=1)
        tickers = list(df.columns[1:])
        for ticker in tickers:
            records.append({"Ticker": ticker, "Asset_Class": asset_class})

    return pd.DataFrame(records).drop_duplicates().sort_values("Ticker").reset_index(drop=True)


def download_single_ticker(ticker: str, start_date: str, end_date: str) -> pd.DataFrame:
    try:
        data = yf.download(ticker, start=start_date, end=end_date, auto_adjust=False, progress=False, actions=True, threads=False)
        if data.empty:
            return pd.DataFrame()
        data.reset_index(inplace=True)
        return data
    except Exception as e:
        logger.error(f"{ticker} download failed : {e}")
        return pd.DataFrame()

def flatten_columns(df: pd.DataFrame) -> pd.DataFrame:
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [col[0] for col in df.columns]
    return df

def clean_and_convert(df: pd.DataFrame, ticker: str, asset_class: str, fx_series: pd.Series) -> pd.DataFrame:
    """
    Standardizes names, drops Close/Adj Close, and applies EUR conversion.
    """
    if df.empty:
        return df

    df = flatten_columns(df)
    
    # 1. Keep only the following columns
    keep = ["Date", "Open", "High", "Low", "Volume", "Dividends", "Stock Splits"]
    existing = [c for c in keep if c in df.columns]
    df = df[existing].copy()
    
    # 2. Get asset currency to check if FX conversion is needed
    try:
        currency = yf.Ticker(ticker).fast_info.get("currency", "EUR")
    except:
        currency = "USD" # default fallback
        
    # 3. Convert USD values to EUR
    if currency == "USD":
        df = df.merge(fx_series, on="Date", how="left")
        # ffill and bfill handles crypto/weekends where FX market is closed
        df["EURUSD"] = df["EURUSD"].ffill().bfill() 
        
        financial_cols = ["Open", "High", "Low", "Dividends"]
        for col in financial_cols:
            if col in df.columns:
                # 1 EUR = X USD  ==> EUR Value = USD Value / X
                df[col] = df[col] / df["EURUSD"]
                
        df.drop(columns=["EURUSD"], inplace=True)
    
    df["Ticker"] = ticker
    df["Asset_Class"] = asset_class
    df["Data_Source"] = "YahooFinance"
    return df

# =============================================================================
# Main Download Function
# =============================================================================

def download_all_features(start_date: pd.Timestamp, end_date: pd.Timestamp) -> pd.DataFrame:
    if start_date >= end_date:
        logger.info("Data is already up to date. No new downloads needed.")
        return pd.DataFrame()

    fx_series = update_and_load_eurusd()
    tickers_df = get_tickers()
    all_data = []

    logger.info(f"Downloading incrementally from {start_date.date()} to {end_date.date()}")

    for _, row in tickers_df.iterrows():
        ticker = row["Ticker"]
        asset_class = row["Asset_Class"]

        if asset_class == "ETF" and ticker in SYNTHETIC_ETFS:
            continue

        df = download_single_ticker(ticker, start_date.strftime("%Y-%m-%d"), (end_date + pd.Timedelta(days=1)).strftime("%Y-%m-%d"))
        
        if not df.empty:
            df = clean_and_convert(df, ticker, asset_class, fx_series)
            all_data.append(df)
            
        time.sleep(YFINANCE_SLEEP)

    if not all_data:
        return pd.DataFrame()

    master = pd.concat(all_data, ignore_index=True)
    return master


def append_synthetic_etfs(yahoo_df: pd.DataFrame, start_date: pd.Timestamp) -> pd.DataFrame:
    """
    Since Close is dropped, we map ETF prices to Open/High/Low for consistency.
    """
    if len(SYNTHETIC_ETFS) == 0:
        return yahoo_df

    etf_prices = pd.read_csv(FILES["etfs"])
    etf_prices.rename(columns={etf_prices.columns[0]: "Date"}, inplace=True)
    etf_prices["Date"] = pd.to_datetime(etf_prices["Date"], errors="coerce")
    
    # Filter for the incremental date range
    etf_prices = etf_prices[etf_prices["Date"] >= start_date]

    synthetic_frames = []
    for ticker in SYNTHETIC_ETFS:
        if ticker not in etf_prices.columns:
            continue

        temp = pd.DataFrame()
        temp["Date"] = etf_prices["Date"]
        temp["Open"] = np.nan #etf_prices[ticker]  # Map price to Open, High, Low
        temp["High"] = np.nan #etf_prices[ticker]
        temp["Low"] = np.nan #etf_prices[ticker]
        temp["Volume"] = np.nan
        temp["Dividends"] = np.nan
        temp["Stock Splits"] = np.nan
        temp["Ticker"] = ticker
        temp["Asset_Class"] = "ETF"
        temp["Data_Source"] = "Synthetic"
        
        temp.dropna(subset=["Open"], inplace=True) # Drop blank dates
        synthetic_frames.append(temp)

    if synthetic_frames:
        yahoo_df = pd.concat([yahoo_df, *synthetic_frames], ignore_index=True)

    yahoo_df.sort_values(["Ticker", "Date"], inplace=True)
    yahoo_df.reset_index(drop=True, inplace=True)
    return yahoo_df

# =============================================================================
# Pipeline Trigger
# =============================================================================

def fetch_yfinance_features() -> pd.DataFrame:
    logger.info("=" * 70)
    logger.info("Starting Yahoo Finance Feature Download (Incremental & EUR Converted)")
    logger.info("=" * 70)
    
    start_date, end_date, existing_df = get_incremental_dates()
    
    new_data = download_all_features(start_date, end_date)
    
    if not new_data.empty:
        new_data = append_synthetic_etfs(new_data, start_date)
        
        if not existing_df.empty:
            # Combine historical and new data
            final_df = pd.concat([existing_df, new_data], ignore_index=True)
            final_df.drop_duplicates(subset=["Ticker", "Date"], keep="last", inplace=True)
        else:
            final_df = new_data
            
        final_df.sort_values(["Ticker", "Date"], inplace=True)
        final_df.reset_index(drop=True, inplace=True)

        final_df.to_parquet(FEATURE_FILE, index=False)
        logger.info(f"Saved total {len(final_df):,} records to : {FEATURE_FILE}")
    else:
        final_df = existing_df
        logger.info("No new data downloaded. Existing file is up to date.")

    logger.info("=" * 70)
    return final_df

if __name__ == "__main__":
    df = fetch_yfinance_features()
    print(df.tail())