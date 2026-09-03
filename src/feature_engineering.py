
"""
Merges Yahoo Finance features with synthetic Bond and Real Estate data into a master market dataset,
computes vectorized technical indicators and statistical features for cross-sectional ML models,
and saves the finalized datasets as Parquet files.

"""

from pathlib import Path
import pandas as pd
import numpy as np
from src.config import FILES, FEATURE_DIR, SYNTHETIC_ETFS, YFINANCE_SLEEP, MASTER_FILE

# =============================================================================
# Helper Functions
# =============================================================================

def get_asset_class_map() -> dict:
    """
    Reads ETF tickers from file headers to build an Asset Class mapping dict.
    """
    asset_class_map = {}
    
    if "etfs" in FILES and Path(FILES["etfs"]).exists():
        df_etf = pd.read_csv(FILES["etfs"], nrows=1)
        etf_tickers = list(df_etf.columns[1:])  # Skip Date column
        for ticker in etf_tickers:
            asset_class_map[ticker] = "ETF"
            
    return asset_class_map

def merge_market_data(master_df: pd.DataFrame, yf_features: pd.DataFrame) -> pd.DataFrame:
    # --------------------------------------------------------
    # 1. Extract Synthetic Assets (Bond & Real Estate)
    # --------------------------------------------------------
    synthetic_configs = [
        {
            "col": "Bond_Total_Holding_Value_EUR",
            "ticker": "BOND",
            "asset_class": "Bond",
        },
        {
            "col": "RealEstate_Total_Holding_Value_EUR",
            "ticker": "REAL_ESTATE",
            "asset_class": "RealEstate",
        },
    ]

    synthetic_dfs = []
    for cfg in synthetic_configs:
        # Extract only Date and the Value column
        df_syn = master_df[["Date", cfg["col"]]].copy()
        
        # Rename to 'Close' to match the new yfinance feature schema
        df_syn.rename(columns={cfg["col"]: "Close"}, inplace=True)
        
        df_syn["Ticker"] = cfg["ticker"]
        df_syn["Asset_Class"] = cfg["asset_class"]
        
        # ML OPTIMIZATION: Use np.nan instead of pd.NA to preserve float64 dtype
        df_syn["Open"] = np.nan
        df_syn["High"] = np.nan
        df_syn["Low"] = np.nan
        df_syn["Volume"] = np.nan
        df_syn["Dividends"] = np.nan
        df_syn["Stock Splits"] = np.nan
        df_syn["Data_Source"] = "Synthetic"
        
        synthetic_dfs.append(df_syn)

    # --------------------------------------------------------
    # 2. Align and Combine All Assets
    # --------------------------------------------------------
    synthetic_market = pd.concat(synthetic_dfs, ignore_index=True)
    
    # Ensure columns strictly match yf_features before concatenation
    for col in yf_features.columns:
        if col not in synthetic_market.columns:
            synthetic_market[col] = np.nan

    synthetic_market = synthetic_market[yf_features.columns]

    # Combine the already up-to-date yf_features with the newly appended Bond/RE data
    master_market = pd.concat([yf_features, synthetic_market], ignore_index=True)
    
    master_market.sort_values(["Ticker", "Date"], inplace=True)
    master_market.reset_index(drop=True, inplace=True)

    return master_market

# =============================================================================
# Save
# =============================================================================

def save_market_data(df: pd.DataFrame):
    """
    Saves merged market dataset.
    """
    output_file = FEATURE_DIR / "merged_market_dataset.parquet"
    df.to_parquet(output_file, index=False)
  #  output_file = FEATURE_DIR / "merged_market_dataset.csv"
  #  df.to_csv(output_file, index=False)
    print(f"\nSaved : {output_file}")


def save_feature_dataset(df: pd.DataFrame):
    """
    Saves engineered feature dataset as Parquet for DuckDB/LightGBM ingestion.
    """
    output_file = MASTER_FILE #FEATURE_DIR / "complete_master_features.parquet"
    df.to_parquet(output_file, index=False)
 #   output_file = FEATURE_DIR / "complete_master_features.csv"
 #   df.to_csv(output_file, index=False)
    print(f"\nSaved master features to: {output_file}")


# =============================================================================
# Pipeline
# =============================================================================

def prepare_market_data(master_df: pd.DataFrame, yf_features: pd.DataFrame) -> pd.DataFrame:
    """
    Creates master market dataset.
    """
    market_df = merge_market_data(master_df, yf_features)
    save_market_data(market_df)
    return market_df

def prepare_master_market_data(master_market: pd.DataFrame) -> pd.DataFrame:
    """
    Creates master market dataset.
    """
    market_master_df = compute_features(master_market)
    save_feature_dataset(market_master_df)
    return market_master_df

# =============================================================================
# Technical Feature Engineering (Vectorized for Cross-Sectional ML)
# =============================================================================

def compute_features(master_market: pd.DataFrame) -> pd.DataFrame:
    """
    Computes all technical and statistical features for a global ML model.
    Uses vectorized operations for maximum efficiency on stacked datasets.
    """
    
    df = master_market.copy()
    
    # Sort strictly by Ticker and Date to ensure sequential integrity
    df.sort_values(["Ticker", "Date"], inplace=True)
    
    # Create groupby object for vectorized grouping
    grp = df.groupby("Ticker")
    
    # Extract base columns for cleaner code
    adj_close = df["Close"]
    volume = df.get("Volume", pd.Series(dtype=float))
    
    # --------------------------------------------------------
    # Log Returns & Rolling Returns (1, 5, 20, 60)
    # --------------------------------------------------------
    for window in [1, 5, 20, 60]:
        df[f"Log_Return_{window}D"] = np.log(adj_close / grp["Close"].shift(window))
        
    # --------------------------------------------------------
    # Rolling Volatility
    # --------------------------------------------------------
    df["Volatility_20D"] = grp["Log_Return_1D"].transform(lambda x: x.rolling(20).std())
    df["Volatility_60D"] = grp["Log_Return_1D"].transform(lambda x: x.rolling(60).std())
    
    # --------------------------------------------------------
    # Momentum (SMA Ratios)
    # --------------------------------------------------------
    sma20 = grp["Close"].transform(lambda x: x.rolling(20).mean())
    sma60 = grp["Close"].transform(lambda x: x.rolling(60).mean())
    
    df["SMA20_Ratio"] = adj_close / sma20
    df["SMA60_Ratio"] = adj_close / sma60
    
    # --------------------------------------------------------
    # Maximum Drawdown (60-day Rolling)
    # --------------------------------------------------------
    # 1. Peak price achieved within the last 60 days (starts on Row 1)
    rolling_max_60 = grp["Close"].transform(lambda x: x.rolling(60, min_periods=1).max())
    
    # 2. Percentage drop from that 60-day peak to today's price
    df["Max_Drawdown_60D"] = (adj_close / rolling_max_60) - 1
    
    # --------------------------------------------------------
    # Technical Indicators (EMA, RSI, MACD, Bollinger)
    # --------------------------------------------------------
    ema20 = grp["Close"].transform(lambda x: x.ewm(span=20, adjust=False).mean())
    ema50 = grp["Close"].transform(lambda x: x.ewm(span=50, adjust=False).mean())
    df["EMA20_Ratio"] = adj_close / ema20
    df["EMA50_Ratio"] = adj_close / ema50
    
    # MACD (12, 26, 9) — expressed as % of EMA26 (this is the standard "Percentage Price Oscillator" formulation of MACD),
    # so the value means the same thing regardless of the asset's price level
    ema12 = grp["Close"].transform(lambda x: x.ewm(span=12, adjust=False).mean())
    ema26 = grp["Close"].transform(lambda x: x.ewm(span=26, adjust=False).mean())
    df["MACD"] = (ema12 - ema26) / ema26
    df["MACD_Signal"] = df.groupby("Ticker")["MACD"].transform(lambda x: x.ewm(span=9, adjust=False).mean())
    df["MACD_Histogram"] = df["MACD"] - df["MACD_Signal"]
   
    # RSI (14-day)
    price_diff = grp["Close"].diff()
    gain = price_diff.clip(lower=0)
    loss = -price_diff.clip(upper=0)
    
    avg_gain = gain.groupby(df["Ticker"]).transform(lambda x: x.ewm(alpha=1/14, adjust=False).mean())
    avg_loss = loss.groupby(df["Ticker"]).transform(lambda x: x.ewm(alpha=1/14, adjust=False).mean())
    
    rs = avg_gain / avg_loss
    df["RSI14"] = 100 - (100 / (1 + rs))
    
    # Bollinger Width ( (Upper - Lower) / Middle ) = (4 * StdDev) / SMA20
    std20 = grp["Close"].transform(lambda x: x.rolling(20).std())
    df["Bollinger_Width"] = (4 * std20) / sma20
    
    # --------------------------------------------------------
    # Volume Features (Handles NaNs)
    # --------------------------------------------------------
    if "Volume" in df.columns:
        vol_sma20 = grp["Volume"].transform(lambda x: x.rolling(20).mean())
        df["Relative_Volume"] = volume / vol_sma20
        
        # FIX: fill_method=None to silence FutureWarning
        df["Volume_Change"] = grp["Volume"].pct_change(1, fill_method=None)
        
        # FIX: Division by 0 causes inf. Replace with np.nan.
        df["Volume_Change"] = df["Volume_Change"].replace([np.inf, -np.inf], np.nan)
        
        # On-Balance Volume (OBV)
        direction = np.sign(price_diff.fillna(0))
        daily_obv = direction * volume
        df["OBV"] = daily_obv.groupby(df["Ticker"]).cumsum()

        obv_sma20 = df.groupby("Ticker")["OBV"].transform(lambda x: x.rolling(20).mean())
        obv_std20 = df.groupby("Ticker")["OBV"].transform(lambda x: x.rolling(20).std())
        
        df["OBV_Z_Score"] = (df["OBV"] - obv_sma20) / obv_std20
        df["OBV_Z_Score"] = df["OBV_Z_Score"].replace([np.inf, -np.inf], np.nan)
    
    return df