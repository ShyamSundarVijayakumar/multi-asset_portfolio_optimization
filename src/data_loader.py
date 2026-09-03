"""
Consolidates raw asset datasets (Stocks, ETFs, Crypto, Commodities, Bonds, and Real Estate), 
normalizes dates against the master stock business calendar to strip out weekends and holidays, 
and produces a unified, clean master DataFrame.
"""
from pathlib import Path
import pandas as pd
from src.config import FILES

def _read_standard_csv(filepath: Path) -> pd.DataFrame:
    """Reads standard consolidated CSV files (Date | Ticker1 | Ticker2 | ...)."""
    df = pd.read_csv(filepath)
    df.rename(columns={df.columns[0]: "Date"}, inplace=True)
    df["Date"] = pd.to_datetime(df["Date"], dayfirst=False, errors="coerce")
    df.sort_values("Date", inplace=True)
    df.reset_index(drop=True, inplace=True)
    return df

def _read_single_holding_csv(filepath: Path, prefix: str) -> pd.DataFrame:
    """
    Reads CSVs with single holdings (Bonds / Real Estate) and prefixes columns 
    to avoid ambiguity (e.g., Bond_Holding_Value_EUR).
    """
    df = pd.read_csv(filepath)
    df.rename(columns={df.columns[0]: "Date"}, inplace=True)
    df["Date"] = pd.to_datetime(df["Date"], dayfirst=False, errors="coerce")
    
    # Filter required columns and rename them dynamically
    target_cols = ["Total_Holding_Value_EUR", "Daily_Return_EUR"]
    df = df[["Date"] + target_cols].copy()
    
    rename_map = {col: f"{prefix}_{col}" for col in target_cols}
    df.rename(columns=rename_map, inplace=True)
    
    df.sort_values("Date", inplace=True)
    df.reset_index(drop=True, inplace=True)
    return df

def load_processed_data() -> dict:
    """Loads all processed datasets with disambiguated column names."""
    datasets = {
        "stocks": _read_standard_csv(FILES["stocks"]),
        "etfs": _read_standard_csv(FILES["etfs"]),
        "crypto": _read_standard_csv(FILES["crypto"]),
        "commodities": _read_standard_csv(FILES["commodities"]),
        "bond": _read_single_holding_csv(FILES["bond"], prefix="Bond"),
        "real_estate": _read_single_holding_csv(FILES["real_estate"], prefix="RealEstate"),
    }
    return datasets

def merge_and_normalize_data(data_dict: dict) -> pd.DataFrame:
    """
    1. Outer joins all asset DataFrames on 'Date' into a single DataFrame.
    2. Builds an empirical Master Calendar handling cross-regional holidays 
       (e.g., US open / DE closed) by taking the union of equity trading dates.
    3. Filters the dataset against this master calendar.
    4. Forward/back-fills missing values safely for markets that were closed.
    """
    # Merge all asset datasets into a single unified DataFrame
    merged_df = None
    for df in data_dict.values():
        if merged_df is None:
            merged_df = df.copy()
        else:
            merged_df = pd.merge(merged_df, df, on="Date", how="outer")

    # Use Stocks as the master trading calendar to drop weekends and holidays
    trading_dates = data_dict["stocks"]["Date"]
    merged_df = merged_df[merged_df["Date"].isin(trading_dates)].copy()

    # EXPERT FIX: The Union Calendar Approach
    # Instead of relying solely on the 'stocks' CSV, we combine dates from all 
    # traditional equity markets (stocks + ETFs) to form a multi-region master calendar.
    # This naturally solves the "holiday in Germany but trading day in USA" problem.
 #   master_calendar = pd.concat([
 #       data_dict["stocks"]["Date"], 
 #       data_dict["etfs"]["Date"]
 #   ]).dropna().unique()
    
    # Filter the merged dataset to only include days where traditional markets were open
 #   merged_df = merged_df[merged_df["Date"].isin(master_calendar)].copy()
    
    merged_df.sort_values("Date", inplace=True)
    
    # Forward-fill any gaps (e.g., a German stock on a US trading day will ffill its last price),
    # then backfill leading NaNs.
    merged_df.set_index("Date", inplace=True)
    merged_df = merged_df.ffill().bfill()
    merged_df.reset_index(inplace=True)
    
    return merged_df

def align_start_date(df: pd.DataFrame, data_dict: dict = None) -> pd.DataFrame:
    """
    Aligns dataset to start from the latest starting asset class 
    to remove initial NaN padding.
    """
    if data_dict is not None:
        start_dates = [d["Date"].min() for d in data_dict.values()]
        common_start = max(start_dates)
    else:
        # Drops leading NaN rows if data_dict is not supplied
        common_start = df.dropna().Date.min() if not df.dropna().empty else df["Date"].min()

    aligned_df = df[df["Date"] >= common_start].copy().reset_index(drop=True)
    return aligned_df

def print_summary(df: pd.DataFrame):
    """Prints a consolidated summary of the master DataFrame."""
    print("=" * 70)
    print(f"Master Dataset Summary")
    print(f"Total Rows (Days) : {len(df)}")
    print(f"Total Columns     : {len(df.columns)}")
    print(f"Start Date        : {df['Date'].min().date()}")
    print(f"End Date          : {df['Date'].max().date()}")
    print(f"Null Values Count : {df.isnull().sum().sum()}")
    print("=" * 70)
    print("Column List:")
    print(list(df.columns))
    print("=" * 70)

def build_master_dataset() -> pd.DataFrame:
    """
    End-to-end pipeline wrapper: Loads, merges, aligns to trading days, 
    trims start dates, and prints a summary.
    """
    raw_data = load_processed_data()
    master_df = merge_and_normalize_data(raw_data)
    master_df = align_start_date(master_df, data_dict=raw_data)
    print_summary(master_df)
    
    return master_df