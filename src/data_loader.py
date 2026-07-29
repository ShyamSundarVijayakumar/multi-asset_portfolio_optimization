"""
Loads all processed asset class datasets.
Returns a clean, unified pandas DataFrame or continuous dictionary.
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
    2. Resamples/reindexes across a complete daily calendar sequence.
    3. Forward-fills missing values (e.g., holidays/weekends for tradable assets).
    """
    # Merge all asset datasets into a single unified DataFrame
    merged_df = None
    for df in data_dict.values():
        if merged_df is None:
            merged_df = df.copy()
        else:
            merged_df = pd.merge(merged_df, df, on="Date", how="outer")
            
    merged_df.sort_values("Date", inplace=True)
    
    # Create a complete daily date range from minimum to maximum date
    min_date = merged_df["Date"].min()
    max_date = merged_df["Date"].max()
    full_date_range = pd.date_range(start=min_date, end=max_date, freq="D", name="Date")
    
    # Set Date as index, reindex to complete calendar, and forward fill
    merged_df.set_index("Date", inplace=True)
    merged_df = merged_df.reindex(full_date_range)
    
    # Forward-fill prices/values for weekends/holidays, then backfill remaining leading NaNs
    merged_df = merged_df.ffill().bfill()
    
    # Reset index to restore 'Date' as a regular column
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