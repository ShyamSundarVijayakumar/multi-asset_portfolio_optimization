"""
Target Engineering (Stage 06)

Builds the two regression targets used by both downstream models:
    - Forward_Return_20D      : 20-day forward log return (Target 1)
    - Forward_Volatility_20D  : 20-day forward annualized realized volatility (Target 2)
"""

import numpy as np
import pandas as pd


def compute_forward_return(df: pd.DataFrame, horizon: int = 20) -> pd.DataFrame:
    """
    Adds Forward_Return_{horizon}D = ln(P_{t+horizon} / P_t), per ticker.
    Target 1. Log return, for consistency with the existing backward-looking
    Log_Return features, and because it composes correctly over time (needed
    below for the volatility target).
    """
    df = df.copy()
    future_price = df.groupby("Ticker")["Close"].shift(-horizon)
    df[f"Forward_Return_{horizon}D"] = np.log(future_price / df["Close"])
    return df


def compute_forward_volatility(df: pd.DataFrame, horizon: int = 20, daily_return_col: str = "Log_Return_1D") -> pd.DataFrame:
    """
    Adds Forward_Volatility_{horizon}D = sqrt(252) * Std(r_(t+1),...,r_(t+horizon)),
    per ticker. Target 2. Shifts the daily-return series forward by one day,
    then runs a reversed rolling window, since pandas rolling() only looks
    backward natively.
    """
    df = df.copy()
    
    def _forward_std(x: pd.Series) -> pd.Series:
        shifted = x.shift(-1)  # shifted[t] = r_{t+1}
        return shifted[::-1].rolling(horizon).std()[::-1]
        
    daily_std = df.groupby("Ticker")[daily_return_col].transform(_forward_std)
    df[f"Forward_Volatility_{horizon}D"] = daily_std * np.sqrt(252)
    return df


def drop_incomplete_rows(df: pd.DataFrame, required_cols: list, target_cols: list) -> pd.DataFrame:
    """
    Drops rows missing any REQUIRED feature or target -- head-of-series
    warm-up and tail-of-series target unavailability. Deliberately excludes
    structurally-optional features (Volume_Change, Relative_Volume,
    OBV_Z_Score) from `required_cols` -- those are legitimately NaN for
    Bond, RealEstate, and the synthetic ETF, and including them here would
    silently delete every row for those asset classes. Leave them NaN;
    LightGBM/XGBoost route around them natively.
    """
    before = len(df)
    clean = df.dropna(subset=required_cols + target_cols).reset_index(drop=True)
    after = len(clean)
    print(f"Dropped {before - after} of {before} rows ({(before - after) / before:.1%}); {after} remain.")
    return clean


def assemble_final_dataset(df: pd.DataFrame, feature_cols: list, target_cols: list) -> pd.DataFrame:
    """Selects Date, Ticker, Asset_Class, features, targets -- the training schema."""
    return df[["Date", "Ticker", "Asset_Class"] + feature_cols + target_cols].copy()


def compute_targets(df: pd.DataFrame, horizon: int = 20) -> pd.DataFrame:
    """Orchestrator: sorts, then runs the full target construction sequence."""
    df = df.sort_values(["Ticker", "Date"]).reset_index(drop=True).copy()
    df = compute_forward_return(df, horizon=horizon)
    df = compute_forward_volatility(df, horizon=horizon)
    return df


def summarize_targets(df: pd.DataFrame, target_cols: list = ["Forward_Return_20D", "Forward_Volatility_20D"]) -> None:
    """Descriptive stats overall and by class, plus a defensive non-negativity check on volatility."""
    print(df[target_cols].describe())
    print("\nBy Asset_Class:")
    print(df.groupby("Asset_Class")[target_cols].describe())
    vol_col = [c for c in target_cols if "Volatility" in c][0]
    assert (df[vol_col].dropna() >= 0).all(), f"{vol_col} has negative values -- should be mathematically impossible"

