"""
Contains functions to perform data quality assurance, missing value analysis, 
and redundancy checks on the master feature dataset.
"""

import pandas as pd
import numpy as np

def get_dataset_overview(df: pd.DataFrame) -> dict:
    """
    Returns a dictionary containing core dataset metrics and asset class balance.
    """
    overview = {
        "Total Rows": len(df),
        "Total Columns": len(df.columns),
        "Unique Assets": df["Ticker"].nunique(),
        "Unique Dates": df["Date"].nunique(),
        "Date Range": f"{df['Date'].min().date()} to {df['Date'].max().date()}",
        "Duplicate Rows": df.duplicated().sum(),
        "Duplicate (Date, Ticker)": df.duplicated(subset=["Date", "Ticker"]).sum()
    }
    
    # Asset Class Balance (Step 7)
    class_counts = df["Asset_Class"].value_counts().to_dict()
    overview["Asset Class Balance"] = class_counts
    
    return overview


def get_missing_values_by_class(df: pd.DataFrame) -> pd.DataFrame:
    """
    Calculates the percentage of missing values per feature, broken down by Asset Class.
    """
    # Calculate % missing per column, grouped by Asset Class
    missing_df = df.groupby("Asset_Class").apply(
        lambda x: (x.isnull().sum() / len(x)) * 100
    ).T
    
    # Add an "Overall" missing percentage column
    missing_df["Overall (%)"] = (df.isnull().sum() / len(df)) * 100
    
    # Sort by Overall missing for easier reading, format to 2 decimal places
    missing_df = missing_df.sort_values(by="Overall (%)", ascending=False).round(2)
    
    return missing_df


def check_logical_bounds(df: pd.DataFrame) -> pd.DataFrame:
    """
    Checks specific features for impossible values (e.g., negative volatility).
    """
    errors = []
    
    # 1. RSI should be strictly between 0 and 100
    if "RSI14" in df.columns:
        invalid_rsi = df[(df["RSI14"] < 0) | (df["RSI14"] > 100)]
        if not invalid_rsi.empty:
            errors.append({"Feature": "RSI14", "Error": "Values outside 0-100", "Count": len(invalid_rsi)})
            
    # 2. Volatility should be >= 0
    vol_cols = [c for c in df.columns if "Volatility" in c]
    for col in vol_cols:
        invalid_vol = df[df[col] < 0]
        if not invalid_vol.empty:
            errors.append({"Feature": col, "Error": "Negative volatility", "Count": len(invalid_vol)})
            
    # 3. Bollinger Width should be >= 0
    if "Bollinger_Width" in df.columns:
        invalid_bw = df[df["Bollinger_Width"] < 0]
        if not invalid_bw.empty:
            errors.append({"Feature": "Bollinger_Width", "Error": "Negative width", "Count": len(invalid_bw)})
            
    if not errors:
        return pd.DataFrame([{"Status": "All logical boundary checks passed. No impossible values found."}])
        
    return pd.DataFrame(errors)

def find_redundant_features(df: pd.DataFrame, corr_threshold: float = 0.90) -> pd.DataFrame:
    """
    Finds zero-variance features and highly correlated feature pairs using Spearman.
    """
    # 1. Zero Variance Features
    numeric_df = df.select_dtypes(include=[np.number])
    variances = numeric_df.var()
    zero_var_cols = variances[variances == 0].index.tolist()
    
    results = {"Zero_Variance_Features": zero_var_cols}
    
    # 2. Highly Correlated Pairs (Spearman)
    # Drop zero variance cols before correlation to avoid NaNs
    valid_numeric = numeric_df.drop(columns=zero_var_cols)
    corr_matrix = valid_numeric.corr(method="spearman").abs()
    
    # Extract upper triangle to avoid duplicate pairs (A-B and B-A)
    upper = corr_matrix.where(np.triu(np.ones(corr_matrix.shape), k=1).astype(bool))
    
    high_corr_pairs = []
    for col in upper.columns:
        highly_correlated_with_col = upper[col][upper[col] > corr_threshold].index.tolist()
        for correlated_col in highly_correlated_with_col:
            high_corr_pairs.append({
                "Feature_1": correlated_col,
                "Feature_2": col,
                "Correlation": round(upper.loc[correlated_col, col], 4)
            })
            
    high_corr_df = pd.DataFrame(high_corr_pairs).sort_values(by="Correlation", ascending=False)
    
    return zero_var_cols, high_corr_df

def compare_feature_distributions_by_class(df: pd.DataFrame, features: list) -> pd.DataFrame:
    """
    Tests whether each feature's distribution differs meaningfully across Asset
    Classes. Kruskal-Wallis runs on one median PER TICKER, not per row, and
    only across classes with 2+ tickers -- single-instrument classes (Bond,
    RealEstate) are excluded from the test itself, not from the calculation.
    Median_Spread_Normalized still uses every class, since it's descriptive
    and carries no such requirement.
    """
    from scipy.stats import kruskal

    results = []
    for feature in features:
        clean = df[["Ticker", "Asset_Class", feature]].dropna()

        per_ticker = clean.groupby(["Ticker", "Asset_Class"])[feature].median().reset_index()
        class_sizes = per_ticker.groupby("Asset_Class")["Ticker"].nunique()
        testable = class_sizes[class_sizes >= 2].index

        groups = [g[feature].values for cls, g in per_ticker.groupby("Asset_Class") if cls in testable]
        stat, p_value = kruskal(*groups) if len(groups) >= 2 else (np.nan, np.nan)

        medians = clean.groupby("Asset_Class")[feature].median()
        overall_iqr = clean[feature].quantile(0.75) - clean[feature].quantile(0.25)
        spread = (medians.max() - medians.min()) / overall_iqr if overall_iqr > 0 else np.nan

        results.append({
            "Feature": feature,
            "H_Statistic": stat,
            "P_Value": p_value,
            "N_Classes_Tested": len(groups),
            "Median_Spread_Normalized": spread,
            "Highest_Median_Class": medians.idxmax(),
            "Lowest_Median_Class": medians.idxmin(),
        })

    return pd.DataFrame(results).sort_values("Median_Spread_Normalized", ascending=False).reset_index(drop=True)

def per_ticker_summary_for_small_classes(df: pd.DataFrame, features: list, max_class_size: int = 5) -> pd.DataFrame:
    """
    Shows each ticker's own median next to its peers, for asset classes with
    few tickers only -- so a class-level number driven by one atypical
    member is visible before you rely on it elsewhere. Descriptive, not a
    test: at 2-4 tickers there usually isn't enough data to statistically
    call one an "outlier" versus just a different instrument -- read it
    against what you actually know about the asset, the same way you did
    for Bond's RSI14 sitting at exactly 100.
    """
    clean = df.dropna(subset=["Ticker", "Asset_Class"])
    per_ticker = clean.groupby(["Ticker", "Asset_Class"])[features].median()
    class_sizes = per_ticker.reset_index().groupby("Asset_Class")["Ticker"].nunique()
    small_classes = class_sizes[class_sizes <= max_class_size].index

    out = per_ticker.reset_index()
    return out[out["Asset_Class"].isin(small_classes)].sort_values(["Asset_Class", "Ticker"]).reset_index(drop=True)