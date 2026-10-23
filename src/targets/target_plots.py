"""
Target Engineering (Stage 06) -- sanity-check plots for the constructed targets.
"""

import matplotlib.pyplot as plt
import seaborn as sns
import pandas as pd


def plot_target_distributions(df: pd.DataFrame, return_col: str = "Forward_Return_20D", vol_col: str = "Forward_Volatility_20D"):
    """
    Histogram of the return target (expect roughly centered near zero, fat
    tails -- same reasoning as in EDA, not normal) and the volatility
    target (expect strictly positive, right-skewed).
    """
    fig, axes = plt.subplots(1, 2, figsize=(14, 4.5))
    sns.histplot(df[return_col].dropna(), bins=100, ax=axes[0])
    axes[0].set_title(return_col)
    axes[0].axvline(0, color="grey", linestyle="--", linewidth=1)
    sns.histplot(df[vol_col].dropna(), bins=100, ax=axes[1])
    axes[1].set_title(vol_col)
    plt.tight_layout()
    plt.show()


def plot_target_by_class(df: pd.DataFrame, return_col: str = "Forward_Return_20D", vol_col: str = "Forward_Volatility_20D"):
    """
    Boxplots of both targets split by Asset_Class -- same lens as EDA,
    now on the targets. Inspect whether target distributions differ materially across asset classes; 
    large differences may reflect genuine cross-asset risk/return characteristics.
    """
    fig, axes = plt.subplots(1, 2, figsize=(14, 4.5))
    order = sorted(df["Asset_Class"].dropna().unique())
    sns.boxplot(data=df, x="Asset_Class", y=return_col, order=order, showfliers=False, ax=axes[0])
    axes[0].tick_params(axis="x", rotation=30)
    sns.boxplot(data=df, x="Asset_Class", y=vol_col, order=order, showfliers=False, ax=axes[1])
    axes[1].tick_params(axis="x", rotation=30)
    plt.tight_layout()
    plt.show()
    