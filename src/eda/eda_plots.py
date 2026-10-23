"""
Contains visualization functions for exploratory data analysis (EDA), 
feature distributions, time-series plotting, and correlation heatmaps.
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import random
import math

# Set professional theme for all plots
sns.set_theme(style="whitegrid", palette="muted")

def plot_feature_distributions(df: pd.DataFrame, features: list):
    """
    Plots a histogram (distribution) and a boxplot (outliers) for each feature.
    """
    # Filter only features that exist in the dataframe
    valid_features = [f for f in features if f in df.columns]
    
    n_features = len(valid_features)
    if n_features == 0:
        print("No valid features provided for distribution plotting.")
        return

    fig, axes = plt.subplots(n_features, 2, figsize=(12, 4 * n_features))
    
    # Handle single feature case (axes is 1D)
    if n_features == 1:
        axes = [axes]
        
    for i, feature in enumerate(valid_features):
        # Drop NaNs for plotting to avoid warnings
        data = df[feature].dropna()
        
        # Histogram
        sns.histplot(data, bins=50, ax=axes[i][0], color="steelblue") #kde=True, 
        axes[i][0].set_title(f"Distribution: {feature}")
        axes[i][0].set_xlabel(feature)
        axes[i][0].set_ylabel("Frequency")
        
        # Boxplot
        sns.boxplot(x=data, ax=axes[i][1], color="lightblue")
        axes[i][1].set_title(f"Outliers: {feature}")
        axes[i][1].set_xlabel(feature)
        
    plt.tight_layout()
    plt.show()

def plot_time_series_sample(df: pd.DataFrame, features: list, n_assets_per_class: int = 1, seed: int = None):
    """
    Randomly selects assets from each class and plots their time-series 
    for the specified features on shared axes.
    """
    rng = random.Random(seed)

    # Randomly select tickers per asset class
    selected_tickers = []
    for asset_class in df['Asset_Class'].unique():
        class_tickers = df[df['Asset_Class'] == asset_class]['Ticker'].unique().tolist()
        sampled = rng.sample(class_tickers, min(n_assets_per_class, len(class_tickers)))
        selected_tickers.extend(sampled)
        
    print(f"Selected assets for Time-Series visualization: {selected_tickers}")
    
    # Filter dataset for selected tickers
    plot_df = df[df['Ticker'].isin(selected_tickers)].copy()
    valid_features = [f for f in features if f in df.columns]
    
    fig, axes = plt.subplots(len(valid_features), 1, figsize=(14, 4 * len(valid_features)), sharex=True)
    if len(valid_features) == 1:
        axes = [axes]
        
    for i, feature in enumerate(valid_features):
        ax = axes[i]
        sns.lineplot(data=plot_df, x="Date", y=feature, hue="Ticker", ax=ax, linewidth=1.5)
        
        ax.set_title(f"Time-Series: {feature}")
        ax.set_ylabel(feature)
        
        # Move legend outside the plot
        ax.legend(bbox_to_anchor=(1.01, 1), loc='upper left')
    plt.xlabel("Date")
    plt.tight_layout()
    plt.show()

def plot_correlation_heatmap(df: pd.DataFrame, method: str = 'spearman', figsize: tuple = (14, 12)):
    """
    Plots a triangular heatmap of feature correlations.
    Uses Spearman by default to capture non-linear monotonic relationships.
    """
    # Select only numeric columns
    numeric_df = df.select_dtypes(include=[np.number])
    
    # Drop zero variance columns to avoid NaNs in correlation
    variances = numeric_df.var()
    valid_numeric = numeric_df.drop(columns=variances[variances == 0].index.tolist())
    
    # Calculate correlation matrix
    corr = valid_numeric.corr(method=method)
    
    # Generate a mask for the upper triangle
    mask = np.triu(np.ones_like(corr, dtype=bool))
    
    plt.figure(figsize=figsize)
    plt.title(f"Feature Correlation Matrix ({method.capitalize()})", fontsize=16)
    
    # Draw the heatmap
    sns.heatmap(
        corr, 
        mask=mask, 
        cmap="coolwarm", 
        vmax=1.0, 
        vmin=-1.0, 
        center=0,
        square=True, 
        linewidths=.5, 
        cbar_kws={"shrink": .75}
    )
    plt.show()

def plot_distributions_by_class(df: pd.DataFrame, features: list, figsize_per_row: tuple = (16, 4)):
    """
    Plots a boxplot per feature, split by Asset Class, to compare distributions
    across the pooled universe. Outlier points are hidden (whiskers stay) since
    fat-tailed financial features otherwise compress into an unreadable sliver.
    """
    n = len(features)
    ncols = 2
    nrows = math.ceil(n / ncols)
    fig, axes = plt.subplots(nrows, ncols, figsize=(figsize_per_row[0], figsize_per_row[1] * nrows))
    axes = axes.flatten() if n > 1 else [axes]
    order = sorted(df["Asset_Class"].dropna().unique())

    for ax, feature in zip(axes, features):
        sns.boxplot(data=df, x="Asset_Class", y=feature, order=order, showfliers=False, ax=ax)
        ax.set_title(feature)
        ax.set_xlabel("")
        ax.tick_params(axis="x", rotation=30)

    for ax in axes[len(features):]:
        ax.axis("off")

    plt.tight_layout()
    plt.show()