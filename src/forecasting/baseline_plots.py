"""
Baseline Model Training (Stage 07) -- IC visualization.
"""

import matplotlib.pyplot as plt
import seaborn as sns
import pandas as pd
from pathlib import Path


def plot_ic_by_fold(results_df: pd.DataFrame, save_dir=None, filename: str = "ic_by_fold.png"):
    """
    Bar chart of Information Coefficient per fold, one panel per target.
    Look for: consistency of sign across folds (a model flipping between
    positive and negative IC fold to fold has no real signal, regardless of
    the average) more than the magnitude of any single fold.
    """
    targets = results_df["Target"].unique()
    fig, axes = plt.subplots(1, len(targets), figsize=(7 * len(targets), 4.5))
    if len(targets) == 1:
        axes = [axes]
    for ax, target in zip(axes, targets):
        subset = results_df[results_df["Target"] == target]
        sns.barplot(data=subset, x="Fold", y="IC", ax=ax)
        ax.axhline(0, color="grey", linestyle="--", linewidth=1)
        ax.set_title(target)
    plt.tight_layout()
    saved = None
    if save_dir is not None:
        saved = Path(save_dir) / filename
        fig.savefig(saved, dpi=150, bbox_inches="tight")
    plt.show()
    return saved