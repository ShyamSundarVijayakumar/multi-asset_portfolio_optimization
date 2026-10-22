"""
Walk-Forward Validation (Stage 07)

Generates expanding-window walk-forward train/test splits for 20-day forward
labels. Each fold trains on every date before its test block and stops early
enough that no training label overlaps the test period (purge), with an
optional wider gap (embargo). Model-agnostic: used identically by the
LightGBM baseline here and any model added later.
"""

import numpy as np
import pandas as pd


def walk_forward_splits(
    df: pd.DataFrame,
    date_col: str = "Date",
    horizon: int = 20,
    n_folds: int = 4,
    embargo_days: int = 20,
):
    """
    Yields (fold_number, train_index, test_index) for `n_folds` consecutive
    test blocks.

    Layout: the unique dates are cut into n_folds + 1 equal blocks. The first
    block is training-only; fold f tests on block f + 1, shifted forward by
    `embargo_days`, so the test blocks are adjacent and the last one is
    clipped at the final date.

    Gap: training ends max(horizon + 1, embargo_days) trading dates before the
    test block starts. horizon + 1 is the purge: the last training label (a
    `horizon`-day forward return) resolves the day before the first test
    date, so no label overlaps the test period. A larger `embargo_days`
    widens the gap further.

    Returns index labels, so `df` must have a unique index.
    """
    if not df.index.is_unique:
        raise ValueError("df must have a unique index: the splits are returned as index labels")

    unique_dates = np.sort(df[date_col].unique())
    fold_size = len(unique_dates) // (n_folds + 1)
    gap = max(horizon + 1, embargo_days)

    for fold in range(1, n_folds + 1):
        test_start_idx = fold_size * fold + embargo_days
        if test_start_idx >= len(unique_dates):
            break
        test_end_idx = min(test_start_idx + fold_size - 1, len(unique_dates) - 1)
        train_end_idx = max(test_start_idx - gap, 0)

        train_end = unique_dates[train_end_idx]
        test_start = unique_dates[test_start_idx]
        test_end = unique_dates[test_end_idx]

        train_idx = df.index[df[date_col] <= train_end]
        test_idx = df.index[(df[date_col] >= test_start) & (df[date_col] <= test_end)]
        yield fold, train_idx, test_idx