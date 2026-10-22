"""
Baseline Model Training (Stage 07)

Trains two independent LightGBM regressors per walk-forward fold -- one for
Forward_Return_20D, one for Forward_Volatility_20D -- with Asset_Class as a
native categorical feature, and scores each fold out of sample:

  - Cross-sectional IC: Spearman rank correlation between predicted and
    realized values within each Date x Asset_Class group of 3+ assets.
  - Pairwise hit rate: share of concordant pairs per class, which also
    covers two-asset classes.

Scores, not R2/RMSE: the optimizer needs assets ranked correctly within a
class. In this run the return IC averages about zero (no stable skill) while
the volatility IC is consistently high. IC does not say how large a
prediction should be; that is measured separately by the calibration slope
in src/calibration/shrinkage.py.

The out-of-sample predictions are returned too, indexed like `df`, because
every later check is computed on them.
"""

from itertools import combinations

import lightgbm as lgb
import numpy as np
import pandas as pd
from scipy.stats import spearmanr

from src.forecasting.walk_forward import walk_forward_splits

# Untuned LightGBM defaults with a fixed seed: a baseline, not a tuned model.
DEFAULT_LGB_PARAMS = {"objective": "regression", "verbosity": -1, "random_state": 42}


def compute_pairwise_concordance(df: pd.DataFrame, preds, test_idx, target_col: str) -> pd.DataFrame:
    """
    Kendall's-tau-style pairwise concordance, by asset class: for every pair
    of assets on the same date, does the prediction order match the realized
    order? Pairs tied on either side are skipped. Works down to n = 2 per
    date -- a single pair still gives one valid judgment, unlike Spearman on
    2 points, which is forced to exactly +-1.

    One-asset classes (Bond, RealEstate) form no pairs and return NaN with
    N_Pairs = 0. They are synthetic series with known parameters and are not
    scored here.

    Returns one row per class: Hit_Rate (0.5 = coin flip) and N_Pairs.
    """
    test_df = df.loc[test_idx, ["Date", "Asset_Class", target_col]].copy()
    test_df["Prediction"] = preds

    results = {}
    for cls, cls_df in test_df.groupby("Asset_Class"):
        concordant, discordant = 0, 0
        for _, group in cls_df.groupby("Date"):
            if len(group) < 2:
                continue
            values = group[["Prediction", target_col]].values
            for (p1, a1), (p2, a2) in combinations(values, 2):
                if p1 == p2 or a1 == a2:
                    continue
                concordant += (p1 > p2) == (a1 > a2)
                discordant += (p1 > p2) != (a1 > a2)
        total = concordant + discordant
        results[cls] = {"Hit_Rate": concordant / total if total else np.nan, "N_Pairs": total}
    return pd.DataFrame(results).T


def compute_cross_sectional_ic(df: pd.DataFrame, preds, test_idx, target_col: str) -> dict:
    """
    Spearman IC computed within each Date x Asset_Class group of at least 3
    assets, then averaged. Pooling across dates or across classes would let
    between-class level differences (Crypto vs Bond) masquerade as ranking
    skill.

    Every group counts equally, for instance a 4-asset class weighs as much as 
    a 37-asset Stocks group on each date. ICIR = mean / std of the group ICs;
    it is descriptive only, because overlapping 20-day labels make
    consecutive ICs correlated.
    """
    test_df = df.loc[test_idx, ["Date", "Asset_Class", target_col]].copy()
    test_df["Prediction"] = preds

    ics = []
    for _, group in test_df.groupby(["Date", "Asset_Class"]):
        if len(group) >= 3 and group["Prediction"].nunique() > 1 and group[target_col].nunique() > 1:
            ic, _ = spearmanr(group["Prediction"], group[target_col])
            if not np.isnan(ic):
                ics.append(ic)

    ics = np.array(ics)
    return {
        "Mean_IC": ics.mean() if len(ics) > 0 else np.nan,
        "Std_IC": ics.std() if len(ics) > 0 else np.nan,
        "ICIR": (ics.mean() / ics.std()) if (len(ics) > 0 and ics.std() > 0) else np.nan,
        "N_Date_Class_Groups": len(ics),
    }


def train_and_evaluate_fold(
    df: pd.DataFrame,
    feature_cols: list,
    categorical_cols: list,
    target_col: str,
    train_idx: pd.Index,
    test_idx: pd.Index,
    lgb_params: dict = None,
):
    """
    Trains one LightGBM model on train_idx and scores it on test_idx.
    Categorical dtype is set on the full feature frame before slicing into
    train/test, so both subsets share the same category encoding.

    Returns (ic_metrics, pairwise_df, model, preds); preds are aligned to test_idx.
    """
    X = df[feature_cols].copy()
    for col in categorical_cols:
        X[col] = X[col].astype("category")

    model = lgb.LGBMRegressor(**(lgb_params or DEFAULT_LGB_PARAMS))
    model.fit(X.loc[train_idx], df.loc[train_idx, target_col], categorical_feature=categorical_cols)
    preds = model.predict(X.loc[test_idx])

    ic_metrics = compute_cross_sectional_ic(df, preds, test_idx, target_col)
    pairwise_df = compute_pairwise_concordance(df, preds, test_idx, target_col)
    return ic_metrics, pairwise_df, model, preds


def run_walk_forward_baseline(
    df: pd.DataFrame,
    feature_cols: list,
    categorical_cols: list,
    target_cols: list,
    n_folds: int = 4,
    embargo_days: int = 20,
    horizon: int = 20,
    lgb_params: dict = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Runs the full walk-forward baseline: for every fold and every target,
    trains a fresh model and scores it out of sample. Two independent models
    per fold (return and volatility), matching the predict-then-optimize
    architecture -- one model is never asked to predict both targets.

    Returns
      results   : one row per fold and target -- IC, ICIR, Std_IC, number of
                  groups scored, row counts, and Hit_Rate_<class>.
      oos_preds : out-of-sample predictions with Date, Asset_Class, Fold,
                  Target and Raw_Prediction, indexed like `df`.
    """
    fold_results = []
    oos_preds = []
    for fold, train_idx, test_idx in walk_forward_splits(
        df, horizon=horizon, n_folds=n_folds, embargo_days=embargo_days
    ):
        for target_col in target_cols:
            ic_metrics, pairwise_df, _, preds = train_and_evaluate_fold(
                df, feature_cols, categorical_cols, target_col, train_idx, test_idx, lgb_params
            )

            current_result = {
                "Fold": fold,
                "Target": target_col,
                "IC": ic_metrics["Mean_IC"],
                "ICIR": ic_metrics["ICIR"],
                "Std_IC": ic_metrics["Std_IC"],
                "Date_Class_Groups": ic_metrics["N_Date_Class_Groups"],
                "Train_Rows": len(train_idx),
                "Test_Rows": len(test_idx),
            }
            for cls in pairwise_df.index:
                current_result[f"Hit_Rate_{cls}"] = pairwise_df.loc[cls, "Hit_Rate"]
            fold_results.append(current_result)

            # Out-of-sample predictions, indexed like df for every later check
            fold_df = df.loc[test_idx, ["Date", "Asset_Class"]].copy()
            fold_df["Fold"] = fold
            fold_df["Target"] = target_col
            fold_df["Raw_Prediction"] = preds
            oos_preds.append(fold_df)
    return pd.DataFrame(fold_results), pd.concat(oos_preds)