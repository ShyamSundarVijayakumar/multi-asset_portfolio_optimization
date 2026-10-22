"""
Baseline Model Training (Stage 07)

Trains two independent LightGBM regressors per walk-forward fold -- one for
Forward_Return_20D, one for Forward_Volatility_20D -- and scores each fold
with Information Coefficient (Spearman rank correlation between predicted
and realized values), not R2/RMSE. Return prediction is expected to have a
low IC (0.02-0.05 is a real, usable signal in this domain); volatility
should score meaningfully higher, since volatility clustering is one of the
most robust patterns in finance. A near-zero return IC is not a failure --
it's the expected result, and the reason downstream shrinkage exists.
"""

import lightgbm as lgb
import pandas as pd
import numpy as np
from scipy.stats import spearmanr
from itertools import combinations

from src.forecasting.walk_forward import walk_forward_splits

DEFAULT_LGB_PARAMS = {"objective": "regression", "verbosity": -1, "random_state": 42}

def compute_pairwise_concordance(df: pd.DataFrame, preds, test_idx, target_col: str) -> pd.DataFrame:
    """
    Kendall's-tau-style pairwise concordance, by asset class. Works down to
    n=2 per date -- a single pair still contributes one valid concordant/
    discordant judgment, unlike Spearman on 2 points, which is mechanically
    forced to exactly +-1 and carries no information on its own. Bond and
    RealEstate (n=1) can't use this either -- no pair can be formed with one
    instrument; time-series correlation remains the only option for those.
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
    IC computed per date AND within asset class only, then aggregated.
    Pooling across dates or across classes within a date both let between-
    class level differences (Crypto vs Bond) masquerade as ranking skill.
    """
    # Grab only the necessary columns for the test indices
    test_df = df.loc[test_idx, ["Date", "Asset_Class", target_col]].copy()
    test_df["Prediction"] = preds
    ics = []
    
    # Group by Date AND Asset_Class to isolate true ranking skill
    for (_, _), group in test_df.groupby(["Date", "Asset_Class"]):
        # Need at least 3 items to calculate a meaningful rank correlation
        if group["Prediction"].nunique() > 1 and len(group) >= 3:
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
    Trains one LightGBM model on train_idx, scores cross-sectional IC on test_idx.
    Categorical dtype is set on the full feature frame before slicing into
    train/test, so both subsets share the same category encoding.
    """
    X = df[feature_cols].copy()
    for col in categorical_cols:
        X[col] = X[col].astype("category")

    X_train, y_train = X.loc[train_idx], df.loc[train_idx, target_col]
    X_test, y_test = X.loc[test_idx], df.loc[test_idx, target_col]

    model = lgb.LGBMRegressor(**(lgb_params or DEFAULT_LGB_PARAMS))
    model.fit(X_train, y_train, categorical_feature=categorical_cols)

    preds = model.predict(X_test)
    
    # Replaced pooled spearmanr with the rigorous cross-sectional function
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
    lgb_params: dict = None,
) -> pd.DataFrame:
    """
    Runs the full walk-forward baseline: for every fold and every target,
    trains a fresh model and records the test-fold IC. Two independent
    models per fold (Model 1: return, Model 2: volatility), matching the
    project's predict-then-optimize architecture -- one model is never
    asked to predict both targets at once.
    """
    fold_results = []
    oos_preds = []
    for fold, train_idx, test_idx in walk_forward_splits(df, n_folds=n_folds, embargo_days=embargo_days):
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
                "Daily_Groups_Scored": ic_metrics["N_Date_Class_Groups"],
                "Train_Rows": len(train_idx),
                "Test_Rows": len(test_idx),
            }
            for cls in pairwise_df.index:
                current_result[f"Hit_Rate_{cls}"] = pairwise_df.loc[cls, "Hit_Rate"]

            fold_results.append(current_result)
            
            # Save the raw predictions and their associated Fold IC
            fold_df = df.loc[test_idx, ["Date", "Asset_Class"]].copy()
            fold_df["Fold"] = fold
            fold_df["Target"] = target_col
            fold_df["Actual"] = df.loc[test_idx, target_col] # CRITICAL FIX: Storing true targets for 'k' calculation
            fold_df["Raw_Prediction"] = preds
            fold_df["Fold_IC"] = ic_metrics["Mean_IC"] 
            oos_preds.append(fold_df)
            
    return pd.DataFrame(fold_results), pd.concat(oos_preds)


def estimate_shrinkage_k_from_prior_oos(
    oos_df: pd.DataFrame, 
    target_col: str, 
    pred_col: str, 
    date_col: str = "Date"
) -> float:
    """
    Analytically derives the optimal shrinkage coefficient 'k' that minimizes MSE.
    Equation: k* = sum(d_i * r_i) / sum(d_i^2)
    """
    # 1. Calculate cross-sectional means per date
    cs_mean_target = oos_df.groupby(date_col)[target_col].transform("mean")
    cs_mean_pred = oos_df.groupby(date_col)[pred_col].transform("mean")
    
    # 2. Center the variables (Target minus Mean, Pred minus Mean)
    r = oos_df[target_col] - cs_mean_target
    d = oos_df[pred_col] - cs_mean_pred
    
    # 3. Calculate optimal k (OLS slope without intercept)
    numerator = np.sum(d * r)
    denominator = np.sum(d ** 2)
    
    # Handle edge cases (zero variance in predictions)
    if denominator == 0 or np.isnan(numerator) or np.isnan(denominator):
        return 0.0
        
    k_star = numerator / denominator
    
    # 4. Constrain k between 0 and 1
    k_optimal = float(np.clip(k_star, 0.0, 1.0))
    
    return k_optimal


def apply_shrinkage_expanding(oos_preds: pd.DataFrame) -> pd.DataFrame:
    """
    Mimics a walk-forward loop on the OOS predictions to calculate 'k' sequentially.
    Only shrinks Return predictions. Leaves Volatility untouched.
    """
    df_shrunk = oos_preds.copy()
    
    # Initialize the new columns
    df_shrunk["Shrunk_Prediction"] = np.nan
    df_shrunk["k_applied"] = np.nan
    
    # 1. Isolate Return predictions (we only shrink returns)
    return_mask = df_shrunk["Target"] == "Forward_Return_20D"
    df_returns = df_shrunk[return_mask].copy()
    
    folds = sorted(df_returns["Fold"].unique())
    historical_oos = pd.DataFrame()
    
    for fold in folds:
        current_fold_mask = df_returns["Fold"] == fold
        df_current_fold = df_returns[current_fold_mask].copy()
        
        # Calculate 'k' based ONLY on history
        if len(historical_oos) > 0:
            current_k = estimate_shrinkage_k_from_prior_oos(
                oos_df=historical_oos,
                target_col="Actual",        # Now points to the true actuals we saved
                pred_col="Raw_Prediction",  # Points to the correct raw prediction column
                date_col="Date"
            )
        else:
            # Fold 0 has no history to learn from. 
            current_k = 0.0 
            
        # Apply 'k' to the cross-sectional mean
        cs_mean = df_current_fold.groupby("Date")["Raw_Prediction"].transform("mean")
        df_current_fold["Shrunk_Prediction"] = (
            cs_mean + current_k * (df_current_fold["Raw_Prediction"] - cs_mean)
        )
        
        # Store results for this fold
        df_returns.loc[current_fold_mask, "Shrunk_Prediction"] = df_current_fold["Shrunk_Prediction"]
        df_returns.loc[current_fold_mask, "k_applied"] = current_k
        
        # Append the current fold to history so the NEXT fold can learn from it
        historical_oos = pd.concat([historical_oos, df_current_fold], ignore_index=True)
        
    # Re-integrate Return predictions back into the main DataFrame
    df_shrunk.loc[return_mask, "Shrunk_Prediction"] = df_returns["Shrunk_Prediction"]
    df_shrunk.loc[return_mask, "k_applied"] = df_returns["k_applied"]
    
    # 2. Volatility predictions remain RAW (unshrunk)
    vol_mask = df_shrunk["Target"] == "Forward_Volatility_20D"
    if vol_mask.any():
        df_shrunk.loc[vol_mask, "Shrunk_Prediction"] = df_shrunk.loc[vol_mask, "Raw_Prediction"]
        df_shrunk.loc[vol_mask, "k_applied"] = 1.0  # 1.0 means 100% of raw signal is used
    
    return df_shrunk