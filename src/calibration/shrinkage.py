"""
Return Shrinkage & Expected Returns (Stage 07)

Decides how much of the model's return predictions to trust, using a
calibration slope learned from out-of-sample predictions rather than an
IC / reference-IC rule, and builds the expected returns the optimizer uses.

For each (Date, Asset_Class) group, predictions and realized returns are
centered on the group mean. The shrinkage weight k minimizes

    sum_i  w_i * (r_centered_i - k * p_centered_i) ** 2

so k* = sum(w p r) / sum(w p^2): the slope of realized on predicted
deviations. It equals rho * sigma_r / sigma_p, i.e. a correlation rescaled
into return units, which a rank IC divided by a fixed constant cannot do.

Groups with one member (Bond, RealEstate) carry no within-class information
and are dropped from the fit; groups with two members contribute.

k only scales deviations from the class mean. It says nothing about class
LEVELS, which come from class_risk_prior (an explicit assumption).

Per-asset expected returns (build_expected_returns) = risk-based prior + tilt:
the class prior is spread across a class's assets by each asset's beta to its
equal-weighted class portfolio (class_betas), and the model's tilt is added
only to the extent the out-of-sample gate supports it.
"""

import numpy as np
import pandas as pd
from scipy.stats import spearmanr


def _center(frame: pd.DataFrame, col: str, group_cols) -> pd.Series:
    """Deviation of `col` from its mean within each group (e.g. Date x Asset_Class)."""
    return frame[col] - frame.groupby(list(group_cols))[col].transform("mean")


def _bartlett_variance(e: np.ndarray, n_lags: int) -> float:
    """Newey-West long-run variance of sum(e), Bartlett kernel."""
    total = float(np.dot(e, e))
    for lag in range(1, min(n_lags, len(e) - 1) + 1):
        weight = 1.0 - lag / (n_lags + 1.0)
        total += 2.0 * weight * float(np.dot(e[lag:], e[:-lag]))
    return max(total, 0.0)


def calibrate_shrinkage_k(
    frame: pd.DataFrame,
    pred_col: str = "Raw_Prediction",
    actual_col: str = "Realized",
    date_col: str = "Date",
    group_cols=("Date", "Asset_Class"),
    weight_col: str = None,
    vol_floor: float = 1e-3,
    winsor_q: float = 0.01,
    horizon: int = 20,
    k_max: float = 1.0,
    n_lags: int = None,
) -> dict:
    """
    Estimates the shrinkage weight k and its standard error.

    weight_col : optional trailing-volatility column (known at time t, so no
        look-ahead). Rows are weighted 1/vol^2, which down-weights the
        high-variance assets that would otherwise dominate the fit.
    vol_floor  : lower bound on that volatility, in the same units as
        weight_col (daily for Volatility_20D).
    winsor_q   : clips centered realized returns at the q / 1-q quantiles so a
        single extreme 20-day move cannot decide k.
    horizon    : label length. Forward-return labels overlap by horizon-1
        days, so the standard error uses a Newey-West correction with this
        many lags on the per-date series; a row-level standard error ignores
        that overlap and is far too small. `n_lags` overrides the lag count.

    Returns k_raw (unclipped), k (clipped to [0, k_max]), se, t_stat,
    k_conservative = max(0, k_raw - se), n_dates.
    """
    f = frame[[date_col, pred_col, actual_col] + [c for c in group_cols if c != date_col]].copy()
    if weight_col is not None:
        f[weight_col] = frame[weight_col].values
    f = f.dropna()
    # Single-member groups center to exactly zero and carry no information.
    # Drop them explicitly so they cannot shift the winsorization quantiles.
    f = f[f.groupby(list(group_cols))[pred_col].transform("size") >= 2]
    if f.empty:
        return {"k_raw": 0.0, "k": 0.0, "se": np.nan, "t_stat": np.nan, "k_conservative": 0.0, "n_dates": 0}

    p = _center(f, pred_col, group_cols)
    r = _center(f, actual_col, group_cols)
    if winsor_q:
        lo, hi = r.quantile([winsor_q, 1.0 - winsor_q])
        r = r.clip(lo, hi)

    w = np.ones(len(f))
    if weight_col is not None:
        w = 1.0 / np.maximum(f[weight_col].to_numpy(dtype=float), vol_floor) ** 2

    num_row = w * p.to_numpy() * r.to_numpy()
    den_row = w * p.to_numpy() ** 2
    per_date = pd.DataFrame({"d": f[date_col].to_numpy(), "num": num_row, "den": den_row}).groupby("d").sum().sort_index()

    den = per_date["den"].sum()
    if den <= 0:
        return {"k_raw": 0.0, "k": 0.0, "se": np.nan, "t_stat": np.nan, "k_conservative": 0.0, "n_dates": 0}

    k_raw = per_date["num"].sum() / den
    e = (per_date["num"] - k_raw * per_date["den"]).to_numpy()
    se = np.sqrt(_bartlett_variance(e, n_lags if n_lags is not None else horizon)) / den
    k = float(np.clip(k_raw, 0.0, k_max))
    k_cons = float(np.clip(k_raw - se, 0.0, k_max))
    return {
        "k_raw": float(k_raw),
        "k": k,
        "se": float(se),
        "t_stat": float(k_raw / se) if se > 0 else np.nan,
        "k_conservative": k_cons,
        "n_dates": int((per_date["den"] > 0).sum()),
    }


def expanding_calibrated_tilt(
    df: pd.DataFrame,
    oos_preds: pd.DataFrame,
    target: str = "Forward_Return_20D",
    weight_col: str = "Volatility_20D",
    group_cols=("Date", "Asset_Class"),
    horizon: int = 20,
    conservative: bool = True,
    date_col: str = "Date",
    **calibrate_kwargs,
):
    """
    Backtest version. For each fold, k is learned only from earlier folds'
    out-of-sample predictions, after dropping the last `horizon + 1` dates
    before the fold starts: their labels resolve on or after the fold's
    first date, so using them would carry test-period returns into k.
    Same cutoff convention as walk_forward_splits. Fold 1 has no history and
    gets k = 0.

    `df` is the model-ready frame (supplies realized targets and volatility);
    `oos_preds` is the walk-forward output, indexed like `df`.

    Returns (tilt, k_table). `tilt` is k * (prediction - group mean), in
    log-return units, aligned to the target's rows of `oos_preds`; k_table
    has one row per fold with k_used and the calibrate_shrinkage_k outputs.
    """
    frame = oos_preds[oos_preds["Target"] == target].copy()
    frame["Realized"] = df.loc[frame.index, target]
    if weight_col is not None:
        frame[weight_col] = df.loc[frame.index, weight_col]

    all_dates = pd.Index(df[date_col].drop_duplicates()).sort_values()
    centered_pred = _center(frame, "Raw_Prediction", group_cols)
    tilt = pd.Series(0.0, index=frame.index, name="Calibrated_Tilt")
    rows = []

    for fold in sorted(frame["Fold"].unique()):
        fold_start = frame.loc[frame["Fold"] == fold, date_col].min()
        start_pos = int(all_dates.searchsorted(fold_start))
        cutoff = all_dates[max(start_pos - horizon - 1, 0)]
        calib = frame[(frame["Fold"] < fold) & (frame[date_col] <= cutoff)]

        if len(calib) == 0:
            res = {"k_raw": 0.0, "k": 0.0, "se": np.nan, "t_stat": np.nan, "k_conservative": 0.0, "n_dates": 0}
        else:
            res = calibrate_shrinkage_k(
                calib, weight_col=weight_col, group_cols=group_cols, horizon=horizon, **calibrate_kwargs
            )
        k_used = res["k_conservative"] if conservative else res["k"]
        mask = frame["Fold"] == fold
        tilt.loc[mask] = k_used * centered_pred.loc[mask]
        rows.append({"Fold": fold, "k_used": k_used, **res, "Calib_Rows": len(calib)})

    return tilt, pd.DataFrame(rows)


def class_risk_prior(
    latest: pd.DataFrame,
    vol_col: str,
    rf_annual: float,
    sr0: float,
    class_col: str = "Asset_Class",
    horizon: int = 20,
    trading_days: int = 252,
    overrides: dict = None,
    class_vol: pd.Series = None,
    ticker_col: str = "Ticker",
) -> pd.Series:
    """
    Class-level expected-return prior on the label horizon (simple return):

        prior_class = rf + sr0 * sigma_class

    An assumption, not a forecast: every class portfolio earns the same Sharpe
    sr0. sigma_class is the class's equal-weighted PORTFOLIO volatility when
    `class_vol` (from class_portfolio_vol) is given -- the consistent choice.
    Without it, the median single-asset predicted vol is used, which overstates
    diversified classes because it includes diversifiable risk. A
    near-zero-volatility class lands on
    rf. `overrides` pins a class to an annual return (e.g. {"Bond": rf_annual}).
    """
    if class_vol is not None:
        sigma = latest[ticker_col].map(class_vol)
        if sigma.isna().any():
            raise ValueError(f"no class vol for: {sorted(latest.loc[sigma.isna(), ticker_col])}")
    else:
        sigma = latest.groupby(class_col)[vol_col].transform("median")
    annual = rf_annual + sr0 * sigma
    if overrides:
        for cls, value in overrides.items():
            annual = annual.where(latest[class_col] != cls, value)
    return annual * horizon / trading_days


def _wide_returns(history, ret_col="Log_Return_1D", date_col="Date", ticker_col="Ticker"):
    """Date x Ticker matrix of trailing daily returns. Fails on duplicate (Date, Ticker) rows."""
    return history.pivot(index=date_col, columns=ticker_col, values=ret_col).sort_index()


def _equal_weight_betas(returns: pd.DataFrame, vol: pd.Series = None, min_obs: int = 126) -> np.ndarray:
    """
    Beta of each column to the equal-weighted portfolio of all columns.

    Built from a covariance Sigma = D R D: R is the sample correlation of the
    returns, D the volatilities (`vol` if given, e.g. the volatility model's
    forecasts, otherwise the sample volatilities). beta = Sigma w / (w' Sigma w)
    with w = 1/n, so the equal-weighted mean of the betas is exactly 1.
    """
    R = returns.corr(min_periods=min_obs)
    if R.isna().to_numpy().any():
        raise ValueError(f"fewer than {min_obs} overlapping returns for some pair in {list(returns.columns)}")
    sd = (vol[returns.columns] if vol is not None else returns.std()).to_numpy(dtype=float)
    S = R.to_numpy() * np.outer(sd, sd)
    w = np.full(len(sd), 1.0 / len(sd))
    return S @ w / (w @ S @ w)


def _shrink_betas(b: np.ndarray, shrink: float, clip: tuple) -> np.ndarray:
    """
    Blume-style shrink toward 1, then bound to `clip` with the mean kept at
    exactly 1: assets that hit a bound are fixed there and only the others are
    rescaled, repeated until no free asset crosses a bound.
    """
    lo, hi = clip
    b = (1.0 - shrink) * np.asarray(b, dtype=float) + shrink
    fixed = np.zeros(len(b), dtype=bool)
    for _ in range(len(b) + 1):
        free = ~fixed
        b[free] *= (len(b) - b[fixed].sum()) / b[free].sum()
        over, under = free & (b > hi), free & (b < lo)
        if not (over.any() or under.any()):
            break
        b[over], b[under] = hi, lo
        fixed |= over | under
    return b


def class_betas(
    history: pd.DataFrame,
    classes: pd.Series,
    vol: pd.Series = None,
    window: int = 252,
    shrink: float = 1 / 3,
    clip: tuple = (0.25, 3.0),
    min_obs: int = 126,
    ret_col: str = "Log_Return_1D",
    date_col: str = "Date",
    ticker_col: str = "Ticker",
) -> pd.Series:
    """
    Beta of every asset to the equal-weighted portfolio of its own class, from
    the last `window` days of `history` (trailing daily returns, known today).

    classes : Series ticker -> asset class for the live universe.
    vol     : optional Series ticker -> predicted annualised volatility. Passing
              the volatility model's forecasts makes the betas consistent with
              the covariance matrix the optimizer will use.
    Classes with one asset get beta 1. Within every class the betas average
    exactly 1, so the class's average expected return is left unchanged.
    """
    wide = _wide_returns(history, ret_col, date_col, ticker_col).iloc[-window:]
    missing = sorted(set(classes.index) - set(wide.columns))
    if missing:
        raise ValueError(f"no return history for: {missing}")
    betas = pd.Series(1.0, index=classes.index, name="Beta_Class")
    for _, members in classes.groupby(classes).groups.items():
        members = list(members)
        if len(members) < 2:
            continue
        b = _equal_weight_betas(wide[members], vol, min_obs)
        betas[members] = _shrink_betas(b, shrink, clip)
    return betas


def class_portfolio_vol(
    history: pd.DataFrame,
    classes: pd.Series,
    vol: pd.Series,
    window: int = 252,
    min_obs: int = 126,
    ret_col: str = "Log_Return_1D",
    date_col: str = "Date",
    ticker_col: str = "Ticker",
) -> pd.Series:
    """
    Annualised volatility of each class's equal-weighted portfolio, from the
    same Sigma = D R D as class_betas (D = predicted vols, R = trailing
    correlation). Returned per ticker. A one-asset class gets that asset's vol.

    This is the risk an investor in the class actually bears; the median
    single-asset vol also contains idiosyncratic risk that diversifies away,
    which would hand large classes (e.g 30+ stocks) a Sharpe above sr0.
    """
    wide = _wide_returns(history, ret_col, date_col, ticker_col).iloc[-window:]
    out = pd.Series(np.nan, index=classes.index, name="Class_Portfolio_Vol")
    for _, members in classes.groupby(classes).groups.items():
        members = list(members)
        if len(members) == 1:
            out[members] = float(vol[members[0]])
            continue
        R = wide[members].corr(min_periods=min_obs)
        if R.isna().to_numpy().any():
            raise ValueError(f"fewer than {min_obs} overlapping returns for some pair in {members}")
        sd = vol[members].to_numpy(dtype=float)
        w = np.full(len(members), 1.0 / len(members))
        out[members] = float(np.sqrt(w @ (R.to_numpy() * np.outer(sd, sd)) @ w))
    return out


def beta_stability(
    history: pd.DataFrame,
    classes: pd.Series,
    window: int = 252,
    step: int = 63,
    shrink: float = 1 / 3,
    clip: tuple = (0.25, 3.0),
    min_obs: int = 126,
    ret_col: str = "Log_Return_1D",
    date_col: str = "Date",
    ticker_col: str = "Ticker",
) -> pd.DataFrame:
    """
    Out-of-sample check that the beta spread is reliable. At anchor dates every
    `step` days, betas estimated on the `window` days BEFORE the anchor are
    compared with betas realised over the `window` days AFTER it.

    Per class: mean rank correlation (classes with >= 3 assets), and the mean
    squared error of three forecasts of the future beta: all ones (no spread),
    the raw estimate, and the shrunk estimate used live.
    MSE_shrunk / MSE_flat below 1 means the spread carries real information.
    Betas here use sample volatilities in both windows; the live betas use the
    predicted volatilities. Anchors overlap in their future windows, so read
    this as descriptive.
    """
    wide = _wide_returns(history, ret_col, date_col, ticker_col)
    rows = []
    for end in range(window, len(wide) - window + 1, step):
        past, future = wide.iloc[end - window:end], wide.iloc[end:end + window]
        for cls, members in classes.groupby(classes).groups.items():
            members = list(members)
            if len(members) < 2:
                continue
            try:
                b_past = _equal_weight_betas(past[members], None, min_obs)
                b_future = _equal_weight_betas(future[members], None, min_obs)
            except ValueError:
                continue
            b_shrunk = _shrink_betas(b_past, shrink, clip)
            rows.append({
                "Anchor": wide.index[end], "Asset_Class": cls, "n": len(members),
                "Rank_Corr": spearmanr(b_past, b_future)[0] if len(members) >= 3 else np.nan,
                "MSE_flat": float(np.mean((1.0 - b_future) ** 2)),
                "MSE_raw": float(np.mean((b_past - b_future) ** 2)),
                "MSE_shrunk": float(np.mean((b_shrunk - b_future) ** 2)),
            })
    per_anchor = pd.DataFrame(rows)
    if per_anchor.empty:
        return per_anchor
    out = per_anchor.groupby("Asset_Class").agg(
        n_assets=("n", "first"), n_anchors=("Anchor", "size"), Mean_Rank_Corr=("Rank_Corr", "mean"),
        MSE_flat=("MSE_flat", "mean"), MSE_raw=("MSE_raw", "mean"), MSE_shrunk=("MSE_shrunk", "mean"),
    )
    out["Shrunk_vs_flat"] = out["MSE_shrunk"] / out["MSE_flat"]
    out["Share_anchors_shrunk_beats_flat"] = (
        per_anchor.assign(win=per_anchor["MSE_shrunk"] < per_anchor["MSE_flat"]).groupby("Asset_Class")["win"].mean()
    )
    return out


def build_expected_returns(
    latest: pd.DataFrame,
    pred_col: str,
    vol_col: str,
    k_tilt: float,
    rf_annual: float,
    sr0: float,
    class_col: str = "Asset_Class",
    horizon: int = 20,
    overrides: dict = None,
    beta: pd.Series = None,
    class_vol: pd.Series = None,
    ticker_col: str = "Ticker",
    trading_days: int = 252,
) -> pd.DataFrame:
    """
    Expected return for the optimizer, on the label horizon (simple return):

        mu_i = rf + beta_i * (class_prior - rf)  +  expm1(k_tilt * (raw_i - class mean of raw))
               '------ risk-based prior ------'     '------- calibrated model tilt -------'

    class_prior : rf + sr0 * class portfolio vol (class_risk_prior with
                  class_vol from class_portfolio_vol). With beta averaging 1,
                  every class's equal-weighted portfolio has Sharpe sr0 exactly.
    beta        : Series ticker -> beta to the asset's class (class_betas). It
                  spreads the class prior across assets by how much class risk
                  each one carries; betas average 1 within a class, so the
                  class average is unchanged. None -> beta 1 for every asset.
    k_tilt      : calibrated shrinkage weight. 0 means the model's return
                  differences failed the out-of-sample gate and add nothing.

    The tilt is a log-return deviation, so expm1 puts it in simple-return
    units before it is added to the simple-return prior.
    """
    out = latest.copy()
    rf_h = rf_annual * horizon / trading_days
    out["Class_Prior_20D"] = class_risk_prior(latest, vol_col, rf_annual, sr0, class_col, horizon,
                                              trading_days, overrides=overrides,
                                              class_vol=class_vol, ticker_col=ticker_col)
    if beta is None:
        out["Beta_Class"] = 1.0
    else:
        out["Beta_Class"] = out[ticker_col].map(beta)
        if out["Beta_Class"].isna().any():
            raise ValueError(f"no beta for: {sorted(out.loc[out['Beta_Class'].isna(), ticker_col])}")
    out["Prior_20D"] = rf_h + out["Beta_Class"] * (out["Class_Prior_20D"] - rf_h)
    deviation = out[pred_col] - out.groupby(class_col)[pred_col].transform("mean")
    out["Tilt_20D"] = np.expm1(k_tilt * deviation)
    out["Expected_Return_20D"] = out["Prior_20D"] + out["Tilt_20D"]
    return out