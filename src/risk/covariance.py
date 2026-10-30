"""Covariance construction for the optimizer (notebook 08).

Sigma_20D = D @ R @ D
    D : diagonal of 20-day predicted volatilities from optimizer_inputs.parquet (notebook 07)
    R : correlation matrix
        - market assets: Ledoit-Wolf shrunk correlation of trailing daily log returns
        - synthetic classes (Bond, RealEstate): zero correlation with every other asset,
          because notebook 03 generates them with independent random shocks
"""
import numpy as np
import pandas as pd
from sklearn.covariance import LedoitWolf


def align_returns(features, tickers, return_col="Log_Return_1D"):
    """Wide Date x Ticker log returns on the dates where every ticker has a return.

    A ticker's returns on dates where some other ticker is missing (e.g. crypto weekends,
    exchange holidays) are summed into the next common date. Log returns add over time,
    so no price move is lost and every row covers the same calendar interval for all tickers.
    Rows after the last common date are dropped.
    """
    wide = (features[features["Ticker"].isin(tickers)]
            .pivot(index="Date", columns="Ticker", values=return_col)
            .sort_index()[list(tickers)])
    common = wide.notna().all(axis=1)
    # stays datetime64 throughout (np.where with NaT turns nanosecond dates into plain integers)
    next_common = wide.index.to_series().where(common).bfill()
    keep = next_common.notna()
    aligned = wide[keep].groupby(next_common[keep]).sum(min_count=1)
    aligned.index.name = "Date"
    n_folded = int((~common & keep).sum())
    return aligned, n_folded


def shrunk_correlation(returns):
    """Ledoit-Wolf shrinkage applied to standardized returns, rescaled to a unit diagonal.

    Returns (sample correlation, shrunk correlation, shrinkage intensity), both as DataFrames.
    """
    z = (returns - returns.mean()) / returns.std(ddof=1)
    lw = LedoitWolf(assume_centered=True).fit(z.to_numpy())
    c = lw.covariance_
    d = np.sqrt(np.diag(c))
    shrunk = pd.DataFrame(c / np.outer(d, d), index=returns.columns, columns=returns.columns)
    sample = returns.corr()
    return sample, shrunk, float(lw.shrinkage_)


def embed_synthetic(corr_market, tickers):
    """Full correlation matrix in `tickers` order; tickers outside corr_market get zero
    correlation with everything else and 1 on the diagonal."""
    full = pd.DataFrame(np.eye(len(tickers)), index=pd.Index(tickers, name="Ticker"), columns=list(tickers))
    m = list(corr_market.index)
    full.loc[m, m] = corr_market.to_numpy()
    return full


def build_covariance(corr, vol):
    """Sigma = D R D, with vol a Series indexed by ticker in the same order as corr."""
    v = vol.loc[corr.index].to_numpy()
    sigma = corr.to_numpy() * np.outer(v, v)
    return pd.DataFrame(sigma, index=corr.index, columns=corr.columns)


def synthetic_correlation_check(features, synthetic_tickers, market_tickers, window, return_col="Log_Return_1D"):
    """Sample correlation of each synthetic asset with every market asset over the same window,
    against the +/- 2/sqrt(T) band expected from pure noise when the true correlation is zero."""
    aligned, _ = align_returns(features, list(market_tickers) + list(synthetic_tickers), return_col)
    tail = aligned.iloc[-window:]
    rows = []
    for t in synthetic_tickers:
        rho = tail[list(market_tickers)].corrwith(tail[t])
        band = 2 / np.sqrt(len(tail))
        rows.append({"Ticker": t, "Max_Abs_Corr": rho.abs().max(), "Mean_Corr": rho.mean(),
                     "Share_Outside_Band": (rho.abs() > band).mean(), "Noise_Band": band})
    return pd.DataFrame(rows).set_index("Ticker")


def validate_covariance(sigma, corr, vol, synthetic_tickers):
    """Properties the optimizer relies on; returns a one-row summary."""
    s = sigma.to_numpy()
    r = corr.to_numpy()
    eig_r = np.linalg.eigvalsh(r)
    off = r.copy()
    np.fill_diagonal(off, 0.0)
    syn = [corr.index.get_loc(t) for t in synthetic_tickers]
    return pd.Series({
        "Max_Asymmetry": np.abs(s - s.T).max(),
        "Min_Eigenvalue_Sigma": np.linalg.eigvalsh(s).min(),
        "Min_Eigenvalue_Corr": eig_r.min(),
        "Corr_Condition_Number": eig_r.max() / eig_r.min(),
        "Max_Diag_Error": np.abs(np.diag(s) - vol.loc[sigma.index].to_numpy() ** 2).max(),
        "Max_Abs_Corr_Synthetic": np.abs(off[syn, :]).max() if syn else 0.0,
    })


def class_portfolio_vol_from_sigma(sigma, classes, horizon=20):
    """Annualized volatility of each class's equal-weighted portfolio implied by Sigma."""
    out = {}
    for cls, members in classes.groupby(classes).groups.items():
        w = pd.Series(1 / len(members), index=members)
        var_h = w @ sigma.loc[members, members] @ w
        out[cls] = np.sqrt(var_h * 252 / horizon)
    return pd.Series(out, name="Class_Portfolio_Vol_Sigma")