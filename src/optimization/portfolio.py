"""Portfolio definition shared by the optimizers (notebook 09) and the backtest (notebook 10).

Everything is on the 20-day horizon of optimizer_inputs.parquet:
    expected return  w' mu
    risk             sqrt(w' Sigma w)
    Sharpe           (w' mu - rf) / risk
Constraints: long-only, per-asset cap on the free assets, fixed weights for the
fixed classes (RealEstate), weights summing to 1.
"""
from dataclasses import dataclass

import numpy as np
import pandas as pd


@dataclass
class PortfolioSpec:
    tickers: list           # optimizer order (rows of optimizer_inputs.parquet)
    classes: pd.Series      # Asset_Class indexed by ticker
    mu: np.ndarray          # 20-day expected returns
    sigma: np.ndarray       # 20-day covariance matrix
    rf: float               # 20-day risk-free rate
    max_weight: float       # cap per free asset
    fixed: pd.Series        # fixed weight per ticker (e.g. the RealEstate ticker)

    @property
    def free_idx(self):
        return np.array([i for i, t in enumerate(self.tickers) if t not in self.fixed.index])

    @property
    def fixed_idx(self):
        return np.array([self.tickers.index(t) for t in self.fixed.index], dtype=int)

    @property
    def budget(self):
        """Weight left for the free assets."""
        return 1.0 - float(self.fixed.sum())


def build_spec(opt, sigma_df, max_weight, fixed_class_weights):
    """PortfolioSpec from optimizer_inputs.parquet and covariance_20d.parquet.

    fixed_class_weights: e.g. {"RealEstate": 0.10}; a class's weight is split
    equally across its tickers.
    """
    tickers = opt["Ticker"].tolist()
    assert opt["Ticker"].is_unique and opt["Date"].nunique() == 1, "one row per ticker on one date"
    assert list(sigma_df.index) == tickers and list(sigma_df.columns) == tickers, \
        "Sigma must be in the ticker order of optimizer_inputs"
    classes = opt.set_index("Ticker")["Asset_Class"]

    fixed = {}
    for cls, w in fixed_class_weights.items():
        members = classes.index[classes == cls].tolist()
        assert members, f"no ticker in fixed class {cls}"
        fixed.update({t: w / len(members) for t in members})
    fixed = pd.Series(fixed, dtype=float)

    sigma = sigma_df.to_numpy(dtype=float)
    spec = PortfolioSpec(tickers=tickers, classes=classes,
                         mu=opt["Expected_Return_20D"].to_numpy(dtype=float), sigma=sigma,
                         rf=float(opt["RF_20D"].iloc[0]), max_weight=max_weight, fixed=fixed)

    assert 0 < spec.budget <= 1, "fixed weights must leave a positive budget"
    assert len(spec.free_idx) * max_weight >= spec.budget - 1e-12, \
        f"infeasible: {len(spec.free_idx)} free assets x cap {max_weight} < budget {spec.budget}"
    assert np.allclose(sigma, sigma.T) and np.linalg.eigvalsh(sigma).min() > 0, \
        "Sigma must be symmetric positive definite"
    return spec


def full_weights(spec, X):
    """Free-asset weights (n x n_free) -> full weights (n x n_assets) with the fixed weights inserted."""
    X = np.atleast_2d(X)
    W = np.zeros((X.shape[0], len(spec.tickers)))
    W[:, spec.free_idx] = X
    W[:, spec.fixed_idx] = spec.fixed.to_numpy()
    return W


def project_capped_simplex(V, budget, cap, n_iter=100):
    """Euclidean projection of each row of V onto {0 <= x <= cap, sum(x) = budget}.

    The projection is clip(v - tau, 0, cap) for the tau that makes the row sum to
    budget; tau is found by bisection (the row sum falls as tau rises).
    Feasible rows are returned unchanged.
    """
    V = np.atleast_2d(np.asarray(V, dtype=float))
    lo = V.min(axis=1, keepdims=True) - cap     # every entry clips to cap: sum = n*cap >= budget
    hi = V.max(axis=1, keepdims=True)           # every entry clips to 0:   sum = 0 <= budget
    for _ in range(n_iter):
        tau = (lo + hi) / 2
        too_big = np.clip(V - tau, 0.0, cap).sum(axis=1, keepdims=True) > budget
        lo = np.where(too_big, tau, lo)
        hi = np.where(too_big, hi, tau)
    return np.clip(V - (lo + hi) / 2, 0.0, cap)


def portfolio_stats(spec, W, horizon=20):
    """Return, risk and Sharpe of each row of W (full weights), 20-day and annualized."""
    W = np.atleast_2d(W)
    ret = W @ spec.mu
    vol = np.sqrt(np.einsum("ij,jk,ik->i", W, spec.sigma, W))
    k = 252 / horizon
    return pd.DataFrame({
        "Return_20D": ret,
        "Volatility_20D": vol,
        "Sharpe_20D": (ret - spec.rf) / vol,
        "Return_Annual": ret * k,               # same convention as notebook 07 (20D x 252/20)
        "Volatility_Annual": vol * np.sqrt(k),
        "Sharpe_Annual": (ret - spec.rf) / vol * np.sqrt(k),
    })


def class_weights(spec, W):
    """Weights summed per asset class, one row per portfolio."""
    W = pd.DataFrame(np.atleast_2d(W), columns=spec.tickers)
    return W.T.groupby(spec.classes.loc[spec.tickers].to_numpy()).sum().T


def equal_weight(spec):
    """Fixed weights as given, the remaining budget split equally over the free assets."""
    x = np.full(len(spec.free_idx), spec.budget / len(spec.free_idx))
    assert x.max() <= spec.max_weight + 1e-12, "equal weight breaks the cap"
    return full_weights(spec, x)[0]


def check_feasible(spec, W, tol=1e-8):
    """Assert that every row of W meets all constraints."""
    W = np.atleast_2d(W)
    free, fixed = W[:, spec.free_idx], W[:, spec.fixed_idx]
    assert (W >= -tol).all(), "negative weight (long-only)"
    assert (free <= spec.max_weight + tol).all(), "weight above the per-asset cap"
    assert np.allclose(fixed, spec.fixed.to_numpy(), atol=tol), "fixed weight changed"
    assert np.allclose(W.sum(axis=1), 1.0, atol=tol), "weights do not sum to 1"
