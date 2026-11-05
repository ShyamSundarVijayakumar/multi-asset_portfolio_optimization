"""Exact (non-evolutionary) solutions on the same PortfolioSpec, used to check NSGA-II
in notebook 09 and as the Markowitz solver for the benchmarks in notebook 10.

min variance and points on the frontier: convex quadratic problems, solved with SLSQP
max return: linear problem, solved exactly by filling the highest-return assets to the cap
max Sharpe: SLSQP on the Sharpe ratio, started from the best frontier point
"""
import numpy as np
import pandas as pd
from scipy.optimize import minimize

from src.optimization.portfolio import full_weights, portfolio_stats


def _slsqp(spec, objective, x0, target_return=None):
    free, fixed = spec.free_idx, spec.fixed_idx
    w_fixed = spec.fixed.to_numpy()
    mu_free = spec.mu[free]
    ret_fixed = float(spec.mu[fixed] @ w_fixed)
    mu_scale = np.abs(mu_free).mean()

    cons = [{"type": "eq", "fun": lambda x: x.sum() - spec.budget, "jac": lambda x: np.ones_like(x)}]
    if target_return is not None:
        cons.append({"type": "ineq",
                     "fun": lambda x: (mu_free @ x + ret_fixed - target_return) / mu_scale,
                     "jac": lambda x: mu_free / mu_scale})
    res = minimize(objective, x0, jac=True, method="SLSQP", bounds=[(0.0, spec.max_weight)] * len(free),
                   constraints=cons, options={"ftol": 1e-14, "maxiter": 2000})
    assert res.success, f"SLSQP failed: {res.message}"
    return np.clip(res.x, 0.0, spec.max_weight)


def _variance_objective(spec):
    free = spec.free_idx
    scale = float(np.diag(spec.sigma).mean())          # keeps the objective near 1 for SLSQP's tolerance

    def f(x):
        w = full_weights(spec, x)[0]
        g = spec.sigma @ w
        return float(w @ g) / scale, 2 * g[free] / scale
    return f


def _equal_start(spec):
    return np.full(len(spec.free_idx), spec.budget / len(spec.free_idx))


def min_variance(spec, target_return=None, x0=None):
    """Lowest-risk portfolio, optionally with expected return >= target_return. Full weights."""
    x0 = _equal_start(spec) if x0 is None else x0
    x = _slsqp(spec, _variance_objective(spec), x0, target_return)
    return full_weights(spec, x)[0]


def max_return(spec):
    """Highest-return portfolio: fill the free assets in order of expected return up to the cap."""
    x = np.zeros(len(spec.free_idx))
    left = spec.budget
    for i in np.argsort(-spec.mu[spec.free_idx]):
        x[i] = min(spec.max_weight, left)
        left -= x[i]
        if left <= 1e-15:
            break
    return full_weights(spec, x)[0]


def efficient_frontier(spec, n_points=50):
    """n_points exact frontier portfolios from min variance to max return.
    Returns full weights (one column per ticker) and portfolio_stats."""
    w_min, w_max = min_variance(spec), max_return(spec)
    r_min, r_max = spec.mu @ w_min, spec.mu @ w_max
    rows, x = [w_min], w_min[spec.free_idx]
    for target in np.linspace(r_min, r_max, n_points)[1:-1]:
        w = min_variance(spec, target, x0=x)
        rows.append(w)
        x = w[spec.free_idx]
    rows.append(w_max)
    W = np.vstack(rows)
    return pd.concat([pd.DataFrame(W, columns=spec.tickers), portfolio_stats(spec, W)], axis=1)


def min_volatility_at(spec, returns):
    """Exact lowest 20-day risk achievable at each expected return in `returns`."""
    w_min, w_max = min_variance(spec), max_return(spec)
    r_min, r_max = spec.mu @ w_min, spec.mu @ w_max
    vol = lambda w: float(np.sqrt(w @ spec.sigma @ w))
    out = []
    for r in returns:
        if r <= r_min:
            out.append(vol(w_min))
        elif r >= r_max - 1e-12:
            out.append(vol(w_max))
        else:
            out.append(vol(min_variance(spec, r)))
    return np.array(out)


def max_sharpe(spec, frontier=None):
    """Highest-Sharpe portfolio. Starts SLSQP from the best point of the exact frontier."""
    frontier = efficient_frontier(spec) if frontier is None else frontier
    x0 = frontier.loc[frontier["Sharpe_20D"].idxmax(), spec.tickers].to_numpy(dtype=float)[spec.free_idx]
    free = spec.free_idx

    def neg_sharpe(x):
        w = full_weights(spec, x)[0]
        g = spec.sigma @ w
        vol = np.sqrt(w @ g)
        excess = w @ spec.mu - spec.rf
        grad = (spec.mu / vol - excess * g / vol ** 3)[free]
        return -excess / vol, -grad

    return full_weights(spec, _slsqp(spec, neg_sharpe, x0))[0]
