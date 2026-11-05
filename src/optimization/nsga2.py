"""NSGA-II for the pipeline portfolio (notebook 09).

Genetic-algorithm view:
    individual   weights of the free assets (the fixed RealEstate weight is not a gene)
    feasibility  every individual is projected onto {0 <= w <= cap, sum = budget}
                 (repair), so no penalties are needed
    objectives   minimize -w'mu (i.e. maximize expected 20-day return)
                 minimize sqrt(w' Sigma w) (20-day risk)
"""
import numpy as np
import pandas as pd
from pymoo.algorithms.moo.nsga2 import NSGA2
from pymoo.core.problem import Problem
from pymoo.core.repair import Repair
from pymoo.core.sampling import Sampling
from pymoo.operators.crossover.sbx import SBX
from pymoo.operators.mutation.pm import PM
from pymoo.optimize import minimize

from src.optimization.portfolio import full_weights, portfolio_stats, project_capped_simplex


class PortfolioProblem(Problem):
    def __init__(self, spec):
        super().__init__(n_var=len(spec.free_idx), n_obj=2, xl=0.0, xu=spec.max_weight)
        self.spec = spec

    def _evaluate(self, X, out, *args, **kwargs):
        W = full_weights(self.spec, X)
        ret = W @ self.spec.mu
        vol = np.sqrt(np.einsum("ij,jk,ik->i", W, self.spec.sigma, W))
        out["F"] = np.column_stack([-ret, vol])


class CappedSimplexRepair(Repair):
    def _do(self, problem, X, **kwargs):
        return project_capped_simplex(X, problem.spec.budget, problem.spec.max_weight)


class DirichletSampling(Sampling):
    """Random portfolios from Dirichlet draws whose concentration varies per individual,
    so the first generation holds both spread-out and concentrated portfolios."""

    def _do(self, problem, n_samples, *args, random_state=None, **kwargs):
        rng = random_state if random_state is not None else np.random.default_rng()
        alpha = 10 ** rng.uniform(-1.0, 1.0, size=(n_samples, 1))
        G = rng.gamma(np.broadcast_to(alpha, (n_samples, problem.n_var)))
        X = G / G.sum(axis=1, keepdims=True) * problem.spec.budget
        return project_capped_simplex(X, problem.spec.budget, problem.spec.max_weight)


def run_nsga2(spec, pop_size=200, n_gen=1000, seed=42):
    """Runs NSGA-II and returns its Pareto front: one row per portfolio with the
    full weights (one column per ticker) and portfolio_stats, sorted by risk."""
    algorithm = NSGA2(pop_size=pop_size, sampling=DirichletSampling(),
                      crossover=SBX(prob=0.9, eta=15), mutation=PM(eta=20),
                      repair=CappedSimplexRepair(), eliminate_duplicates=True)
    res = minimize(PortfolioProblem(spec), algorithm, ("n_gen", n_gen), seed=seed, verbose=False)
    W = full_weights(spec, res.X)
    front = pd.concat([pd.DataFrame(W, columns=spec.tickers), portfolio_stats(spec, W)], axis=1)
    return front.sort_values("Volatility_20D").reset_index(drop=True)


def select_max_sharpe(front):
    """Index of the front portfolio with the highest predicted Sharpe ratio."""
    return int(front["Sharpe_20D"].idxmax())
