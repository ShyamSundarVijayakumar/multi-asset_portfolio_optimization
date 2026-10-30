"""Plots for covariance construction (notebook 08)."""
import numpy as np
import matplotlib.pyplot as plt


def plot_correlation_heatmap(corr, classes, save_dir=None):
    """Final correlation matrix, tickers grouped by asset class, class boundaries drawn."""
    order = sorted(corr.index, key=lambda t: (classes[t], t))
    c = corr.loc[order, order]
    cls = classes.loc[order]

    fig, ax = plt.subplots(figsize=(11, 9.5))
    im = ax.imshow(c.to_numpy(), cmap="RdBu_r", vmin=-1, vmax=1)
    edges = np.flatnonzero(cls.to_numpy()[1:] != cls.to_numpy()[:-1]) + 0.5
    for e in edges:
        ax.axhline(e, color="#334155", lw=1)
        ax.axvline(e, color="#334155", lw=1)
    ax.set_xticks(range(len(order)))
    ax.set_xticklabels(order, rotation=90, fontsize=7)
    ax.set_yticks(range(len(order)))
    ax.set_yticklabels(order, fontsize=7)
    starts = np.r_[0, np.ceil(edges).astype(int)]
    ends = np.r_[np.ceil(edges).astype(int), len(order)]
    for s, e in zip(starts, ends):
        ax.text(len(order) + 0.8, (s + e - 1) / 2, cls.iloc[s], va="center", fontsize=9, color="#334155")
    fig.colorbar(im, ax=ax, fraction=0.035, pad=0.12, label="Correlation")
    ax.set_title("Shrunk correlation matrix, grouped by asset class", loc="left", fontsize=12)
    fig.tight_layout()
    if save_dir is not None:
        fig.savefig(save_dir / "correlation_heatmap.png", dpi=150, bbox_inches="tight")
    plt.show()


def plot_eigenvalue_spectrum(sample_corr, shrunk_corr, save_dir=None):
    """Sorted eigenvalues of the sample and shrunk market correlation matrices (log scale).
    The smallest eigenvalues are where an optimizer finds spurious low-risk portfolios."""
    ev_s = np.sort(np.linalg.eigvalsh(sample_corr.to_numpy()))[::-1]
    ev_h = np.sort(np.linalg.eigvalsh(shrunk_corr.to_numpy()))[::-1]
    x = np.arange(1, len(ev_s) + 1)

    fig, ax = plt.subplots(figsize=(9, 4.5))
    ax.plot(x, ev_s, marker="o", ms=4, lw=2, color="#94A3B8", label="Sample")
    ax.plot(x, ev_h, marker="o", ms=4, lw=2, color="#1D4ED8", label="Ledoit-Wolf shrunk")
    ax.set_yscale("log")
    ax.set_xlabel("Eigenvalue rank")
    ax.set_ylabel("Eigenvalue (log scale)")
    ax.grid(True, which="major", axis="y", color="#E2E8F0", lw=0.8)
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    ax.legend(frameon=False)
    ax.set_title("Market correlation eigenvalues: sample vs shrunk", loc="left", fontsize=12)
    fig.tight_layout()
    if save_dir is not None:
        fig.savefig(save_dir / "eigenvalue_spectrum.png", dpi=150, bbox_inches="tight")
    plt.show()
