"""Plots for NSGA-II portfolio optimization (notebook 09)."""
import matplotlib.pyplot as plt
import numpy as np

# categorical palette in fixed order; classes get slots alphabetically so a class keeps its colour
PALETTE = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#6250d6", "#e34948"]
SLATE, SLATE_LIGHT, GRID = "#334155", "#94A3B8", "#E2E8F0"


def class_colors(classes):
    return {c: PALETTE[i % len(PALETTE)] for i, c in enumerate(sorted(set(classes)))}


def _style(ax):
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    ax.grid(True, axis="y", color=GRID, lw=0.8)
    ax.set_axisbelow(True)


def plot_pareto_front(front, exact, compare, save_dir=None):
    """NSGA-II front against the exact efficient frontier, annualized, with the
    selected max-Sharpe portfolios and equal weight marked."""
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.plot(exact["Volatility_Annual"] * 100, exact["Return_Annual"] * 100, color=SLATE, lw=2,
            label="Exact efficient frontier", zorder=2)
    ax.scatter(front["Volatility_Annual"] * 100, front["Return_Annual"] * 100, s=14, color=PALETTE[0],
               alpha=0.8, label="NSGA-II front", zorder=3)
    # marker, colour, size, label offset (the two max-Sharpe points sit close together)
    markers = {"NSGA-II max Sharpe": ("*", PALETTE[1], 260, (12, 6)),
               "Exact max Sharpe": ("D", SLATE, 70, (12, -18)),
               "Equal weight": ("s", SLATE_LIGHT, 70, (8, -14))}
    for name, (m, c, s, offset) in markers.items():
        x, y = compare.loc[name, "Volatility_Annual"] * 100, compare.loc[name, "Return_Annual"] * 100
        ax.scatter(x, y, marker=m, s=s, color=c, edgecolor="white", linewidth=1.2, zorder=4, label=name)
        ax.annotate(name, (x, y), xytext=offset, textcoords="offset points", fontsize=9, color=SLATE)
    ax.set_xlabel("Predicted volatility (% a year)")
    ax.set_ylabel("Expected return (% a year)")
    _style(ax)
    ax.legend(frameon=False, loc="lower right")
    ax.set_title("Pareto front: NSGA-II vs exact efficient frontier", loc="left", fontsize=12)
    fig.tight_layout()
    if save_dir is not None:
        fig.savefig(save_dir / "pareto_front.png", dpi=150, bbox_inches="tight")
    plt.show()


def plot_class_allocation_along_front(front_classes, front_vol_annual, selected_vol_annual, save_dir=None):
    """How the class mix changes from the lowest-risk to the highest-return end of the front."""
    order = np.argsort(front_vol_annual.to_numpy())
    x = front_vol_annual.to_numpy()[order] * 100
    cols = sorted(front_classes.columns)
    colors = class_colors(cols)
    fig, ax = plt.subplots(figsize=(10, 5.5))
    ax.stackplot(x, *(front_classes[c].to_numpy()[order] * 100 for c in cols),
                 labels=cols, colors=[colors[c] for c in cols], edgecolor="white", linewidth=0.6)
    ax.axvline(selected_vol_annual * 100, color=SLATE, lw=1.5, ls="--")
    ax.annotate("max Sharpe", (selected_vol_annual * 100, 101), ha="center", fontsize=9, color=SLATE,
                annotation_clip=False)
    ax.set_xlim(x.min(), x.max())
    ax.set_ylim(0, 100)
    ax.set_xlabel("Predicted volatility of the front portfolio (% a year)")
    ax.set_ylabel("Class weight (%)")
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    ax.legend(frameon=False, loc="upper left", bbox_to_anchor=(1.01, 1))
    ax.set_title("Class allocation along the Pareto front", loc="left", fontsize=12, pad=18)
    fig.tight_layout()
    if save_dir is not None:
        fig.savefig(save_dir / "class_allocation_along_front.png", dpi=150, bbox_inches="tight")
    plt.show()


def plot_selected_allocation(class_table, save_dir=None):
    """Class weights of the compared portfolios (columns of class_table), side by side."""
    classes = list(class_table.index)
    names = list(class_table.columns)
    y = np.arange(len(classes))
    h = 0.8 / len(names)
    fig, ax = plt.subplots(figsize=(10, 5.5))
    for j, name in enumerate(names):
        vals = class_table[name].to_numpy() * 100
        bars = ax.barh(y + (j - (len(names) - 1) / 2) * h, vals, height=h * 0.9, color=PALETTE[j], label=name)
        if j == 0:
            for b, v in zip(bars, vals):
                ax.text(b.get_width() + 0.4, b.get_y() + b.get_height() / 2, f"{v:.1f}%", va="center",
                        fontsize=8, color=SLATE)
    ax.set_yticks(y)
    ax.set_yticklabels(classes)
    ax.invert_yaxis()
    ax.set_xlabel("Weight (%)")
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    ax.grid(True, axis="x", color=GRID, lw=0.8)
    ax.set_axisbelow(True)
    ax.legend(frameon=False, loc="upper right")
    ax.set_title("Class allocation of the selected portfolio", loc="left", fontsize=12)
    fig.tight_layout()
    if save_dir is not None:
        fig.savefig(save_dir / "selected_class_allocation.png", dpi=150, bbox_inches="tight")
    plt.show()
