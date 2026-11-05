"""Current holdings from Consolidated_Portfolio_Positions.csv (PROCESSED_DIR).

Used for the RealEstate weight in notebook 09 and the portfolio value in notebook 10,
so both follow the real portfolio automatically when the file is updated.
"""
import pandas as pd

REAL_ESTATE_LABEL = "real estate"      # value in the ISIN or Security Name column
NUMERIC_COLS = ["Current Quantity", "Average Buy Price (EUR)", "Current Price (EUR)", "Unrealized PnL (EUR)"]


def load_positions(path):
    """Positions with Value_EUR = Current Quantity x Current Price (EUR) per row."""
    pos = pd.read_csv(path)
    missing = [c for c in ["ISIN", "Security Name"] + NUMERIC_COLS if c not in pos.columns]
    assert not missing, f"missing columns in {path.name}: {missing}"
    for col in NUMERIC_COLS:
        values = pd.to_numeric(pos[col], errors="coerce")
        bad = pos.loc[values.isna(), "Security Name"].tolist()
        assert not bad, f"non-numeric '{col}' for: {bad}"
        pos[col] = values
    pos["Value_EUR"] = pos["Current Quantity"] * pos["Current Price (EUR)"]
    assert (pos["Value_EUR"] >= 0).all(), "negative position value"
    return pos


def is_real_estate(pos):
    """Rows whose ISIN or Security Name is 'real estate' (case and spaces ignored)."""
    label = lambda col: pos[col].astype(str).str.strip().str.lower() == REAL_ESTATE_LABEL
    return label("ISIN") | label("Security Name")


def portfolio_value(pos):
    """Total value of all positions in EUR."""
    return float(pos["Value_EUR"].sum())


def real_estate_weight(pos, decimals=3):
    """RealEstate value / total value, rounded."""
    mask = is_real_estate(pos)
    assert mask.any(), f"no row with ISIN or Security Name '{REAL_ESTATE_LABEL}'"
    return round(float(pos.loc[mask, "Value_EUR"].sum()) / portfolio_value(pos), decimals)


def pnl_gap(pos):
    """Largest difference (EUR) between the file's Unrealized PnL and
    Current Quantity x (Current Price - Average Buy Price). Near 0 means the
    quantity and price columns are read as intended."""
    implied = pos["Current Quantity"] * (pos["Current Price (EUR)"] - pos["Average Buy Price (EUR)"])
    return float((implied - pos["Unrealized PnL (EUR)"]).abs().max())
