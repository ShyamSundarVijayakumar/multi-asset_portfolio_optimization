import pandas as pd

def cut_to_as_of(features: pd.DataFrame, as_of, date_col: str = "Date") -> pd.DataFrame:
    """Drop rows after the as-of date so every stage works on the same cross-section."""
    as_of = pd.Timestamp(as_of)
    if pd.Timestamp(features[date_col].max()) < as_of:
        raise ValueError(f"features end before the as-of date {as_of:%Y-%m-%d}")
    return features[features[date_col] <= as_of].copy()