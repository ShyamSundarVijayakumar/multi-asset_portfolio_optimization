import json
import pandas as pd
from pathlib import Path
from datetime import datetime

# Dynamically locate the project root directory from src/utils.py
ROOT_DIR = Path(__file__).resolve().parent.parent
P_PATH = ROOT_DIR / "data/processed/Consolidated_Portfolio_Positions.csv"
M_PATH = ROOT_DIR / "config/yfinance_ticker_mapper.csv"
CACHE_LOG_PATH = ROOT_DIR / "data/processed/.engine_cache.json"

def read_clean_dataframe(file_path: Path) -> pd.DataFrame:
    """
    Reads a target data CSV file and aggressively strips out lingering 
    simulation records or duplicate assets across critical identification axes.
    """
    if not file_path.exists():
        return pd.DataFrame()
    
    df = pd.read_csv(file_path)
    
    # Remove rows explicitly marked as simulations
    if 'Is_Simulation' in df.columns:
        df = df[~df['Is_Simulation'].astype(str).str.lower().isin(['true', '1', '1.0'])]
        
    # Completely drop duplicates on unique identifier columns to prevent engine crashes
    if 'ISIN' in df.columns:
        df = df.drop_duplicates(subset=['ISIN'], keep='first')
    elif 'Security Name' in df.columns:
        df = df.drop_duplicates(subset=['Security Name'], keep='first')
    else:
        df = df.drop_duplicates()
        
    return df

def purge_simulations() -> bool:
    """
    Purges temporary simulation rows and formats the core persistence 
    CSV files back to their clean base state.
    """
    cleaned = False
    if P_PATH.exists():
        df_p = read_clean_dataframe(P_PATH)
        df_p.to_csv(P_PATH, index=False)
        cleaned = True
    if M_PATH.exists():
        df_m = read_clean_dataframe(M_PATH)
        df_m.to_csv(M_PATH, index=False)
        cleaned = True
    return cleaned

def sanitize_portfolio_data():
    """
    Finds and removes any temporary simulation elements left over from crashes.
    Acts as a wrapper pointing to our robust de-duplication cleaning layout.
    """
    purge_simulations()

def check_is_cached(engine_key: str) -> bool:
    """
    Verifies whether a specific background analytical engine run has 
    already successfully processed metrics for the current calendar day.
    """
    if not CACHE_LOG_PATH.exists():
        return False
    try:
        with open(CACHE_LOG_PATH, 'r', encoding='utf-8') as f:
            cache_data = json.load(f)
        today_str = datetime.today().strftime('%Y-%m-%d')
        return cache_data.get(engine_key) == today_str
    except Exception:
        return False

def update_cache_timestamp(engine_key: str):
    """
    Logs a successful modern calendar date execution stamp for the target calculation layer.
    """
    cache_data = {}
    if CACHE_LOG_PATH.exists():
        try:
            with open(CACHE_LOG_PATH, 'r', encoding='utf-8') as f:
                cache_data = json.load(f)
        except Exception:
            pass
            
    cache_data[engine_key] = datetime.today().strftime('%Y-%m-%d')
    CACHE_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(CACHE_LOG_PATH, 'w', encoding='utf-8') as f:
        json.dump(cache_data, f, indent=4)