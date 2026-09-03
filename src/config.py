"""
Central configuration file for Hybrid ML + Multi-Objective Portfolio Optimization
"""

from pathlib import Path
import os
from dotenv import load_dotenv
# Load environment variables
load_dotenv()

# =============================================================================
# Project Directories
# =============================================================================

PROJECT_ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = PROJECT_ROOT / "data"
ASSET_CLASS_DIR = DATA_DIR / "processed" / "Asset class"
FEATURE_DIR = DATA_DIR / "features"
FINAL_DIR = DATA_DIR / "final"
OUTPUT_DIR = PROJECT_ROOT / "outputs"
MODEL_DIR = PROJECT_ROOT / "models"

# Create directories if they don't exist
FEATURE_DIR.mkdir(parents=True, exist_ok=True)
FINAL_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
MODEL_DIR.mkdir(parents=True, exist_ok=True)

# =============================================================================
# Processed Files
# =============================================================================

FILES = {
    "stocks": ASSET_CLASS_DIR / "Equity_Stock_Consolidated.csv",
    "etfs": ASSET_CLASS_DIR / "ETF_Consolidated.csv",
    "crypto": ASSET_CLASS_DIR / "Cryptocurrency_Consolidated.csv",
    "commodities": ASSET_CLASS_DIR / "Commodity_Consolidated.csv",
    "bond": ASSET_CLASS_DIR / "bonds.csv",
    "real_estate": ASSET_CLASS_DIR / "real_estate.csv",
}

# =============================================================================
# Yahoo Finance
# =============================================================================

YFINANCE_SLEEP = 1.0

# These will NOT be downloaded from YFinance.
SYNTHETIC_ETFS = os.getenv("Etf1_ticker")

# =============================================================================
# Output Files
# =============================================================================

YF_RAW_FILE = FEATURE_DIR / "yfinance_raw_features.parquet"
MASTER_FILE = FEATURE_DIR / "master_dataset.parquet"
