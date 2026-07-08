import sys
import os
from pathlib import Path
from datetime import date

from dotenv import load_dotenv
# Load environment variables
load_dotenv()

os.environ['KMP_DUPLICATE_LIB_OK'] = 'TRUE'
sys.path.append(os.path.abspath(".."))


from src.portfolio_analytics import run_portfolio_analytics
from src.real_estate_bonds_data_builder import PortfolioMarketSimulator
from src.asset_class_builder import consolidate_portfolio_data
from src.market_data_engine import update_and_fetch_market_data
from src.index_builder import build_equal_weight_indices

base_dir = Path(os.getenv("data_dir", "."))
config_dir = Path(os.getenv("config_dir", "."))
market_data_dir = Path(os.getenv("market_data_dir", "."))
processed_dir = base_dir / "processed"
chart_dir = base_dir / "charts"
asset_class_dir = processed_dir / "Asset class"

def step1_synthetic_generation():
    input_csv = Path(processed_dir) / "portfolio_platform5_input.csv"
    simulator = PortfolioMarketSimulator(
        input_csv_path = input_csv,
        processed_dir = processed_dir,
        market_data_dir = market_data_dir,
        start_date = '2021-06-20',
        end_date = date.today().strftime('%Y-%m-%d')
    )
    simulator.process_portfolio()

def step2_fx_sync():
    update_and_fetch_market_data()

def step3_data_consolidation():
    consolidate_portfolio_data(
        positions_path = Path(processed_dir) / "Consolidated_Portfolio_Positions.csv",
        mapper_path = Path(config_dir) / "yfinance_ticker_mapper.csv",
        market_data_dir = market_data_dir,
        output_dir = asset_class_dir,
        price_column = 'Adjusted close price'
    )

def step4_index_construction():
    build_equal_weight_indices(processed_dir / "Asset class")

def step5_generate_dashboard():
    run_portfolio_analytics(
        processed_dir, chart_dir,
        Path(processed_dir) / "Consolidated_Portfolio_Positions.csv",
        Path(config_dir) / "yfinance_ticker_mapper.csv"
    )