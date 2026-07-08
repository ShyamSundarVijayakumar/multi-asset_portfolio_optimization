import os
import datetime
from pathlib import Path
import pandas as pd
import sys

sys.path.append(os.path.abspath(".."))
from src.ETF_holdings_enriched import process_universal_etf
from src.look_through_analysis import run_look_through_analysis
from src.portfolio_visualizer import generate_master_dashboard

from dotenv import load_dotenv
load_dotenv()

# --- Configurations ---
base_dir = Path(os.getenv("data_dir", "."))
output_dir = os.path.join(base_dir, "processed")
consolidated_file = Path(output_dir) / "Consolidated_Portfolio_Positions.csv"
chart_dir = os.path.join(base_dir, "charts")

# input paths/ETF's and output path + file name from .env file -- please also check description.txt for more details
etf1_output_csv = Path(output_dir) / "ETF1_holdings_enriched.csv"
etf2_output_csv = Path(output_dir) / "ETF2_holdings_enriched.csv"
etf3_output_csv = Path(output_dir) / "ETF3_holdings_enriched.csv"
etf4_output_csv = Path(output_dir) / "ETF4_holdings_enriched.csv"
input_dir = Path(os.getenv("ETF1_path"))
etf1_input_csv = Path(os.getenv("ETF1"))
etf2_input_csv = Path(os.getenv("ETF2"))
etf3_input_csv = Path(os.getenv("ETF3"))
etf4_input_csv = Path(os.getenv("ETF4"))
etf1_map_file = Path(input_dir) / "name_ticker_manual_map.csv"

# Configuration for orchestrator
# Source 1
etf1_ticker = os.getenv("etf1_ticker")
etf1_name = os.getenv("etf1_name")
etf1_config = {
    "index_name": etf1_name,      
    "index_ticker": etf1_ticker,             
    "col_name": "Component Name",                
    "col_isin": None,                
    "col_weight": "Weight %", 
    "col_country": None
}

# Source 2
etf2_ticker = os.getenv("etf2_ticker")
etf2_name = os.getenv("etf2_name")
etf2_config = {
    "index_name": etf2_name,      
    "index_ticker": etf2_ticker,             
    "col_name": "Name",                
    "col_isin": "ISIN",                
    "col_weight": "Weighting", 
    "col_country": "Country"
}

# Source 3
etf3_ticker = os.getenv("etf3_ticker")
etf3_name = os.getenv("etf3_name")
etf3_config = {
    "index_name": etf3_name,      
    "index_ticker": etf3_ticker,             
    "col_name": "Name",                
    "col_isin": "ISIN",                
    "col_weight": "Weighting", 
    "col_country": "Country"
}

# Source 4
etf4_ticker = os.getenv("etf4_ticker")
etf4_name = os.getenv("etf4_name")
Country = os.getenv("etf4_country")
etf4_config = {
    "index_name": etf4_name,      
    "index_ticker": etf4_ticker,             
    "col_name": "Security Name",                
    "col_isin": "ISIN",                
    "col_weight": "Weight (%)", 
    "col_country": Country
}

def execute_portfolio_engine():
    """Runs the base portfolio engine script."""
    import src.portfolio_engine as pe
    pe.run_pipeline() 

def run_update_orchestrator(selected_targets):
    """
    selected_targets: A list passed from Streamlit, e.g., ['1', '4'] or [] for none.
    """
    date_file = "last_run.txt"
    today = datetime.datetime.now()
    
    etf_registry = {
        '1': {'config': etf1_config, 'input': etf1_input_csv, 'output': etf1_output_csv, 'map': etf1_map_file},
        '2': {'config': etf2_config, 'input': etf2_input_csv, 'output': etf2_output_csv, 'map': None},
        '3': {'config': etf3_config, 'input': etf3_input_csv, 'output': etf3_output_csv, 'map': None},
        '4': {'config': etf4_config, 'input': etf4_input_csv, 'output': etf4_output_csv, 'map': None}
    }
    
    if not selected_targets:
        print("No updates selected for today. Passing through...")
        return

    ran_update = False
    for t in selected_targets:
        if t in etf_registry:
            data = etf_registry[t]
            process_universal_etf(
                input_path=data['input'],
                output_path=data['output'],
                map_path=data['map'],
                config=data['config']
            )
            ran_update = True
            
    if ran_update:
        with open(date_file, "w") as f:
            f.write(today.strftime("%Y-%m-%d"))

def generate_look_through():
    """Generates the Master Look Through Exposure."""
    etf_mapping = {
        etf1_name: etf1_output_csv,
        etf2_name: etf2_output_csv,
        etf3_name: etf3_output_csv,
        etf4_name: etf4_output_csv
    }    
    run_look_through_analysis(
        consolidated_file_path=consolidated_file, 
        etf_file_mapping=etf_mapping, 
        output_dir=output_dir
    )

def generate_dash():
    """Generates the Master Dashboard HTML."""
    etf_vars = [
        os.getenv("etf1_name"), os.getenv("etf1_ticker"),
        os.getenv("etf2_name"), os.getenv("etf2_ticker"),
        os.getenv("etf3_name"), os.getenv("etf3_ticker"),
        os.getenv("etf4_name"), os.getenv("etf4_ticker")
    ]
    etf_vars = [str(x).strip() for x in etf_vars if x]
    
    FILE_POSITIONS = Path(output_dir) / "Consolidated_Portfolio_Positions.csv"
    FILE_SUMMARY = Path(output_dir) / "Overall_PnL_and_Tax_Summary.xlsx"
    FILE_LOOK_THROUGH = Path(output_dir) / "Master_Look_Through_Exposure.csv"

    if not all([FILE_POSITIONS.exists(), FILE_SUMMARY.exists(), FILE_LOOK_THROUGH.exists()]):
        print("Error: Required files for dashboard are missing.")
        print(f"Check if these exist: {FILE_POSITIONS}, {FILE_SUMMARY}, {FILE_LOOK_THROUGH}")
        return None # Stop execution if files are missing
    generate_master_dashboard(FILE_POSITIONS, FILE_SUMMARY, FILE_LOOK_THROUGH, etf_vars, chart_dir)