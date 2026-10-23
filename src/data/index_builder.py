import pandas as pd
from pathlib import Path
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

def build_equal_weight_indices(output_dir: Path):
    """
    Reads consolidated asset class files, handles missing data, 
    and constructs a base-100 equal-weighted index for each class.
    """
   
    if not output_dir.exists():
        logging.error(f"Directory not found: {output_dir}")
        return

    # Find all files ending with _Consolidated.csv
    consolidated_files = list(output_dir.glob("*_Consolidated.csv"))
    
    if not consolidated_files:
        logging.warning("No consolidated files found in the directory.")
        return

    for file_path in consolidated_files:
        logging.info(f"Processing {file_path.name}...")
        
        df = pd.read_csv(file_path, index_col=0, parse_dates=True)
        
        # 1. Handle missing values: Forward fill first, then backward fill for the start
        df = df.ffill().bfill()
        
        # 2. Calculate daily percentage returns
        returns = df.pct_change()
        
        # Set the very first row to 0 to act as our starting point (base day)
        returns.iloc[0] = 0 
        
        # 3. Calculate equal-weighted daily portfolio return (cross-sectional mean)
        # Using mean(axis=1) automatically applies the 1/N weight to all columns
        daily_index_return = returns.mean(axis=1)
        
        # 4. Construct the Cumulative Index (Rebased to 100)
        index_series = (1 + daily_index_return).cumprod() * 100
        
        # 5. Format the output DataFrame
        # e.g., "ETF_Consolidated" becomes "ETF_EW_Index"
        base_name = file_path.stem.replace("_Consolidated", "")
        index_name = f"{base_name}_EW_Index"
        
        index_df = index_series.to_frame(name="Index_Value")
        index_df.index.name = 'Date'
        
        # 6. Save back to the same directory
        output_file = output_dir / f"{index_name}.csv"
        index_df.to_csv(output_file)
        
        logging.info(f"Saved equal-weight index to {output_file.name}\n")
