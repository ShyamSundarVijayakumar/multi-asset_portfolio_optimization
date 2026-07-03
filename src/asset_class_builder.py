import pandas as pd
from pathlib import Path
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

def consolidate_portfolio_data(
    positions_path: Path, 
    mapper_path: Path, 
    market_data_dir: Path, 
    output_dir: Path,
    price_column: str = 'Adj Close'
):
    """
    Automated data builder: Maps tickers and asset classes directly from the mapper file,
    extracts the specified price column, and aligns dates across all assets.
    """
    logging.info("Loading positions and mapper files...")
    positions = pd.read_csv(positions_path)
    mapper = pd.read_csv(mapper_path)

    # 1. Create mapping dictionaries from the mapper file
    # We use these dicts to map ISIN or Security Name to both the Ticker (filename) and Asset Class
    isin_to_ticker = mapper.dropna(subset=['ISIN']).set_index('ISIN')['yfinance_tickters'].to_dict()
    name_to_ticker = mapper.dropna(subset=['Security Name']).set_index('Security Name')['yfinance_tickters'].to_dict()
    
    isin_to_class = mapper.dropna(subset=['ISIN']).set_index('ISIN')['asset_class'].to_dict()
    name_to_class = mapper.dropna(subset=['Security Name']).set_index('Security Name')['asset_class'].to_dict()

    # 2. Apply mappings to the positions dataframe
    positions['File_Identifier'] = positions['ISIN'].map(isin_to_ticker).fillna(positions['Security Name'].map(name_to_ticker))
    positions['Asset Class'] = positions['ISIN'].map(isin_to_class).fillna(positions['Security Name'].map(name_to_class))
    
    # Drop rows that couldn't be mapped
    unmapped = positions[positions['File_Identifier'].isna() | positions['Asset Class'].isna()]
    if not unmapped.empty:
        logging.warning(f"Could not fully map {len(unmapped)} assets. Listing unmapped items:")
        for index, row in unmapped.iterrows():
            logging.warning(f"-> ISIN: {row.get('ISIN', 'N/A')} | Security Name: {row.get('Security Name', 'N/A')}")
    
    positions = positions.dropna(subset=['File_Identifier', 'Asset Class']).copy()

    # 3. Load Market Data for All Mapped Assets
    logging.info("Loading market data...")
    all_asset_data = {}
    
    for identifier in positions['File_Identifier'].unique():
        file_path = market_data_dir / f"{identifier}.csv"
        
        if file_path.exists():
            df = pd.read_csv(file_path, parse_dates=['Date'], index_col='Date')
            if price_column in df.columns:
                all_asset_data[identifier] = df[price_column].rename(identifier)
            else:
                logging.warning(f"Column '{price_column}' not found in {identifier}.csv")
        else:
            logging.warning(f"Market data missing for {identifier}. Expected at {file_path}")

    if not all_asset_data:
        logging.error("No market data found. Exiting.")
        return

    # 4. Find the Lowest Common Timeframe (Inner Join)
    global_consolidated_df = pd.concat(all_asset_data.values(), axis=1, join='inner')
    logging.info(f"Global dataset trimmed to {len(global_consolidated_df)} overlapping trading days.")

    # 5. Save separated datasets per Asset Class
    output_dir.mkdir(parents=True, exist_ok=True)
    
    for asset_class, group in positions.groupby('Asset Class'):
        class_identifiers = group['File_Identifier'].unique()
        available_identifiers = [t for t in class_identifiers if t in global_consolidated_df.columns]
        
        if available_identifiers:
            asset_class_df = global_consolidated_df[available_identifiers]
            
            # Clean filename to prevent OS errors
            safe_filename = "".join([c for c in str(asset_class) if c.isalnum() or c==' ']).rstrip()
            output_file = output_dir / f"{safe_filename.replace(' ', '_')}_Consolidated.csv"
            
            asset_class_df.to_csv(output_file)
            logging.info(f"Saved {asset_class} data to {output_file.name} with shape {asset_class_df.shape}")