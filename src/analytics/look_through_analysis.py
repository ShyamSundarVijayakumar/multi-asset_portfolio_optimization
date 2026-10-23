import os
import re
from pathlib import Path
import pandas as pd
import numpy as np

def clean_company_name(name):
    """Standardizes company names by stripping legal suffixes."""
    if not isinstance(name, str) or name.lower() == 'nan':
        return ""
    name = name.lower().strip()
    suffixes = [r'\binc\b\.?', r'\bcorp\b\.?', r'\bco\b\.?', r'\bltd\b\.?', r'\bplc\b\.?', r'\bsa\b\.?', r'\bag\b\.?', r'\bgmbh\b\.?']
    for suffix in suffixes:
        name = re.sub(suffix, '', name)
    return re.sub(r'[^a-z0-9 ]', '', name).strip()

def load_asset_mapping(mapping_path):
    """Loads ISIN to Unified Key mapping from config/asset_groups.csv."""
    if not os.path.exists(mapping_path):
        print(f"Warning: {mapping_path} not found. Skipping custom groupings.")
        return {}
    
    mapping_df = pd.read_csv(mapping_path, sep=None, engine='python')
    mapping_df.columns = mapping_df.columns.str.strip()
    
    if 'ISIN' not in mapping_df.columns or 'Unified Key' not in mapping_df.columns:
        print(f"Error: Required columns 'ISIN' and 'Unified Key' not found in {mapping_path}")
        return {}
        
    return dict(zip(
        mapping_df['ISIN'].astype(str).str.strip().str.upper(), 
        mapping_df['Unified Key'].astype(str).str.strip()
    ))

def run_look_through_analysis(consolidated_file_path, etf_file_mapping, output_dir):
    print("Loading consolidated portfolio...")
    portfolio_df = pd.read_csv(consolidated_file_path)
    portfolio_df['Total Value (EUR)'] = portfolio_df['Current Quantity'] * portfolio_df['Current Price (EUR)']
    
    mapping_path = Path("../config/asset_groups.csv")
    asset_group_map = load_asset_mapping(mapping_path)

    portfolio_df['isin_clean'] = portfolio_df['ISIN'].astype(str).str.strip().str.upper()
    portfolio_df['name_clean'] = portfolio_df['Security Name'].apply(clean_company_name)
    
    # Separate ETFs using the passed-in mapping dict
    is_etf = portfolio_df['Security Name'].isin(etf_file_mapping.keys())
    direct_holdings = portfolio_df[~is_etf].copy()
    etf_holdings = portfolio_df[is_etf].copy()
    
    # Shatter ETFs
    look_through_records = []
    for _, etf_row in etf_holdings.iterrows():
        components_df = pd.read_csv(etf_file_mapping.get(etf_row['Security Name']), encoding='utf-8-sig')
        components_df['Weight Decimal'] = pd.to_numeric(components_df['Weight %'], errors='coerce') / 100
        components_df['Look-Through Value (EUR)'] = components_df['Weight Decimal'] * etf_row['Total Value (EUR)']
        
        for _, comp in components_df.iterrows():
            if pd.isna(comp['Look-Through Value (EUR)']): continue
            look_through_records.append({
                'Security Name': comp['Component Name'], 'ISIN': comp['ISIN'], 'Sector': comp['Sector'],
                'Industry': comp['Industry'], 'Country': comp['Country'], 'Total Value (EUR)': comp['Look-Through Value (EUR)'],
                'Source': f"Underlying ({etf_row['Security Name']})"
            })
            
    # Combine and Normalize
    raw_exposure_df = pd.concat([direct_holdings, pd.DataFrame(look_through_records)], ignore_index=True)
    raw_exposure_df['isin_clean'] = raw_exposure_df['ISIN'].astype(str).str.strip().str.upper()
    
    # Apply Unified Group Key
    raw_exposure_df['Group Key'] = raw_exposure_df['isin_clean'].apply(lambda x: asset_group_map.get(x, x))
    
    missing_or_nan = (raw_exposure_df['Group Key'] == 'NAN') | (raw_exposure_df['Group Key'].isna()) | (raw_exposure_df['Group Key'] == '')
    raw_exposure_df.loc[missing_or_nan, 'Group Key'] = raw_exposure_df.loc[missing_or_nan, 'Security Name'].apply(clean_company_name)

    # Final Aggregation
    final_exposure_df = raw_exposure_df.groupby(['Group Key', 'Sector', 'Industry', 'Country']).agg({
        'Security Name': 'first',
        'Total Value (EUR)': 'sum',
        'Source': lambda x: ", ".join(sorted(list(set(str(i) for i in x if pd.notna(i) and str(i).strip() != ""))))
    }).reset_index()
    
    final_exposure_df.to_csv(Path(output_dir) / "Master_Look_Through_Exposure.csv", index=False)
    print("Look-through analysis complete. File saved successfully.")
    return final_exposure_df