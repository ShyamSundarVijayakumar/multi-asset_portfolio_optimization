import pandas as pd
import numpy as np
import plotly.express as px
import gc
from pathlib import Path
import logging
import faulthandler
faulthandler.enable()

logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
CHART_TEMPLATE = "plotly_white"

# ==========================================
# 1. DATA PROCESSING & MATH ENGINE
# ==========================================

def clean_data(df):
    """Safely handles missing gaps (NaNs) via forward-filling and forces numeric types."""
    df = df.ffill()
    df = df.apply(pd.to_numeric, errors='coerce')
    df = df.replace([np.inf, -np.inf], np.nan).fillna(0)
    return df

def infer_trading_days(returns_col, asset_label=""):
    """Dynamically detects if a specific asset trades 252 days or 365 days."""
    try:
        returns_col = returns_col.dropna()
        n = len(returns_col)

        if n < 20:
            return 252

        start = returns_col.index.min()
        end = returns_col.index.max()
        days_span = (end - start).days

        if days_span == 0:
            return 252

        obs_per_year = n / (days_span / 365.25)
        result = 365 if obs_per_year > 300 else 252
        return result
    except Exception as e:
        print(f"[infer_trading_days] asset={asset_label} ERROR: {e} returning 247")
        return 252

def _calculate_metrics_core(items_iterator, aligned_df, risk_free_rate=0.0):
    """
    Calculates summary metrics independently per asset, and covariance/correlation 
    matrices utilizing dynamically tracked trading day frequencies.
    """
    summary_records = []
    # ==========================================
    # 1. CALCULATE SUMMARY METRICS INDEPENDENTLY
    # ==========================================
    for asset_name, series in items_iterator:
        clean_series = series.dropna().astype(float)
        
        if len(clean_series) < 2:
            continue
            
        days_in_year = infer_trading_days(clean_series, asset_label=asset_name)
        years_of_data = len(clean_series) / days_in_year if days_in_year else 0
        
        if years_of_data > 0:
            safe_series = clean_series.clip(lower=-0.999) 
            total_log_return = np.log1p(safe_series).sum()
            ann_return = float(np.expm1(total_log_return / years_of_data))
        else:
            ann_return = 0.0
            
        daily_vol = float(clean_series.std())
        ann_volatility = daily_vol * np.sqrt(days_in_year)
        
        if ann_volatility > 0.001:
            sharpe_ratio = float((ann_return - risk_free_rate) / ann_volatility)
        else:
            sharpe_ratio = 0.0  
            
        summary_records.append({
            'Asset': asset_name,
            'Annualized Return': ann_return,
            'Annualized Volatility': ann_volatility,
            'Sharpe Ratio': sharpe_ratio
        })
        
    summary_df = pd.DataFrame(summary_records)
    
    # ==========================================
    # 2. CALCULATE MATRICES USING ALIGNED DATA
    # ==========================================
    n_obs = len(aligned_df)
    n_cols = len(aligned_df.columns)
    
    # Safeguard: Prevent kernel crash from massive matrices
    if n_cols > 5000:
        print(f"Warning: Attempting to calculate covariance on {n_cols} columns. This requires massive memory. Skipping matrix computation.")
        cov_matrix = pd.DataFrame()
        corr_matrix = pd.DataFrame()
        return cov_matrix, corr_matrix, summary_df

    if n_obs > 1:
        cov_matrix = aligned_df.cov() * 252 
        corr_matrix = aligned_df.corr()
    else:
        columns = [rec['Asset'] for rec in summary_records]
        cov_matrix = pd.DataFrame(0.0, index=columns, columns=columns)
        corr_matrix = pd.DataFrame(1.0, index=columns, columns=columns)
        
    return cov_matrix, corr_matrix, summary_df

def compile_macro_metrics(data_input, risk_free_rate=0.0):
    if isinstance(data_input, pd.DataFrame):
        df_safe = data_input.loc[~data_input.index.duplicated(keep='last')].copy()
        
        for col in df_safe.columns:
            df_safe[col] = pd.to_numeric(df_safe[col], errors='coerce')
        
        aligned_df = df_safe.dropna()
        items_iterator = ((str(col), aligned_df[col]) for col in aligned_df.columns)

    return _calculate_metrics_core(items_iterator, aligned_df, risk_free_rate=risk_free_rate)
    
def compute_portfolio_weights(positions_path, mapper_path):
    try:
        pos_df = pd.read_csv(positions_path)
        map_df = pd.read_csv(mapper_path)
        
        pos_df.columns = [c.strip().lower() for c in pos_df.columns]
        map_df.columns = [c.strip().lower() for c in map_df.columns]
        
        merged = pd.merge(pos_df, map_df, on='isin', how='left')
        merged['total_value'] = merged['current quantity'] * merged['current price (eur)']
        total_portfolio_value = merged['total_value'].sum()
        
        weights_dict = {}
        for _, row in merged.iterrows():
            weight = row['total_value'] / total_portfolio_value if total_portfolio_value > 0 else 0
            ticker = str(row.get('yfinance_tickters', '')).strip()
            isin = str(row.get('isin', '')).strip().lower()
            
            if pd.notna(ticker) and ticker not in ['', 'nan']:
                weights_dict[ticker] = weights_dict.get(ticker, 0) + weight
            elif 'bonds' in isin:
                weights_dict['bonds'] = weights_dict.get('bonds', 0) + weight
            elif 'real_estate' in isin or 'real estate' in isin:
                weights_dict['real_estate'] = weights_dict.get('real_estate', 0) + weight
                
        return weights_dict
    except Exception as e:
        logging.error(f"Failed to compute weights from files: {e}")
        return {}


# ==========================================
# 2. DASHBOARD UI ENGINE
# ==========================================

def get_vol_highlight(val):
    """Returns HTML annotation highlighting the standard deviation risk bracket."""
    if val < 0.10: return f"{val:.2%} (🟢 Low Risk)"
    elif val <= 0.20: return f"{val:.2%} (🟡 Moderate Risk)"
    elif val <= 0.30: return f"{val:.2%} (🟠 High Risk)"
    else: return f"{val:.2%} (🔴 Very High Risk)"

def get_sharpe_highlight(val):
    """Returns HTML annotation highlighting the risk-adjusted return bracket."""
    if val < 1.0: return f"{val:.2f} (Sub-par / Needs adjustment)"
    elif val < 2.0: return f"<b>{val:.2f} (🟢 Good / Balanced)</b>"
    elif val < 3.0: return f"<b>{val:.2f} (🔵 Very Good / Efficient)</b>"
    else: return f"<b>{val:.2f} (⭐ Excellent / Exceptional)</b>"

def create_html_table(df):
    df_styled = df.copy()
    df_styled['Annualized Return'] = df_styled['Annualized Return'].apply(lambda x: f"{x:.2%}")
    
    df_styled['Annualized Volatility'] = df_styled['Annualized Volatility'].apply(
        lambda x: f'<span title="Measures dispersion around the mean (historical volatility brackets:\n<10% Low Risk\n10-20% Moderate\n20-30% High\n>30% Very High)">{get_vol_highlight(x)}</span>'
    )
    df_styled['Sharpe Ratio'] = df_styled['Sharpe Ratio'].apply(
        lambda x: f'<span title="Excess return per unit of volatility relative to a risk-free rate.\nInterpretation Scale:\n<1.0: Sub-par / Weak return-to-risk\n1.0-1.99: Good / Balanced\n2.0-2.99: Very Good / Highly efficient\n>=3.0: Excellent / Outstanding">{get_sharpe_highlight(x)}</span>'
    )
        
    html_table = df_styled.to_html(classes="custom-table display", index=False, justify="left", border=0, escape=False)
    return f'<div class="table-container" style="margin-top: 20px;">\n{html_table}\n</div>'


def get_plotly_divs(corr_matrix, summary_df, returns_df, matrix_css):
    n_assets = len(corr_matrix.columns)
    show_text = n_assets <= 15
    
    if matrix_css == "full-width":
        matrix_height = max(720, n_assets * 28)
        left_margin = max(110, min(n_assets * 7, 200))
        bottom_margin = max(110, min(n_assets * 7, 200))
    else:
        matrix_height = max(600, n_assets * 28)
        left_margin = 60
        bottom_margin = 110

    fig_corr = px.imshow(
        corr_matrix.round(2), text_auto=show_text, 
        color_continuous_scale='RdBu_r', zmin=-1, zmax=1,
        title="Correlation Matrix", template=CHART_TEMPLATE
    )
    
    fig_corr.update_traces(hovertemplate="<b>Asset 1:</b> %{x}<br><b>Asset 2:</b> %{y}<br><b>Correlation:</b> %{z:.2f}<extra></extra>")
    
    fig_corr.update_layout(
      #  title_x=0.5,
        title=dict(text="Correlation Matrix", x=0.5, font=dict(size=20, weight='bold')),
        height=matrix_height,
        margin=dict(l=left_margin, r=15, t=70, b=bottom_margin),
        xaxis=dict(tickmode='linear', dtick=1, tickfont=dict(size=11)),
        yaxis=dict(tickmode='linear', dtick=1, tickfont=dict(size=11)),
        coloraxis_colorbar=dict(thickness=25, lenmode="pixels", len=400, title="Correlation")
    )
    
    corr_div = fig_corr.to_html(full_html=False, include_plotlyjs=False, config={'responsive': True})
    del fig_corr
    
    fig_scatter = px.scatter(
        summary_df, x='Annualized Volatility', y='Annualized Return', 
        text='Asset', color='Sharpe Ratio', size_max=15,
        title="Risk vs. Return", template=CHART_TEMPLATE
    )
    fig_scatter.update_traces(textposition='top center')
    fig_scatter.update_layout(title=dict(text="Risk vs. Return", x=0.5, font=dict(size=20, weight='bold')), xaxis=dict(tickformat=".2%"), yaxis=dict(tickformat=".2%"), height=550)
    scatter_div = fig_scatter.to_html(full_html=False, include_plotlyjs=False, config={'responsive': True})
    del fig_scatter

    flat_returns = returns_df.values.flatten()
    flat_returns = flat_returns[~np.isnan(flat_returns)]
    
    fig_dist = px.histogram(
        x=flat_returns, nbins=100, 
        title="Daily Return Distribution", 
        template=CHART_TEMPLATE,
        labels={'x': 'Daily Return', 'y': 'Frequency'}
    )
    fig_dist.update_layout(title=dict(text="Daily Return Distribution", x=0.5, font=dict(size=20, weight='bold')), xaxis=dict(tickformat=".2%"), height=550, showlegend=False)
    fig_dist.update_traces(marker_color='#3498DB', opacity=0.75)
    dist_div = fig_dist.to_html(full_html=False, include_plotlyjs=False, config={'responsive': True})
    del fig_dist
    
    return corr_div, scatter_div, dist_div


def build_master_dashboard(tabs_data, kpi_data, out_path: Path):
    html_divs = ""
    buttons = ""
    
    for i, (tab_id, content) in enumerate(tabs_data.items()):
        display_style = "block" if i == 0 else "none"
        active_class = "active" if i == 0 else ""
        
        buttons += f'<button class="tablinks {active_class}" onclick="openTab(event, \'{tab_id}\')">{tab_id}</button>\n'
        
        tab_content = f"""
        <div class="chart-grid">
            <div class="chart-box {content['matrix_css']}">{content['corr']}</div>
            <div class="chart-box {content['scatter_css']}">{content['scatter']}</div>
            <div class="chart-box {content['scatter_css']}">{content['dist']}</div>
        </div>
        {content['table']}
        """
        html_divs += f'<div id="{tab_id}" class="tabcontent" style="display:{display_style};">{tab_content}</div>\n'

    master_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Portfolio Risk Master Dashboard</title>
        <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
        <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css">
        <script src="https://code.jquery.com/jquery-3.7.0.min.js"></script>
        <script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
        
        <style>
            :root {{ --bg-color: #f4f6f9; --card-bg: #ffffff; --text-color: #2C3E50; --table-alt: #F8F9F9; }}
            body.dark-mode {{ --bg-color: #1a1a1a; --card-bg: #2d2d2d; --text-color: #ffffff; --table-alt: #383838; }}
            
            body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 20px; background: var(--bg-color); color: var(--text-color); margin: 0; transition: background 0.2s, color 0.2s; }}
            
            .header-container {{ text-align: center; margin-bottom: 30px; position: relative; }}
            .header-title {{ margin: 0 0 20px 0; font-size: 32px; color: var(--text-color); font-weight: 700; }}
            
            #theme-toggle {{ position: absolute; top: 0; right: 0; cursor: pointer; padding: 8px 16px; border-radius: 6px; border: 1px solid #ccc; background: var(--card-bg); color: var(--text-color); font-weight: bold; }}
            
            .kpi-container {{ display: flex; justify-content: center; flex-wrap: wrap; gap: 20px; }}
            /* MODIFIED: Added informative title attributes to the asset risk benchmark flash cards */
            .kpi-card {{ background: var(--card-bg); padding: 20px 30px; border-radius: 10px; border: 1px solid #e0e0e0; text-align: center; min-width: 200px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }}
            .kpi-card h3 {{ margin: 0; font-size: 13px; color: #7f8c8d; text-transform: uppercase; letter-spacing: 1px; }}
            .kpi-card p {{ margin: 10px 0 0; font-size: 24px; font-weight: bold; color: var(--text-color); }}
            
            .tab {{ display: flex; justify-content: center; flex-wrap: wrap; background: var(--card-bg); padding: 10px; border-radius: 8px 8px 0 0; border: 1px solid #ccc; gap: 5px; }}
            .tab button {{ padding: 12px 24px; cursor: pointer; border: none; background: none; font-weight: bold; color: #566573; font-size: 15px; transition: 0.2s; border-bottom: 3px solid transparent; }}
            .tab button:hover {{ color: #2C3E50; background: #f0f3f4; border-radius: 4px; }}
            .tab button.active {{ background: #3498DB; color: white; border-radius: 4px; border-bottom: 3px solid #2980B9; }}
            
            .tabcontent {{ background: var(--card-bg); padding: 20px; border: 1px solid #ccc; border-top: none; border-radius: 0 0 8px 8px; }}
            
            .chart-grid {{ display: flex; flex-wrap: wrap; gap: 20px; justify-content: center; margin-bottom: 20px; align-items: stretch; }}
            .chart-box {{ background: var(--card-bg); border-radius: 8px; border: 1px solid #eee; overflow: hidden; padding: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.02); }}
            .chart-box > div {{ width: 100% !important; height: 100% !important; }}
            .chart-box .js-plotly-plot {{ width: 100% !important; height: 100% !important; }}
            
            .full-width {{ flex: 1 1 100%; }}
            .half-width {{ flex: 1 1 48%; min-width: 500px; }}
            
            .custom-table {{ width: 100% !important; border-collapse: collapse !important; color: #2C3E50; margin-top: 20px; }}
            .custom-table th {{ background-color: #2C3E50 !important; color: white !important; padding: 12px !important; text-align: left; }}
            .custom-table td {{ padding: 10px !important; border-bottom: 1px solid #EAECEE !important; }}
            .custom-table tbody tr:nth-child(even) {{ background-color: var(--table-alt) !important; }}
            .custom-table tbody tr:hover td {{ background-color: #3498DB !important; color: white !important; cursor: pointer; }}
            .dataTables_wrapper {{ color: var(--text-color) !important; }}
        </style>
    </head>
    <body>
        <div class="header-container">
            <button id="theme-toggle" onclick="toggleTheme()">Toggle Dark Mode</button>
            <h2 class="header-title">Portfolio Risk Analytics</h2>
            
            <div class="kpi-container">
                <div class="kpi-card" title="Measures the dispersion of portfolio returns around the mean (historical volatility brackets:\n<10% Low Risk\n10-20% Moderate\n20-30% High\n>30% Very High)">
                    <h3>Portfolio Volatility</h3>
                    <p>{kpi_data['vol']:.2%}</p>
                </div>
                <div class="kpi-card" title="Excess return per unit of volatility relative to a risk-free rate.\nInterpretation Scale:\n<1.0: Sub-par / Weak return-to-risk\n1.0-1.99: Good / Balanced\n2.0-2.99: Very Good / Highly efficient\n>=3.0: Excellent / Outstanding">
                    <h3>Sharpe Ratio</h3>
                    <p>{kpi_data['sharpe']:.2f}</p>
                </div>
                <div class="kpi-card" title="Worst historical peak-to-trough drop"><h3>Max Drawdown</h3><p style="color:#e74c3c;">{kpi_data['mdd']:.2%}</p></div>
                <div class="kpi-card" title="Maximum expected daily loss (95% confidence)"><h3>Daily VaR (95%)</h3><p style="color:#e74c3c;">{kpi_data['var']:.2%}</p></div>
            </div>
        </div>
        
        <div class="tab">{buttons}</div>
        {html_divs}
    
        <script>
            $(document).ready(function() {{ 
                $('.custom-table').DataTable({{ "pageLength": 15 }});
                setTimeout(function() {{
                    window.dispatchEvent(new Event('resize'));
                }}, 200);
            }});
            
            function toggleTheme() {{ 
                var body = document.body;
                body.classList.toggle('dark-mode'); 
                if (body.classList.contains('dark-mode')) {{
                    $('.custom-table td').css('color', '#ffffff');
                    $('.chart-box').css('border-color', '#444');
                }} else {{
                    $('.custom-table td').css('color', '#2C3E50');
                    $('.chart-box').css('border-color', '#eee');
                }}
            }}
            
            function openTab(evt, tabName) {{
                document.querySelectorAll('.tabcontent').forEach(el => el.style.display = 'none');
                document.querySelectorAll('.tablinks').forEach(el => el.className = el.className.replace(" active", ""));
                document.getElementById(tabName).style.display = 'block';
                evt.currentTarget.className += " active";
                window.dispatchEvent(new Event('resize')); 
            }}
        </script>
    </body>
    </html>
    """
    
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(master_html)


# ==========================================
# 3. MAIN EXECUTION PIPELINE
# ==========================================

def run_portfolio_analytics(processed_dir: Path, chart_dir: Path, positions_csv: Path, mapper_csv: Path):
    target_dir = Path(processed_dir) / "Asset class"
    
    export_dir = Path(chart_dir) / "data_exports"
    export_dir.mkdir(parents=True, exist_ok=True)
    
    dashboard_tabs = {}
    
    weights_dict = compute_portfolio_weights(positions_csv, mapper_csv)
    all_individual_returns = pd.DataFrame()

    # 1. MICRO ANALYSIS (Intra-Asset Class)
    for file_path in target_dir.glob("*_Consolidated.csv"):
        class_name = file_path.stem.replace("_Consolidated", "").replace("_", " ")
        
        df = pd.read_csv(file_path, index_col=0, parse_dates=True)
        df.index = pd.to_datetime(df.index).tz_localize(None).normalize()
        df = clean_data(df)
        
        returns = df.pct_change().dropna(how='all')
        all_individual_returns = pd.concat([all_individual_returns, returns], axis=1)
        
        if len(returns) < 2 or len(returns.columns) < 2:
            continue
            
        try:
            cov, corr, summary = compile_macro_metrics(returns, risk_free_rate=0.0)
            
            safe_name = class_name.replace(" ", "_")
            summary.to_csv(export_dir / f"IntraClass_{safe_name}_Metrics.csv", index=False)
            cov.to_csv(export_dir / f"IntraClass_{safe_name}_Covariance.csv")
            
            n_assets = len(corr.columns)
            matrix_css = "full-width" if n_assets >= 12 else "half-width"
            scatter_css = "half-width"
            
            corr_div, scatter_div, dist_div = get_plotly_divs(corr, summary, returns, matrix_css)
            table_html = create_html_table(summary)
            
            dashboard_tabs[class_name] = {
                'corr': corr_div, 'scatter': scatter_div, 'dist': dist_div, 
                'table': table_html, 'matrix_css': matrix_css, 'scatter_css': scatter_css
            }
        except Exception as ex:
            logging.error(f"[{class_name}] ERROR: {ex}")

    # 2. MACRO ANALYSIS
    macro_returns_dict = {}

    for file_path in target_dir.glob("*_EW_Index.csv"):
        df = pd.read_csv(file_path, index_col=0, parse_dates=True)
        df.index = pd.to_datetime(df.index).tz_localize(None).normalize()
        class_name = file_path.stem.replace("_EW_Index", "").replace("_", " ")
        ret = df['Index_Value'].pct_change()
        macro_returns_dict[class_name] = pd.to_numeric(ret, errors='coerce')

    standalone = {'Bonds': 'bonds.csv', 'Real Estate': 'real_estate.csv'}
    for name, filename in standalone.items():
        file_path = target_dir / filename
        if file_path.exists():
            df = pd.read_csv(file_path, index_col=0)
            df.index = pd.to_datetime(df.index, dayfirst=True, format='mixed').tz_localize(None).normalize()
            ret = pd.to_numeric(df['Daily_Return_EUR'], errors='coerce')
            macro_returns_dict[name] = ret
            key = 'bonds' if name == 'Bonds' else 'real_estate'
            all_individual_returns[key] = ret

    # 3. TRUE PORTFOLIO METRICS
    kpi_data = {'vol': 0.0, 'sharpe': 0.0, 'mdd': 0.0, 'var': 0.0}
    
    if weights_dict and not all_individual_returns.empty:
        all_individual_returns = clean_data(all_individual_returns)
        true_portfolio_returns = pd.Series(0.0, index=all_individual_returns.index)
        
        for ticker, weight in weights_dict.items():
            if ticker in all_individual_returns.columns:
                true_portfolio_returns += all_individual_returns[ticker] * weight
                
        port_days = infer_trading_days(true_portfolio_returns)
        safe_series = true_portfolio_returns.clip(lower=-0.999)
        total_log_return = np.log1p(safe_series).sum()
        years = len(true_portfolio_returns.dropna()) / port_days
        ann_return = float(np.expm1(total_log_return / years))
        kpi_data['vol'] = true_portfolio_returns.std() * np.sqrt(port_days)
        
        kpi_data['sharpe'] = (ann_return - 0) / kpi_data['vol'] if kpi_data['vol'] > 0 else 0
        
        kpi_data['var'] = np.percentile(true_portfolio_returns.dropna(), 5)
        non_na = true_portfolio_returns.dropna()
        cumulative_returns = (1 + true_portfolio_returns).cumprod()
        peak = cumulative_returns.cummax()
        drawdown = (cumulative_returns - peak) / peak
        kpi_data['mdd'] = drawdown.min()     

    # 4. COMPILE MACRO TABS
    if macro_returns_dict:
        series_dict = {}
        
        for asset_name, s in macro_returns_dict.items():
            # Standardize and deduplicate dates
            s = s.loc[~s.index.duplicated(keep='last')]
            series_dict[asset_name] = pd.to_numeric(s, errors='coerce')
        
        aligned_df = pd.concat(series_dict, axis=1, join='inner').dropna()
        try:
            cov, corr, summary = _calculate_metrics_core(series_dict.items(), aligned_df, risk_free_rate=0.0)
            summary.to_csv(export_dir / "Macro_Allocation_Metrics.csv", index=False)
            cov.to_csv(export_dir / "Macro_Allocation_Covariance.csv")
            
            n_assets = len(corr.columns)
            matrix_css = "full-width" if n_assets >= 12 else "half-width"
            
            corr_div, scatter_div, dist_div = get_plotly_divs(corr, summary, aligned_df, matrix_css)
            table_html = create_html_table(summary)
            
            dashboard_tabs = {"Macro Allocation": {
                'corr': corr_div, 'scatter': scatter_div, 'dist': dist_div,
                'table': table_html, 'matrix_css': matrix_css, 'scatter_css': 'half-width'
            }} | dashboard_tabs
        except Exception as ex:
            logging.error(f"[MACRO] ERROR: {ex}")
    
    if dashboard_tabs:
        build_master_dashboard(dashboard_tabs, kpi_data, Path(chart_dir) / "Master_Risk_Dashboard.html")
        gc.collect()