import os
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path

PROFESSIONAL_PALETTE = [
    '#E6194B', '#3CB44B', '#FFE119', '#4363D8', '#F58231', 
    '#911EB4', '#46F0F0', '#F032E6', '#BCF60C', '#FABEBE', 
    '#008080', '#E6BEFF', '#9A6324', '#FFFAC8', '#800000'
]
CHART_TEMPLATE = "plotly_white"

def determine_asset_class(row, etf_vars):
    """Classifies assets dynamically using custom rules."""
    isin = str(row.get('ISIN', '')).upper().strip()
    name = str(row.get('Security Name', '')).strip()
    name_lower = name.lower()
    sector = str(row.get('Sector', '')).strip().lower()
    
    if any(x and (name == x or isin == str(x).upper().strip()) for x in etf_vars):
        return "ETF"
    elif isin.startswith("CRYPTO CURRENCY") or sector == "cryptocurrency":
        return "Cryptocurrency"
    elif isin.startswith("bonds") or name == "bonds":
        return "Bonds"
    elif sector == "real estate":
        return "Real Estate"
    elif sector == "commodities" or "gold" in name_lower or "silver" in name_lower or isin in ["GC=F", "SI=F"]:
        return f"Commodity ({name})"
    else:
        return "Equity (Stock)"

def create_html_table(df, drop_cols=None):
    """Generates a native HTML table to allow hover effects and perfect sizing."""
    if drop_cols:
        df = df.drop(columns=[c for c in drop_cols if c in df.columns], errors='ignore')
    
    df = df.round(2).fillna("N/A")
    html_table = df.to_html(classes="custom-table", index=False, justify="left", border=0)
    
    # Wrap in a scrollable container
    return f'<div class="table-container">\n{html_table}\n</div>'

def generate_master_dashboard(file_positions, file_summary, file_look_through, etf_vars, chart_dir):
    print("Compiling Master Interactive Dashboard...")
    
    df_pos = pd.read_csv(file_positions)
    df_look = pd.read_csv(file_look_through)
    
    val_col = 'Total Value (EUR)'
    df_pos['Total Value (EUR)'] = df_pos['Current Quantity'] * df_pos['Current Price (EUR)']
    df_pos['Asset Class'] = df_pos.apply(lambda row: determine_asset_class(row, etf_vars), axis=1)

    df_pos['Invested Amount (EUR)'] = df_pos['Current Quantity'] * df_pos['Average Buy Price (EUR)']
    inv_col = 'Invested Amount (EUR)'

    figures = {}
    html_tables = {}

    # --- 1. TREEMAP ---
    plot_df = df_look.copy()
    
    plot_df = plot_df[plot_df[val_col].notna() & (plot_df[val_col] > 0)].copy()
    plot_df[val_col] = plot_df[val_col].round(1)

    columns_to_clean = ['Sector', 'Industry', 'Group Key', 'Security Name']
    for col in columns_to_clean:
        if col in plot_df.columns:
            plot_df[col] = plot_df[col].fillna(f"Unknown {col}").astype(str).str.strip()
            plot_df.loc[plot_df[col] == "", col] = f"Unknown {col}"

    if not plot_df.empty:
        fig_tree = px.treemap(
            plot_df, 
            path=[px.Constant("Total Portfolio"), 'Sector', 'Industry', 'Security Name'], 
            values=val_col, 
            color='Sector',
            color_discrete_sequence=PROFESSIONAL_PALETTE,
            title="Portfolio Look-Through Exposure Summary",
            hover_data=['Group Key', 'Source'] if 'Group Key' in plot_df.columns else None
        )
        fig_tree.update_traces(
            textinfo="label+value+percent parent",
            hovertemplate="<b>%{label}</b><br>Value: €%{value:,.1f}<br>Share: %{percentParent:.2%}<extra></extra>"
        )
        fig_tree.update_layout(
            template=CHART_TEMPLATE, 
            margin=dict(t=50, l=10, r=10, b=10), 
            height=750, 
            font=dict(color="white", family="Segoe UI, Tahoma, Geneva, Verdana, sans-serif", size=12)
        )
        figures['Treemap'] = fig_tree
    else:
        print("Warning: Look-through data is empty after filtering positive values.")

    # --- 2. ASSET CLASS (Donut) ---
    df_asset_class = df_pos.groupby('Asset Class')['Total Value (EUR)'].sum().reset_index()

    fig_asset = px.pie(
        df_asset_class, values='Total Value (EUR)', names='Asset Class', hole=0.65,
        title='Asset Class Allocation Summary', color_discrete_sequence=PROFESSIONAL_PALETTE
    )
    fig_asset.update_traces(textinfo='percent+label', textposition='outside')
    fig_asset.update_layout(showlegend=False, title_x=0.5,template=CHART_TEMPLATE, height=700)
    figures['Asset_Class'] = fig_asset

    # --- 3. ASSET CLASS PERFORMANCE ---
    # Perform the groupby
    df_perf = df_pos.groupby('Asset Class')[[inv_col, 'Total Value (EUR)']].sum().reset_index()
    
    # Calculate % Change
    df_perf['% Change'] = ((df_perf['Total Value (EUR)'] - df_perf[inv_col]) / df_perf[inv_col] * 100)
    df_perf['Hover Info'] = df_perf['% Change'].apply(lambda x: f"Change: {x:+.2f}%")
    
    # Melt for Grouped Bar
    df_melt = df_perf.melt(id_vars=['Asset Class', 'Hover Info'], value_vars=[inv_col, 'Total Value (EUR)'], 
                           var_name='Metric', value_name='Amount (EUR)')
    
    fig_perf = px.bar(
        df_melt, x='Asset Class', y='Amount (EUR)', color='Metric', barmode='group',
        title='Invested Amount vs Present Value by Asset Class',
        text_auto='.2s', color_discrete_sequence=['#FCF3CF', '#2ECC71'],
        hover_data={'Hover Info': True}
    )
    fig_perf.update_layout(title_x=0.5, template=CHART_TEMPLATE, height=600)
    figures['Asset_Class_Performance'] = fig_perf

    # --- 4. SECTOR-WISE ALLOCATION FOR STOCKS ---
    df_stocks = df_pos[df_pos['Asset Class'] == 'Equity (Stock)'].copy()
    if 'Sector' in df_stocks.columns and not df_stocks.empty:
        df_stock_sector = df_stocks.groupby('Sector')['Total Value (EUR)'].sum().reset_index()
        fig_stock_sector = px.pie(
            df_stock_sector, values='Total Value (EUR)', names='Sector', hole=0.45,
            title='Sector Allocation (Equities Only)', color_discrete_sequence=PROFESSIONAL_PALETTE
        )
        fig_stock_sector.update_traces(textinfo='percent+label', textposition='outside')
        fig_stock_sector.update_layout(showlegend=False, title_x=0.5, template=CHART_TEMPLATE, height=700)
        figures['Stock_Sectors'] = fig_stock_sector
    
    # --- 5. COUNTRY ALLOCATION (Look-Through Based) ---
    if 'Country' in df_look.columns:
        df_country = df_look.groupby('Country')[val_col].sum().reset_index().sort_values(by=val_col)
        
        fig_country = px.bar(
            df_country, x=val_col, y='Country', orientation='h',
            title='Geographic Capital Allocation (Underlying Holdings)', 
            color=val_col, 
            color_continuous_scale=['#FCF3CF', '#2ECC71']
        )
        fig_country.update_layout(title_x=0.5, template=CHART_TEMPLATE, coloraxis_showscale=False, height=600)
        figures['Country_Bar'] = fig_country
    else:
        print("Note: 'Country' column not found in look-through data; skipping Geographic chart.")

    # --- 6. WATERFALL PNL ---
    if file_summary.exists():
        try:
            df_metrics = pd.read_excel(file_summary, sheet_name='Performance Metrics')
            gains = df_metrics.loc[df_metrics['Performance Metric'] == 'Realized Gains (EUR)', 'Value'].values[0]
            losses = df_metrics.loc[df_metrics['Performance Metric'] == 'Realized Losses / Harvestable (EUR)', 'Value'].values[0]
            unrealized = df_metrics.loc[df_metrics['Performance Metric'] == 'Unrealized PnL (EUR)', 'Value'].values[0]
            fees = df_metrics.loc[df_metrics['Performance Metric'] == 'Total Broker Fees Paid (EUR)', 'Value'].values[0]
            
            fig_pnl = go.Figure(go.Waterfall(
                name="PnL", orientation="v", measure=["relative", "relative", "relative", "relative", "total"],
                x=["Realized Gains", "Realized Losses", "Unrealized PnL", "Broker Fees Paid", "Net Portfolio PnL"],
                textposition="outside", y=[gains, losses, unrealized, -fees, 0],
                connector={"line":{"color":"#BDC3C7"}}, decreasing={"marker":{"color":"#E74C3C"}},
                increasing={"marker":{"color":"#FCF3CF"}}, totals={"marker":{"color":"#2ECC71"}} ##2C3E50 
            ))
            fig_pnl.update_layout(title="Portfolio Capital Value Drivers", title_x=0.5, template=CHART_TEMPLATE, height=600)
            figures['Waterfall_PnL'] = fig_pnl
        except Exception as e:
            print(f"Metrics notation skip: {e}")

    # --- 7. YEARLY PNL SUMMARY ---
    try:
        df_closed = pd.read_excel(file_summary, sheet_name='Closed Positions Log')
        df_closed['Sell Date'] = pd.to_datetime(df_closed['Sell Date'])
        df_closed['Year'] = df_closed['Sell Date'].dt.year
        df_yearly = df_closed.groupby('Year')['Total Profit or Loss (EUR)'].sum().reset_index()
        df_yearly['Year_Str'] = df_yearly['Year'].astype(str)

        fig_yearly = px.bar(
            df_yearly, 
            x='Year_Str', 
            y='Total Profit or Loss (EUR)',
            title="Realized Profit/Loss by Year",
            text_auto='.2s',
            color='Total Profit or Loss (EUR)',
            color_continuous_scale=['#FCF3CF', '#2ECC71']
        )
        
        fig_yearly.update_layout(
            title_x=0.5, 
            template=CHART_TEMPLATE, 
            height=600,
            xaxis_title="Year"
        )

        fig_yearly.update_xaxes(type='category')   
        figures['Yearly_Performance'] = fig_yearly
    except Exception as e:
        print(f"Could not generate Yearly PnL chart: {e}")
        
    # --- 8. YEAR-OVER-YEAR DIVIDEND TRACKING ---
    try:
        df_div = pd.read_excel(file_summary, sheet_name='Performance Metrics')
        
        if 'Performance Metric' in df_div.columns and 'Value' in df_div.columns:
            # Filter rows that look like "Dividends Collected in XXXX (EUR)"
            df_div_filtered = df_div[df_div['Performance Metric'].str.contains('Dividends Collected in', na=False, case=False)].copy()
            
            # Extract the 4-digit year from the text string using regex
            df_div_filtered['Year'] = df_div_filtered['Performance Metric'].str.extract(r'(\d{4})')
            df_div_filtered['Value'] = pd.to_numeric(df_div_filtered['Value'], errors='coerce')
            df_div_filtered = df_div_filtered.dropna(subset=['Year', 'Value'])
            df_div_filtered = df_div_filtered.sort_values('Year')
            
            df_div_filtered['Year_Str'] = df_div_filtered['Year'].astype(str)

            fig_div = px.bar(
                df_div_filtered, 
                x='Year_Str', 
                y='Value',
                title="Year-over-Year Dividend Income (Bonds yield not considered)",
                text_auto='.2s',
                labels={'Value': 'Dividends (EUR)', 'Year_Str': 'Year'},
                color='Value', 
                color_continuous_scale=['#FCF3CF', '#2ECC71']
            )
            
            fig_div.update_layout(
                title_x=0.5, 
                template=CHART_TEMPLATE, 
                height=600
            )
    
            fig_div.update_xaxes(type='category')
            
            figures['Dividend_Growth'] = fig_div
        else:
            print("Warning: 'Performance Metric' or 'Value' column missing in Dividend Log.")
    except Exception as e:
        print(f"Could not generate Dividend Tracking chart: {e}")
    
    # --- 9. NATIVE HTML DATA TABLES ---
    if 'Unrealized PnL (EUR)' in df_pos.columns:
        df_pos_sorted = df_pos.sort_values(by='Unrealized PnL (EUR)', ascending=False)
    else:
        df_pos_sorted = df_pos
        
    val_col_look = 'Total Value (EUR)'
    df_look_sorted = df_look.sort_values(by=val_col_look, ascending=False)

    html_tables['Table_Positions'] = create_html_table(df_pos_sorted, drop_cols=['Native Currency', 'Asset Class'])
    html_tables['Table_LookThrough'] = create_html_table(df_look_sorted, drop_cols=['Source', 'Group Key'])

    # --- HTML WEB PAGE GENERATOR ---
    html_divs = ""
    buttons = ""
    
    # Process Plotly Figures
    for i, (name, fig) in enumerate(figures.items()):
        display_style = "block" if i == 0 else "none"
        active_class = "active" if i == 0 else ""
        div_html = fig.to_html(full_html=False, include_plotlyjs='cdn' if i==0 else False)
        buttons += f'<button class="tablinks {active_class}" onclick="openTab(event, \'{name}\')">{name.replace("_", " ")}</button>\n'
        html_divs += f'<div id="{name}" class="tabcontent" style="display:{display_style};">{div_html}</div>\n'
        
    # Process HTML Tables
    for name, table_html in html_tables.items():
        # They default to hidden
        buttons += f'<button class="tablinks" onclick="openTab(event, \'{name}\')">{name.replace("_", " ")}</button>\n'
        html_divs += f'<div id="{name}" class="tabcontent" style="display:none;">{table_html}</div>\n'

    total_val_fmt = f"{df_pos['Total Value (EUR)'].sum():,.0f}"
    total_pnl_fmt = f"{pd.to_numeric(df_pos['Unrealized PnL (EUR)'], errors='coerce').sum():,.0f}"

    master_html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <title>Portfolio Master Dashboard</title>
            <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
            <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css">
            <script src="https://code.jquery.com/jquery-3.7.0.min.js"></script>
            <script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
            
            <style>
                :root {{ 
                    --bg-color: #f4f6f9; 
                    --card-bg: #ffffff; 
                    --text-color: #2C3E50;
                    --table-row-even: #F8F9F9; 
                    --table-row-odd: #FFFFFF; 
                }}
                body.dark-mode {{ 
                    --bg-color: #1a1a1a; 
                    --card-bg: #2d2d2d; 
                    --text-color: #ffffff; 
                    --table-row-even: #383838; 
                    --table-row-odd: #2d2d2d; 
                }}
                
                body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 20px; background: var(--bg-color); color: var(--text-color); margin: 0; transition: background 0.2s, color 0.2s; }}
                
                /* Header & KPI layout matching original code */
                .header-container {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 25px; width: 100%; }}
                .header-title {{ flex: 1; text-align: center; margin: 0; font-size: 28px; color: var(--text-color); }}
                
                .kpi-container {{ display: flex; gap: 20px; }}
                .kpi-card {{ background: var(--card-bg); padding: 15px 30px; border-radius: 8px; border: 1px solid #ccc; text-align: center; }}
                .kpi-card h3 {{ margin: 0; font-size: 14px; color: #7f8c8d; }}
                .kpi-card p {{ margin: 5px 0 0; font-size: 20px; font-weight: bold; color: var(--text-color); }}
                
                /* Tab Management */
                .tab {{ display: flex; justify-content: center; flex-wrap: wrap; background: #fff; padding: 10px; border-radius: 8px 8px 0 0; border: 1px solid #ccc; gap: 5px; }}
                .tab button {{ padding: 12px 24px; cursor: pointer; border: none; background: none; font-weight: bold; color: #566573; font-size: 15px; transition: 0.2s; border-bottom: 3px solid transparent; }}
                .tab button:hover {{ color: #2C3E50; background: #f0f3f4; border-radius: 4px; }}
                .tab button.active {{ background: #3498DB; color: white; border-radius: 4px; border-bottom: 3px solid #2980B9; }}
                
                .tabcontent {{ background: var(--card-bg); padding: 15px; min-height: 760px; border: 1px solid #ccc; border-top: none; border-radius: 0 0 8px 8px; width: 98%; margin: 0 auto; color: var(--text-color); }}
                
                /* High-priority Table Styles */
                .custom-table {{ width: 100% !important; border-collapse: collapse !important; }}
                .custom-table th {{ background-color: #2C3E50 !important; color: white !important; padding: 14px !important; text-align: left; }}
                .custom-table td {{ padding: 12px 14px !important; border-bottom: 1px solid #EAECEE !important; }}
                
                .custom-table tbody tr.even {{ background-color: var(--table-row-even) !important; }}
                .custom-table tbody tr.odd {{ background-color: var(--table-row-odd) !important; }}
                .custom-table tbody tr:hover td {{ background-color: #3498DB !important; color: white !important; cursor: pointer; }}
                
                /* DataTables Text Control elements */
                .dataTables_wrapper .dataTables_length, .dataTables_wrapper .dataTables_filter, 
                .dataTables_wrapper .dataTables_info, .dataTables_wrapper .dataTables_paginate {{ 
                    color: var(--text-color) !important; 
                }}
                
                #theme-toggle {{ cursor: pointer; padding: 10px 15px; border-radius: 6px; border: 1px solid #ccc; background: var(--card-bg); color: var(--text-color); font-weight: bold; }}
            </style>
        </head>
        <body>
            <div class="header-container">
                <button id="theme-toggle" onclick="toggleTheme()">Toggle Dark Mode</button>
                <h2 class="header-title">Portfolio Analytics Dashboard</h2>
                <div class="kpi-container">
                    <div class="kpi-card"><h3>Total Value</h3><p>€{total_val_fmt}</p></div>
                    <div class="kpi-card"><h3>Total PnL</h3><p>€{total_pnl_fmt}</p></div>
                </div>
            </div>
            
            <div class="tab">{buttons}</div>
            {html_divs}
        
            <script>
                $(document).ready(function() {{ 
                    $('.custom-table').DataTable();
                    // Set default initial light mode font color for security
                    $('.custom-table td').css('color', '#2C3E50');
                }});
                
                function toggleTheme() {{ 
                    var body = document.body;
                    body.classList.toggle('dark-mode'); 
                    
                    // Active JavaScript override to guarantee font color updates instantly
                    if (body.classList.contains('dark-mode')) {{
                        $('.custom-table td').css('color', '#ffffff');
                    }} else {{
                        $('.custom-table td').css('color', '#2C3E50');
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
    
    out_path = Path(chart_dir) / "Master_Dashboard.html"
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f: 
        f.write(master_html)
        
    print(f"SUCCESS: Master Dashboard generated at {out_path}")
    return out_path