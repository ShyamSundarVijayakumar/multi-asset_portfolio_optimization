import os
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path

PROFESSIONAL_PALETTE = ['#2C3E50', '#18BC9C', '#3498DB', '#95A5A6', '#F39C12', '#8E44AD', '#34495E']
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
    
    # Convert dataframe directly to an HTML table with our custom CSS class
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
            textfont=dict(color="white"),
            hovertemplate="<b>%{label}</b><br>Value: €%{value:,.1f}<br>Share: %{percentParent:.2%}<extra></extra>"
        )
        fig_tree.update_layout(template=CHART_TEMPLATE, margin=dict(t=50, l=10, r=10, b=10), height=750)
        figures['Treemap'] = fig_tree
    else:
        print("Warning: Look-through data is empty after filtering positive values.")

    # --- 2. ASSET CLASS (Donut) ---
    df_asset_class = df_pos.groupby('Asset Class')['Total Value (EUR)'].sum().reset_index()
    fig_asset = px.pie(
        df_asset_class, values='Total Value (EUR)', names='Asset Class', hole=0.70,
        title='Asset Class Allocation Summary', color_discrete_sequence=PROFESSIONAL_PALETTE
    )
    fig_asset.update_traces(textinfo='percent+label', textposition='outside')
    fig_asset.update_layout(showlegend=False, title_x=0.5, template=CHART_TEMPLATE, height=600)
    figures['Asset_Class'] = fig_asset

    # --- 3. COUNTRY ALLOCATION (Look-Through Based) ---
    if 'Country' in df_look.columns:
        df_country = df_look.groupby('Country')[val_col].sum().reset_index().sort_values(by=val_col)
        
        fig_country = px.bar(
            df_country, x=val_col, y='Country', orientation='h',
            title='Geographic Capital Allocation (Underlying Holdings)', 
            color=val_col, 
            color_continuous_scale='Teal'
        )
        fig_country.update_layout(title_x=0.5, template=CHART_TEMPLATE, coloraxis_showscale=False, height=600)
        figures['Country_Bar'] = fig_country
    else:
        print("Note: 'Country' column not found in look-through data; skipping Geographic chart.")

    # --- 4. WATERFALL PNL ---
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
                increasing={"marker":{"color":"#18BC9C"}}, totals={"marker":{"color":"#2C3E50"}}
            ))
            fig_pnl.update_layout(title="Portfolio Capital Value Drivers", title_x=0.5, template=CHART_TEMPLATE, height=600)
            figures['Waterfall_PnL'] = fig_pnl
        except Exception as e:
            print(f"Metrics notation skip: {e}")

    # --- 5. YEARLY PNL SUMMARY ---
    try:
        # Load the sheet
        df_closed = pd.read_excel(file_summary, sheet_name='Closed Positions Log')
        
        # Convert Sell Date to datetime, extract year
        df_closed['Sell Date'] = pd.to_datetime(df_closed['Sell Date'])
        df_closed['Year'] = df_closed['Sell Date'].dt.year
        
        # Group by Year
        df_yearly = df_closed.groupby('Year')['Total Profit or Loss (EUR)'].sum().reset_index()
        
        # Create bar chart
        fig_yearly = px.bar(
            df_yearly, x='Year', y='Total Profit or Loss (EUR)',
            title="Realized Profit/Loss by Year",
            text_auto='.2s',
            color='Total Profit or Loss (EUR)',
            color_continuous_scale=['#E74C3C', '#18BC9C']
        )
        fig_yearly.update_layout(title_x=0.5, template=CHART_TEMPLATE, height=600)
        figures['Yearly_Performance'] = fig_yearly
    except Exception as e:
        print(f"Could not generate Yearly PnL chart: {e}")

    # --- 6. YEAR-OVER-YEAR DIVIDEND TRACKING ---
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

            fig_div = px.bar(
                df_div_filtered, x='Year', y='Value',
                title="Year-over-Year Dividend Income",
                text_auto='.2s',
                labels={'Value': 'Dividends (EUR)'},
                color_discrete_sequence=['#48C9B0']
            )
            fig_div.update_layout(title_x=0.5, template=CHART_TEMPLATE, height=600)
            figures['Dividend_Growth'] = fig_div
        else:
            print("Warning: 'Performance Metric' or 'Value' column missing in Dividend Log.")
    except Exception as e:
        print(f"Could not generate Dividend Tracking chart: {e}")
    
    # --- 7. NATIVE HTML DATA TABLES ---
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

    master_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Portfolio Master Dashboard</title>
        <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
        <style>
            body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 20px; background: #f4f6f9; margin: 0; }}
            h2 {{ color: #2C3E50; text-align: center; margin-bottom: 20px; font-size: 28px; }}
            .tab {{ display: flex; justify-content: center; background: #fff; padding: 10px; border-radius: 8px 8px 0 0; border: 1px solid #ccc; }}
            .tab button {{ padding: 12px 24px; cursor: pointer; border: none; background: none; font-weight: bold; color: #566573; font-size: 15px; transition: 0.2s; }}
            .tab button:hover {{ color: #2C3E50; background: #f0f3f4; border-radius: 4px; }}
            .tab button.active {{ background: #3498DB; color: white; border-radius: 4px; }}
            .tabcontent {{ background: white; padding: 15px; min-height: 760px; border: 1px solid #ccc; border-top: none; border-radius: 0 0 8px 8px; width: 98%; margin: 0 auto; }}
            .plotly-graph-div {{ width: 100% !important; height: 700px !important; }}
            
            /* Enhanced Interactive Table Styling */
            .table-container {{ height: 730px; overflow-y: auto; border: 1px solid #EAECEE; border-radius: 6px; }}
            .custom-table {{ width: 100%; border-collapse: collapse; font-size: 14px; text-align: left; }}
            .custom-table th {{ background-color: #2C3E50; color: white; padding: 14px; position: sticky; top: 0; z-index: 1; font-weight: 600; }}
            .custom-table td {{ padding: 12px 14px; border-bottom: 1px solid #EAECEE; color: #2C3E50; }}
            .custom-table tbody tr:nth-child(even) {{ background-color: #F8F9F9; }}
            .custom-table tbody tr:nth-child(odd) {{ background-color: #FFFFFF; }}
            
            /* Highlight row on hover */
            .custom-table tbody tr:hover {{ background-color: #D6EAF8; cursor: pointer; transition: background-color 0.2s ease; }}
        </style>
    </head>
    <body>
        <h2>Portfolio Analytics Dashboard</h2>
        <div class="tab">{buttons}</div>
        {html_divs}
        <script>
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