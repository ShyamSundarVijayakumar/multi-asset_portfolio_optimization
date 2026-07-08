import streamlit as st
import base64
from pathlib import Path
import pandas as pd
import yfinance as yf
from datetime import datetime
import importlib

from src.utils import (
    sanitize_portfolio_data, 
    purge_simulations, 
    check_is_cached, 
    update_cache_timestamp,
    read_clean_dataframe,
    P_PATH,
    M_PATH
)

portfolio_tracker = importlib.import_module("src.02_portfolio_tracker")
risk_analysis = importlib.import_module("src.03_risk_analysis")

st.set_page_config(layout="wide", page_title="Portfolio Control Center")

if "sanitized" not in st.session_state:
    sanitize_portfolio_data()
    st.session_state["sanitized"] = True

if "current_view" not in st.session_state:
    st.session_state["current_view"] = "MAIN"

def apply_background(image_path):
    if Path(image_path).exists():
        with open(image_path, "rb") as f:
            b64_str = base64.b64encode(f.read()).decode()
        st.markdown(f"""
            <style>
            .stApp {{
                background: linear-gradient(rgba(0,0,0,0.65), rgba(0,0,0,0.65)), url("data:image/jpeg;base64,{b64_str}");
                background-size: cover;
                background-position: center;
            }}
            </style>
        """, unsafe_allow_html=True)

apply_background("assets/background.jpg")

# --- SUB-PAGE VIEW ROUTER ---
if st.session_state["current_view"] == "PORTFOLIO_DASH":
    if st.button("⬅️ Back to Control Center"):
        purge_simulations()
        st.session_state["current_view"] = "MAIN"
        st.rerun()
    try:
        with open("data/charts/Master_Dashboard.html", "r", encoding="utf-8") as f:
            st.components.v1.html(f.read(), height=900, scrolling=True)
    except FileNotFoundError:
        st.error("Portfolio dashboard file not found. Run tracking calculations first.")

elif st.session_state["current_view"] == "RISK_DASH":
    if st.button("⬅️ Back to Control Center"):
        purge_simulations()
        st.session_state["current_view"] = "MAIN"
        st.rerun()
    try:
        with open("data/charts/Master_Risk_Dashboard.html", "r", encoding="utf-8") as f:
            st.components.v1.html(f.read(), height=900, scrolling=True)
    except FileNotFoundError:
        st.error("Risk dashboard report file not found. Run the analysis engine first.")

elif st.session_state["current_view"] == "MAIN":
    st.markdown("<h1 style='text-align: center; color: white; margin-bottom: 2rem;'>Portfolio Control Center</h1>", unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("### Core Analytics Engines")

        # --- CORE PORTFOLIO ENGINE (ONCE PER DAY CACHED) ---
        if st.button("Run Portfolio Tracking Dashboard", use_container_width=True):
            purge_simulations()
            dash_file = Path("data/charts/Master_Dashboard.html")
            
            if check_is_cached("portfolio_tracker") and dash_file.exists():
                st.toast("🚀 Displaying cached portfolio layout for today...", icon="⚡")
            else:
                with st.status("Executing Analytics Sync Engine...", expanded=True) as status:
                    st.write("Executing Step 1: Base Portfolio Engine...")
                    portfolio_tracker.execute_portfolio_engine()
                    
                    st.write("Executing Step 2: Orchestrating ETF Updates...")
                    portfolio_tracker.run_update_orchestrator([])
                    
                    st.write("Executing Step 3: Look-Through Analysis...")
                    portfolio_tracker.generate_look_through()
                    
                    st.write("Executing Step 4: Generating HTML Dashboard...")
                    portfolio_tracker.generate_dash()
                    
                    status.update(label="Portfolio Calculations Completed!", state="complete")
                update_cache_timestamp("portfolio_tracker")
                    
            st.session_state["current_view"] = "PORTFOLIO_DASH"
            st.rerun()
            
        # --- CORE BASELINE RISK ENGINE (ONCE PER DAY CACHED) ---
        if st.button("Run Baseline Risk Analytics Dashboard", use_container_width=True):
            purge_simulations()
            risk_file = Path("data/charts/Master_Risk_Dashboard.html")
            
            if check_is_cached("risk_analytics") and risk_file.exists():
                st.toast("🚀 Displaying cached baseline risk matrix for today...", icon="⚡")
            else:
                with st.status("Running Risk Assessment Engine...", expanded=True) as status:
                    st.write("Processing Step 1: Synthesizing asset gaps...")
                    risk_analysis.step1_synthetic_generation()
                    st.write("Processing Step 2: Querying historical returns & FX Sync...")
                    risk_analysis.step2_fx_sync()
                    st.write("Processing Step 3: Indexing multi-asset alignment...")
                    risk_analysis.step3_data_consolidation()
                    st.write("Processing Step 4: Structuring benchmark indexes...")
                    risk_analysis.step4_index_construction()
                    st.write("Processing Step 5: Finalizing report generation...")
                    risk_analysis.step5_generate_dashboard()
                    
                    status.update(label="Risk Matrix Compiled!", state="complete")
                update_cache_timestamp("risk_analytics")
                    
            st.session_state["current_view"] = "RISK_DASH"
            st.rerun()

    with col2:
        st.markdown("### Strategic Portfolio Simulations")
        search_term = st.text_input("Search Asset Name (Yahoo Finance Engine):", placeholder="e.g. Apple or Vanguard")
        
        if search_term:
            try:
                results = yf.Search(search_term, max_results=3).quotes
                options = {f"{r.get('shortname', 'Unknown')} ({r.get('symbol')})": r.get('symbol') for r in results if 'symbol' in r}
            except Exception:
                options = {}
                
            if options:
                selected_label = st.selectbox("Confirm Target Instrument Selection:", list(options.keys()))
                target_ticker = options[selected_label]
                
                asset_class = st.selectbox("Classification Vector:", ["Equity (Stock)", "ETF", "Cryptocurrency", "Commodity"])
                isin_code = st.text_input("Instrument Verification Identification Vector (ISIN):").strip()
                investment_capital = st.number_input("Target Allocation Target Metric (EUR):", min_value=250.0, step=100.0)
                
                # --- STRATEGIC SIMULATION ENGINE (ALWAYS RUNS DYNAMICALLY) ---
                if st.button("⚡ Run Comparative Risk Evaluation Matrix", use_container_width=True):
                    if P_PATH.exists() and M_PATH.exists():
                        df_p = read_clean_dataframe(P_PATH)
                        df_m = read_clean_dataframe(M_PATH)
                        
                        isin_clean = isin_code.strip()
                        name_clean = selected_label.strip()
                        
                        if isin_clean in df_p['ISIN'].astype(str).values or isin_clean in df_m['ISIN'].astype(str).values:
                            st.error(f"❌ Processing halted: ISIN '{isin_clean}' already exists.")
                        elif name_clean in df_p['Security Name'].astype(str).values or name_clean in df_m['Security Name'].astype(str).values:
                            st.error(f"❌ Processing halted: Security Name '{name_clean}' already exists.")
                        elif not isin_clean:
                            st.error("❌ Please enter a valid ISIN code before running the simulation.")
                        elif investment_capital <= 0:
                            st.error("❌ Allocation must be greater than 0.")
                        else:
                            with st.spinner(f"Fetching real-time market data for {target_ticker}..."):
                                try:
                                    ticker_obj = yf.Ticker(target_ticker)
                                    live_price = ticker_obj.fast_info.get('lastPrice', None)
                                    if live_price is None or live_price <= 0:
                                        hist = ticker_obj.history(period="1d")
                                        if not hist.empty:
                                            live_price = hist['Close'].iloc[-1]
                                        else:
                                            raise ValueError("No live data found.")
                                except Exception as e:
                                    st.error(f"❌ Failed to fetch market data: {e}")
                                    live_price = None
                                
                            if live_price and live_price > 0:
                                calculated_quantity = float(investment_capital / live_price)
                                
                                sim_p = pd.DataFrame([{
                                    "ISIN": isin_clean, 
                                    "Security Name": name_clean, 
                                    "Last Bought Date": datetime.today().strftime('%Y-%m-%d'),
                                    "Current Quantity": calculated_quantity, 
                                    "Average Buy Price (EUR)": float(live_price), 
                                    "Current Price (EUR)": float(live_price),
                                    "Native Currency": "EUR", 
                                    "Sector": "Simulated", "Industry": "Simulated", "Country": "Simulated", 
                                    "Unrealized PnL (EUR)": 0.0, 
                                    "Is_Simulation": True
                                }])
                                
                                sim_m = pd.DataFrame([{
                                    "ISIN": isin_clean, 
                                    "Security Name": name_clean, 
                                    "yfinance_tickters": target_ticker, 
                                    "asset_class": asset_class, 
                                    "Is_Simulation": True
                                }])
                                
                                pd.concat([df_p, sim_p], ignore_index=True).to_csv(P_PATH, index=False)
                                pd.concat([df_m, sim_m], ignore_index=True).to_csv(M_PATH, index=False)
                                
                                with st.status("Calculating Risk Analytics Simulation...", expanded=True) as status:
                                    st.write("Processing Step 1: Synthesizing asset gaps...")
                                    risk_analysis.step1_synthetic_generation()
                                    st.write("Processing Step 2: Querying historical returns...")
                                    risk_analysis.step2_fx_sync()
                                    st.write("Processing Step 3: Indexing multi-asset alignment...")
                                    risk_analysis.step3_data_consolidation()
                                    st.write("Processing Step 4: Structuring benchmark calculations...")
                                    risk_analysis.step4_index_construction()
                                    st.write("Processing Step 5: Finalizing reporting visualizations...")
                                    risk_analysis.step5_generate_dashboard()
                                    
                                    status.update(label="Simulation calculations verified!", state="complete")
                                
                                sanitize_portfolio_data()
                                st.session_state["current_view"] = "RISK_DASH"
                                st.rerun()

        # --- CLEAR ALL BUTTON ---
        st.markdown("---")
        if st.button("🧹 Clear All Simulation Data", use_container_width=True):
            if purge_simulations():
                st.success("Cleaned all simulation metrics successfully!")
                sanitize_portfolio_data()
                st.rerun()
            else:
                st.info("No active simulation rows found to clear.")