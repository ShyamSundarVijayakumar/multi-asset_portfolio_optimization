import numpy as np
import pandas as pd
import yfinance as yf
from pathlib import Path
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class PortfolioMarketSimulator:
    def __init__(self, input_csv_path: str, processed_dir: str,market_data_dir: str, start_date: str, end_date: str):
        self.input_csv_path = Path(input_csv_path)
        self.output_dir = Path(processed_dir) / "Asset class"
        self.market_data_dir = Path(market_data_dir)
        self.start_date = pd.to_datetime(start_date)
        self.end_date = pd.to_datetime(end_date)
        self.dates = pd.date_range(start=self.start_date, end=self.end_date, freq='D')
        
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.fx_series = None
        # Initialize modern random number generator for true stochastic modeling
        self.rng = np.random.default_rng()

    def _fetch_historical_fx(self, ticker: str = 'EURINR=X') -> pd.Series:
        if self.fx_series is not None:
            return self.fx_series
    
        logger.info(f"Checking historical FX data for {ticker}...")
        file_path = self.market_data_dir / f"{ticker}_history.csv"
    
        if file_path.exists():
            existing_df = pd.read_csv(file_path, index_col=0, parse_dates=True)
            existing_series = existing_df.iloc[:, 0]
            
            last_date = existing_series.index.max()
            
            if last_date >= self.end_date:
                logger.info("Local FX data is already up to date.")
                combined_series = existing_series
            else:
                # Fetch only the gap (from the day after last local record up to requested end_date)
                fetch_start = last_date + pd.Timedelta(days=1)
                logger.info(f"Local data ends at {last_date.strftime('%Y-%m-%d')}. Fetching gap from {fetch_start.strftime('%Y-%m-%d')} to {self.end_date.strftime('%Y-%m-%d')}...")
                
                new_data = self._download_yf_data(ticker, fetch_start, self.end_date)
                
                if not new_data.empty:
                    combined_series = pd.concat([existing_series, new_data]).drop_duplicates()
                else:
                    combined_series = existing_series
                    
                combined_series.sort_index(inplace=True)
                combined_series.to_frame(name='Close').to_csv(file_path)
        else:
            logger.info(f"No local FX cache found for {ticker}. Downloading full range from {self.start_date} to {self.end_date}...")
            combined_series = self._download_yf_data(ticker, self.start_date, self.end_date)
            combined_series.to_frame(name='Close').to_csv(file_path)
    
        if combined_series.index.duplicated().any():
            combined_series = combined_series[~combined_series.index.duplicated(keep='first')]
        
        series = combined_series.reindex(self.dates).ffill().bfill()
    
        if isinstance(series, pd.DataFrame):
            series = series.iloc[:, 0]
            
        self.fx_series = series
        return self.fx_series


    def _download_yf_data(self, ticker: str, start: pd.Timestamp, end: pd.Timestamp) -> pd.Series:
        """Helper method to download and extract the Close series from yfinance."""
        
        fx_data = yf.download(ticker, start=start, end=end + pd.Timedelta(days=1), progress=False)
        
        if fx_data.empty:
            logger.warning(f"No data returned from yfinance for ticker {ticker} between {start.strftime('%Y-%m-%d')} and {end.strftime('%Y-%m-%d')}. Relying on ffill.")
            return pd.Series(dtype=float)
                
        if isinstance(fx_data.columns, pd.MultiIndex):
            if 'Close' in fx_data.columns.levels[0]:
                close_series = fx_data['Close']
            else:
                raise KeyError(f"'Close' column not found in MultiIndex columns: {fx_data.columns.levels[0]}")
        else:
            if 'Close' in fx_data.columns:
                close_series = fx_data['Close']
            else:
                raise KeyError(f"'Close' column not found in columns: {fx_data.columns.tolist()}")
            
        if isinstance(close_series, pd.DataFrame):
            close_series = close_series.iloc[:, 0]
            
        return close_series


    def _generate_synthetic_path(self, annual_return: float, annual_volatility: float, 
                                 target_local_value: float, anchor_date: pd.Timestamp, 
                                 is_risk_free: bool = False) -> pd.Series:
        """
        Generates a synthetic daily asset price path using Geometric Brownian Motion (GBM).
        
        This implementation follows the discrete-time GBM framework outlined in Reddy & Clinton (2016,
        Australasian Accounting, Business and Finance Journal), specifically
        mirroring Equations (3) and (4) for daily step simulation. It is designed to generate realistic
        historical daily valuation data for assets (e.g., real estate) when only the annual appreciation
        rate (CAGR) and annual volatility are known.
        
        Parameters
        ----------
        annual_return : float
            The expected annual price appreciation rate, expressed as a percentage (e.g., 13.5 for 13.5%).
            **CRITICAL ASSUMPTION**: This value is treated as a Geometric Average (CAGR). 
            This is the standard way real estate indices and long-term appraisals quote returns.
        annual_volatility : float
            The expected annual price volatility, expressed as a percentage (e.g., 6.0 for 6%).
        target_local_value : float
            The exact known value of the asset in local currency on the anchor_date. The generated
            path will be forced to hit this exact value on that specific day.
        anchor_date : pd.Timestamp
            The specific valuation/purchase date where the path is pinned to the target_local_value.
        is_risk_free : bool, default=False
            If True, the stochastic (random) component is removed. The path grows purely 
            deterministically at the exact CAGR provided, simulating a risk-free asset (like for e.g Bonds).
            
        Returns
        -------
        pd.Series
            A time series of daily synthetic asset values (local currency) indexed from self.start_date 
            to self.end_date. The series exactly equals target_local_value on the anchor_date.
    
        Notes
        -----
        **Mathematical Derivation & Implementation Details**
        
        1.  **Time Discretization (dt)**:
            Since real estate does not trade on "business days" like stocks, we use a calendar-day
            convention: dt = 1 / 365.0.
        
        2.  **Handling the Input Return (Crucial Distinction)**:
            In the classic GBM formula (Reddy & Clinton, Eq. 4), the log-return is defined as:
                ln(S_{t+dt} / S_t) = (mu - sigma^2/2) * dt + sigma * epsilon * sqrt(dt)
            where `mu` is the **Arithmetic Mean** of the log-returns.
            
            However, our `annual_return` input is a **Geometric Mean (CAGR)**. 
            The mathematical relationship between CAGR and Arithmetic Mean (mu) is:
                CAGR = exp(mu - sigma^2/2) - 1  =>  (mu - sigma^2/2) = ln(1 + CAGR)
            
            Therefore, to ensure our simulated path perfectly targets the input CAGR (e.g., 13.5%)
            over the long run, we substitute the drift term as:
                Drift (log-return) = ln(1 + annual_return) * dt
            
            **Why we DON'T use `(annual_return - sigma^2/2) * dt`**:
            Doing so would imply that `annual_return` is the Arithmetic Mean (mu). This would
            artificially inflate the resulting CAGR to exp(annual_return - sigma^2/2) - 1, 
            overshooting the user's intended target by approximately (sigma^2 / 2) annually.
            
        3.  **Stochastic Shock (Volatility)**:
            The random component is implemented exactly as per standard GBM:
                Shock = sigma * epsilon * sqrt(dt)
            where epsilon is a standard normal random variable (mean=0, std=1), generated via
            `self.rng.normal(0, 1.0)`.
            
        4.  **Daily Growth Factor**:
            Combining the drift and shock, the daily multiplicative factor is:
                Growth_Factor = exp( Drift + Shock )
            This ensures the asset's log-returns are normally distributed, making the prices
            themselves log-normally distributed (bounded at zero, right-skewed), which is the
            standard assumption for asset prices, including real estate.
            
        5.  **Bidirectional Pinning (Forward & Backward Simulation)**:
            Standard GBM (Reddy & Clinton, Eq. 4) only simulates forward from a starting point.
            However, in portfolio backfilling, we need the path to EXACTLY equal the 
            `target_local_value` on the `anchor_date`.
            
            To achieve this, the implementation performs a "Brownian Bridge" style adjustment:
                - **Forward Pass**: Starting from the anchor date, we project into the future.
                    S_{t+1} = S_t * Growth_Factor_{t+1}
                - **Backward Pass**: Starting from the anchor date, we project into the past.
                    Since S_{t+1} = S_t * G_{t+1}, rearranging gives S_t = S_{t+1} / G_{t+1}
            
            This ensures the resulting time-series is continuous, preserves the statistical 
            properties of GBM, and strictly adheres to the known valuation at the anchor point.
            
        6.  **Risk-Free Logic**:
            If `is_risk_free` is True or volatility is 0, the shock vector is set to zero.
            The path grows purely exponentially via `exp(ln(1+CAGR)*dt)`, resulting in a smooth
            curve that exactly hits the CAGR target without any random deviations.
        """
       
        # Convert percentages to decimals
        r_annual = annual_return / 100.0
        vol_annual = annual_volatility / 100.0
        
        # Define daily time increment (dt) based on a calendar-year convention
        dt = 1.0 / 365.0
        num_days = len(self.dates)
        
        if is_risk_free or vol_annual == 0:
            # Risk-free: exact deterministic growth using the CAGR
            drift_vector = np.full(num_days, np.log(1 + r_annual) * dt)
            shock_vector = np.zeros(num_days)
        else:
            # Risky: log-return drift is exactly ln(1 + CAGR) (for all the formulars refer to Reddy & Clinton (2016) paper)
            # Since r_annual is a Geometric Mean, it intrinsically accounts for -(sigma^2)/2
            drift_constant = np.log(1 + r_annual) * dt
            drift_vector = np.full(num_days, drift_constant)
            
            # Stochastic shock component: sigma * epsilon * sqrt(dt)
            # self.rng.normal(0, 1) generates the random epsilon array
            epsilon = self.rng.normal(0, 1.0, num_days)
            shock_vector = vol_annual * epsilon * np.sqrt(dt)
            
        # Compute exponential daily growth factors: exp(drift + shock)
        growth_factors = np.exp(drift_vector + shock_vector)
        
        # Locate the index of the anchor date
        anchor_idx = self.dates.get_indexer([anchor_date], method='nearest')[0]
        local_values = np.zeros(num_days)
        
        # Step 1: Hard-pin the value on the specific purchase/valuation date
        local_values[anchor_idx] = target_local_value
        
        # Step 2: Forward simulation (from anchor date to the end of the time series)
        for i in range(anchor_idx + 1, num_days):
            local_values[i] = local_values[i - 1] * growth_factors[i]
            
        # Step 3: Backward simulation (from anchor date to the beginning of the time series)
        for i in range(anchor_idx - 1, -1, -1):
            # Since S_t = S_{t-1} * G_t, then S_{t-1} = S_t / G_t
            local_values[i] = local_values[i + 1] / growth_factors[i + 1]
            
        return pd.Series(local_values, index=self.dates)

    def _extract_value(self, row: pd.Series, possible_names: list, default_val: float) -> float:
        for name in possible_names:
            if name in row and pd.notna(row[name]):
                val = str(row[name]).strip().replace(',', '')
                if val != '':
                    try:
                        return float(val)
                    except ValueError:
                        continue
        return default_val

    def process_portfolio(self):
            try:
                df_input = pd.read_csv(self.input_csv_path)
                df_input.columns = df_input.columns.str.strip().str.lower()
            except Exception as e:
                logger.error(f"Could not read input CSV file: {e}")
                return
                
            fx_history = self._fetch_historical_fx()
    
            for index, row in df_input.iterrows():
                raw_id = str(row.get('isin', row.get('security_name', f'asset_{index}'))).strip().lower()
                if not ('real_estate' in raw_id or 'bonds' in raw_id):
                    continue
                
                currency = str(row.get('currency', '')).strip().upper()
                if currency != 'EUR':
                    raise ValueError(f"Asset '{row.get('isin')}' in file '{self.input_csv_path.name}' must be 'EUR'. Current currency: '{currency}'")
    
                try:
                    quantity = float(row['quantity'])
                    unit_price_eur = float(row['price'])
                    ann_return_local = float(row['annual_return'])
                    volatility_local = float(row['volatility_rate'])
                    fx_rate_flag = float(row['eurinr_rate'])
                except (ValueError, KeyError, TypeError) as e:
                    logger.error(f"Skipping row {index} in file '{self.input_csv_path.name}'due to data error: {e}")
                    continue
    
                asset_id = str(row.get('isin', f'asset_{index}'))
                safe_filename = str(asset_id).strip().replace(" ", "_").lower()
                total_invested_eur = quantity * unit_price_eur
                
                parsed_date = None
                for date_col in ['purchase_date', 'date', 'transaction_date']:
                    if date_col in row and pd.notna(row[date_col]):
                        try:
                            parsed_date = pd.to_datetime(row[date_col], dayfirst=True).normalize()
                            break
                        except: continue
                
                if parsed_date is None:
                    anchor_date = self.start_date
                else:
                    anchor_date = max(self.start_date, min(self.end_date, parsed_date))
    
                if fx_rate_flag > 0:
                    fx_on_purchase = fx_history.loc[anchor_date]
                    initial_local_principal = total_invested_eur * fx_on_purchase
                    
                    local_value_path = self._generate_synthetic_path(
                        ann_return_local, volatility_local, initial_local_principal, anchor_date
                    )
                    df_asset = pd.DataFrame({'Total_Holding_Value_Local': local_value_path}, index=self.dates)
                    df_asset['Exchange_Rate_EUR_INR'] = fx_history
                    df_asset['Total_Holding_Value_EUR'] = df_asset['Total_Holding_Value_Local'] / df_asset['Exchange_Rate_EUR_INR']
                else:
                    local_value_path = self._generate_synthetic_path(
                        ann_return_local, volatility_local, total_invested_eur, anchor_date
                    )
                    df_asset = pd.DataFrame({'Total_Holding_Value_Local': local_value_path}, index=self.dates)
                    df_asset['Total_Holding_Value_EUR'] = df_asset['Total_Holding_Value_Local']
    
                df_asset['Daily_Return_EUR'] = df_asset['Total_Holding_Value_EUR'].pct_change().fillna(0)
                df_asset.to_csv(self.output_dir / f"{safe_filename}.csv")
                logger.info(f"Successfully generated 5-year path for {safe_filename}")