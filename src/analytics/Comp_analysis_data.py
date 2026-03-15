import sys
from pathlib import Path

import pandas as pd

_project_root = Path(__file__).resolve().parents[2]
_data_fetchers = _project_root / "src" / "data fetchers"
for path in (str(_project_root), str(_data_fetchers)):
    if path not in sys.path:
        sys.path.insert(0, path)

from fmp_data import get_fy1_estimates_fmp
from yfin_data import get_yfinance_metrics, check_ticker_exists

def create_valuation_table(tickers):
    """
    Combines LTM data from Yahoo Finance and FY1 Estimates from FMP
    into a single data table for analysis.
    """
    # Filter tickers to only valid ones
    valid_tickers = [t for t in tickers if check_ticker_exists(t)]
    
    table_data = []

    for ticker in valid_tickers:
        print(f"Fetching data for {ticker}...")
        try:
            # Get LTM and Market Data from yfinance
            yfin_metrics = get_yfinance_metrics(ticker)
            
            # Get Forward Estimates from FMP
            fmp_estimates = get_fy1_estimates_fmp(ticker)
            
            # Merge dictionaries
            combined_data = {**yfin_metrics, **fmp_estimates}
            table_data.append(combined_data)
        except Exception as e:
            print(f"Could not fetch data for {ticker}: {e}")

    # Create DataFrame
    df = pd.DataFrame(table_data)
    
    # Optional: Calculate basic multiples for analysis
    if not df.empty:
        # Calculate Market Cap and Enterprise Value
        df['Market Cap'] = df['Stock Price'] * df['Shares Outstanding']
        df['EV'] = df['Market Cap'] + df['Net Debt']
        
        # Helper to handle 'nm' (not meaningful) in calculations
        def safe_calc(numerator, denominator):
            if numerator == 'nm' or denominator == 'nm' or denominator <= 0 or denominator is None:
                return 'nm'
            return numerator / denominator

        # Helper to calculate EPS
        def safe_eps(net_income, shares):
            if net_income == 'nm' or shares is None or shares <= 0:
                return 'nm'
            return net_income / shares

        # Calculate Multiples
        df['EV/LTM EBITDA'] = df.apply(lambda x: safe_calc(x['EV'], x['LTM EBITDA']), axis=1)
        df['EV/FY1 EBITDA'] = df.apply(lambda x: safe_calc(x['EV'], x['fy1_ebitda']), axis=1)
        df['EV/LTM Revenue'] = df.apply(lambda x: safe_calc(x['EV'], x['LTM Revenue']), axis=1)
        df['EV/FY1 Revenue'] = df.apply(lambda x: safe_calc(x['EV'], x['fy1_revenue']), axis=1)
        df['P/E (LTM)'] = df.apply(lambda x: safe_calc(x['Stock Price'], safe_eps(x['LTM Net Income'], x['Shares Outstanding'])), axis=1)
        df['P/E (FY1)'] = df.apply(lambda x: safe_calc(x['Stock Price'], safe_eps(x['fy1_net_income'], x['Shares Outstanding'])), axis=1)

        # Convert large financial figures to Millions for readability
        cols_to_scale = ['Shares Outstanding', 'Net Debt', 'LTM EBITDA', 'LTM Revenue', 'LTM Net Income', 
                         'fy1_revenue', 'fy1_ebitda', 'fy1_net_income', 'EV', 'Market Cap']
        
        for col in cols_to_scale:
            if col in df.columns:
                df[col] = df[col].apply(lambda x: x / 1_000_000 if isinstance(x, (int, float)) and x != 'nm' else x)

    # Create Summary Statistics Table
    summary_cols = ['EV/LTM EBITDA', 'EV/FY1 EBITDA', 'P/E (LTM)', 'P/E (FY1)']
    summary_df = pd.DataFrame(index=['Average', 'Median', 'Min', 'Max'], columns=summary_cols)

    for col in summary_cols:
        if col in df.columns:
            clean_series = pd.to_numeric(df[col], errors='coerce').dropna()
            summary_df.loc['Average', col] = clean_series.mean()
            summary_df.loc['Median', col] = clean_series.median()
            summary_df.loc['Min', col] = clean_series.min()
            summary_df.loc['Max', col] = clean_series.max()

    return df, summary_df, valid_tickers





    