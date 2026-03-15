import yfinance as yf

def check_ticker_exists(ticker_symbol):
    """
    Checks if a ticker symbol is valid and has data available.
    """
    ticker = yf.Ticker(ticker_symbol)
    try:
        # Attempt to fetch history to verify existence
        hist = ticker.history(period="1d")
        return not hist.empty
    except Exception:
        return False

def get_company_name(ticker_symbol):
    """
    Fetches the company name for a given ticker symbol.
    Returns the long name if available, otherwise short name, or None if not found.
    """
    try:
        ticker = yf.Ticker(ticker_symbol)
        info = ticker.info
        
        # Try to get long name first, fallback to short name
        company_name = info.get('longName') or info.get('shortName')
        return company_name
    except Exception as e:
        print(f"Error fetching company name for {ticker_symbol}: {e}")
        return None

def get_yfinance_metrics(ticker_symbol):
    """
    Fetches key valuation metrics for a given ticker using yfinance.
    Returns a dictionary with price, shares, EV, and LTM metrics.
    """
    ticker = yf.Ticker(ticker_symbol)
    info = ticker.info
    
    # Get quarterly income statement and sum last 4 quarters for LTM
    q_inc = ticker.quarterly_income_stmt
    ltm_revenue = q_inc.loc['Total Revenue'].iloc[:4].sum() if 'Total Revenue' in q_inc.index else None
    ltm_ebitda = q_inc.loc['EBITDA'].iloc[:4].sum() if 'EBITDA' in q_inc.index else None
    ltm_net_income = q_inc.loc['Net Income'].iloc[:4].sum() if 'Net Income' in q_inc.index else None

    # Fetching metrics
    metrics = {
        "Ticker": ticker_symbol.upper(),
        "Stock Price": info.get("currentPrice"),
        "Shares Outstanding": info.get("sharesOutstanding"),
        "Net Debt": info.get("totalDebt", 0) - info.get("totalCash", 0),
        "LTM EBITDA": "nm" if ltm_ebitda is not None and ltm_ebitda <= 0 else ltm_ebitda,
        "LTM Revenue": ltm_revenue,
        "LTM Net Income": "nm" if ltm_net_income is not None and ltm_net_income <= 0 else ltm_net_income
    }
    
    return metrics

# Example usage:
if __name__ == "__main__":
    ticker = "VZ"
    
    # Get company name
    company_name = get_company_name(ticker)
    print(f"Company Name: {company_name}\n")
    
    # Get stock data
    stock_data = get_yfinance_metrics(ticker)
    for key, value in stock_data.items():
        print(f"{key}: {value}")
