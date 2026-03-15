import sys
from pathlib import Path
from datetime import datetime

import requests

_project_root = Path(__file__).resolve().parents[2]
if str(_project_root) not in sys.path:
    sys.path.insert(0, str(_project_root))

try:
    from src.config import API_KEY, BASE_URL
except ModuleNotFoundError:
    from config import API_KEY, BASE_URL

def get_fy1_estimates_fmp(ticker):
    """
    Fetches FY1 estimates for Revenue, EBITDA, and Net Income from Financial Modeling Prep
    """
    url = f"{BASE_URL}/analyst-estimates?symbol={ticker}&period=annual&page=0&limit=10&apikey={API_KEY}"
    response = requests.get(url)
    
    if response.status_code != 200:
        raise Exception(f"API Error: {response.status_code}")
    
    data = response.json()
    
    if not data:
        return {"fy1_revenue": None, "fy1_ebitda": None, "fy1_net_income": None}

    # Filter for dates in the future and pick the one closest to today
    today = datetime.now().date()
    future_estimates = []
    for entry in data:
        estimate_date = datetime.strptime(entry.get("date"), "%Y-%m-%d").date()
        if estimate_date > today:
            future_estimates.append(entry)

    if not future_estimates:
        return {"fy1_revenue": None, "fy1_ebitda": None, "fy1_net_income": None}

    # Sort by date ascending and pick the first one (closest to today)
    fy1 = sorted(future_estimates, key=lambda x: x.get("date"))[0]

    fy1_ebitda = fy1.get("ebitdaAvg")
    fy1_net_income = fy1.get("netIncomeAvg")

    return {
        "fy1_revenue": fy1.get("revenueAvg"),
        "fy1_ebitda": "nm" if fy1_ebitda is not None and fy1_ebitda <= 0 else fy1_ebitda,
        "fy1_net_income": "nm" if fy1_net_income is not None and fy1_net_income <= 0 else fy1_net_income
    }

# Example usage:
if __name__ == "__main__":
    stock_data = get_fy1_estimates_fmp("AAPL")
    print(stock_data)