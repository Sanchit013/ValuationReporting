"""
Comparable Company Valuation Module
"""


def comp_ebitda(ebitda, net_debt, shares_outstanding, ebitda_multiple=None):
    """
    Calculate equity value and implied share price using comparable company multiples.
    
    Args:
        ebitda (float): EBITDA value (LTM or FY1)
        net_debt (float): Net Debt (Total Debt - Cash)
        shares_outstanding (float): Number of shares outstanding
        ebitda_multiple (float): EV/EBITDA median multiple from comps
    
    Returns:
        dict: Dictionary containing:
            - enterprise_value: Calculated enterprise value
            - equity_value: Calculated equity value
            - implied_share_price: Implied share price per share
    """
    # Validate multiple
    if ebitda_multiple is None:
        raise ValueError("ebitda_multiple is required")

    multiple = ebitda_multiple
    
    # Calculate Enterprise Value
    enterprise_value = ebitda * multiple
    
    # Calculate Equity Value: EV - Net Debt
    equity_value = enterprise_value - net_debt
    
    # Calculate Implied Share Price
    implied_share_price = equity_value / shares_outstanding
    
    return {
        "enterprise_value": round(enterprise_value, 2),
        "equity_value": round(equity_value, 2),
        "implied_share_price": round(implied_share_price, 2),
    }


def comp_pe(net_income, net_debt, shares_outstanding, pe_multiple=None):
    """
    Calculate equity value and implied share price using P/E multiples.
    
    Args:
        net_income (float): Net Income (LTM or FY1)
        net_debt (float): Net Debt (Total Debt - Cash)
        shares_outstanding (float): Number of shares outstanding
        pe_multiple (float): P/E median multiple from comps
    
    Returns:
        dict: Dictionary containing:
            - enterprise_value: Calculated enterprise value
            - equity_value: Calculated equity value
            - implied_share_price: Implied share price per share
    """
    # Validate multiple
    if pe_multiple is None:
        raise ValueError("pe_multiple is required")

    multiple = pe_multiple
    
    # Calculate Equity Value: P/E Multiple * Net Income
    equity_value = net_income * multiple
    
    # Calculate Implied Share Price
    implied_share_price = equity_value / shares_outstanding
    
    # Calculate Enterprise Value: Equity Value + Net Debt
    enterprise_value = equity_value + net_debt
    
    return {
        "enterprise_value": round(enterprise_value, 2),
        "equity_value": round(equity_value, 2),
        "implied_share_price": round(implied_share_price, 2),
    }


if __name__ == "__main__":
    # Example usage
    result_ltm = comp_ebitda(
        ebitda=1000,
        net_debt=500,
        shares_outstanding=100,
        ebitda_multiple=8.5,
    )
    print("LTM Valuation:", result_ltm)
    
    result_fy1 = comp_ebitda(
        ebitda=1200,
        net_debt=500,
        shares_outstanding=100,
        ebitda_multiple=9.0,
    )
    print("FY1 Valuation:", result_fy1)
    
    # P/E Multiple Valuation Examples
    result_pe_ltm = comp_pe(
        net_income=500,
        net_debt=300,
        shares_outstanding=100,
        pe_multiple=15.0,
    )
    print("P/E LTM Valuation:", result_pe_ltm)
    
    result_pe_fy1 = comp_pe(
        net_income=600,
        net_debt=300,
        shares_outstanding=100,
        pe_multiple=16.0,
    )
    print("P/E FY1 Valuation:", result_pe_fy1)
