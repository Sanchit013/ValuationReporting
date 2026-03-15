"""
FCFF (Free Cash Flow to Firm) Calculation Module
"""


def calc_fcff_ebit_method(ebit, d_and_a, tax_rate, capex, change_in_nwc):
    """
    Calculate Free Cash Flow to Firm (FCFF) using the EBIT approach.
    
    FCFF = EBIT * (1 - Tax Rate) + D&A - CapEx - Change in NWC
    
    Args:
        ebit (float): Earnings Before Interest and Taxes
        d_and_a (float): Depreciation & Amortization (takes absolute value)
        tax_rate (float): Corporate tax rate (as decimal, e.g., 0.25 for 25%)
        capex (float): Capital Expenditure (takes absolute value)
        change_in_nwc (float): Change in Net Working Capital
    
    Returns:
        float: Free Cash Flow to Firm (FCFF), rounded to 2 decimal places
    """
    # Calculate NOPAT (Net Operating Profit After Tax)
    nopat = ebit * (1 - tax_rate)
    
    # Add back D&A (use absolute value)
    add_back_da = abs(d_and_a)
    
    # Subtract CapEx (use absolute value)
    subtract_capex = abs(capex)
    
    # Calculate FCFF
    fcff = nopat + add_back_da - subtract_capex - change_in_nwc
    
    return round(fcff, 2)

def calc_fcff_net_income_method(net_income, tax_rate, net_int_expense, d_and_a, capex, change_in_nwc):
    """
    Calculate Free Cash Flow to Firm (FCFF) using the Net Income approach.
    
    FCFF = Net Income + Net Interest Expense * (1 - Tax Rate) + D&A - CapEx - Change in NWC
    
    Args:
        net_income (float): Net Income
        tax_rate (float): Corporate tax rate (as decimal, e.g., 0.25 for 25%)
        net_int_expense (float): Net Interest Expense
        d_and_a (float): Depreciation & Amortization (takes absolute value)
        capex (float): Capital Expenditure (takes absolute value)
        change_in_nwc (float): Change in Net Working Capital
    
    Returns:
        float: Free Cash Flow to Firm (FCFF), rounded to 2 decimal places
    """
    # Add back after-tax interest expense (unlevered, use absolute value)
    after_tax_int = abs(net_int_expense) * (1 - tax_rate)
    
    # Add back D&A (use absolute value)
    add_back_da = abs(d_and_a)
    
    # Subtract CapEx (use absolute value)
    subtract_capex = abs(capex)
    
    # Calculate FCFF
    fcff = net_income + after_tax_int + add_back_da - subtract_capex - change_in_nwc
    
    return round(fcff, 2)


if __name__ == "__main__":
    # Example usage for EBIT method
    ebit = 1000
    d_and_a = 100
    tax_rate = 0.25
    capex = 150
    change_in_nwc = 50
    
    fcff_ebit = calc_fcff_ebit_method(ebit, d_and_a, tax_rate, capex, change_in_nwc)
    print(f"FCFF (EBIT Method): {fcff_ebit}")
    
    # Example usage for Net Income method
    net_income = 750
    net_int_expense = 50
    
    fcff_ni = calc_fcff_net_income_method(net_income, tax_rate, net_int_expense, d_and_a, capex, change_in_nwc)
    print(f"FCFF (Net Income Method): {fcff_ni}")
