import pandas as pd

# This Function Calcs the Cost of Equity using CAPM
def calculate_cost_of_equity(risk_free_rate, beta, market_return):
    # CAPM formula: Re = Rf + β * (Rm - Rf)
    cost_of_equity = risk_free_rate + (beta * (market_return - risk_free_rate))
    cost_of_equity = round(cost_of_equity, 4)
    return cost_of_equity

# This Function Calcs the WACC
def calculate_wacc(equity, debt, cost_of_equity, cost_of_debt, tax_rate):
    total_capital = equity + debt
    weight_of_equity = equity / total_capital
    weight_of_debt = debt / total_capital
    
    # WACC formula: (E/V * Re) + (D/V * Rd * (1 - Tax))
    wacc = (weight_of_equity * cost_of_equity) + (weight_of_debt * cost_of_debt * (1 - tax_rate))
    wacc = round(wacc, 4)
    return wacc

# This Function Calcs the Terminal Value using the Perpetuity Growth Method
def calculate_terminal_value(last_ufcf, growth_rate, wacc):
    # Formula: TV = (Last UFCF * (1 + g)) / (WACC - g)
    terminal_value = (last_ufcf * (1 + growth_rate)) / (wacc - growth_rate)
    terminal_value = round(terminal_value, 2)
    return terminal_value

# This Function Calcs the Terminal Value using Exit Multiplier (EV/EBITDA)
def calculate_terminal_value_exit_multiplier(fy1_ebitda, exit_multiplier):
    # Formula: TV = FY1 EBITDA * Exit Multiplier
    # Exit multiplier is typically the mean EV/FY1 EBITDA from comparable companies
    if fy1_ebitda == 'nm' or exit_multiplier == 'nm' or exit_multiplier is None:
        return 'nm'
    terminal_value = fy1_ebitda * exit_multiplier
    terminal_value = round(terminal_value, 2)
    return terminal_value

# This Function Calcs the EV from UFCC, Termianal Value using WACC
def calculate_enterprise_value(ufcf_list, terminal_value, wacc):
    enterprise_value = 0
    total_pv_ufcf = 0
    
    # Discount each Unlevered Free Cash Flow to the present value
    for i, ufcf in enumerate(ufcf_list):
        year = i + 1
        pv_ufcf = ufcf / ((1 + wacc) ** year)
        total_pv_ufcf += pv_ufcf
        
    # Discount the Terminal Value to the present value (at the end of year n)
    pv_terminal_value = terminal_value / ((1 + wacc) ** len(ufcf_list))
    enterprise_value = total_pv_ufcf + pv_terminal_value
    
    percent_ev_from_ufcf = round(total_pv_ufcf / enterprise_value, 4)
    percent_ev_from_tv = round(pv_terminal_value / enterprise_value, 4)
    
    enterprise_value = round(enterprise_value, 2)
    
    return enterprise_value, percent_ev_from_ufcf, percent_ev_from_tv

# This Function Calcs the Equity Value from Enterprise Value
def calculate_equity_value(enterprise_value, total_debt, cash_and_equivalents):
    # Formula: Equity Value = Enterprise Value - Debt + Cash
    equity_value = enterprise_value - total_debt + cash_and_equivalents
    return round(equity_value, 2)

# This Function Performs Sensitivity Analysis on WACC and Growth Rate, then Returns 3 Tables: Enterprise Value, Equity Value, and Per Share Values
def sensitivity_analysis(last_ufcf, ufcf_list, base_wacc, base_growth, total_debt, cash, shares):
    # WACC: +/- 2% in 0.5% increments
    wacc_range = [round(base_wacc + i, 4) for i in [-0.02, -0.015, -0.01, -0.005, 0, 0.005, 0.01, 0.015, 0.02]]
    # Growth: +/- 1% in 0.5% increments
    growth_range = [round(base_growth + i, 4) for i in [-0.01, -0.005, 0, 0.005, 0.01]]
    
    ev_results = {}
    eq_results = {}
    share_price_results = {}
    
    for w in wacc_range:
        if w <= 0: continue
        ev_results[w] = {}
        eq_results[w] = {}
        share_price_results[w] = {}
        for g in growth_range:
            if g < 0 or g >= w: continue # g must be >= 0 and < wacc for perpetuity formula
            
            tv = calculate_terminal_value(last_ufcf, g, w)
            ev, _, _ = calculate_enterprise_value(ufcf_list, tv, w)
            eq_val = calculate_equity_value(ev, total_debt, cash)
            share_price = round(eq_val / shares, 2)
            
            ev_results[w][g] = round(ev, 2)
            eq_results[w][g] = round(eq_val, 2)
            share_price_results[w][g] = share_price

    return ev_results, eq_results, share_price_results

# This Function Performs Sensitivity Analysis using Exit Multiplier for Terminal Value, then Returns 3 Tables: Enterprise Value, Equity Value, and Per Share Values
def sensitivity_analysis_exit_multiplier(fy1_ebitda, ufcf_list, base_wacc, base_exit_mult, total_debt, cash, shares):
    # WACC: +/- 2% in 0.5% increments
    wacc_range = [round(base_wacc + i, 4) for i in [-0.02, -0.015, -0.01, -0.005, 0, 0.005, 0.01, 0.015, 0.02]]
    # Exit Multiplier: +/- 1x in 0.5x increments
    exit_mult_range = [round(base_exit_mult + i, 2) for i in [-1, -0.5, 0, 0.5, 1]]
    
    ev_results = {}
    eq_results = {}
    share_price_results = {}
    
    for w in wacc_range:
        if w <= 0: continue
        ev_results[w] = {}
        eq_results[w] = {}
        share_price_results[w] = {}
        for em in exit_mult_range:
            if em <= 0: continue # Exit multiplier must be positive
            
            tv = calculate_terminal_value_exit_multiplier(fy1_ebitda, em)
            if tv == 'nm': 
                ev_results[w][em] = 'N/A'
                eq_results[w][em] = 'N/A'
                share_price_results[w][em] = 'N/A'
                continue
            ev, _, _ = calculate_enterprise_value(ufcf_list, tv, w)
            eq_val = calculate_equity_value(ev, total_debt, cash)
            share_price = round(eq_val / shares, 2)
            
            ev_results[w][em] = round(ev, 2)
            eq_results[w][em] = round(eq_val, 2)
            share_price_results[w][em] = share_price
            
    return ev_results, eq_results, share_price_results


# Example Usage
"""
sensitivity_analysis(last_ufcf=519.76, ufcf_list=[355, 390.5, 429.55, 472.50, 519.76], base_wacc=0.10, base_growth=0.01, total_debt=5000, cash=10000, shares=1000)
wacc = calculate_wacc(equity=10000, debt=5000, cost_of_equity=0.10, cost_of_debt=0.05, tax_rate=0.25)
ev, pct_ufcf, pct_tv = calculate_enterprise_value([355, 390.5, 429.55, 472.50, 519.76], 8156.23, wacc)
print(ev)
print(wacc)
print(f"EV from UFCF: {pct_ufcf*100}%, EV from TV: {pct_tv*100}%")
"""