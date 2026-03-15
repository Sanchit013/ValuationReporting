"""
DCF Model Dataset
"""

import pandas as pd

class DCFDataset:
    """
    An empty dataset class for the DCF model.
    """
    def __init__(self):
        # Input configuration variables
        self.target_excel = None
        self.forecast_period = None
        self.figures_stated_in = None
        
        # Stock and company variables
        self.ticker_sym = None
        self.company_name = None  # Target company name
        self.ticker_list = []  # List of comparable company tickers
        self.company_name_list = []  # List of comparable company names
        self.share_price = None
        self.share_outstanding = None
        
        # Last Actuals (LA)
        self.LA_EBITDA = None
        self.LA_NI = None
        self.LA_CA = None
        self.LA_CL = None
        
        # Debt and Equity
        self.Tax_Rate = None
        self.TE = None  # Total Equity
        self.TSTD = None  # Total Short Term Debt
        self.TLTD = None  # Total Long Term Debt
        self.Cash_Equivalents = None
        
        # DCF parameters
        self.Growth_Rate = None
        self.Exit_Mult = None
        
        # CAPM variables
        self.Beta = None
        self.RFR = None  # Risk Free Rate
        self.Market_Return = None
        
        # Calculated costs
        self.Ke = None  # Cost of Equity
        self.Kd = None  # Cost of Debt
        self.WACC = None  # Weighted Average Cost of Capital
        
        # Scenario variables (3 scenarios: Base, Bull, Bear or Scenario 1, 2, 3)
        # Each variable is a list of 3 lists, where each sub-list holds values for forecast periods
        # Structure: [Scenario1_values, Scenario2_values, Scenario3_values]
        self.NI = [[], [], []]  # Net Income
        self.D_A = [[], [], []]  # Depreciation & Amortization
        self.Net_IntExp = [[], [], []]  # Net Interest Expense
        self.EBITDA = [[], [], []]  # EBITDA
        self.EBIT = [[], [], []]  # EBIT
        self.Capex = [[], [], []]  # Capital Expenditure
        self.CA = [[], [], []]  # Current Assets
        self.CL = [[], [], []]  # Current Liabilities
        self.chng_nwc = [[], [], []]  # Change in Net Working Capital
        self.fcff = [[], [], []]  # Free Cash Flow to Firm
        
        # Terminal Value variables (one value per scenario)
        self.terminal_value = [None, None, None]  # Terminal Value (perpetuity growth method)
        self.terminal_value_exit_mult = [None, None, None]  # Terminal Value (exit multiple method)
        
        # Enterprise Value variables (one value per scenario)
        self.enterprise_value = [None, None, None]  # Enterprise Value
        self.pct_ev_from_fcff = [None, None, None]  # Percentage of EV from FCFF
        self.pct_ev_from_tv = [None, None, None]  # Percentage of EV from Terminal Value
        self.equity_value = [None, None, None]  # Equity Value
        
        # Exit Multiple Valuation variables (one value per scenario)
        self.exit_mult_enterprise_val = [None, None, None]  # Enterprise Value (exit multiple method)
        self.exit_mult_equity_val = [None, None, None]  # Equity Value (exit multiple method)
        self.exit_mult_pct_ev_from_fcff = [None, None, None]  # Percentage of EV from FCFF (exit multiple method)
        self.exit_mult_pct_ev_from_tv = [None, None, None]  # Percentage of EV from Terminal Value (exit multiple method)
        
        # Share Price variables (one value per scenario)
        self.dcf_sharePrice = [None, None, None]  # Implied Share Price (DCF method)
        self.dcf_exitMult_sharePrice = [None, None, None]  # Implied Share Price (exit multiple method)
        
        # Sensitivity Analysis variables (Enterprise Value, Equity Value, Share Price for each scenario)
        # Best scenario
        self.best_ev_results = {}  # Enterprise Value sensitivity table (best case)
        self.best_eq_results = {}  # Equity Value sensitivity table (best case)
        self.best_share_price_results = {}  # Share Price sensitivity table (best case)
        # Base scenario
        self.base_ev_results = {}  # Enterprise Value sensitivity table (base case)
        self.base_eq_results = {}  # Equity Value sensitivity table (base case)
        self.base_share_price_results = {}  # Share Price sensitivity table (base case)
        # Bear scenario
        self.bear_ev_results = {}  # Enterprise Value sensitivity table (bear case)
        self.bear_eq_results = {}  # Equity Value sensitivity table (bear case)
        self.bear_share_price_results = {}  # Share Price sensitivity table (bear case)
        
        # Sensitivity Analysis variables - Exit Multiple Method (Enterprise Value, Equity Value, Share Price for each scenario)
        # Best scenario
        self.best_ev_results_exitMult = {}  # Enterprise Value sensitivity table (best case, exit multiple)
        self.best_eq_results_exitMult = {}  # Equity Value sensitivity table (best case, exit multiple)
        self.best_share_price_results_exitMult = {}  # Share Price sensitivity table (best case, exit multiple)
        # Base scenario
        self.base_ev_results_exitMult = {}  # Enterprise Value sensitivity table (base case, exit multiple)
        self.base_eq_results_exitMult = {}  # Equity Value sensitivity table (base case, exit multiple)
        self.base_share_price_results_exitMult = {}  # Share Price sensitivity table (base case, exit multiple)
        # Bear scenario
        self.bear_ev_results_exitMult = {}  # Enterprise Value sensitivity table (bear case, exit multiple)
        self.bear_eq_results_exitMult = {}  # Equity Value sensitivity table (bear case, exit multiple)
        self.bear_share_price_results_exitMult = {}  # Share Price sensitivity table (bear case, exit multiple)
        
        # Comparable Company Analysis variables
        self.valuation_df = None  # Valuation table from comparable companies
        self.summary_df = None  # Summary statistics table
        self.comp_valid_tickers = []  # List of valid tickers used in comp analysis
        comp_index = [
            "ev_ebitda_ltm",
            "ev_ebidta_fy",
            "pe_ltm",
            "pe_fy",
        ]
        comp_columns = ["share_price", "equity_val", "enterprise_val"]
        self.comp_best = pd.DataFrame(index=comp_index, columns=comp_columns)
        self.comp_base = pd.DataFrame(index=comp_index, columns=comp_columns)
        self.comp_bear = pd.DataFrame(index=comp_index, columns=comp_columns)
        
        # Calculated variables (boolean flags)
        self.wacc_calculated = False
        self.Ke_calculated = False
        self.exit_mult_calculated = False
        
        # Output configuration
        self.output_file_name = None
