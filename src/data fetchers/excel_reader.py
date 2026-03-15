import sys
from pathlib import Path

import xlwings as xw
import pandas as pd
import time

_project_root = Path(__file__).resolve().parents[2]
if str(_project_root) not in sys.path:
    sys.path.insert(0, str(_project_root))

try:
    from src.dataset import DCFDataset
except ModuleNotFoundError:
    from dataset import DCFDataset

def open_inputs_workbook():
    """
    Opens the inputs.xlsx file using xlwings and ensures it is visible.
    """
    try:
        # app = xw.App(visible=True) is the default, but we ensure it here
        inputs_path = _project_root / "data" / "inputs.xlsx"
        if not inputs_path.exists():
            inputs_path = _project_root / "inputs.xlsx"
        wb = xw.Book(str(inputs_path))
        return wb
    except Exception as e:
        print(f"Error opening inputs.xlsx: {e}")
        return None
   
def open_workbook(file_path):
    """
    Opens any Excel file specified by the file_path using xlwings.
    """
    try:
        path = Path(file_path)
        if not path.is_absolute():
            data_path = _project_root / "data" / path
            root_path = _project_root / path
            if data_path.exists():
                path = data_path
            elif root_path.exists():
                path = root_path
        wb = xw.Book(str(path))
        return wb
    except Exception as e:
        print(f"Error opening {file_path}: {e}")
        return None
    
def close_inputs_workbook(wb):
    """
    Closes the provided xlwings workbook object.
    """
    try:
        if wb:
            wb.close()
            print("inputs.xlsx closed successfully.")
    except Exception as e:
        print(f"Error closing inputs.xlsx: {e}")

def extract_meta_data(wb, ds):
    """Populate dataset `ds` with values from the meta_data sheet."""
    sheet = wb.sheets['meta_data']
    ds.target_excel = sheet.range('E5').value

    # Extract integer from forecast period (e.g., "5 yrs" -> 5)
    raw_period = sheet.range('E6').value
    ds.forecast_period = int(''.join(filter(str.isdigit, str(raw_period)))) if raw_period else None

    # Figures stated in cell
    ds.figures_stated_in = sheet.range('E7').value

    # Flags read from Yes/No cells
    def yes_no_to_bool(val):
        if isinstance(val, str) and val.strip().lower() in ('yes', 'y', 'true', '1'):
            return True
        return False

    ds.wacc_calculated = yes_no_to_bool(sheet.range('E9').value)
    ds.Ke_calculated = yes_no_to_bool(sheet.range('E10').value)
    ds.exit_mult_calculated = yes_no_to_bool(sheet.range('E12').value)

    # Optional output file name (assumed in E14)
    ds.output_file_name = sheet.range('E14').value

def read_comp_tickers(wb, ds):
    """Populate ds.ticker_list from the Comp_Tickers sheet.

    Reads count in E6 and then values starting E9 downward.
    """
    sheet = wb.sheets['Comp_Tickers']
    try:
        num = int(sheet.range('E6').value)
    except Exception:
        num = 0
    ds.ticker_list = []
    row = 9
    for i in range(num):
        val = sheet.range(f'E{row + i}').value
        if val:
            ds.ticker_list.append(str(val))

def scenario_switch(wb, ds):
    """Loop through scenarios 1, 2, 3 and extract FCFF components for each scenario.
    
    For each scenario:
    1. Waits 0.5s, then sets scenario control cell to that scenario number
    2. Saves the workbook
    3. Waits 0.5s for Excel recalculation
    4. Extracts FCFF components (NI, D_A, Net_IntExp, EBITDA, EBIT, Capex, AR, AP)
       from the fcff_components sheet
    5. Stores the values in the corresponding dataset lists indexed by scenario
    """
    if not ds.target_excel:
        print("No target_excel specified. Cannot switch scenario.")
        return

    sheet = wb.sheets['Target_Variables']
    scenario_sheet = sheet.range('E6').value
    scenario_cell = sheet.range('F6').value

    if not scenario_sheet or not scenario_cell:
        print("Scenario control location not found in Target_Variables sheet.")
        return

    # Map of FCFF components to their row numbers in fcff_components sheet
    # Column E (sheet name) and Column F (cell reference) at these rows
    component_map = {
        'NI': 6,
        'D_A': 7,
        'Net_IntExp': 8,
        'EBITDA': 9,
        'EBIT': 10,
        'Capex': 11,
        'CA': 12,
        'CL': 13
    }
    
    # Order matters for dataset attributes
    component_order = ['NI', 'D_A', 'Net_IntExp', 'EBITDA', 'EBIT', 'Capex', 'CA', 'CL']
    
    # Open the target workbook once
    twb = open_workbook(ds.target_excel)
    if not twb:
        print(f"Could not open target workbook: {ds.target_excel}")
        return
    
    try:
        # Loop through scenarios 1-3, writing the value and saving each time
        for scenario_num in range(1, 4):
            # Change scenario control cell to this scenario number
            twb.sheets[scenario_sheet].range(scenario_cell).value = scenario_num
            print(f"Set scenario value to {scenario_num} at {scenario_sheet}!{scenario_cell}")
            
            # Save workbook to persist the change
            twb.save()
            print(f"Saved workbook for scenario {scenario_num}")
            
            # Wait for Excel to recalculate
            time.sleep(0.5)
            
            # Extract FCFF components for this scenario
            fcff_sheet = wb.sheets['fcff_components']
            scenario_idx = scenario_num - 1  # 0-indexed for dataset lists
            
            for component in component_order:
                row_num = component_map[component]
                
                # Read target sheet name from column E and starting cell from column F
                target_sheet_name = fcff_sheet.range(f'E{row_num}').value
                starting_cell = fcff_sheet.range(f'F{row_num}').value
                
                if not target_sheet_name or not starting_cell:
                    continue
                
                # Extract values from target workbook
                # Parse starting cell to get row and column
                import re
                match = re.match(r'([A-Z]+)(\d+)', starting_cell)
                if not match:
                    continue
                
                start_col = match.group(1)
                start_row = int(match.group(2))
                
                # Extract values going right (same row, next columns) for forecast_period length
                values = []
                for i in range(ds.forecast_period):
                    # Calculate column letter for each position
                    col_offset = i
                    # Convert column letter to number, add offset, convert back
                    col_num = sum((ord(c) - 64) * (26 ** idx) for idx, c in enumerate(reversed(start_col)))
                    new_col_num = col_num + col_offset
                    
                    # Convert number back to letter
                    new_col = ''
                    while new_col_num > 0:
                        new_col_num, remainder = divmod(new_col_num - 1, 26)
                        new_col = chr(65 + remainder) + new_col
                    
                    cell_ref = f'{new_col}{start_row}'
                    try:
                        val = twb.sheets[target_sheet_name].range(cell_ref).value
                        # Round to 2 decimal places
                        rounded_val = round(val, 2) if val is not None else 0
                        values.append(rounded_val)
                    except Exception as e:
                        values.append(0)
                
                # Store values in dataset based on component
                if component == 'NI':
                    ds.NI[scenario_idx] = values
                elif component == 'D_A':
                    ds.D_A[scenario_idx] = values
                elif component == 'Net_IntExp':
                    ds.Net_IntExp[scenario_idx] = values
                elif component == 'EBITDA':
                    ds.EBITDA[scenario_idx] = values
                elif component == 'EBIT':
                    ds.EBIT[scenario_idx] = values
                elif component == 'Capex':
                    ds.Capex[scenario_idx] = values
                elif component == 'CA':
                    ds.CA[scenario_idx] = values
                elif component == 'CL':
                    ds.CL[scenario_idx] = values


            # and store them in ds.NI[scenario_num-1], ds.D_A[scenario_num-1], etc.
    
    except Exception as e:
        print(f"Error in scenario_switch: {e}")
    finally:
        twb.close()

def extract_target_variable(wb, ds):
    """Read various target workbook values via Target_Variables sheet.

    Cells in the inputs workbook specify where to pull each value from the
    target workbook:
      E7/F7  -> ticker_sym
      E8/F8  -> share_price
      E9/F9  -> share_outstanding
      E11/F11 -> LA_EBITDA
      E12/F12 -> LA_NI
      E13/F13 -> LA_CA
      E14/F14 -> LA_CL
      E15/F15 -> Tax_Rate
    E16/F16 -> Cash_Equivalents

    The function opens the target workbook once and assigns the retrieved
    values to the corresponding attributes on `ds`.
    """
    sheet = wb.sheets['Target_Variables']
    # mapping attribute -> (input sheet cell, input cell reference)
    mapping = {
        'ticker_sym': ('E7', 'F7'),
        'share_price': ('E8', 'F8'),
        'share_outstanding': ('E9', 'F9'),
        'LA_EBITDA': ('E11', 'F11'),
        'LA_NI': ('E12', 'F12'),
        'LA_CA': ('E13', 'F13'),
        'LA_CL': ('E14', 'F14'),
        'Tax_Rate': ('E15', 'F15'),
        'Cash_Equivalents': ('E16', 'F16'),
        'TE': ('E17', 'F17'),
        'TSTD': ('E18', 'F18'),
        'TLTD': ('E19', 'F19'),
        'Growth_Rate': ('E21', 'F21'),
        'exit_mult': ('E22', 'F22'),
        'Beta': ('E24', 'F24'),
        'RFR': ('E25', 'F25'),
        'Market_Return': ('E26', 'F26'),
        'Ke': ('E27', 'F27'),
        'WACC': ('E29', 'F29'),
        'Kd': ('E28', 'F28'),
    }

    # default values
    for attr in mapping:
        setattr(ds, attr, None)

    if not ds.target_excel:
        return None

    twb = open_workbook(ds.target_excel)
    if not twb:
        return None

    for attr, (cell_sheet, cell_ref) in mapping.items():
        # only fetch exit_mult if the exit multiplier calculation flag is set
        if attr == 'exit_mult' and not ds.exit_mult_calculated:
            setattr(ds, attr, None)
            continue
        # only fetch WACC if wacc_calculated flag is set
        if attr == 'WACC' and not ds.wacc_calculated:
            setattr(ds, attr, None)
            continue
        # only fetch beta/RFR/Market_Return if both wacc_calculated and Ke_calculated are False
        if attr in ('Beta', 'RFR', 'Market_Return') and (ds.wacc_calculated or ds.Ke_calculated):
            setattr(ds, attr, None)
            continue
        # only fetch Ke value if wacc is not calculated and Ke_calculated flag is true
        if attr == 'Ke' and (ds.wacc_calculated or not ds.Ke_calculated):
            setattr(ds, attr, None)
            continue
        tgt_sheet = sheet.range(cell_sheet).value
        tgt_cell = sheet.range(cell_ref).value
        if tgt_sheet and tgt_cell:
            try:
                val = twb.sheets[tgt_sheet].range(tgt_cell).value
                setattr(ds, attr, val)
            except Exception as e:
                print(f"Error reading {tgt_sheet}!{tgt_cell} from {ds.target_excel}: {e}")

    twb.close()
    return ds




#example usage:
if __name__ == "__main__":
    # 1. open the inputs workbook
    wb = open_inputs_workbook()
    ds = DCFDataset()

    if not wb:
        print("Failed to open inputs.xlsx.")
    else:
        print("inputs.xlsx opened successfully.")

        # 2. extract metadata values into dataset
        extract_meta_data(wb, ds)
        # 3. read comparables ticker list
        read_comp_tickers(wb, ds)
        # 4. immediately extract target workbook variables
        ds = extract_target_variable(wb, ds)

        # 5. switch scenario to 3 in target workbook and extract FCFF components
        scenario_switch(wb, ds)
        
        # Print extracted scenario switch values
        print("\n--- FCFF Components by Scenario ---")
        print(f"Net Income (NI): Scenario 1: {ds.NI[0]}, Scenario 2: {ds.NI[1]}, Scenario 3: {ds.NI[2]}")
        print(f"Depreciation & Amortization (D_A): Scenario 1: {ds.D_A[0]}, Scenario 2: {ds.D_A[1]}, Scenario 3: {ds.D_A[2]}")
        print(f"Net Interest Expense (Net_IntExp): Scenario 1: {ds.Net_IntExp[0]}, Scenario 2: {ds.Net_IntExp[1]}, Scenario 3: {ds.Net_IntExp[2]}")
        print(f"EBITDA: Scenario 1: {ds.EBITDA[0]}, Scenario 2: {ds.EBITDA[1]}, Scenario 3: {ds.EBITDA[2]}")
        print(f"EBIT: Scenario 1: {ds.EBIT[0]}, Scenario 2: {ds.EBIT[1]}, Scenario 3: {ds.EBIT[2]}")
        print(f"Capital Expenditure (Capex): Scenario 1: {ds.Capex[0]}, Scenario 2: {ds.Capex[1]}, Scenario 3: {ds.Capex[2]}")
        print(f"Current Assets (CA): Scenario 1: {ds.CA[0]}, Scenario 2: {ds.CA[1]}, Scenario 3: {ds.CA[2]}")
        print(f"Current Liabilities (CL): Scenario 1: {ds.CL[0]}, Scenario 2: {ds.CL[1]}, Scenario 3: {ds.CL[2]}")

        # 6. open the target workbook (if specified)
        target_wb = None
        if ds.target_excel:
            target_wb = open_workbook(ds.target_excel)
            if target_wb:
                print(f"Successfully opened target workbook: {target_wb.name}")

        # 7. close the target workbook
        close_inputs_workbook(target_wb)

        # 8. close the inputs workbook
        close_inputs_workbook(wb)
