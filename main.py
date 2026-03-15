from __future__ import annotations
from dataclasses import dataclass
from importlib.machinery import SourceFileLoader
from importlib.util import module_from_spec, spec_from_loader
from pathlib import Path


@dataclass
class DatasetLoadResult:
	dataset: object
	inputs_workbook_opened: bool


def _load_excel_reader_module():
	root = Path(__file__).resolve().parent
	module_path = root / "src" / "data fetchers" / "excel_reader.py"
	loader = SourceFileLoader("excel_reader", str(module_path))
	spec = spec_from_loader(loader.name, loader)
	if spec is None:
		raise ImportError("Failed to create module spec for excel_reader")
	module = module_from_spec(spec)
	loader.exec_module(module)
	return module


def _load_dcf_engine_module():
	root = Path(__file__).resolve().parent
	module_path = root / "src" / "analytics" / "dcf_engine.py"
	loader = SourceFileLoader("dcf_engine", str(module_path))
	spec = spec_from_loader(loader.name, loader)
	if spec is None:
		raise ImportError("Failed to create module spec for dcf_engine")
	module = module_from_spec(spec)
	loader.exec_module(module)
	return module


def _load_fcff_module():
	root = Path(__file__).resolve().parent
	module_path = root / "src" / "analytics" / "fcff.py"
	loader = SourceFileLoader("fcff", str(module_path))
	spec = spec_from_loader(loader.name, loader)
	if spec is None:
		raise ImportError("Failed to create module spec for fcff")
	module = module_from_spec(spec)
	loader.exec_module(module)
	return module


def _load_comp_analysis_module():
	root = Path(__file__).resolve().parent
	module_path = root / "src" / "analytics" / "Comp_analysis_data.py"
	loader = SourceFileLoader("comp_analysis_data", str(module_path))
	spec = spec_from_loader(loader.name, loader)
	if spec is None:
		raise ImportError("Failed to create module spec for comp_analysis_data")
	module = module_from_spec(spec)
	loader.exec_module(module)
	return module


def _load_comp_valuation_module():
	root = Path(__file__).resolve().parent
	module_path = root / "src" / "analytics" / "comp_valuation.py"
	loader = SourceFileLoader("comp_valuation", str(module_path))
	spec = spec_from_loader(loader.name, loader)
	if spec is None:
		raise ImportError("Failed to create module spec for comp_valuation")
	module = module_from_spec(spec)
	loader.exec_module(module)
	return module

def _load_report_generator_module():
	root = Path(__file__).resolve().parent
	module_path = root / "src" / "report_generator.py"
	loader = SourceFileLoader("report_generator", str(module_path))
	spec = spec_from_loader(loader.name, loader)
	if spec is None:
		raise ImportError("Failed to create module spec for report_generator")
	module = module_from_spec(spec)
	loader.exec_module(module)
	return module

def _load_yfin_data_module():
	root = Path(__file__).resolve().parent
	module_path = root / "src" / "data fetchers" / "yfin_data.py"
	loader = SourceFileLoader("yfin_data", str(module_path))
	spec = spec_from_loader(loader.name, loader)
	if spec is None:
		raise ImportError("Failed to create module spec for yfin_data")
	module = module_from_spec(spec)
	loader.exec_module(module)
	return module

def load_dataset():
	"""Load the DCF dataset using the Excel input workbook."""
	excel_reader = _load_excel_reader_module()
	dataset = excel_reader.DCFDataset()
	wb = excel_reader.open_inputs_workbook()
	if not wb:
		return DatasetLoadResult(dataset=dataset, inputs_workbook_opened=False)

	excel_reader.extract_meta_data(wb, dataset)
	excel_reader.read_comp_tickers(wb, dataset)
	excel_reader.extract_target_variable(wb, dataset)
	excel_reader.scenario_switch(wb, dataset)

	excel_reader.close_inputs_workbook(wb)
	return DatasetLoadResult(dataset=dataset, inputs_workbook_opened=True)


if __name__ == "__main__":
	result = load_dataset()
	if not result.inputs_workbook_opened:
		print("inputs.xlsx failed to open.")
	else:
		dcf_engine = _load_dcf_engine_module()
		ds = result.dataset

		if not ds.wacc_calculated and not ds.Ke_calculated:
			ds.Ke = dcf_engine.calculate_cost_of_equity(ds.RFR, ds.Beta, ds.Market_Return)
			ds.Ke_calculated = True

		if not ds.wacc_calculated and ds.Ke_calculated:
			total_debt = abs(ds.TSTD or 0) + abs(ds.TLTD or 0)
			ds.WACC = dcf_engine.calculate_wacc(
				equity=ds.TE or 0,
				debt=total_debt,
				cost_of_equity=ds.Ke,
				cost_of_debt=ds.Kd,
				tax_rate=ds.Tax_Rate or 0,
			)
			ds.wacc_calculated = True

		# Calculate change in NWC for each scenario
		la_nwc = (ds.LA_CA or 0) - (ds.LA_CL or 0)
		for scenario_idx in range(3):
			prev_nwc = la_nwc
			for year_idx in range(len(ds.CA[scenario_idx])):
				current_nwc = ds.CA[scenario_idx][year_idx] - ds.CL[scenario_idx][year_idx]
				change_nwc = round(current_nwc - prev_nwc, 2)
				ds.chng_nwc[scenario_idx].append(change_nwc)
				prev_nwc = current_nwc

		# Calculate FCFF for each scenario using both methods
		fcff_module = _load_fcff_module()
		for scenario_idx in range(3):
			for year_idx in range(len(ds.EBIT[scenario_idx])):
				# FCFF using EBIT method
				fcff_ebit = fcff_module.calc_fcff_ebit_method(
					ebit=ds.EBIT[scenario_idx][year_idx],
					d_and_a=ds.D_A[scenario_idx][year_idx],
					tax_rate=ds.Tax_Rate or 0,
					capex=ds.Capex[scenario_idx][year_idx],
					change_in_nwc=ds.chng_nwc[scenario_idx][year_idx]
				)
				
				# FCFF using Net Income method
				fcff_ni = fcff_module.calc_fcff_net_income_method(
					net_income=ds.NI[scenario_idx][year_idx],
					tax_rate=ds.Tax_Rate or 0,
					net_int_expense=ds.Net_IntExp[scenario_idx][year_idx],
					d_and_a=ds.D_A[scenario_idx][year_idx],
					capex=ds.Capex[scenario_idx][year_idx],
					change_in_nwc=ds.chng_nwc[scenario_idx][year_idx]
				)
				
				# Compare and store
				if fcff_ebit == fcff_ni:
					ds.fcff[scenario_idx].append(fcff_ebit)
				else:
					print(f"WARNING: FCFF mismatch in Scenario {scenario_idx + 1}, Year {year_idx + 1}. EBIT method: {fcff_ebit}, Net Income method: {fcff_ni}. Using EBIT method value.")
					ds.fcff[scenario_idx].append(fcff_ebit)

		# Calculate Terminal Value for each scenario
		for scenario_idx in range(3):
			if ds.fcff[scenario_idx]:  # Check if fcff list is not empty
				last_fcff = ds.fcff[scenario_idx][-1]
				tv = dcf_engine.calculate_terminal_value(
					last_ufcf=last_fcff,
					growth_rate=ds.Growth_Rate or 0,
					wacc=ds.WACC
				)
				ds.terminal_value[scenario_idx] = tv

		# Calculate Enterprise Value for each scenario
		for scenario_idx in range(3):
			if ds.fcff[scenario_idx] and ds.terminal_value[scenario_idx] is not None:
				ev, pct_fcff, pct_tv = dcf_engine.calculate_enterprise_value(
					ufcf_list=ds.fcff[scenario_idx],
					terminal_value=ds.terminal_value[scenario_idx],
					wacc=ds.WACC
				)
				ds.enterprise_value[scenario_idx] = ev
				ds.pct_ev_from_fcff[scenario_idx] = pct_fcff
				ds.pct_ev_from_tv[scenario_idx] = pct_tv

		# Calculate Equity Value for each scenario
		total_debt = abs(ds.TSTD or 0) + abs(ds.TLTD or 0)
		cash = ds.Cash_Equivalents or 0
		for scenario_idx in range(3):
			if ds.enterprise_value[scenario_idx] is not None:
				eq_val = dcf_engine.calculate_equity_value(
					enterprise_value=ds.enterprise_value[scenario_idx],
					total_debt=total_debt,
					cash_and_equivalents=cash,
				)
				ds.equity_value[scenario_idx] = eq_val
				# Run comparable company analysis
		if ds.ticker_list:
			comp_analysis = _load_comp_analysis_module()
			ds.valuation_df, ds.summary_df, ds.comp_valid_tickers = comp_analysis.create_valuation_table(ds.ticker_list)

			if ds.summary_df is not None and "EV/LTM EBITDA" in ds.summary_df.columns:
				median_multiple = ds.summary_df.loc["Median", "EV/LTM EBITDA"]
				if median_multiple is not None:
					comp_valuation = _load_comp_valuation_module()
					net_debt = (abs(ds.TLTD or 0) + abs(ds.TSTD or 0)) - (ds.Cash_Equivalents or 0)
					comp_result = comp_valuation.comp_ebitda(
						ebitda=ds.LA_EBITDA or 0,
						net_debt=net_debt,
						shares_outstanding=ds.share_outstanding or 0,
						ebitda_multiple=median_multiple,
					)
					for comp_df in (ds.comp_best, ds.comp_base, ds.comp_bear):
						comp_df.loc["ev_ebitda_ltm", "share_price"] = comp_result["implied_share_price"]
						comp_df.loc["ev_ebitda_ltm", "equity_val"] = comp_result["equity_value"]
						comp_df.loc["ev_ebitda_ltm", "enterprise_val"] = comp_result["enterprise_value"]

					if not ds.exit_mult_calculated:
						ds.Exit_Mult = median_multiple
					ds.exit_mult_calculated = True

				median_multiple = ds.summary_df.loc["Median", "P/E (LTM)"]
				if median_multiple is not None:
					comp_valuation = _load_comp_valuation_module()
					net_debt = (abs(ds.TLTD or 0) + abs(ds.TSTD or 0)) - (ds.Cash_Equivalents or 0)
					comp_result = comp_valuation.comp_pe(
						net_income=ds.LA_NI or 0,
						net_debt=net_debt,
						shares_outstanding=ds.share_outstanding or 0,
						pe_multiple=median_multiple,
					)
					for comp_df in (ds.comp_best, ds.comp_base, ds.comp_bear):
						comp_df.loc["pe_ltm", "share_price"] = comp_result["implied_share_price"]
						comp_df.loc["pe_ltm", "equity_val"] = comp_result["equity_value"]
						comp_df.loc["pe_ltm", "enterprise_val"] = comp_result["enterprise_value"]

			if ds.summary_df is not None and "EV/FY1 EBITDA" in ds.summary_df.columns:
				median_multiple = ds.summary_df.loc["Median", "EV/FY1 EBITDA"]
				if median_multiple is not None:
					comp_valuation = _load_comp_valuation_module()
					net_debt = (abs(ds.TLTD or 0) + abs(ds.TSTD or 0)) - (ds.Cash_Equivalents or 0)
					scenario_targets = (
						(ds.comp_best, ds.EBITDA[0]),
						(ds.comp_base, ds.EBITDA[1]),
						(ds.comp_bear, ds.EBITDA[2]),
					)
					for comp_df, scenario_ebitda in scenario_targets:
						ebitda_value = scenario_ebitda[0] if scenario_ebitda else 0
						comp_result = comp_valuation.comp_ebitda(
							ebitda=ebitda_value,
							net_debt=net_debt,
							shares_outstanding=ds.share_outstanding or 0,
							ebitda_multiple=median_multiple,
						)
						comp_df.loc["ev_ebidta_fy", "share_price"] = comp_result["implied_share_price"]
						comp_df.loc["ev_ebidta_fy", "equity_val"] = comp_result["equity_value"]
						comp_df.loc["ev_ebidta_fy", "enterprise_val"] = comp_result["enterprise_value"]

			if ds.summary_df is not None and "P/E (FY1)" in ds.summary_df.columns:
				median_multiple = ds.summary_df.loc["Median", "P/E (FY1)"]
				if median_multiple is not None:
					comp_valuation = _load_comp_valuation_module()
					net_debt = (abs(ds.TLTD or 0) + abs(ds.TSTD or 0)) - (ds.Cash_Equivalents or 0)
					scenario_targets = (
						(ds.comp_best, ds.NI[0]),
						(ds.comp_base, ds.NI[1]),
						(ds.comp_bear, ds.NI[2]),
					)
					for comp_df, scenario_ni in scenario_targets:
						net_income_value = scenario_ni[0] if scenario_ni else 0
						comp_result = comp_valuation.comp_pe(
							net_income=net_income_value,
							net_debt=net_debt,
							shares_outstanding=ds.share_outstanding or 0,
							pe_multiple=median_multiple,
						)
						comp_df.loc["pe_fy", "share_price"] = comp_result["implied_share_price"]
						comp_df.loc["pe_fy", "equity_val"] = comp_result["equity_value"]
						comp_df.loc["pe_fy", "enterprise_val"] = comp_result["enterprise_value"]

			if ds.Exit_Mult is not None:
				for scenario_idx in range(3):
					scenario_ebitda = ds.EBITDA[scenario_idx]
					if scenario_ebitda:
						last_ebitda = scenario_ebitda[-1]

						tv_exit_mult = dcf_engine.calculate_terminal_value_exit_multiplier(
							fy1_ebitda=last_ebitda,
							exit_multiplier=ds.Exit_Mult,
						)
						ds.terminal_value_exit_mult[scenario_idx] = tv_exit_mult
						
			# Calculate Enterprise Value (exit multiple method) for each scenario
			for scenario_idx in range(3):
				if ds.fcff[scenario_idx] and ds.terminal_value_exit_mult[scenario_idx] is not None:
					ev, pct_fcff, pct_tv = dcf_engine.calculate_enterprise_value(
						ufcf_list=ds.fcff[scenario_idx],
						terminal_value=ds.terminal_value_exit_mult[scenario_idx],
						wacc=ds.WACC
					)
					ds.exit_mult_enterprise_val[scenario_idx] = ev
					ds.exit_mult_pct_ev_from_fcff[scenario_idx] = pct_fcff
					ds.exit_mult_pct_ev_from_tv[scenario_idx] = pct_tv

			# Calculate Equity Value (exit multiple method) for each scenario
			for scenario_idx in range(3):
				if ds.exit_mult_enterprise_val[scenario_idx] is not None:
					eq_val = dcf_engine.calculate_equity_value(
						enterprise_value=ds.exit_mult_enterprise_val[scenario_idx],
						total_debt=total_debt,
						cash_and_equivalents=cash,
					)
					ds.exit_mult_equity_val[scenario_idx] = eq_val

			# Calculate Implied Share Price (DCF method) for each scenario
			for scenario_idx in range(3):
				if ds.equity_value[scenario_idx] is not None:
					implied_share_price = ds.equity_value[scenario_idx] / (ds.share_outstanding or 1)
					ds.dcf_sharePrice[scenario_idx] = round(implied_share_price, 2)

			# Calculate Implied Share Price (exit multiple method) for each scenario
			for scenario_idx in range(3):
				if ds.exit_mult_equity_val[scenario_idx] is not None:
					implied_share_price = ds.exit_mult_equity_val[scenario_idx] / (ds.share_outstanding or 1)
					ds.dcf_exitMult_sharePrice[scenario_idx] = round(implied_share_price, 2)

		# Sensitivity Analysis for DCF method (perpetuity growth) for each scenario
		scenario_names = ["best", "base", "bear"]
		for scenario_idx in range(3):
			if ds.fcff[scenario_idx] and ds.terminal_value[scenario_idx] is not None:
				last_fcff = ds.fcff[scenario_idx][-1]
				ev_results, eq_results, share_price_results = dcf_engine.sensitivity_analysis(
					last_ufcf=last_fcff,
					ufcf_list=ds.fcff[scenario_idx],
					base_wacc=ds.WACC,
					base_growth=ds.Growth_Rate or 0,
					total_debt=total_debt,
					cash=cash,
					shares=ds.share_outstanding or 1
				)
				
				# Store results based on scenario
				if scenario_idx == 0:
					ds.best_ev_results = ev_results
					ds.best_eq_results = eq_results
					ds.best_share_price_results = share_price_results
				elif scenario_idx == 1:
					ds.base_ev_results = ev_results
					ds.base_eq_results = eq_results
					ds.base_share_price_results = share_price_results
				else:
					ds.bear_ev_results = ev_results
					ds.bear_eq_results = eq_results
					ds.bear_share_price_results = share_price_results

		# Sensitivity Analysis for Exit Multiple method for each scenario
		for scenario_idx in range(3):
			if ds.fcff[scenario_idx] and ds.terminal_value_exit_mult[scenario_idx] is not None:
				last_ebitda = ds.EBITDA[scenario_idx][-1] if ds.EBITDA[scenario_idx] else 0
				ev_results, eq_results, share_price_results = dcf_engine.sensitivity_analysis_exit_multiplier(
					fy1_ebitda=last_ebitda,
					ufcf_list=ds.fcff[scenario_idx],
					base_wacc=ds.WACC,
					base_exit_mult=ds.Exit_Mult,
					total_debt=total_debt,
					cash=cash,
					shares=ds.share_outstanding or 1
				)
				
				# Store results based on scenario
				if scenario_idx == 0:
					ds.best_ev_results_exitMult = ev_results
					ds.best_eq_results_exitMult = eq_results
					ds.best_share_price_results_exitMult = share_price_results
				elif scenario_idx == 1:
					ds.base_ev_results_exitMult = ev_results
					ds.base_eq_results_exitMult = eq_results
					ds.base_share_price_results_exitMult = share_price_results
				else:
					ds.bear_ev_results_exitMult = ev_results
					ds.bear_eq_results_exitMult = eq_results
					ds.bear_share_price_results_exitMult = share_price_results
		
		# Fetch company names from yfinance
		yfin_data = _load_yfin_data_module()
		
		# Get target company name
		if ds.ticker_sym:
			ds.company_name = yfin_data.get_company_name(ds.ticker_sym)
			print(f"Target Company: {ds.company_name} ({ds.ticker_sym})")
		
		# Get comparable company names
		if ds.comp_valid_tickers:
			for ticker in ds.comp_valid_tickers:
				company_name = yfin_data.get_company_name(ticker)
				ds.company_name_list.append(company_name)
			print(f"Comparable Companies: {len(ds.company_name_list)} names fetched")
		
		# Generate PDF Report
		report_generator = _load_report_generator_module()
		doc, elements, file_path = report_generator.create_pdf(
			document_name=ds.output_file_name or "Valuation_Report",
			title=f"{ds.ticker_sym} Valuation Report" if ds.ticker_sym else "Valuation Report",
			author="DCF Model"
		)
		cover_page_callback = report_generator.add_cover_page(
			doc=doc,
			elements=elements,
			title="Financial Valuation Report",
			subtitle=f"{ds.company_name} ({ds.ticker_sym})" if ds.company_name and ds.ticker_sym else "Prepared by DCF Model"
		)
		
		# Add Executive Summary Page
		def _fmt_price(value):
			try:
				if value is None:
					return "$0.00"
				num = float(value)
				if num != num:
					return "$0.00"
				return f"${num:,.2f}"
			except (TypeError, ValueError):
				return "$0.00"
		
		intrinsic_valuation_data = {
			'Terminal Growth': {
				'Bull Case': _fmt_price(ds.dcf_sharePrice[0]),
				'Base Case': _fmt_price(ds.dcf_sharePrice[1]),
				'Bear Case': _fmt_price(ds.dcf_sharePrice[2])
			},
			'Exit Multiple': {
				'Bull Case': _fmt_price(ds.dcf_exitMult_sharePrice[0]),
				'Base Case': _fmt_price(ds.dcf_exitMult_sharePrice[1]),
				'Bear Case': _fmt_price(ds.dcf_exitMult_sharePrice[2])
			}
		}
		
		relative_valuation_data = {
			'EV/EBITDA LTM': {
				'Bull Case': _fmt_price(ds.comp_best.loc['ev_ebitda_ltm', 'share_price']),
				'Base Case': _fmt_price(ds.comp_base.loc['ev_ebitda_ltm', 'share_price']),
				'Bear Case': _fmt_price(ds.comp_bear.loc['ev_ebitda_ltm', 'share_price'])
			},
			'EV/EBITDA FY': {
				'Bull Case': _fmt_price(ds.comp_best.loc['ev_ebidta_fy', 'share_price']),
				'Base Case': _fmt_price(ds.comp_base.loc['ev_ebidta_fy', 'share_price']),
				'Bear Case': _fmt_price(ds.comp_bear.loc['ev_ebidta_fy', 'share_price'])
			},
			'P/E LTM': {
				'Bull Case': _fmt_price(ds.comp_best.loc['pe_ltm', 'share_price']),
				'Base Case': _fmt_price(ds.comp_base.loc['pe_ltm', 'share_price']),
				'Bear Case': _fmt_price(ds.comp_bear.loc['pe_ltm', 'share_price'])
			},
			'P/E FY': {
				'Bull Case': _fmt_price(ds.comp_best.loc['pe_fy', 'share_price']),
				'Base Case': _fmt_price(ds.comp_base.loc['pe_fy', 'share_price']),
				'Bear Case': _fmt_price(ds.comp_bear.loc['pe_fy', 'share_price'])
			}
		}
		
		report_generator.add_executive_summary_page(
			doc=doc,
			elements=elements,
			company_name=ds.company_name or "Company",
			ticker_symbol=ds.ticker_sym or "N/A",
			current_share_price=ds.share_price or 0,
			shares_outstanding=ds.share_outstanding or 0,
			market_capitalization=(ds.share_price or 0) * (ds.share_outstanding or 0),
			intrinsic_valuation_data=intrinsic_valuation_data,
			relative_valuation_data=relative_valuation_data,
			figures_stated_in=ds.figures_stated_in
		)
		
		report_generator.add_model_assumptions_page(
			doc=doc,
			elements=elements,
			forecast_period=ds.forecast_period,
			effective_tax_rate=ds.Tax_Rate,
			cost_of_equity=ds.Ke,
			beta=ds.Beta,
			risk_free_rate=ds.RFR,
			market_return=ds.Market_Return,
			cost_of_debt=ds.Kd,
			wacc=ds.WACC,
			terminal_growth_rate=ds.Growth_Rate,
			exit_multiple=ds.Exit_Mult 
		)
		
		base_idx = 1
		ebit_list = ds.EBIT[base_idx] or []
		da_list = ds.D_A[base_idx] or []
		capex_list = ds.Capex[base_idx] or []
		nwc_list = ds.chng_nwc[base_idx] or []
		fcff_list = ds.fcff[base_idx] or []
		year_count = min(
			len(ebit_list),
			len(da_list),
			len(capex_list),
			len(nwc_list),
			len(fcff_list)
		)
		forecast_years_data = []
		tax_rate = ds.Tax_Rate or 0
		for i in range(year_count):
			forecast_years_data.append({
				'year': i + 1,
				'ebit_1_minus_t': ebit_list[i] * (1 - tax_rate),
				'da': da_list[i],
				'capex': capex_list[i],
				'change_nwc': nwc_list[i],
				'fcff': fcff_list[i]
			})
		
		terminal_growth_summary = {
			'terminal_value': ds.terminal_value[base_idx],
			'enterprise_value': ds.enterprise_value[base_idx],
			'net_debt': (total_debt or 0) - (cash or 0),
			'total_equity_value': ds.equity_value[base_idx],
			'shares_outstanding': ds.share_outstanding,
			'implied_share_price': ds.dcf_sharePrice[base_idx]
		}
		
		terminal_ebitda = None
		if ds.EBITDA[base_idx]:
			terminal_ebitda = ds.EBITDA[base_idx][-1]
		
		exit_multiple_summary = {
			'terminal_ebitda': terminal_ebitda,
			'exit_multiple': ds.Exit_Mult,
			'terminal_value': ds.terminal_value_exit_mult[base_idx],
			'enterprise_value': ds.exit_mult_enterprise_val[base_idx],
			'net_debt': (total_debt or 0) - (cash or 0),
			'total_equity_value': ds.exit_mult_equity_val[base_idx],
			'shares_outstanding': ds.share_outstanding,
			'implied_share_price': ds.dcf_exitMult_sharePrice[base_idx]
		}
		
		def _build_sensitivity_table(result_dict, header_label, col_label_formatter):
			if not result_dict:
				return [[header_label]]
			wacc_values = sorted(result_dict.keys())
			first_row = result_dict[wacc_values[0]] if wacc_values else {}
			col_keys = sorted(first_row.keys()) if isinstance(first_row, dict) else []
			header = [header_label] + [col_label_formatter(k) for k in col_keys]
			table = [header]
			for w in wacc_values:
				row = [f"{w*100:.1f}%" if isinstance(w, (int, float)) else str(w)]
				row_dict = result_dict.get(w, {}) if isinstance(result_dict, dict) else {}
				for k in col_keys:
					val = row_dict.get(k, "N/A") if isinstance(row_dict, dict) else "N/A"
					if isinstance(val, (int, float)):
						row.append(f"${val:,.2f}")
					else:
						row.append(str(val))
				table.append(row)
			return table
		
		sensitivity_tgm = _build_sensitivity_table(
			ds.base_share_price_results,
			"WACC / TGR",
			lambda k: f"{k*100:.1f}%"
		)
		sensitivity_em = _build_sensitivity_table(
			ds.base_share_price_results_exitMult,
			"WACC / Multiple",
			lambda k: f"{k:.1f}x"
		)
		
		report_generator.add_intrinsic_valuation_page(
			doc=doc,
			elements=elements,
			forecast_years_data=forecast_years_data,
			terminal_growth_summary=terminal_growth_summary,
			exit_multiple_summary=exit_multiple_summary,
			sensitivity_tgm=sensitivity_tgm,
			sensitivity_em=sensitivity_em,
			base_wacc=ds.WACC,
			base_terminal_growth_rate=ds.Growth_Rate,
			base_exit_multiple=ds.Exit_Mult,
			figures_stated_in=ds.figures_stated_in or "millions"
		)
		
		# Add Relative Valuation Page
		if ds.valuation_df is not None and ds.summary_df is not None:
			# Build peer_data from valuation_df
			peer_data = []
			for i, ticker in enumerate(ds.comp_valid_tickers):
				company_name = ds.company_name_list[i] if i < len(ds.company_name_list) else ticker
				peer_dict = {
					'ticker': ticker,
					'company_name': company_name,
					'ev_ebitda_ltm': ds.valuation_df.iloc[i]['EV/LTM EBITDA'] if 'EV/LTM EBITDA' in ds.valuation_df.columns and i < len(ds.valuation_df) else None,
					'ev_ebitda_fy': ds.valuation_df.iloc[i]['EV/FY1 EBITDA'] if 'EV/FY1 EBITDA' in ds.valuation_df.columns and i < len(ds.valuation_df) else None,
					'pe_ltm': ds.valuation_df.iloc[i]['P/E (LTM)'] if 'P/E (LTM)' in ds.valuation_df.columns and i < len(ds.valuation_df) else None,
					'pe_fy': ds.valuation_df.iloc[i]['P/E (FY1)'] if 'P/E (FY1)' in ds.valuation_df.columns and i < len(ds.valuation_df) else None
				}
				peer_data.append(peer_dict)
			
			# Build summary_stats from summary_df
			def _safe_get_summary(row_label, col_label):
				"""Safely get summary statistic, trying common variations of row labels."""
				if col_label not in ds.summary_df.columns:
					return 0
				
				# Try different common variations of row labels
				row_variations = {
					'average': ['Average', 'Mean', 'Avg'],
					'median': ['Median', 'Med'],
					'highest': ['Highest', 'High', 'Max', 'Maximum'],
					'lowest': ['Lowest', 'Low', 'Min', 'Minimum']
				}
				
				for variant in row_variations.get(row_label.lower(), [row_label]):
					if variant in ds.summary_df.index:
						return ds.summary_df.loc[variant, col_label]
				
				return 0
			
			summary_stats = {
				'ev_ebitda_ltm': {
					'average': _safe_get_summary('average', 'EV/LTM EBITDA'),
					'median': _safe_get_summary('median', 'EV/LTM EBITDA'),
					'highest': _safe_get_summary('highest', 'EV/LTM EBITDA'),
					'lowest': _safe_get_summary('lowest', 'EV/LTM EBITDA')
				},
				'ev_ebitda_fy': {
					'average': _safe_get_summary('average', 'EV/FY1 EBITDA'),
					'median': _safe_get_summary('median', 'EV/FY1 EBITDA'),
					'highest': _safe_get_summary('highest', 'EV/FY1 EBITDA'),
					'lowest': _safe_get_summary('lowest', 'EV/FY1 EBITDA')
				},
				'pe_ltm': {
					'average': _safe_get_summary('average', 'P/E (LTM)'),
					'median': _safe_get_summary('median', 'P/E (LTM)'),
					'highest': _safe_get_summary('highest', 'P/E (LTM)'),
					'lowest': _safe_get_summary('lowest', 'P/E (LTM)')
				},
				'pe_fy': {
					'average': _safe_get_summary('average', 'P/E (FY1)'),
					'median': _safe_get_summary('median', 'P/E (FY1)'),
					'highest': _safe_get_summary('highest', 'P/E (FY1)'),
					'lowest': _safe_get_summary('lowest', 'P/E (FY1)')
				}
			}
			
			# Build implied_valuations from base scenario comp data
			implied_valuations = {
				'ev_ebitda_ltm': {
					'share_price': ds.comp_base.loc['ev_ebitda_ltm', 'share_price'],
					'equity_value': ds.comp_base.loc['ev_ebitda_ltm', 'equity_val'],
					'enterprise_value': ds.comp_base.loc['ev_ebitda_ltm', 'enterprise_val']
				},
				'ev_ebitda_fy': {
					'share_price': ds.comp_base.loc['ev_ebidta_fy', 'share_price'],
					'equity_value': ds.comp_base.loc['ev_ebidta_fy', 'equity_val'],
					'enterprise_value': ds.comp_base.loc['ev_ebidta_fy', 'enterprise_val']
				},
				'pe_ltm': {
					'share_price': ds.comp_base.loc['pe_ltm', 'share_price'],
					'equity_value': ds.comp_base.loc['pe_ltm', 'equity_val'],
					'enterprise_value': ds.comp_base.loc['pe_ltm', 'enterprise_val']
				},
				'pe_fy': {
					'share_price': ds.comp_base.loc['pe_fy', 'share_price'],
					'equity_value': ds.comp_base.loc['pe_fy', 'equity_val'],
					'enterprise_value': ds.comp_base.loc['pe_fy', 'enterprise_val']
				}
			}
			
			report_generator.add_relative_valuation_page(
				doc=doc,
				elements=elements,
				peer_data=peer_data,
				summary_stats=summary_stats,
				implied_valuations=implied_valuations,
				figures_stated_in=ds.figures_stated_in or "millions"
			)
		
		# Add Conclusion Page
		# Extract min/max from DCF Terminal Growth sensitivity table
		def _extract_min_max_from_sensitivity(sensitivity_dict):
			"""Extract minimum and maximum values from sensitivity analysis results."""
			if not sensitivity_dict:
				return None, None
			
			all_values = []
			for wacc_key, growth_dict in sensitivity_dict.items():
				if isinstance(growth_dict, dict):
					for growth_key, price in growth_dict.items():
						if isinstance(price, (int, float)) and price > 0:
							all_values.append(price)
			
			if all_values:
				return min(all_values), max(all_values)
			return None, None
		
		# Get min/max from Terminal Growth sensitivity table
		tgr_min, tgr_max = _extract_min_max_from_sensitivity(ds.base_share_price_results)
		
		# Get min/max from Exit Multiple sensitivity table
		em_min, em_max = _extract_min_max_from_sensitivity(ds.base_share_price_results_exitMult)
		
		# Build valuation summary
		valuation_summary = {
			'dcf_terminal_growth': {
				'low': tgr_min if tgr_min else ds.dcf_sharePrice[base_idx],
				'base': ds.dcf_sharePrice[base_idx],
				'high': tgr_max if tgr_max else ds.dcf_sharePrice[base_idx]
			},
			'dcf_exit_multiple': {
				'low': em_min if em_min else ds.dcf_exitMult_sharePrice[base_idx],
				'base': ds.dcf_exitMult_sharePrice[base_idx],
				'high': em_max if em_max else ds.dcf_exitMult_sharePrice[base_idx]
			},
			'ev_ebitda_ltm': ds.comp_base.loc['ev_ebitda_ltm', 'share_price'],
			'ev_ebitda_fy': ds.comp_base.loc['ev_ebidta_fy', 'share_price'],
			'pe_ltm': ds.comp_base.loc['pe_ltm', 'share_price'],
			'pe_fy': ds.comp_base.loc['pe_fy', 'share_price']
		}
		
		report_generator.add_conclusion_page(
			doc=doc,
			elements=elements,
			valuation_summary=valuation_summary,
			current_share_price=ds.share_price or 0
		)
		
		# Add Appendix A - Bull Case (Scenario 1)
		bull_idx = 0
		
		# Extract min/max from Bull Case sensitivity tables
		bull_tgr_min, bull_tgr_max = _extract_min_max_from_sensitivity(ds.best_share_price_results)
		bull_em_min, bull_em_max = _extract_min_max_from_sensitivity(ds.best_share_price_results_exitMult)
		
		# Build Bull Case forecast data
		bull_ebit_list = ds.EBIT[bull_idx] or []
		bull_da_list = ds.D_A[bull_idx] or []
		bull_capex_list = ds.Capex[bull_idx] or []
		bull_nwc_list = ds.chng_nwc[bull_idx] or []
		bull_fcff_list = ds.fcff[bull_idx] or []
		bull_year_count = min(
			len(bull_ebit_list),
			len(bull_da_list),
			len(bull_capex_list),
			len(bull_nwc_list),
			len(bull_fcff_list)
		)
		bull_forecast_years_data = []
		for i in range(bull_year_count):
			bull_forecast_years_data.append({
				'year': i + 1,
				'ebit_1_minus_t': bull_ebit_list[i] * (1 - tax_rate),
				'da': bull_da_list[i],
				'capex': bull_capex_list[i],
				'change_nwc': bull_nwc_list[i],
				'fcff': bull_fcff_list[i]
			})
		
		# Build Bull Case terminal growth summary
		bull_terminal_growth_summary = {
			'terminal_value': ds.terminal_value[bull_idx],
			'enterprise_value': ds.enterprise_value[bull_idx],
			'net_debt': (total_debt or 0) - (cash or 0),
			'total_equity_value': ds.equity_value[bull_idx],
			'shares_outstanding': ds.share_outstanding,
			'implied_share_price': ds.dcf_sharePrice[bull_idx]
		}
		
		# Build Bull Case exit multiple summary
		bull_terminal_ebitda = None
		if ds.EBITDA[bull_idx]:
			bull_terminal_ebitda = ds.EBITDA[bull_idx][-1]
		
		bull_exit_multiple_summary = {
			'terminal_ebitda': bull_terminal_ebitda,
			'exit_multiple': ds.Exit_Mult,
			'terminal_value': ds.terminal_value_exit_mult[bull_idx],
			'enterprise_value': ds.exit_mult_enterprise_val[bull_idx],
			'net_debt': (total_debt or 0) - (cash or 0),
			'total_equity_value': ds.exit_mult_equity_val[bull_idx],
			'shares_outstanding': ds.share_outstanding,
			'implied_share_price': ds.dcf_exitMult_sharePrice[bull_idx]
		}
		
		# Build Bull Case implied valuations from comp data
		bull_implied_valuations = {
			'ev_ebitda_ltm': {
				'share_price': ds.comp_best.loc['ev_ebitda_ltm', 'share_price'],
				'equity_value': ds.comp_best.loc['ev_ebitda_ltm', 'equity_val'],
				'enterprise_value': ds.comp_best.loc['ev_ebitda_ltm', 'enterprise_val']
			},
			'ev_ebitda_fy': {
				'share_price': ds.comp_best.loc['ev_ebidta_fy', 'share_price'],
				'equity_value': ds.comp_best.loc['ev_ebidta_fy', 'equity_val'],
				'enterprise_value': ds.comp_best.loc['ev_ebidta_fy', 'enterprise_val']
			},
			'pe_ltm': {
				'share_price': ds.comp_best.loc['pe_ltm', 'share_price'],
				'equity_value': ds.comp_best.loc['pe_ltm', 'equity_val'],
				'enterprise_value': ds.comp_best.loc['pe_ltm', 'enterprise_val']
			},
			'pe_fy': {
				'share_price': ds.comp_best.loc['pe_fy', 'share_price'],
				'equity_value': ds.comp_best.loc['pe_fy', 'equity_val'],
				'enterprise_value': ds.comp_best.loc['pe_fy', 'enterprise_val']
			}
		}
		
		# Build Bull Case valuation summary
		bull_valuation_summary = {
			'dcf_terminal_growth': {
				'low': bull_tgr_min if bull_tgr_min else ds.dcf_sharePrice[bull_idx],
				'base': ds.dcf_sharePrice[bull_idx],
				'high': bull_tgr_max if bull_tgr_max else ds.dcf_sharePrice[bull_idx]
			},
			'dcf_exit_multiple': {
				'low': bull_em_min if bull_em_min else ds.dcf_exitMult_sharePrice[bull_idx],
				'base': ds.dcf_exitMult_sharePrice[bull_idx],
				'high': bull_em_max if bull_em_max else ds.dcf_exitMult_sharePrice[bull_idx]
			},
			'ev_ebitda_ltm': ds.comp_best.loc['ev_ebitda_ltm', 'share_price'],
			'ev_ebitda_fy': ds.comp_best.loc['ev_ebidta_fy', 'share_price'],
			'pe_ltm': ds.comp_best.loc['pe_ltm', 'share_price'],
			'pe_fy': ds.comp_best.loc['pe_fy', 'share_price']
		}
		
		report_generator.appendix_a(
			doc=doc,
			elements=elements,
			forecast_years_data=bull_forecast_years_data,
			terminal_growth_summary=bull_terminal_growth_summary,
			exit_multiple_summary=bull_exit_multiple_summary,
			peer_data=peer_data,
			summary_stats=summary_stats,
			implied_valuations=bull_implied_valuations,
			valuation_summary=bull_valuation_summary,
			current_share_price=ds.share_price or 0,
			figures_stated_in=ds.figures_stated_in or "millions"
		)
		
		# Add Appendix B - Bear Case (Scenario 2)
		bear_idx = 2
		
		# Extract min/max from Bear Case sensitivity tables
		bear_tgr_min, bear_tgr_max = _extract_min_max_from_sensitivity(ds.bear_share_price_results)
		bear_em_min, bear_em_max = _extract_min_max_from_sensitivity(ds.bear_share_price_results_exitMult)
		
		# Build Bear Case forecast data
		bear_ebit_list = ds.EBIT[bear_idx] or []
		bear_da_list = ds.D_A[bear_idx] or []
		bear_capex_list = ds.Capex[bear_idx] or []
		bear_nwc_list = ds.chng_nwc[bear_idx] or []
		bear_fcff_list = ds.fcff[bear_idx] or []
		bear_year_count = min(
			len(bear_ebit_list),
			len(bear_da_list),
			len(bear_capex_list),
			len(bear_nwc_list),
			len(bear_fcff_list)
		)
		bear_forecast_years_data = []
		for i in range(bear_year_count):
			bear_forecast_years_data.append({
				'year': i + 1,
				'ebit_1_minus_t': bear_ebit_list[i] * (1 - tax_rate),
				'da': bear_da_list[i],
				'capex': bear_capex_list[i],
				'change_nwc': bear_nwc_list[i],
				'fcff': bear_fcff_list[i]
			})
		
		# Build Bear Case terminal growth summary
		bear_terminal_growth_summary = {
			'terminal_value': ds.terminal_value[bear_idx],
			'enterprise_value': ds.enterprise_value[bear_idx],
			'net_debt': (total_debt or 0) - (cash or 0),
			'total_equity_value': ds.equity_value[bear_idx],
			'shares_outstanding': ds.share_outstanding,
			'implied_share_price': ds.dcf_sharePrice[bear_idx]
		}
		
		# Build Bear Case exit multiple summary
		bear_terminal_ebitda = None
		if ds.EBITDA[bear_idx]:
			bear_terminal_ebitda = ds.EBITDA[bear_idx][-1]
		
		bear_exit_multiple_summary = {
			'terminal_ebitda': bear_terminal_ebitda,
			'exit_multiple': ds.Exit_Mult,
			'terminal_value': ds.terminal_value_exit_mult[bear_idx],
			'enterprise_value': ds.exit_mult_enterprise_val[bear_idx],
			'net_debt': (total_debt or 0) - (cash or 0),
			'total_equity_value': ds.exit_mult_equity_val[bear_idx],
			'shares_outstanding': ds.share_outstanding,
			'implied_share_price': ds.dcf_exitMult_sharePrice[bear_idx]
		}
		
		# Build Bear Case implied valuations from comp data
		bear_implied_valuations = {
			'ev_ebitda_ltm': {
				'share_price': ds.comp_bear.loc['ev_ebitda_ltm', 'share_price'],
				'equity_value': ds.comp_bear.loc['ev_ebitda_ltm', 'equity_val'],
				'enterprise_value': ds.comp_bear.loc['ev_ebitda_ltm', 'enterprise_val']
			},
			'ev_ebitda_fy': {
				'share_price': ds.comp_bear.loc['ev_ebidta_fy', 'share_price'],
				'equity_value': ds.comp_bear.loc['ev_ebidta_fy', 'equity_val'],
				'enterprise_value': ds.comp_bear.loc['ev_ebidta_fy', 'enterprise_val']
			},
			'pe_ltm': {
				'share_price': ds.comp_bear.loc['pe_ltm', 'share_price'],
				'equity_value': ds.comp_bear.loc['pe_ltm', 'equity_val'],
				'enterprise_value': ds.comp_bear.loc['pe_ltm', 'enterprise_val']
			},
			'pe_fy': {
				'share_price': ds.comp_bear.loc['pe_fy', 'share_price'],
				'equity_value': ds.comp_bear.loc['pe_fy', 'equity_val'],
				'enterprise_value': ds.comp_bear.loc['pe_fy', 'enterprise_val']
			}
		}
		
		# Build Bear Case valuation summary
		bear_valuation_summary = {
			'dcf_terminal_growth': {
				'low': bear_tgr_min if bear_tgr_min else ds.dcf_sharePrice[bear_idx],
				'base': ds.dcf_sharePrice[bear_idx],
				'high': bear_tgr_max if bear_tgr_max else ds.dcf_sharePrice[bear_idx]
			},
			'dcf_exit_multiple': {
				'low': bear_em_min if bear_em_min else ds.dcf_exitMult_sharePrice[bear_idx],
				'base': ds.dcf_exitMult_sharePrice[bear_idx],
				'high': bear_em_max if bear_em_max else ds.dcf_exitMult_sharePrice[bear_idx]
			},
			'ev_ebitda_ltm': ds.comp_bear.loc['ev_ebitda_ltm', 'share_price'],
			'ev_ebitda_fy': ds.comp_bear.loc['ev_ebidta_fy', 'share_price'],
			'pe_ltm': ds.comp_bear.loc['pe_ltm', 'share_price'],
			'pe_fy': ds.comp_bear.loc['pe_fy', 'share_price']
		}
		
		report_generator.appendix_b(
			doc=doc,
			elements=elements,
			forecast_years_data=bear_forecast_years_data,
			terminal_growth_summary=bear_terminal_growth_summary,
			exit_multiple_summary=bear_exit_multiple_summary,
			peer_data=peer_data,
			summary_stats=summary_stats,
			implied_valuations=bear_implied_valuations,
			valuation_summary=bear_valuation_summary,
			current_share_price=ds.share_price or 0,
			figures_stated_in=ds.figures_stated_in or "millions"
		)
		
		# Build the PDF
		doc.build(elements, onFirstPage=cover_page_callback)
		print(f"✓ PDF report successfully generated: {file_path}")


