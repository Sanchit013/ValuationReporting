**Automated Equity Valuation Reporting Model**
**Author: **Sanchit Gupta****

**Overview**

This project is a Python-based automated valuation model designed to streamline equity research workflows. The system integrates a user-provided financial forecast model with market data retrieved from Financial Modeling Prep (FMP) and Yahoo Finance to generate a professional, investor-ready valuation report in PDF format.
The model performs:

•	Intrinsic Valuation: DCF via Terminal Growth and Exit Multiple methods.

•	Relative Valuation: Comparable Company Analysis (EV/EBITDA, P/E).

•	Sensitivity Analysis: WACC vs. Terminal Growth and WACC vs. Exit Multiple.

•	Visualizations: Automated "Football Field" valuation charts.

**Environment Setup**

1. Python Version

•	Developed on: Python 3.13.12

•	Minimum Requirement: Python 3.10+
2. Dependencies
Install the required libraries via the provided requirements.txt:
Bash
pip install -r requirements.txt
Core libraries include xlwings for Excel integration, reportlab for PDF generation, and yfinance/requests for data retrieval.
3. API Configuration
This model requires an API key from Financial Modeling Prep (FMP).
1.	Open the .env file in the project root.
2.	Enter your key: FMP_API_KEY=your_api_key_here.

**Workflow & Data Integration**

The Two-File Excel System
The model uses a structured approach to stay flexible across different companies:
1.	inputs.xlsx (The Control Center): The user must fill this file. It contains sheets for Meta_Data, Target_Variables, FCFF_Components, and Comp_Tickers.
2.	User Forecast Workbook: Your personal DCF or 3-statement model. It can have any name, provided it is placed in the /data folder and the filename is specified in the Meta_Data sheet of inputs.xlsx.
Analyst-Driven Assumptions
To maintain professional differentiation, key variables are manually set by the analyst:

•	Beta & Risk-Free Rate: These are intentionally kept manual, allowing analysts to reflect specific views on market volatility and duration risk.

•	Scenario Switch: The user provides the cell location of the "Scenario Switch" in their forecast model; the script automatically toggles between Base, Bull, and Bear cases to extract values.

**How to Run**

1.	Place your forecast Excel workbook in the /data folder.
2.	Update inputs.xlsx with the forecast filename and your valuation parameters (WACC, Terminal Growth, etc.).
3.	Run the script:
Bash
python main.py
4.	Collect the Report: A PDF will be generated in data/reports/.

**Sample Output & Analysis**

The model generates a professional, multi-page PDF report that synthesizes complex financial data into actionable insights.
1. Valuation Summary
The report aggregates multiple methodologies to calculate a weighted fair value across different outlooks:

•	Scenario Analysis: Displays specific price targets for Base, Bull, and Bear cases based on the user’s forecast model.

•	Implied Upside/Downside: Calculates the percentage difference between the current market price and the derived intrinsic value.

•	Methodology Weighting: Summarizes the contribution of DCF vs. Comparable Analysis to the final target price.

2. Visualizations
The report includes an automated Football Field Chart to visualize valuation ranges and central tendencies:

•	Value Ranges (Floating Bars): Represents the spread of valuations for each method (e.g., Exit Multiple vs. Terminal Growth).

•	Base Case Indicator: A distinct marker showing the most likely valuation scenario.

•	Average Target Line: A vertical reference line indicating the blended fair value across all methodologies.

3. Investment Recommendation Logic

The model assigns a rating by comparing the current share price to the calculated upside:

•	BUY: Implied upside greater than 20%.

•	ACCUMULATE: Implied upside between 10% and 20%.

•	HOLD: Implied upside/downside between -10% and 10%.

•	REDUCE: Implied downside greater than -10%.

**Error Handling**

•	Critical Errors: Missing forecast files or required inputs will stop execution to ensure data integrity.

•	Non-Critical Errors: If an API request fails for a specific peer company, the model skips that ticker and continues to generate the report with available data.
