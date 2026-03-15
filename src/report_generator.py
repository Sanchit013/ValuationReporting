# Report Generator Module
# Simple PDF creation with duplicate filename handling

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.graphics.shapes import Drawing, Rect, String, Line
from reportlab.graphics import renderPDF
from pathlib import Path


def add_cover_page(doc, elements, title="Financial Valuation Report", subtitle="Prepared by DCF Model"):
    """
    Configure and add a full-sheet cover page with a centered report title.

    Parameters:
    -----------
    doc : SimpleDocTemplate
        The PDF document object
    elements : list
        List to store document elements
    title : str, optional
        Main cover title (default: "Financial Valuation Report")
    subtitle : str, optional
        Secondary line shown below title
    Returns:
    --------
    function
        onFirstPage callback for doc.build(...)
    """

    def _draw_cover_page(canvas, document):
        page_width, page_height = document.pagesize

        # Layered full-page background
        canvas.saveState()
        canvas.setFillColor(colors.HexColor('#F7F8F5'))
        canvas.rect(0, 0, page_width, page_height, fill=1, stroke=0)

        canvas.setFillColor(colors.HexColor('#124734'))
        canvas.rect(0, page_height * 0.80, page_width, page_height * 0.20, fill=1, stroke=0)
        canvas.rect(0, 0, page_width, page_height * 0.12, fill=1, stroke=0)

        center_x = page_width / 2
        center_y = page_height / 2

        # Decorative accent lines
        line_half_width = page_width * 0.22
        canvas.setStrokeColor(colors.HexColor('#2E6F40'))
        canvas.setLineWidth(1.5)
        canvas.line(center_x - line_half_width, center_y + 24, center_x + line_half_width, center_y + 24)
        canvas.line(center_x - line_half_width, center_y - 50, center_x + line_half_width, center_y - 50)

        # Main centered title
        canvas.setFillColor(colors.HexColor('#0A0A0A'))
        canvas.setFont('Helvetica-Bold', 30)
        canvas.drawCentredString(center_x, center_y - 12, title)

        # Subtitle
        canvas.setFillColor(colors.HexColor('#4A4A4A'))
        canvas.setFont('Helvetica', 13)
        canvas.drawCentredString(center_x, center_y - 80, subtitle)
        canvas.restoreState()

    # First flowable starts on page 2; page 1 is reserved for cover rendering.
    elements.append(PageBreak())
    return _draw_cover_page


def create_pdf(document_name, title=None, author="DCF Model"):
    """
    Create a PDF document object and initialize it for adding pages.
    If the filename already exists, appends a counter until a unique name is found.
    
    Parameters:
    -----------
    document_name : str
        Name of the document (without .pdf extension)
    title : str, optional
        Title for the PDF document (appears in browser tab)
    author : str, optional
        Author name for the PDF metadata (default: "DCF Model")
    
    Returns:
    --------
    tuple : (doc, elements, file_path)
        - doc: SimpleDocTemplate object for building PDF
        - elements: list to store document elements
        - file_path: path where PDF will be saved
    """
    
    # Ensure data/reports directory exists
    report_dir = Path("data/reports")
    report_dir.mkdir(parents=True, exist_ok=True)
    
    # Create full file path with .pdf extension
    file_path = report_dir / f"{document_name}.pdf"
    
    # Handle duplicate file names by adding counter
    counter = 1
    while file_path.exists():
        file_path = report_dir / f"{document_name} {counter}.pdf"
        counter += 1
    
    # Create PDF document
    doc = SimpleDocTemplate(
        str(file_path),
        pagesize=A4,
        rightMargin=0.75*inch,
        leftMargin=0.75*inch,
        topMargin=0.75*inch,
        bottomMargin=0.75*inch,
        title=title or document_name,
        author=author
    )
    
    # List to store document elements
    elements = []
    
    return doc, elements, file_path


def add_executive_summary_page(
    doc,
    elements,
    company_name,
    ticker_symbol,
    current_share_price,
    shares_outstanding,
    market_capitalization,
    intrinsic_valuation_data=None,
    relative_valuation_data=None,
    figures_stated_in="millions"
):
    """
    Add an Executive Summary page to the PDF with company data and valuation tables.
    
    Parameters:
    -----------
    doc : SimpleDocTemplate
        The PDF document object
    elements : list
        List to store document elements
    company_name : str
        Name of the company
    ticker_symbol : str
        Company ticker symbol
    current_share_price : float
        Current share price
    shares_outstanding : float
        Shares outstanding (in millions)
    market_capitalization : float
        Market capitalization (in billions)
    intrinsic_valuation_data : dict
        Dictionary with intrinsic valuation data (optional)
    relative_valuation_data : dict
        Dictionary with relative valuation data (optional)
    """
    
    styles = getSampleStyleSheet()
    
    # Define custom styles
    heading_style = ParagraphStyle(
        'CustomHeading',
        parent=styles['Heading1'],
        fontSize=18,
        textColor=colors.HexColor('#0A0A0A'),
        spaceAfter=0.3*inch,
        alignment=TA_LEFT
    )
    
    subheading_style = ParagraphStyle(
        'CustomSubHeading',
        parent=styles['Heading2'],
        fontSize=14,
        textColor=colors.HexColor('#0A0A0A'),
        spaceAfter=0.2*inch,
        spaceBefore=0.2*inch,
        alignment=TA_LEFT
    )
    
    table_heading_style = ParagraphStyle(
        'TableHeading',
        parent=styles['Normal'],
        fontSize=12,
        textColor=colors.HexColor('#FFFFFF'),
        alignment=TA_CENTER,
        fontName='Helvetica-Bold'
    )
    
    body_style = ParagraphStyle(
        'CustomBody',
        parent=styles['Normal'],
        fontSize=10,
        textColor=colors.HexColor('#4A4A4A'),
        spaceAfter=0.1*inch
    )
    
    # Helper function for getting suffix
    def _get_suffix(figures_stated_in):
        """Get appropriate suffix based on figures_stated_in parameter."""
        if isinstance(figures_stated_in, str):
            if 'thousand' in figures_stated_in.lower():
                return 'Th'
            elif 'billion' in figures_stated_in.lower():
                return 'B'
            elif 'million' in figures_stated_in.lower():
                return 'M'
        return 'M'  # Default to millions
    
    def _fmt_float_with_suffix(value, decimals=1):
        """Format float with dynamic suffix."""
        try:
            if value is None:
                return "N/A"
            num = float(value)
            if num != num:
                return "N/A"
            suffix = _get_suffix(figures_stated_in)
            return f"{num:,.{decimals}f} {suffix}"
        except (TypeError, ValueError):
            return "N/A"
    
    def _fmt_currency_with_suffix(value, decimals=2):
        """Format currency with dynamic suffix."""
        try:
            if value is None:
                return "N/A"
            num = float(value)
            if num != num:
                return "N/A"
            suffix = _get_suffix(figures_stated_in)
            return f"${num:,.{decimals}f} {suffix}"
        except (TypeError, ValueError):
            return "N/A"
    
    # Add main heading
    heading = Paragraph("Executive Summary", heading_style)
    elements.append(heading)
    elements.append(Spacer(1, 0.2*inch))
    
    # Add Market Snapshot subheading
    market_snapshot_heading = Paragraph("Market Snapshot", subheading_style)
    elements.append(market_snapshot_heading)
    
    # Add company information
    company_info_data = [
        ['Company Name:', company_name],
        ['Ticker Symbol:', ticker_symbol],
        ['Current Share Price:', f"${current_share_price:,.2f}"],
        ['Shares Outstanding:', _fmt_float_with_suffix(shares_outstanding, decimals=1)],
        ['Market Capitalization:', _fmt_currency_with_suffix(market_capitalization, decimals=2)]
    ]
    
    company_info_table = Table(company_info_data, colWidths=[2.5*inch, 3*inch])
    company_info_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.HexColor('#4A4A4A')),
        ('ROWBACKGROUNDS', (0, 0), (-1, -1), [colors.HexColor('#F2F2EB'), colors.white]),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#CCCCCC')),
    ]))
    elements.append(company_info_table)
    elements.append(Spacer(1, 0.3*inch))
    
    # Add valuation results summary heading
    valuation_heading = Paragraph("Valuation Result Summary", subheading_style)
    elements.append(valuation_heading)
    elements.append(Spacer(1, 0.15*inch))
    
    # Summary Table: Average Valuation Across All Methods
    summary_table_label = Paragraph("Valuation Summary - Average Price Target", ParagraphStyle(
        'TableLabel',
        parent=styles['Normal'],
        fontSize=11,
        fontName='Helvetica-Bold',
        textColor=colors.HexColor('#0A0A0A'),
        spaceAfter=0.1*inch
    ))
    elements.append(summary_table_label)
    
    # Calculate average prices from all valuation methods
    def extract_price(price_str):
        """Extract numeric value from price string like '$168.50'"""
        try:
            return float(price_str.replace('$', '').replace(',', ''))
        except:
            return 0.0
    
    # Initialize values for each scenario
    avg_base = 0
    avg_bull = 0
    avg_bear = 0
    
    # Step 1: Extract intrinsic valuation values (DCF Terminal Growth and Exit Multiple)
    intrinsic_base = []
    intrinsic_bull = []
    intrinsic_bear = []
    
    if intrinsic_valuation_data:
        for method, scenarios in intrinsic_valuation_data.items():
            intrinsic_base.append(extract_price(scenarios.get('Base Case', '$0')))
            intrinsic_bull.append(extract_price(scenarios.get('Bull Case', '$0')))
            intrinsic_bear.append(extract_price(scenarios.get('Bear Case', '$0')))
    
    # Step 2: Extract relative valuation values by method
    pe_ltm_base = pe_ltm_bull = pe_ltm_bear = 0
    pe_fy_base = pe_fy_bull = pe_fy_bear = 0
    ev_ltm_base = ev_ltm_bull = ev_ltm_bear = 0
    ev_fy_base = ev_fy_bull = ev_fy_bear = 0
    
    if relative_valuation_data:
        # P/E LTM
        if 'P/E LTM' in relative_valuation_data:
            pe_ltm_base = extract_price(relative_valuation_data['P/E LTM'].get('Base Case', '$0'))
            pe_ltm_bull = extract_price(relative_valuation_data['P/E LTM'].get('Bull Case', '$0'))
            pe_ltm_bear = extract_price(relative_valuation_data['P/E LTM'].get('Bear Case', '$0'))
        
        # P/E FY
        if 'P/E FY' in relative_valuation_data:
            pe_fy_base = extract_price(relative_valuation_data['P/E FY'].get('Base Case', '$0'))
            pe_fy_bull = extract_price(relative_valuation_data['P/E FY'].get('Bull Case', '$0'))
            pe_fy_bear = extract_price(relative_valuation_data['P/E FY'].get('Bear Case', '$0'))
        
        # EV/EBITDA LTM
        if 'EV/EBITDA LTM' in relative_valuation_data:
            ev_ltm_base = extract_price(relative_valuation_data['EV/EBITDA LTM'].get('Base Case', '$0'))
            ev_ltm_bull = extract_price(relative_valuation_data['EV/EBITDA LTM'].get('Bull Case', '$0'))
            ev_ltm_bear = extract_price(relative_valuation_data['EV/EBITDA LTM'].get('Bear Case', '$0'))
        
        # EV/EBITDA FY
        if 'EV/EBITDA FY' in relative_valuation_data:
            ev_fy_base = extract_price(relative_valuation_data['EV/EBITDA FY'].get('Base Case', '$0'))
            ev_fy_bull = extract_price(relative_valuation_data['EV/EBITDA FY'].get('Bull Case', '$0'))
            ev_fy_bear = extract_price(relative_valuation_data['EV/EBITDA FY'].get('Bear Case', '$0'))
    
    # Step 3: Calculate averages for P/E and EV/EBITDA
    pe_avg_base = (pe_ltm_base + pe_fy_base) / 2 if (pe_ltm_base or pe_fy_base) else 0
    pe_avg_bull = (pe_ltm_bull + pe_fy_bull) / 2 if (pe_ltm_bull or pe_fy_bull) else 0
    pe_avg_bear = (pe_ltm_bear + pe_fy_bear) / 2 if (pe_ltm_bear or pe_fy_bear) else 0
    
    ev_avg_base = (ev_ltm_base + ev_fy_base) / 2 if (ev_ltm_base or ev_fy_base) else 0
    ev_avg_bull = (ev_ltm_bull + ev_fy_bull) / 2 if (ev_ltm_bull or ev_fy_bull) else 0
    ev_avg_bear = (ev_ltm_bear + ev_fy_bear) / 2 if (ev_ltm_bear or ev_fy_bear) else 0
    
    # Step 4: Calculate final average (2 DCF + PE avg + EV/EBITDA avg) / 4
    all_base_values = intrinsic_base + [pe_avg_base, ev_avg_base]
    all_bull_values = intrinsic_bull + [pe_avg_bull, ev_avg_bull]
    all_bear_values = intrinsic_bear + [pe_avg_bear, ev_avg_bear]
    
    avg_base = sum(all_base_values) / len(all_base_values) if all_base_values else 0
    avg_bull = sum(all_bull_values) / len(all_bull_values) if all_bull_values else 0
    avg_bear = sum(all_bear_values) / len(all_bear_values) if all_bear_values else 0
    
    # Create summary table
    summary_table_data = [
        [Paragraph('Scenario', table_heading_style), 
         Paragraph('Average Price Target', table_heading_style),
         Paragraph('Upside/Downside', table_heading_style)]
    ]
    
    # Calculate upside/downside percentages
    base_upside = ((avg_base - current_share_price) / current_share_price * 100)
    bull_upside = ((avg_bull - current_share_price) / current_share_price * 100)
    bear_upside = ((avg_bear - current_share_price) / current_share_price * 100)
    
    summary_table_data.append(['Base Case', f'${avg_base:,.2f}', f'{base_upside:+.2f}%'])
    summary_table_data.append(['Bull Case', f'${avg_bull:,.2f}', f'{bull_upside:+.2f}%'])
    summary_table_data.append(['Bear Case', f'${avg_bear:,.2f}', f'{bear_upside:+.2f}%'])
    
    summary_table = Table(summary_table_data, colWidths=[1.8*inch, 2*inch, 2*inch])
    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#124734')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),  # Bold Base Case row
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.HexColor('#4A4A4A')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#F2F2EB'), colors.white]),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#CCCCCC')),
    ]))
    elements.append(summary_table)
    elements.append(Spacer(1, 0.2*inch))
    
    # Table A: Intrinsic Valuation
    table_a_label = Paragraph("Table A. Intrinsic Valuation", ParagraphStyle(
        'TableLabel',
        parent=styles['Normal'],
        fontSize=11,
        fontName='Helvetica-Bold',
        textColor=colors.HexColor('#0A0A0A'),
        spaceAfter=0.1*inch
    ))
    elements.append(table_a_label)
    
    # Create intrinsic valuation table with DCF methods and scenarios
    if intrinsic_valuation_data is None:
        intrinsic_valuation_data = {
            'Terminal Growth Model': {
                'Base Case': '$XX.XX',
                'Bull Case': '$XX.XX',
                'Bear Case': '$XX.XX'
            },
            'EBITDA Exit Multiple': {
                'Base Case': '$XX.XX',
                'Bull Case': '$XX.XX',
                'Bear Case': '$XX.XX'
            }
        }
    
    intrinsic_table_data = [
        [Paragraph('DCF Method', table_heading_style), 
         Paragraph('Base Case', table_heading_style),
         Paragraph('Bull Case', table_heading_style),
         Paragraph('Bear Case', table_heading_style)]
    ]
    
    for method, scenarios in intrinsic_valuation_data.items():
        row = [method]
        for scenario in ['Base Case', 'Bull Case', 'Bear Case']:
            row.append(scenarios.get(scenario, '$XX.XX'))
        intrinsic_table_data.append(row)
    
    intrinsic_table = Table(intrinsic_table_data, colWidths=[2*inch, 1.3*inch, 1.3*inch, 1.3*inch])
    intrinsic_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#124734')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTNAME', (1, 1), (1, -1), 'Helvetica-Bold'),  # Bold Base Case column
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.HexColor('#4A4A4A')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#F2F2EB'), colors.white]),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#CCCCCC')),
    ]))
    elements.append(intrinsic_table)
    elements.append(Spacer(1, 0.2*inch))
    
    # Table B: Relative Valuation
    table_b_label = Paragraph("Table B. Relative Valuation", ParagraphStyle(
        'TableLabel',
        parent=styles['Normal'],
        fontSize=11,
        fontName='Helvetica-Bold',
        textColor=colors.HexColor('#0A0A0A'),
        spaceAfter=0.1*inch
    ))
    elements.append(table_b_label)
    
    # Create relative valuation table with 4 methods and 3 scenarios
    if relative_valuation_data is None:
        relative_valuation_data = {
            'EV/EBITDA LTM': {
                'Base Case': '$XX.XX',
                'Bull Case': '$XX.XX',
                'Bear Case': '$XX.XX'
            },
            'EV/EBITDA FY': {
                'Base Case': '$XX.XX',
                'Bull Case': '$XX.XX',
                'Bear Case': '$XX.XX'
            },
            'P/E LTM': {
                'Base Case': '$XX.XX',
                'Bull Case': '$XX.XX',
                'Bear Case': '$XX.XX'
            },
            'P/E FY': {
                'Base Case': '$XX.XX',
                'Bull Case': '$XX.XX',
                'Bear Case': '$XX.XX'
            }
        }
    
    relative_table_data = [
        [Paragraph('Valuation Method', table_heading_style), 
         Paragraph('Base Case', table_heading_style),
         Paragraph('Bull Case', table_heading_style),
         Paragraph('Bear Case', table_heading_style)]
    ]
    
    for method, scenarios in relative_valuation_data.items():
        row = [method]
        for scenario in ['Base Case', 'Bull Case', 'Bear Case']:
            row.append(scenarios.get(scenario, '$XX.XX'))
        relative_table_data.append(row)
    
    relative_table = Table(relative_table_data, colWidths=[1.8*inch, 1.3*inch, 1.3*inch, 1.3*inch])
    relative_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#124734')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTNAME', (1, 1), (1, -1), 'Helvetica-Bold'),  # Bold Base Case column
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.HexColor('#4A4A4A')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#F2F2EB'), colors.white]),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#CCCCCC')),
    ]))
    elements.append(relative_table)
    elements.append(PageBreak())


def add_model_assumptions_page(
    doc,
    elements,
    forecast_period=5,
    effective_tax_rate=0.20,
    cost_of_equity=0.10,
    beta=1.10,
    risk_free_rate=0.045,
    market_return=0.10,
    cost_of_debt=0.042,
    wacc=0.092,
    terminal_growth_rate=0.03,
    exit_multiple=12.0
):
    """
    Add a Model Assumptions & Parameters page to the PDF with dynamic values.
    
    Parameters:
    -----------
    doc : SimpleDocTemplate
        The PDF document object
    elements : list
        List to store document elements
    forecast_period : int
        Number of years for forecast period (default: 5)
    effective_tax_rate : float
        Effective tax rate as decimal (default: 0.20 = 20%)
    cost_of_equity : float
        Cost of equity as decimal (default: 0.10 = 10%)
    beta : float
        Beta coefficient (default: 1.10)
    risk_free_rate : float
        Risk-free rate as decimal (default: 0.045 = 4.5%)
    market_return : float
        Market return as decimal (default: 0.10 = 10%)
    cost_of_debt : float
        Cost of debt as decimal (default: 0.042 = 4.2%)
    wacc : float
        Weighted average cost of capital as decimal (default: 0.092 = 9.2%)
    terminal_growth_rate : float
        Terminal growth rate as decimal (default: 0.03 = 3%)
    exit_multiple : float
        Exit multiple for terminal value (default: 12.0x)
    """
    
    styles = getSampleStyleSheet()
    
    # Define custom styles
    heading_style = ParagraphStyle(
        'ModelHeading',
        parent=styles['Heading1'],
        fontSize=18,
        textColor=colors.HexColor('#0A0A0A'),
        spaceAfter=0.3*inch,
        alignment=TA_LEFT
    )
    
    table_heading_style = ParagraphStyle(
        'ModelTableHeading',
        parent=styles['Normal'],
        fontSize=11,
        textColor=colors.HexColor('#FFFFFF'),
        alignment=TA_CENTER,
        fontName='Helvetica-Bold'
    )
    
    # Add main heading
    heading = Paragraph("Model Assumptions & Parameters", heading_style)
    elements.append(heading)
    elements.append(Spacer(1, 0.2*inch))
    
    # Create assumptions data table
    assumptions_data = [
        [Paragraph('Parameter', table_heading_style), Paragraph('Value', table_heading_style)]
    ]

    def _fmt_years(value):
        try:
            if value is None:
                return "N/A"
            num = float(value)
            if num != num:
                return "N/A"
            return f"{int(num)}" if num.is_integer() else f"{num:.2f}"
        except (TypeError, ValueError):
            return "N/A"

    def _fmt_percent(value):
        try:
            if value is None:
                return "N/A"
            num = float(value)
            if num != num:
                return "N/A"
            return f"{num * 100:.2f}%"
        except (TypeError, ValueError):
            return "N/A"

    def _fmt_float(value, suffix=""):
        try:
            if value is None:
                return "N/A"
            num = float(value)
            if num != num:
                return "N/A"
            return f"{num:.2f}{suffix}"
        except (TypeError, ValueError):
            return "N/A"
    
    # Add all assumptions with formatted values
    assumptions_data.append(['Forecast Period (Years)', _fmt_years(forecast_period)])
    assumptions_data.append(['Effective Tax Rate', _fmt_percent(effective_tax_rate)])
    assumptions_data.append(['Cost of Equity', _fmt_percent(cost_of_equity)])
    assumptions_data.append(['Beta', _fmt_float(beta)])
    assumptions_data.append(['Risk-Free Rate', _fmt_percent(risk_free_rate)])
    assumptions_data.append(['Market Return', _fmt_percent(market_return)])
    assumptions_data.append(['Cost of Debt', _fmt_percent(cost_of_debt)])
    assumptions_data.append(['WACC (Discount Rate)', _fmt_percent(wacc)])
    assumptions_data.append(['Terminal Growth Rate', _fmt_percent(terminal_growth_rate)])
    assumptions_data.append(['Exit Multiple (EV/EBITDA)', _fmt_float(exit_multiple, 'x')])
    
    # Create table
    assumptions_table = Table(assumptions_data, colWidths=[3.5*inch, 2*inch])
    assumptions_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#124734')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('ALIGN', (1, 0), (1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.HexColor('#4A4A4A')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#F2F2EB'), colors.white]),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#CCCCCC')),
    ]))
    
    elements.append(assumptions_table)
    elements.append(PageBreak())


def add_intrinsic_valuation_page(
    doc,
    elements,
    forecast_years_data,
    terminal_growth_summary,
    exit_multiple_summary,
    sensitivity_tgm,
    sensitivity_em,
    base_wacc=None,
    base_terminal_growth_rate=None,
    base_exit_multiple=None,
    figures_stated_in="millions"
):
    """
    Add an Intrinsic Valuation page using precomputed inputs only.

    Parameters:
    -----------
    doc : SimpleDocTemplate
        The PDF document object
    elements : list
        List to store document elements
    forecast_years_data : list of dict
        Each dict should include: 'year', 'ebit_1_minus_t', 'da', 'capex',
        'change_nwc', and 'fcff' (all precomputed values).
    terminal_growth_summary : dict
        Precomputed values for the terminal growth method:
        'terminal_value', 'enterprise_value', 'net_debt', 'total_equity_value',
        'shares_outstanding', 'implied_share_price'.
    exit_multiple_summary : dict
        Precomputed values for the exit multiple method:
        'terminal_ebitda', 'exit_multiple', 'terminal_value', 'enterprise_value',
        'net_debt', 'total_equity_value', 'shares_outstanding', 'implied_share_price'.
    sensitivity_tgm : list of lists
        Precomputed table for WACC vs Terminal Growth (already formatted as strings).
    sensitivity_em : list of lists
        Precomputed table for WACC vs Exit Multiple (already formatted as strings).
    base_wacc : float, optional
        Base WACC value used to highlight the sensitivity tables (decimal).
    base_terminal_growth_rate : float, optional
        Base terminal growth rate used to highlight the sensitivity table (decimal).
    base_exit_multiple : float, optional
        Base exit multiple used to highlight the sensitivity table.
    """
    
    styles = getSampleStyleSheet()
    
    # Define custom styles
    heading_style = ParagraphStyle(
        'IntrinsicHeading',
        parent=styles['Heading1'],
        fontSize=18,
        textColor=colors.HexColor('#0A0A0A'),
        spaceAfter=0.3*inch,
        alignment=TA_LEFT
    )
    
    subheading_style = ParagraphStyle(
        'IntrinsicSubHeading',
        parent=styles['Heading2'],
        fontSize=12,
        textColor=colors.HexColor('#0A0A0A'),
        spaceAfter=0.15*inch,
        spaceBefore=0.15*inch,
        alignment=TA_LEFT
    )
    
    table_heading_style = ParagraphStyle(
        'IntrinsicTableHeading',
        parent=styles['Normal'],
        fontSize=10,
        textColor=colors.HexColor('#FFFFFF'),
        alignment=TA_CENTER,
        fontName='Helvetica-Bold'
    )
    
    table_data_style = ParagraphStyle(
        'TableData',
        parent=styles['Normal'],
        fontSize=9,
        alignment=TA_CENTER
    )
    
    # Add main heading
    heading = Paragraph("Intrinsic Valuation", heading_style)
    elements.append(heading)
    elements.append(Spacer(1, 0.2*inch))
    
    if forecast_years_data is None:
        forecast_years_data = []

    def _fmt_money(value, decimals=0):
        try:
            if value is None:
                return "N/A"
            num = float(value)
            if num != num:
                return "N/A"
            return f"${num:,.{decimals}f}"
        except (TypeError, ValueError):
            return "N/A"

    def _get_suffix(figures_stated_in):
        """Get appropriate suffix based on figures_stated_in parameter."""
        if isinstance(figures_stated_in, str):
            if 'thousand' in figures_stated_in.lower():
                return 'Th'
            elif 'billion' in figures_stated_in.lower():
                return 'B'
            elif 'million' in figures_stated_in.lower():
                return 'M'
        return 'M'  # Default to millions

    def _fmt_money_m(value, decimals=0):
        formatted = _fmt_money(value, decimals)
        suffix = _get_suffix(figures_stated_in)
        return formatted if formatted == "N/A" else f"{formatted} {suffix}"

    def _fmt_price(value):
        return _fmt_money(value, decimals=2)

    def _fmt_float(value, decimals=2, suffix=""):
        try:
            if value is None:
                return "N/A"
            num = float(value)
            if num != num:
                return "N/A"
            return f"{num:.{decimals}f}{suffix}"
        except (TypeError, ValueError):
            return "N/A"

    def _fmt_float_with_suffix(value, decimals=1):
        """Format float with dynamic suffix based on figures_stated_in."""
        if value is None:
            return "N/A"
        try:
            num = float(value)
            if num != num:
                return "N/A"
            suffix = _get_suffix(figures_stated_in)
            return f"{num:,.{decimals}f} {suffix}"
        except (TypeError, ValueError):
            return "N/A"

    def _stringify_table(table_data):
        if not table_data:
            return [["N/A"]]
        return [["N/A" if cell is None else str(cell) for cell in row] for row in table_data]

    def _normalize_table(table_data, target_cols):
        if not table_data:
            return [["N/A" for _ in range(target_cols)]]
        normalized = []
        for row in table_data:
            row_vals = list(row[:target_cols])
            if len(row_vals) < target_cols:
                row_vals.extend(["N/A"] * (target_cols - len(row_vals)))
            normalized.append(row_vals)
        return normalized

    def _fmt_percent_label(value):
        try:
            if value is None:
                return None
            num = float(value)
            if num != num:
                return None
            return f"{num * 100:.1f}%"
        except (TypeError, ValueError):
            return None

    def _fmt_multiple_label(value):
        try:
            if value is None:
                return None
            num = float(value)
            if num != num:
                return None
            return f"{num:.1f}x"
        except (TypeError, ValueError):
            return None

    def _find_index(values, target):
        if not target:
            return None
        try:
            return values.index(target)
        except ValueError:
            return None
    
    # Create Free Cash Flow Forecast Table
    heading_fcf = Paragraph("Free Cash Flow Forecast", subheading_style)
    elements.append(heading_fcf)
    
    # Build FCF table with dynamic columns based on forecast_period
    years = [item.get("year", "N/A") for item in forecast_years_data]
    if not years:
        years = ["N/A"]

    fcf_table_data = [
        ['Metric'] + [f'Year {year}' for year in years]
    ]
    
    # EBIT(1-T) row
    ebit_row = ['EBIT(1-T)'] + [
        _fmt_money(i.get("ebit_1_minus_t"), decimals=0) for i in forecast_years_data
    ]
    fcf_table_data.append(ebit_row)
    
    # D&A row
    da_row = ['+ D&A'] + [_fmt_money(i.get("da"), decimals=0) for i in forecast_years_data]
    fcf_table_data.append(da_row)
    
    # CapEx row
    capex_row = ['- Capital Expenditures'] + [
        _fmt_money(i.get("capex"), decimals=0) for i in forecast_years_data
    ]
    fcf_table_data.append(capex_row)
    
    # Change in NWC row
    nwc_row = ['- Change in NWC'] + [
        _fmt_money(i.get("change_nwc"), decimals=0) for i in forecast_years_data
    ]
    fcf_table_data.append(nwc_row)
    
    # FCFF row (calculated)
    fcff_row = ['FCFF'] + [_fmt_money(i.get("fcff"), decimals=0) for i in forecast_years_data]
    fcf_table_data.append(fcff_row)
    
    # Create column widths dynamically
    years_count = len(years)
    col_width = 5.5 / (years_count + 1)
    fcf_col_widths = [1.5*inch] + [col_width*inch for _ in range(years_count)]
    
    fcf_table = Table(fcf_table_data, colWidths=fcf_col_widths)
    fcf_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#124734')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.HexColor('#4A4A4A')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#F2F2EB'), colors.white]),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#CCCCCC')),
    ]))
    elements.append(fcf_table)
    elements.append(Spacer(1, 0.25*inch))
    
    # ============================================================================
    # DCF Valuation - Terminal Growth Method
    # ============================================================================
    heading_tgm = Paragraph("DCF Valuation - Terminal Growth Method", subheading_style)
    elements.append(heading_tgm)
    
    terminal_growth_summary = terminal_growth_summary or {}
    
    # Valuation summary table
    valuation_tgm_data = [
        [Paragraph('Metric', table_heading_style), Paragraph('Value', table_heading_style)]
    ]
    valuation_tgm_data.append(['Terminal Value', _fmt_money_m(terminal_growth_summary.get('terminal_value'), decimals=0)])
    valuation_tgm_data.append(['Enterprise Value', _fmt_money_m(terminal_growth_summary.get('enterprise_value'), decimals=0)])
    valuation_tgm_data.append(['Less: Net Debt', _fmt_money_m(terminal_growth_summary.get('net_debt'), decimals=0)])
    valuation_tgm_data.append(['Total Equity Value', _fmt_money_m(terminal_growth_summary.get('total_equity_value'), decimals=0)])
    valuation_tgm_data.append(['Shares Outstanding', _fmt_float_with_suffix(terminal_growth_summary.get('shares_outstanding'), decimals=1)])
    valuation_tgm_data.append(['Implied Share Price', _fmt_price(terminal_growth_summary.get('implied_share_price'))])
    
    valuation_tgm_table = Table(valuation_tgm_data, colWidths=[3.5*inch, 2*inch])
    valuation_tgm_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#124734')),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('ALIGN', (1, 0), (1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.HexColor('#4A4A4A')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#F2F2EB'), colors.white]),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#CCCCCC')),
    ]))
    elements.append(valuation_tgm_table)
    elements.append(Spacer(1, 0.2*inch))
    
    # Sensitivity Table for Terminal Growth Method
    sensitivity_heading = Paragraph("Sensitivity Analysis - WACC vs Terminal Growth Rate", ParagraphStyle(
        'SensitivityHeading',
        parent=styles['Normal'],
        fontSize=10,
        fontName='Helvetica-Bold',
        textColor=colors.HexColor('#0A0A0A'),
        spaceAfter=0.1*inch
    ))
    elements.append(sensitivity_heading)
    
    sensitivity_data = _normalize_table(_stringify_table(sensitivity_tgm), 6)
    def _parse_percent_label(value):
        try:
            if value is None:
                return None
            text = str(value).strip().replace('%', '')
            return float(text) / 100
        except (TypeError, ValueError):
            return None

    if sensitivity_data and len(sensitivity_data) > 1:
        tgr_values = [_parse_percent_label(val) for val in sensitivity_data[0][1:]]
        for row_idx in range(1, len(sensitivity_data)):
            wacc_value = _parse_percent_label(sensitivity_data[row_idx][0])
            for col_idx, tgr_value in enumerate(tgr_values, start=1):
                if wacc_value is None or tgr_value is None:
                    continue
                if wacc_value < tgr_value:
                    sensitivity_data[row_idx][col_idx] = "N/A"
    tgm_cols = len(sensitivity_data[0])
    tgm_first_col = 1.2 * inch
    tgm_total_width = 5.5 * inch
    tgm_other_col = (tgm_total_width - tgm_first_col) / max(tgm_cols - 1, 1)
    tgm_col_widths = [tgm_first_col] + [tgm_other_col for _ in range(max(tgm_cols - 1, 0))]
    sensitivity_table = Table(sensitivity_data, colWidths=tgm_col_widths)
    sensitivity_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#124734')),
        ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#124734')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('TEXTCOLOR', (0, 0), (0, -1), colors.white),
        ('TEXTCOLOR', (0, 1), (0, -1), colors.white),  # First column values (WACC)
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('TEXTCOLOR', (1, 1), (-1, -1), colors.HexColor('#4A4A4A')),
        ('ROWBACKGROUNDS', (1, 1), (-1, -1), [colors.HexColor('#F2F2EB'), colors.white]),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#CCCCCC')),
    ]))
    tgm_wacc_label = _fmt_percent_label(base_wacc)
    tgm_tgr_label = _fmt_percent_label(base_terminal_growth_rate)
    tgm_header = sensitivity_data[0] if sensitivity_data else []
    tgm_wacc_rows = [row[0] for row in sensitivity_data[1:]]
    tgm_wacc_row = _find_index(tgm_wacc_rows, tgm_wacc_label)
    tgm_tgr_col = _find_index(tgm_header, tgm_tgr_label)
    if tgm_wacc_row is not None:
        row_idx = tgm_wacc_row + 1
        sensitivity_table.setStyle(TableStyle([
            ('FONTNAME', (0, row_idx), (0, row_idx), 'Helvetica-Bold'),
        ]))
    if tgm_tgr_col is not None:
        sensitivity_table.setStyle(TableStyle([
            ('FONTNAME', (tgm_tgr_col, 0), (tgm_tgr_col, 0), 'Helvetica-Bold'),
        ]))
    if tgm_wacc_row is not None and tgm_tgr_col is not None:
        row_idx = tgm_wacc_row + 1
        sensitivity_table.setStyle(TableStyle([
            ('FONTNAME', (tgm_tgr_col, row_idx), (tgm_tgr_col, row_idx), 'Helvetica-Bold'),
        ]))
    elements.append(sensitivity_table)
    elements.append(Spacer(1, 0.25*inch))
    
    # ============================================================================
    # DCF Valuation - EBITDA Exit Multiple Method
    # ============================================================================
    heading_em = Paragraph("DCF Valuation - EBITDA Exit Multiple Method", subheading_style)
    elements.append(heading_em)
    
    exit_multiple_summary = exit_multiple_summary or {}
    
    # Valuation summary table for Exit Multiple method
    valuation_em_data = [
        [Paragraph('Metric', table_heading_style), Paragraph('Value', table_heading_style)]
    ]
    valuation_em_data.append(['Terminal EBITDA', _fmt_money_m(exit_multiple_summary.get('terminal_ebitda'), decimals=0)])
    valuation_em_data.append(['Exit Multiple (EV/EBITDA)', _fmt_float(exit_multiple_summary.get('exit_multiple'), decimals=1, suffix="x")])
    valuation_em_data.append(['Terminal Value', _fmt_money_m(exit_multiple_summary.get('terminal_value'), decimals=0)])
    valuation_em_data.append(['Enterprise Value', _fmt_money_m(exit_multiple_summary.get('enterprise_value'), decimals=0)])
    valuation_em_data.append(['Less: Net Debt', _fmt_money_m(exit_multiple_summary.get('net_debt'), decimals=0)])
    valuation_em_data.append(['Total Equity Value', _fmt_money_m(exit_multiple_summary.get('total_equity_value'), decimals=0)])
    valuation_em_data.append(['Shares Outstanding', _fmt_float_with_suffix(exit_multiple_summary.get('shares_outstanding'), decimals=1)])
    valuation_em_data.append(['Implied Share Price', _fmt_price(exit_multiple_summary.get('implied_share_price'))])
    
    valuation_em_table = Table(valuation_em_data, colWidths=[3.5*inch, 2*inch])
    valuation_em_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#124734')),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('ALIGN', (1, 0), (1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.HexColor('#4A4A4A')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#F2F2EB'), colors.white]),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#CCCCCC')),
    ]))
    elements.append(valuation_em_table)
    elements.append(Spacer(1, 0.2*inch))
    
    # Sensitivity Table for Exit Multiple Method
    sensitivity_em_heading = Paragraph("Sensitivity Analysis - WACC vs Exit Multiple", ParagraphStyle(
        'SensitivityEMHeading',
        parent=styles['Normal'],
        fontSize=10,
        fontName='Helvetica-Bold',
        textColor=colors.HexColor('#0A0A0A'),
        spaceAfter=0.1*inch
    ))
    elements.append(sensitivity_em_heading)
    
    sensitivity_em_data = _normalize_table(_stringify_table(sensitivity_em), 6)
    em_cols = len(sensitivity_em_data[0])
    em_first_col = 1.2 * inch
    em_total_width = 5.5 * inch
    em_other_col = (em_total_width - em_first_col) / max(em_cols - 1, 1)
    em_col_widths = [em_first_col] + [em_other_col for _ in range(max(em_cols - 1, 0))]
    sensitivity_em_table = Table(sensitivity_em_data, colWidths=em_col_widths)
    sensitivity_em_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#124734')),
        ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#124734')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('TEXTCOLOR', (0, 0), (0, -1), colors.white),
        ('TEXTCOLOR', (0, 1), (0, -1), colors.white),  # First column values (WACC)
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('TEXTCOLOR', (1, 1), (-1, -1), colors.HexColor('#4A4A4A')),
        ('ROWBACKGROUNDS', (1, 1), (-1, -1), [colors.HexColor('#F2F2EB'), colors.white]),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#CCCCCC')),
    ]))
    em_wacc_label = _fmt_percent_label(base_wacc)
    em_mult_label = _fmt_multiple_label(base_exit_multiple)
    em_header = sensitivity_em_data[0] if sensitivity_em_data else []
    em_wacc_rows = [row[0] for row in sensitivity_em_data[1:]]
    em_wacc_row = _find_index(em_wacc_rows, em_wacc_label)
    em_mult_col = _find_index(em_header, em_mult_label)
    if em_wacc_row is not None:
        row_idx = em_wacc_row + 1
        sensitivity_em_table.setStyle(TableStyle([
            ('FONTNAME', (0, row_idx), (0, row_idx), 'Helvetica-Bold'),
        ]))
    if em_mult_col is not None:
        sensitivity_em_table.setStyle(TableStyle([
            ('FONTNAME', (em_mult_col, 0), (em_mult_col, 0), 'Helvetica-Bold'),
        ]))
    if em_wacc_row is not None and em_mult_col is not None:
        row_idx = em_wacc_row + 1
        sensitivity_em_table.setStyle(TableStyle([
            ('FONTNAME', (em_mult_col, row_idx), (em_mult_col, row_idx), 'Helvetica-Bold'),
        ]))
    elements.append(sensitivity_em_table)
    
    elements.append(PageBreak())


def add_relative_valuation_page(doc, elements, peer_data, summary_stats, implied_valuations, figures_stated_in="millions"):
    """
    Create a Relative Valuation Analysis page with peer group comparison.
    
    Parameters:
    -----------
    doc : SimpleDocTemplate
        The PDF document object
    elements : list
        List to append PDF elements to
    peer_data : list of dicts
        List of peer company data. Each dict should have:
        - 'ticker': str
        - 'company_name': str
        - 'ev_ebitda_ltm': float
        - 'ev_ebitda_fy': float
        - 'pe_ltm': float
        - 'pe_fy': float
    summary_stats : dict
        Summary statistics for each multiple. Each key should have:
        - 'ev_ebitda_ltm': {'average': float, 'median': float, 'highest': float, 'lowest': float}
        - 'ev_ebitda_fy': {'average': float, 'median': float, 'highest': float, 'lowest': float}
        - 'pe_ltm': {'average': float, 'median': float, 'highest': float, 'lowest': float}
        - 'pe_fy': {'average': float, 'median': float, 'highest': float, 'lowest': float}
    implied_valuations : dict
        Pre-calculated implied valuations for each method. Should contain:
        - 'ev_ebitda_ltm': {'share_price': float, 'equity_value': float, 'enterprise_value': float}
        - 'ev_ebitda_fy': {'share_price': float, 'equity_value': float, 'enterprise_value': float}
        - 'pe_ltm': {'share_price': float, 'equity_value': float, 'enterprise_value': float}
        - 'pe_fy': {'share_price': float, 'equity_value': float, 'enterprise_value': float}
    """
    
    styles = getSampleStyleSheet()
    
    # Define custom styles
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=18,
        textColor=colors.HexColor('#0A0A0A'),
        spaceAfter=0.3*inch,
        alignment=TA_CENTER,
        fontName='Helvetica-Bold'
    )
    
    subheading_style = ParagraphStyle(
        'CustomSubHeading',
        parent=styles['Heading2'],
        fontSize=14,
        fontName='Helvetica-Bold',
        textColor=colors.HexColor('#0A0A0A'),
        spaceAfter=0.15*inch,
        spaceBefore=0.2*inch,
        alignment=TA_LEFT
    )
    
    table_heading_style = ParagraphStyle(
        'TableHeading',
        parent=styles['Normal'],
        fontSize=12,
        textColor=colors.HexColor('#FFFFFF'),
        alignment=TA_CENTER,
        fontName='Helvetica-Bold'
    )
    
    # Page title
    title = Paragraph("Relative Valuation Analysis", title_style)
    elements.append(title)
    elements.append(Spacer(1, 0.2*inch))
    
    # Helper function for dynamic suffix
    def _get_suffix(figures_stated_in):
        """Get appropriate suffix based on figures_stated_in parameter."""
        if isinstance(figures_stated_in, str):
            if 'thousand' in figures_stated_in.lower():
                return 'Th'
            elif 'billion' in figures_stated_in.lower():
                return 'B'
            elif 'million' in figures_stated_in.lower():
                return 'M'
        return 'M'  # Default to millions
    
    def _fmt_currency_with_suffix(value, decimals=0):
        """Format currency with dynamic suffix."""
        try:
            if value is None:
                return "N/A"
            num = float(value)
            if num != num:
                return "N/A"
            suffix = _get_suffix(figures_stated_in)
            return f"${num:,.{decimals}f}{suffix}"
        except (TypeError, ValueError):
            return "N/A"
    
    # ============================================================================
    # Peer Group Metrics Table
    # ============================================================================
    peer_heading = Paragraph("Peer Group Metrics", subheading_style)
    elements.append(peer_heading)
    
    # Create peer group table
    peer_table_data = [
        [
            Paragraph('Ticker', table_heading_style),
            Paragraph('Company Name', table_heading_style),
            Paragraph('EV/EBITDA<br/>LTM', table_heading_style),
            Paragraph('EV/EBITDA<br/>FY', table_heading_style),
            Paragraph('P/E<br/>LTM', table_heading_style),
            Paragraph('P/E<br/>FY', table_heading_style)
        ]
    ]
    
    # Add peer data rows
    for peer in peer_data:
        row = [
            peer['ticker'],
            peer['company_name'],
            f"{peer['ev_ebitda_ltm']:.2f}x" if peer['ev_ebitda_ltm'] is not None else 'N/A',
            f"{peer['ev_ebitda_fy']:.2f}x" if peer['ev_ebitda_fy'] is not None else 'N/A',
            f"{peer['pe_ltm']:.2f}x" if peer['pe_ltm'] is not None else 'N/A',
            f"{peer['pe_fy']:.2f}x" if peer['pe_fy'] is not None else 'N/A'
        ]
        peer_table_data.append(row)
    
    peer_table = Table(peer_table_data, colWidths=[0.8*inch, 2.2*inch, 1.0*inch, 1.0*inch, 0.9*inch, 0.9*inch])
    peer_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#124734')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('ALIGN', (1, 0), (1, -1), 'LEFT'),  # Company name left-aligned
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.HexColor('#4A4A4A')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#F2F2EB'), colors.white]),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#CCCCCC')),
    ]))
    elements.append(peer_table)
    elements.append(Spacer(1, 0.3*inch))
    
    # ============================================================================
    # Summary Statistics
    # ============================================================================
    summary_heading = Paragraph("Summary Statistics", subheading_style)
    elements.append(summary_heading)
    
    # Create summary statistics table
    summary_table_data = [
        [
            Paragraph('Statistic', table_heading_style),
            Paragraph('EV/EBITDA<br/>LTM', table_heading_style),
            Paragraph('EV/EBITDA<br/>FY', table_heading_style),
            Paragraph('P/E<br/>LTM', table_heading_style),
            Paragraph('P/E<br/>FY', table_heading_style)
        ]
    ]
    
    # Add summary statistics rows
    summary_table_data.append([
        'Average',
        f"{summary_stats['ev_ebitda_ltm']['average']:.2f}x",
        f"{summary_stats['ev_ebitda_fy']['average']:.2f}x",
        f"{summary_stats['pe_ltm']['average']:.2f}x",
        f"{summary_stats['pe_fy']['average']:.2f}x"
    ])
    
    summary_table_data.append([
        'Median',
        f"{summary_stats['ev_ebitda_ltm']['median']:.2f}x",
        f"{summary_stats['ev_ebitda_fy']['median']:.2f}x",
        f"{summary_stats['pe_ltm']['median']:.2f}x",
        f"{summary_stats['pe_fy']['median']:.2f}x"
    ])
    
    summary_table_data.append([
        'Highest',
        f"{summary_stats['ev_ebitda_ltm']['highest']:.2f}x",
        f"{summary_stats['ev_ebitda_fy']['highest']:.2f}x",
        f"{summary_stats['pe_ltm']['highest']:.2f}x",
        f"{summary_stats['pe_fy']['highest']:.2f}x"
    ])
    
    summary_table_data.append([
        'Lowest',
        f"{summary_stats['ev_ebitda_ltm']['lowest']:.2f}x",
        f"{summary_stats['ev_ebitda_fy']['lowest']:.2f}x",
        f"{summary_stats['pe_ltm']['lowest']:.2f}x",
        f"{summary_stats['pe_fy']['lowest']:.2f}x"
    ])
    
    summary_table = Table(summary_table_data, colWidths=[1.5*inch, 1.3*inch, 1.3*inch, 1.3*inch, 1.3*inch])
    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#124734')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.HexColor('#4A4A4A')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#F2F2EB'), colors.white]),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#CCCCCC')),
    ]))
    elements.append(summary_table)
    elements.append(Spacer(1, 0.3*inch))
    
    # ============================================================================
    # Implied Valuation
    # ============================================================================
    valuation_heading = Paragraph("Implied Valuation", subheading_style)
    elements.append(valuation_heading)
    
    # Use pre-calculated implied valuations
    # Create implied valuation table
    valuation_table_data = [
        [
            Paragraph('Method', table_heading_style),
            Paragraph('Share Price', table_heading_style),
            Paragraph('Equity Value', table_heading_style),
            Paragraph('Enterprise Value', table_heading_style)
        ]
    ]
    
    valuation_table_data.append([
        'EV/EBITDA LTM',
        f'${implied_valuations["ev_ebitda_ltm"]["share_price"]:.2f}',
        _fmt_currency_with_suffix(implied_valuations["ev_ebitda_ltm"]["equity_value"], decimals=0),
        _fmt_currency_with_suffix(implied_valuations["ev_ebitda_ltm"]["enterprise_value"], decimals=0)
    ])
    
    valuation_table_data.append([
        'EV/EBITDA FY',
        f'${implied_valuations["ev_ebitda_fy"]["share_price"]:.2f}',
        _fmt_currency_with_suffix(implied_valuations["ev_ebitda_fy"]["equity_value"], decimals=0),
        _fmt_currency_with_suffix(implied_valuations["ev_ebitda_fy"]["enterprise_value"], decimals=0)
    ])
    
    valuation_table_data.append([
        'P/E LTM',
        f'${implied_valuations["pe_ltm"]["share_price"]:.2f}',
        _fmt_currency_with_suffix(implied_valuations["pe_ltm"]["equity_value"], decimals=0),
        _fmt_currency_with_suffix(implied_valuations["pe_ltm"]["enterprise_value"], decimals=0)
    ])
    
    valuation_table_data.append([
        'P/E FY',
        f'${implied_valuations["pe_fy"]["share_price"]:.2f}',
        _fmt_currency_with_suffix(implied_valuations["pe_fy"]["equity_value"], decimals=0),
        _fmt_currency_with_suffix(implied_valuations["pe_fy"]["enterprise_value"], decimals=0)
    ])
    
    valuation_table = Table(valuation_table_data, colWidths=[1.8*inch, 1.5*inch, 1.8*inch, 1.8*inch])
    valuation_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#124734')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.HexColor('#4A4A4A')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#F2F2EB'), colors.white]),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#CCCCCC')),
    ]))
    elements.append(valuation_table)
    
    elements.append(PageBreak())


def add_conclusion_page(doc, elements, valuation_summary, current_share_price):
    """
    Create a Conclusion page with football field valuation chart and average price.
    
    Parameters:
    -----------
    doc : SimpleDocTemplate
        The PDF document object
    elements : list
        List to append PDF elements to
    valuation_summary : dict
        Dictionary containing valuation results from different methods:
        - 'dcf_terminal_growth': {'low': float, 'base': float, 'high': float}
          where 'low' = min value from sensitivity table, 
                'base' = base case valuation,
                'high' = max value from sensitivity table
        - 'dcf_exit_multiple': {'low': float, 'base': float, 'high': float}
          where 'low' = min value from sensitivity table,
                'base' = base case valuation,
                'high' = max value from sensitivity table
        - 'ev_ebitda_ltm': float (EV/EBITDA LTM valuation)
        - 'ev_ebitda_fy': float (EV/EBITDA FY valuation)
        - 'pe_ltm': float (P/E LTM valuation)
        - 'pe_fy': float (P/E FY valuation)
    current_share_price : float
        Current market share price for comparison
    """
    
    styles = getSampleStyleSheet()
    
    # Define custom styles
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=18,
        textColor=colors.HexColor('#0A0A0A'),
        spaceAfter=0.3*inch,
        alignment=TA_CENTER,
        fontName='Helvetica-Bold'
    )
    
    subheading_style = ParagraphStyle(
        'CustomSubHeading',
        parent=styles['Heading2'],
        fontSize=14,
        fontName='Helvetica-Bold',
        textColor=colors.HexColor('#0A0A0A'),
        spaceAfter=0.15*inch,
        spaceBefore=0.2*inch,
        alignment=TA_LEFT
    )
    
    table_heading_style = ParagraphStyle(
        'TableHeading',
        parent=styles['Normal'],
        fontSize=12,
        textColor=colors.HexColor('#FFFFFF'),
        alignment=TA_CENTER,
        fontName='Helvetica-Bold'
    )
    
    body_style = ParagraphStyle(
        'CustomBody',
        parent=styles['Normal'],
        fontSize=11,
        textColor=colors.HexColor('#4A4A4A'),
        spaceAfter=0.1*inch,
        alignment=TA_LEFT
    )
    
    # Page title
    title = Paragraph("Valuation Summary & Conclusion", title_style)
    elements.append(title)
    elements.append(Spacer(1, 0.2*inch))
    
    # Introduction text
    intro_text = Paragraph(
        "The following football field chart summarizes the valuation ranges derived from multiple methodologies, "
        "providing a comprehensive view of the target company's fair value range.",
        body_style
    )
    elements.append(intro_text)
    elements.append(Spacer(1, 0.3*inch))
    
    # ============================================================================
    # Football Field Valuation Chart (Visual)
    # ============================================================================
    chart_heading = Paragraph("Valuation Football Field", subheading_style)
    elements.append(chart_heading)
    
    # Extract valuation data
    dcf_tg_low = valuation_summary['dcf_terminal_growth']['low']
    dcf_tg_base = valuation_summary['dcf_terminal_growth']['base']
    dcf_tg_high = valuation_summary['dcf_terminal_growth']['high']
    
    dcf_em_low = valuation_summary['dcf_exit_multiple']['low']
    dcf_em_base = valuation_summary['dcf_exit_multiple']['base']
    dcf_em_high = valuation_summary['dcf_exit_multiple']['high']
    
    # Calculate EV/EBITDA range from LTM and FY
    ev_ebitda_ltm = valuation_summary['ev_ebitda_ltm']
    ev_ebitda_fy = valuation_summary['ev_ebitda_fy']
    ev_ebitda_low = min(ev_ebitda_ltm, ev_ebitda_fy)
    ev_ebitda_high = max(ev_ebitda_ltm, ev_ebitda_fy)
    ev_ebitda_avg = (ev_ebitda_ltm + ev_ebitda_fy) / 2
    
    # Calculate P/E range from LTM and FY
    pe_ltm = valuation_summary['pe_ltm']
    pe_fy = valuation_summary['pe_fy']
    pe_low = min(pe_ltm, pe_fy)
    pe_high = max(pe_ltm, pe_fy)
    pe_avg = (pe_ltm + pe_fy) / 2
    
    # Calculate average share price from all base case values
    all_base_values = [dcf_tg_base, dcf_em_base, ev_ebitda_avg, pe_avg]
    average_share_price = sum(all_base_values) / len(all_base_values)
    
    # Find min and max values for scaling
    all_values = [dcf_tg_low, dcf_tg_high, dcf_em_low, dcf_em_high, ev_ebitda_low, ev_ebitda_high, pe_low, pe_high, average_share_price]
    min_val = min(all_values) * 0.9  # 10% padding
    max_val = max(all_values) * 1.1  # 10% padding
    value_range = max_val - min_val
    
    # Create drawing for football field chart
    drawing_width = 6.5 * inch
    drawing_height = 3.5 * inch
    d = Drawing(drawing_width, drawing_height)
    
    # Chart parameters
    chart_left = 2.0 * inch
    chart_width = 4.0 * inch
    bar_height = 0.35 * inch
    bar_spacing = 0.65 * inch
    chart_top = drawing_height - 0.5 * inch
    
    # Helper function to convert value to x position
    def value_to_x(value):
        proportion = (value - min_val) / value_range
        return chart_left + (proportion * chart_width)
    
    # Draw vertical grid lines and labels
    grid_values = [min_val, (min_val + max_val) / 2, max_val]
    for grid_val in grid_values:
        x_pos = value_to_x(grid_val)
        d.add(Line(x_pos, 0.3 * inch, x_pos, chart_top + 0.3 * inch, 
                   strokeColor=colors.HexColor('#CCCCCC'), strokeWidth=0.5, strokeDashArray=[2, 2]))
        d.add(String(x_pos, 0.1 * inch, f'${grid_val:.0f}', 
                    fontSize=8, fillColor=colors.HexColor('#4A4A4A'), textAnchor='middle'))
    
    # Draw average target price line
    avg_x = value_to_x(average_share_price)
    d.add(Line(avg_x, 0.3 * inch, avg_x, chart_top + 0.3 * inch,
               strokeColor=colors.HexColor('#FF6B6B'), strokeWidth=2))
    d.add(String(avg_x, chart_top + 0.45 * inch, f'Avg Target: ${average_share_price:.2f}',
                fontSize=8, fillColor=colors.HexColor('#FF6B6B'), textAnchor='middle', fontName='Helvetica-Bold'))
    
    # Method 1: DCF Terminal Growth
    y_pos = chart_top - (0 * bar_spacing)
    x_start = value_to_x(dcf_tg_low)
    x_end = value_to_x(dcf_tg_high)
    x_base = value_to_x(dcf_tg_base)
    
    d.add(Rect(x_start, y_pos - bar_height/2, x_end - x_start, bar_height,
               fillColor=colors.HexColor('#124734'), strokeColor=colors.HexColor('#0A0A0A'), strokeWidth=1))
    d.add(Line(x_base, y_pos - bar_height/2 - 0.05*inch, x_base, y_pos + bar_height/2 + 0.05*inch,
               strokeColor=colors.HexColor('#FFD700'), strokeWidth=3))
    d.add(String(0.2 * inch, y_pos, 'DCF - Terminal Growth',
                fontSize=9, fillColor=colors.HexColor('#0A0A0A'), textAnchor='start'))
    d.add(String(x_end + 0.15 * inch, y_pos, f'${dcf_tg_low:.0f} - ${dcf_tg_high:.0f}',
                fontSize=8, fillColor=colors.HexColor('#4A4A4A'), textAnchor='start'))
    
    # Method 2: DCF Exit Multiple
    y_pos = chart_top - (1 * bar_spacing)
    x_start = value_to_x(dcf_em_low)
    x_end = value_to_x(dcf_em_high)
    x_base = value_to_x(dcf_em_base)
    
    d.add(Rect(x_start, y_pos - bar_height/2, x_end - x_start, bar_height,
               fillColor=colors.HexColor('#124734'), strokeColor=colors.HexColor('#0A0A0A'), strokeWidth=1))
    d.add(Line(x_base, y_pos - bar_height/2 - 0.05*inch, x_base, y_pos + bar_height/2 + 0.05*inch,
               strokeColor=colors.HexColor('#FFD700'), strokeWidth=3))
    d.add(String(0.2 * inch, y_pos, 'DCF - Exit Multiple',
                fontSize=9, fillColor=colors.HexColor('#0A0A0A'), textAnchor='start'))
    d.add(String(x_end + 0.15 * inch, y_pos, f'${dcf_em_low:.0f} - ${dcf_em_high:.0f}',
                fontSize=8, fillColor=colors.HexColor('#4A4A4A'), textAnchor='start'))
    
    # Method 3: EV/EBITDA (range from LTM and FY)
    y_pos = chart_top - (2 * bar_spacing)
    x_start = value_to_x(ev_ebitda_low)
    x_end = value_to_x(ev_ebitda_high)
    x_avg = value_to_x(ev_ebitda_avg)
    
    d.add(Rect(x_start, y_pos - bar_height/2, x_end - x_start, bar_height,
               fillColor=colors.HexColor('#2E6F40'), strokeColor=colors.HexColor('#0A0A0A'), strokeWidth=1))
    d.add(Line(x_avg, y_pos - bar_height/2 - 0.05*inch, x_avg, y_pos + bar_height/2 + 0.05*inch,
               strokeColor=colors.HexColor('#FFD700'), strokeWidth=3))
    d.add(String(0.2 * inch, y_pos, 'EV/EBITDA (LTM & FY)',
                fontSize=9, fillColor=colors.HexColor('#0A0A0A'), textAnchor='start'))
    d.add(String(x_end + 0.15 * inch, y_pos, f'${ev_ebitda_low:.0f} - ${ev_ebitda_high:.0f}',
                fontSize=8, fillColor=colors.HexColor('#4A4A4A'), textAnchor='start'))
    
    # Method 4: P/E (range from LTM and FY)
    y_pos = chart_top - (3 * bar_spacing)
    x_start = value_to_x(pe_low)
    x_end = value_to_x(pe_high)
    x_avg = value_to_x(pe_avg)
    
    d.add(Rect(x_start, y_pos - bar_height/2, x_end - x_start, bar_height,
               fillColor=colors.HexColor('#2E6F40'), strokeColor=colors.HexColor('#0A0A0A'), strokeWidth=1))
    d.add(Line(x_avg, y_pos - bar_height/2 - 0.05*inch, x_avg, y_pos + bar_height/2 + 0.05*inch,
               strokeColor=colors.HexColor('#FFD700'), strokeWidth=3))
    d.add(String(0.2 * inch, y_pos, 'P/E (LTM & FY)',
                fontSize=9, fillColor=colors.HexColor('#0A0A0A'), textAnchor='start'))
    d.add(String(x_end + 0.4 * inch, y_pos, f'${pe_low:.0f} - ${pe_high:.0f}',
                fontSize=8, fillColor=colors.HexColor('#4A4A4A'), textAnchor='start'))
    
    elements.append(d)
    elements.append(Spacer(1, 0.3*inch))
    
    # Add legend
    legend_text = Paragraph(
        '<b>Legend:</b> Green bars = Valuation ranges | Yellow line = Base case | Red line = Average target price',
        body_style
    )
    elements.append(legend_text)
    elements.append(Spacer(1, 0.4*inch))
    
    # Create football field summary table
    football_field_data = [
        [
            Paragraph('Valuation Method', table_heading_style),
            Paragraph('Low', table_heading_style),
            Paragraph('Base Case', table_heading_style),
            Paragraph('High', table_heading_style),
            Paragraph('Range', table_heading_style)
        ]
    ]
    
    # Data already extracted above
    dcf_tg_range = f"${dcf_tg_low:.2f} - ${dcf_tg_high:.2f}"
    
    football_field_data.append([
        'DCF - Terminal Growth',
        f'${dcf_tg_low:.2f}',
        f'${dcf_tg_base:.2f}',
        f'${dcf_tg_high:.2f}',
        dcf_tg_range
    ])
    
    # Data already extracted above
    dcf_em_range = f"${dcf_em_low:.2f} - ${dcf_em_high:.2f}"
    
    football_field_data.append([
        'DCF - Exit Multiple',
        f'${dcf_em_low:.2f}',
        f'${dcf_em_base:.2f}',
        f'${dcf_em_high:.2f}',
        dcf_em_range
    ])
    
    ev_ebitda_range = f"${ev_ebitda_low:.2f} - ${ev_ebitda_high:.2f}"
    
    football_field_data.append([
        'EV/EBITDA (LTM & FY)',
        f'${ev_ebitda_low:.2f}',
        f'${ev_ebitda_avg:.2f}',
        f'${ev_ebitda_high:.2f}',
        ev_ebitda_range
    ])
    
    pe_range = f"${pe_low:.2f} - ${pe_high:.2f}"
    
    football_field_data.append([
        'P/E (LTM & FY)',
        f'${pe_low:.2f}',
        f'${pe_avg:.2f}',
        f'${pe_high:.2f}',
        pe_range
    ])
    
    football_field_table = Table(football_field_data, colWidths=[2.2*inch, 1.2*inch, 1.2*inch, 1.2*inch, 1.5*inch])
    football_field_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#124734')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),  # Method names left-aligned
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTNAME', (2, 1), (2, -1), 'Helvetica-Bold'),  # Bold Base Case column
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.HexColor('#4A4A4A')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#F2F2EB'), colors.white]),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#CCCCCC')),
    ]))
    elements.append(football_field_table)
    elements.append(Spacer(1, 0.4*inch))
    
    # ============================================================================
    # Average Valuation Summary
    # ============================================================================
    summary_heading = Paragraph("Weighted Average Valuation", subheading_style)
    elements.append(summary_heading)
    
    # average_share_price already calculated above from all base case values
    
    # Calculate upside/downside vs current price
    upside_downside = ((average_share_price - current_share_price) / current_share_price) * 100
    
    # Create summary table
    summary_data = [
        [
            Paragraph('Metric', table_heading_style),
            Paragraph('Value', table_heading_style)
        ]
    ]
    
    summary_data.append(['Current Share Price', f'${current_share_price:.2f}'])
    summary_data.append(['Average Target Price', f'${average_share_price:.2f}'])
    summary_data.append(['Implied Upside / (Downside)', f'{upside_downside:+.2f}%'])
    
    summary_table = Table(summary_data, colWidths=[3*inch, 2.5*inch])
    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#124734')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTNAME', (1, 2), (1, 2), 'Helvetica-Bold'),  # Bold the upside/downside
        ('FONTSIZE', (0, 0), (-1, -1), 11),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.HexColor('#4A4A4A')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#F2F2EB'), colors.white]),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
        ('TOPPADDING', (0, 0), (-1, -1), 10),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#CCCCCC')),
    ]))
    elements.append(summary_table)
    elements.append(Spacer(1, 0.3*inch))
    
    # ============================================================================
    # Conclusion Note
    # ============================================================================
    conclusion_heading = Paragraph("Investment Recommendation", subheading_style)
    elements.append(conclusion_heading)
    
    # Determine recommendation based on upside/downside
    if upside_downside > 20:
        recommendation = "BUY"
        recommendation_text = f"With an implied upside of {upside_downside:.1f}%, the stock appears <b>undervalued</b> based on our analysis."
    elif upside_downside > 10:
        recommendation = "ACCUMULATE"
        recommendation_text = f"With an implied upside of {upside_downside:.1f}%, the stock shows <b>moderate upside potential</b>."
    elif upside_downside > -10:
        recommendation = "HOLD"
        recommendation_text = f"With an implied return of {upside_downside:+.1f}%, the stock is trading near <b>fair value</b>."
    else:
        recommendation = "REDUCE"
        recommendation_text = f"With an implied downside of {upside_downside:.1f}%, the stock appears <b>overvalued</b> based on our analysis."
    
    recommendation_para = Paragraph(
        f"<b>Recommendation: {recommendation}</b><br/><br/>{recommendation_text}",
        body_style
    )
    elements.append(recommendation_para)
    
    elements.append(PageBreak())


def appendix_a(
    doc,
    elements,
    forecast_years_data,
    terminal_growth_summary,
    exit_multiple_summary,
    peer_data,
    summary_stats,
    implied_valuations,
    valuation_summary,
    current_share_price,
    figures_stated_in="millions"
):
    """
    Create Appendix A (Bull Case) with detailed valuation analysis.
    
    Parameters:
    -----------
    doc : SimpleDocTemplate
        The PDF document object
    elements : list
        List to store document elements
    forecast_years_data : list of dict
        Each dict should include: 'year', 'ebit_1_minus_t', 'da', 'capex',
        'change_nwc', and 'fcff' (all precomputed values).
    terminal_growth_summary : dict
        Precomputed values for terminal growth method
    exit_multiple_summary : dict
        Precomputed values for exit multiple method
    peer_data : list of dicts
        List of peer company data
    summary_stats : dict
        Summary statistics for each multiple
    implied_valuations : dict
        Pre-calculated implied valuations
    valuation_summary : dict
        Dictionary containing valuation results from different methods
    current_share_price : float
        Current market share price
    figures_stated_in : str
        Format of figures stated in (thousands, millions, billions)
    """
    
    styles = getSampleStyleSheet()
    
    # Define custom styles
    title_style = ParagraphStyle(
        'AppendixTitle',
        parent=styles['Heading1'],
        fontSize=20,
        textColor=colors.HexColor('#0A0A0A'),
        spaceAfter=0.3*inch,
        alignment=TA_CENTER,
        fontName='Helvetica-Bold'
    )
    
    appendix_heading_style = ParagraphStyle(
        'AppendixHeading',
        parent=styles['Heading2'],
        fontSize=16,
        fontName='Helvetica-Bold',
        textColor=colors.HexColor('#0A0A0A'),
        spaceAfter=0.2*inch,
        spaceBefore=0.3*inch,
        alignment=TA_LEFT
    )
    
    subheading_style = ParagraphStyle(
        'AppendixSubHeading',
        parent=styles['Heading3'],
        fontSize=13,
        fontName='Helvetica-Bold',
        textColor=colors.HexColor('#0A0A0A'),
        spaceAfter=0.15*inch,
        spaceBefore=0.2*inch,
        alignment=TA_LEFT
    )
    
    table_heading_style = ParagraphStyle(
        'TableHeading',
        parent=styles['Normal'],
        fontSize=10,
        textColor=colors.HexColor('#FFFFFF'),
        alignment=TA_CENTER,
        fontName='Helvetica-Bold'
    )
    
    body_style = ParagraphStyle(
        'CustomBody',
        parent=styles['Normal'],
        fontSize=10,
        textColor=colors.HexColor('#4A4A4A'),
        spaceAfter=0.1*inch,
        alignment=TA_LEFT
    )
    
    # Helper functions
    def _fmt_money(value, decimals=0):
        try:
            if value is None:
                return "N/A"
            num = float(value)
            if num != num:
                return "N/A"
            return f"${num:,.{decimals}f}"
        except (TypeError, ValueError):
            return "N/A"

    def _get_suffix(figures_stated_in):
        """Get appropriate suffix based on figures_stated_in parameter."""
        if isinstance(figures_stated_in, str):
            if 'thousand' in figures_stated_in.lower():
                return 'Th'
            elif 'billion' in figures_stated_in.lower():
                return 'B'
            elif 'million' in figures_stated_in.lower():
                return 'M'
        return 'M'  # Default to millions

    def _fmt_money_m(value, decimals=0):
        formatted = _fmt_money(value, decimals)
        suffix = _get_suffix(figures_stated_in)
        return formatted if formatted == "N/A" else f"{formatted} {suffix}"

    def _fmt_price(value):
        return _fmt_money(value, decimals=2)

    def _fmt_float(value, decimals=2, suffix=""):
        try:
            if value is None:
                return "N/A"
            num = float(value)
            if num != num:
                return "N/A"
            return f"{num:.{decimals}f}{suffix}"
        except (TypeError, ValueError):
            return "N/A"

    def _fmt_currency_with_suffix(value, decimals=0):
        """Format currency with dynamic suffix."""
        try:
            if value is None:
                return "N/A"
            num = float(value)
            if num != num:
                return "N/A"
            suffix = _get_suffix(figures_stated_in)
            return f"${num:,.{decimals}f} {suffix}"
        except (TypeError, ValueError):
            return "N/A"

    def _fmt_float_with_suffix(value, decimals=1):
        """Format float with dynamic suffix based on figures_stated_in."""
        if value is None:
            return "N/A"
        try:
            num = float(value)
            if num != num:
                return "N/A"
            suffix = _get_suffix(figures_stated_in)
            return f"{num:,.{decimals}f} {suffix}"
        except (TypeError, ValueError):
            return "N/A"
    
    # ============================================================================
    # Main Heading
    # ============================================================================
    title = Paragraph("Appendices", title_style)
    elements.append(title)
    elements.append(Spacer(1, 0.3*inch))
    
    # ============================================================================
    # Appendix A - Bull Case Heading
    # ============================================================================
    appendix_heading = Paragraph("Appendix A - Bull Case", appendix_heading_style)
    elements.append(appendix_heading)
    elements.append(Spacer(1, 0.2*inch))
    
    # ============================================================================
    # Intrinsic Valuation - FCFF Only
    # ============================================================================
    intrinsic_heading = Paragraph("Intrinsic Valuation", subheading_style)
    elements.append(intrinsic_heading)
    
    if forecast_years_data is None:
        forecast_years_data = []
    
    # Build FCF table with FCFF only
    years = [item.get("year", "N/A") for item in forecast_years_data]
    if not years:
        years = ["N/A"]
    
    fcf_table_data = [
        ['Metric'] + [f'Year {year}' for year in years]
    ]
    
    # FCFF row only - no intermediate calculations
    fcff_row = ['FCFF'] + [_fmt_money(i.get("fcff"), decimals=0) for i in forecast_years_data]
    fcf_table_data.append(fcff_row)
    
    # Create column widths dynamically
    years_count = len(years)
    col_width = 5.5 / (years_count + 1)
    fcf_col_widths = [1.5*inch] + [col_width*inch for _ in range(years_count)]
    
    fcf_table = Table(fcf_table_data, colWidths=fcf_col_widths)
    fcf_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#124734')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.HexColor('#4A4A4A')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#F2F2EB'), colors.white]),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#CCCCCC')),
    ]))
    elements.append(fcf_table)
    elements.append(Spacer(1, 0.25*inch))
    
    # ============================================================================
    # DCF Valuation - Terminal Growth Method
    # ============================================================================
    tgm_heading = Paragraph("DCF Valuation - Terminal Growth Method", subheading_style)
    elements.append(tgm_heading)
    
    terminal_growth_summary = terminal_growth_summary or {}
    
    # Valuation summary table for Terminal Growth Method
    valuation_tgm_data = [
        [Paragraph('Metric', table_heading_style), Paragraph('Value', table_heading_style)]
    ]
    valuation_tgm_data.append(['Terminal Value', _fmt_money_m(terminal_growth_summary.get('terminal_value'), decimals=0)])
    valuation_tgm_data.append(['Enterprise Value', _fmt_money_m(terminal_growth_summary.get('enterprise_value'), decimals=0)])
    valuation_tgm_data.append(['Less: Net Debt', _fmt_money_m(terminal_growth_summary.get('net_debt'), decimals=0)])
    valuation_tgm_data.append(['Total Equity Value', _fmt_money_m(terminal_growth_summary.get('total_equity_value'), decimals=0)])
    valuation_tgm_data.append(['Shares Outstanding', _fmt_float_with_suffix(terminal_growth_summary.get('shares_outstanding'), decimals=1)])
    valuation_tgm_data.append(['Implied Share Price', _fmt_price(terminal_growth_summary.get('implied_share_price'))])
    
    valuation_tgm_table = Table(valuation_tgm_data, colWidths=[3.5*inch, 2*inch])
    valuation_tgm_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#124734')),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('ALIGN', (1, 0), (1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.HexColor('#4A4A4A')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#F2F2EB'), colors.white]),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#CCCCCC')),
    ]))
    elements.append(valuation_tgm_table)
    elements.append(Spacer(1, 0.25*inch))
    
    # ============================================================================
    # DCF Valuation - Exit Multiple Method
    # ============================================================================
    em_heading = Paragraph("DCF Valuation - Exit Multiple Method", subheading_style)
    elements.append(em_heading)
    
    exit_multiple_summary = exit_multiple_summary or {}
    
    # Valuation summary table for Exit Multiple Method
    valuation_em_data = [
        [Paragraph('Metric', table_heading_style), Paragraph('Value', table_heading_style)]
    ]
    valuation_em_data.append(['Terminal Value', _fmt_money_m(exit_multiple_summary.get('terminal_value'), decimals=0)])
    valuation_em_data.append(['Enterprise Value', _fmt_money_m(exit_multiple_summary.get('enterprise_value'), decimals=0)])
    valuation_em_data.append(['Less: Net Debt', _fmt_money_m(exit_multiple_summary.get('net_debt'), decimals=0)])
    valuation_em_data.append(['Total Equity Value', _fmt_money_m(exit_multiple_summary.get('total_equity_value'), decimals=0)])
    valuation_em_data.append(['Shares Outstanding', _fmt_float_with_suffix(exit_multiple_summary.get('shares_outstanding'), decimals=1)])
    valuation_em_data.append(['Implied Share Price', _fmt_price(exit_multiple_summary.get('implied_share_price'))])
    
    valuation_em_table = Table(valuation_em_data, colWidths=[3.5*inch, 2*inch])
    valuation_em_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#124734')),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('ALIGN', (1, 0), (1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.HexColor('#4A4A4A')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#F2F2EB'), colors.white]),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#CCCCCC')),
    ]))
    elements.append(valuation_em_table)
    elements.append(Spacer(1, 0.25*inch))
    
    # ============================================================================
    # Relative Valuation - Comp Analysis
    # ============================================================================
    comp_heading = Paragraph("Relative Valuation", subheading_style)
    elements.append(comp_heading)
    
    # Create comp analysis table showing share price, equity value, and enterprise value
    comp_table_data = [
        [
            Paragraph('Method', table_heading_style),
            Paragraph('Share Price', table_heading_style),
            Paragraph('Equity Value', table_heading_style),
            Paragraph('Enterprise Value', table_heading_style)
        ]
    ]
    
    comp_table_data.append([
        'EV/EBITDA LTM',
        f'${implied_valuations["ev_ebitda_ltm"]["share_price"]:.2f}',
        _fmt_currency_with_suffix(implied_valuations["ev_ebitda_ltm"]["equity_value"], decimals=0),
        _fmt_currency_with_suffix(implied_valuations["ev_ebitda_ltm"]["enterprise_value"], decimals=0)
    ])
    
    comp_table_data.append([
        'EV/EBITDA FY',
        f'${implied_valuations["ev_ebitda_fy"]["share_price"]:.2f}',
        _fmt_currency_with_suffix(implied_valuations["ev_ebitda_fy"]["equity_value"], decimals=0),
        _fmt_currency_with_suffix(implied_valuations["ev_ebitda_fy"]["enterprise_value"], decimals=0)
    ])
    
    comp_table_data.append([
        'P/E LTM',
        f'${implied_valuations["pe_ltm"]["share_price"]:.2f}',
        _fmt_currency_with_suffix(implied_valuations["pe_ltm"]["equity_value"], decimals=0),
        _fmt_currency_with_suffix(implied_valuations["pe_ltm"]["enterprise_value"], decimals=0)
    ])
    
    comp_table_data.append([
        'P/E FY',
        f'${implied_valuations["pe_fy"]["share_price"]:.2f}',
        _fmt_currency_with_suffix(implied_valuations["pe_fy"]["equity_value"], decimals=0),
        _fmt_currency_with_suffix(implied_valuations["pe_fy"]["enterprise_value"], decimals=0)
    ])
    
    comp_table = Table(comp_table_data, colWidths=[1.8*inch, 1.5*inch, 1.8*inch, 1.8*inch])
    comp_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#124734')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.HexColor('#4A4A4A')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#F2F2EB'), colors.white]),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#CCCCCC')),
    ]))
    elements.append(comp_table)
    elements.append(Spacer(1, 0.25*inch))
    
    # ============================================================================
    # Summary - Football Field Chart and Average Price
    # ============================================================================
    summary_heading = Paragraph("Summary", subheading_style)
    elements.append(summary_heading)
    
    # Extract valuation data for football field
    dcf_tg_low = valuation_summary['dcf_terminal_growth']['low']
    dcf_tg_base = valuation_summary['dcf_terminal_growth']['base']
    dcf_tg_high = valuation_summary['dcf_terminal_growth']['high']
    
    dcf_em_low = valuation_summary['dcf_exit_multiple']['low']
    dcf_em_base = valuation_summary['dcf_exit_multiple']['base']
    dcf_em_high = valuation_summary['dcf_exit_multiple']['high']
    
    # Calculate EV/EBITDA range
    ev_ebitda_ltm = valuation_summary['ev_ebitda_ltm']
    ev_ebitda_fy = valuation_summary['ev_ebitda_fy']
    ev_ebitda_low = min(ev_ebitda_ltm, ev_ebitda_fy)
    ev_ebitda_high = max(ev_ebitda_ltm, ev_ebitda_fy)
    ev_ebitda_avg = (ev_ebitda_ltm + ev_ebitda_fy) / 2
    
    # Calculate P/E range
    pe_ltm = valuation_summary['pe_ltm']
    pe_fy = valuation_summary['pe_fy']
    pe_low = min(pe_ltm, pe_fy)
    pe_high = max(pe_ltm, pe_fy)
    pe_avg = (pe_ltm + pe_fy) / 2
    
    # Calculate average share price from all base case values
    all_base_values = [dcf_tg_base, dcf_em_base, ev_ebitda_avg, pe_avg]
    average_share_price = sum(all_base_values) / len(all_base_values)
    
    # Find min and max values for scaling
    all_values = [dcf_tg_low, dcf_tg_high, dcf_em_low, dcf_em_high, ev_ebitda_low, ev_ebitda_high, pe_low, pe_high, average_share_price]
    min_val = min(all_values) * 0.9
    max_val = max(all_values) * 1.1
    value_range = max_val - min_val
    
    # Create football field chart
    drawing_width = 6.5 * inch
    drawing_height = 3.0 * inch
    d = Drawing(drawing_width, drawing_height)
    
    # Chart parameters
    chart_left = 2.0 * inch
    chart_width = 4.0 * inch
    bar_height = 0.30 * inch
    bar_spacing = 0.55 * inch
    chart_top = drawing_height - 0.4 * inch
    
    # Helper function to convert value to x position
    def value_to_x(value):
        proportion = (value - min_val) / value_range
        return chart_left + (proportion * chart_width)
    
    # Draw vertical grid lines
    grid_values = [min_val, (min_val + max_val) / 2, max_val]
    for grid_val in grid_values:
        x_pos = value_to_x(grid_val)
        d.add(Line(x_pos, 0.25 * inch, x_pos, chart_top + 0.25 * inch,
                   strokeColor=colors.HexColor('#CCCCCC'), strokeWidth=0.5, strokeDashArray=[2, 2]))
        d.add(String(x_pos, 0.08 * inch, f'${grid_val:.0f}',
                    fontSize=8, fillColor=colors.HexColor('#4A4A4A'), textAnchor='middle'))
    
    # Draw average target price line
    avg_x = value_to_x(average_share_price)
    d.add(Line(avg_x, 0.25 * inch, avg_x, chart_top + 0.25 * inch,
               strokeColor=colors.HexColor('#FF6B6B'), strokeWidth=2))
    d.add(String(avg_x, chart_top + 0.35 * inch, f'Avg: ${average_share_price:.2f}',
                fontSize=8, fillColor=colors.HexColor('#FF6B6B'), textAnchor='middle', fontName='Helvetica-Bold'))
    
    # Method 1: DCF Terminal Growth
    y_pos = chart_top - (0 * bar_spacing)
    x_start = value_to_x(dcf_tg_low)
    x_end = value_to_x(dcf_tg_high)
    x_base = value_to_x(dcf_tg_base)
    
    d.add(Rect(x_start, y_pos - bar_height/2, x_end - x_start, bar_height,
               fillColor=colors.HexColor('#124734'), strokeColor=colors.HexColor('#0A0A0A'), strokeWidth=1))
    d.add(Line(x_base, y_pos - bar_height/2 - 0.04*inch, x_base, y_pos + bar_height/2 + 0.04*inch,
               strokeColor=colors.HexColor('#FFD700'), strokeWidth=2.5))
    d.add(String(0.2 * inch, y_pos, 'DCF-TG',
                fontSize=8, fillColor=colors.HexColor('#0A0A0A'), textAnchor='start'))
    
    # Method 2: DCF Exit Multiple
    y_pos = chart_top - (1 * bar_spacing)
    x_start = value_to_x(dcf_em_low)
    x_end = value_to_x(dcf_em_high)
    x_base = value_to_x(dcf_em_base)
    
    d.add(Rect(x_start, y_pos - bar_height/2, x_end - x_start, bar_height,
               fillColor=colors.HexColor('#124734'), strokeColor=colors.HexColor('#0A0A0A'), strokeWidth=1))
    d.add(Line(x_base, y_pos - bar_height/2 - 0.04*inch, x_base, y_pos + bar_height/2 + 0.04*inch,
               strokeColor=colors.HexColor('#FFD700'), strokeWidth=2.5))
    d.add(String(0.2 * inch, y_pos, 'DCF-EM',
                fontSize=8, fillColor=colors.HexColor('#0A0A0A'), textAnchor='start'))
    
    # Method 3: EV/EBITDA
    y_pos = chart_top - (2 * bar_spacing)
    x_start = value_to_x(ev_ebitda_low)
    x_end = value_to_x(ev_ebitda_high)
    x_avg = value_to_x(ev_ebitda_avg)
    
    d.add(Rect(x_start, y_pos - bar_height/2, x_end - x_start, bar_height,
               fillColor=colors.HexColor('#2E6F40'), strokeColor=colors.HexColor('#0A0A0A'), strokeWidth=1))
    d.add(Line(x_avg, y_pos - bar_height/2 - 0.04*inch, x_avg, y_pos + bar_height/2 + 0.04*inch,
               strokeColor=colors.HexColor('#FFD700'), strokeWidth=2.5))
    d.add(String(0.2 * inch, y_pos, 'EV/EBITDA',
                fontSize=8, fillColor=colors.HexColor('#0A0A0A'), textAnchor='start'))
    
    # Method 4: P/E
    y_pos = chart_top - (3 * bar_spacing)
    x_start = value_to_x(pe_low)
    x_end = value_to_x(pe_high)
    x_avg = value_to_x(pe_avg)
    
    d.add(Rect(x_start, y_pos - bar_height/2, x_end - x_start, bar_height,
               fillColor=colors.HexColor('#2E6F40'), strokeColor=colors.HexColor('#0A0A0A'), strokeWidth=1))
    d.add(Line(x_avg, y_pos - bar_height/2 - 0.04*inch, x_avg, y_pos + bar_height/2 + 0.04*inch,
               strokeColor=colors.HexColor('#FFD700'), strokeWidth=2.5))
    d.add(String(0.2 * inch, y_pos, 'P/E',
                fontSize=8, fillColor=colors.HexColor('#0A0A0A'), textAnchor='start'))
    
    elements.append(d)
    elements.append(Spacer(1, 0.2*inch))
    
    # Add summary text with average price
    summary_text = Paragraph(
        f"<b>Average Target Price: ${average_share_price:.2f}</b><br/>"
        f"Implied Upside / (Downside) vs Current Price: {((average_share_price - current_share_price) / current_share_price) * 100:+.2f}%",
        body_style
    )
    elements.append(summary_text)
    elements.append(Spacer(1, 0.3*inch))


def appendix_b(
    doc,
    elements,
    forecast_years_data,
    terminal_growth_summary,
    exit_multiple_summary,
    peer_data,
    summary_stats,
    implied_valuations,
    valuation_summary,
    current_share_price,
    figures_stated_in="millions"
):
    """
    Create Appendix B (Bear Case) with detailed valuation analysis.
    
    Parameters:
    -----------
    Same as appendix_a, but for Bear Case scenario with figures_stated_in parameter
    """
    
    styles = getSampleStyleSheet()
    
    # Define custom styles
    appendix_heading_style = ParagraphStyle(
        'AppendixHeading',
        parent=styles['Heading2'],
        fontSize=16,
        fontName='Helvetica-Bold',
        textColor=colors.HexColor('#0A0A0A'),
        spaceAfter=0.2*inch,
        spaceBefore=0.3*inch,
        alignment=TA_LEFT
    )
    
    subheading_style = ParagraphStyle(
        'AppendixSubHeading',
        parent=styles['Heading3'],
        fontSize=13,
        fontName='Helvetica-Bold',
        textColor=colors.HexColor('#0A0A0A'),
        spaceAfter=0.15*inch,
        spaceBefore=0.2*inch,
        alignment=TA_LEFT
    )
    
    table_heading_style = ParagraphStyle(
        'TableHeading',
        parent=styles['Normal'],
        fontSize=10,
        textColor=colors.HexColor('#FFFFFF'),
        alignment=TA_CENTER,
        fontName='Helvetica-Bold'
    )
    
    body_style = ParagraphStyle(
        'CustomBody',
        parent=styles['Normal'],
        fontSize=10,
        textColor=colors.HexColor('#4A4A4A'),
        spaceAfter=0.1*inch,
        alignment=TA_LEFT
    )
    
    # Helper functions
    def _fmt_money(value, decimals=0):
        try:
            if value is None:
                return "N/A"
            num = float(value)
            if num != num:
                return "N/A"
            return f"${num:,.{decimals}f}"
        except (TypeError, ValueError):
            return "N/A"

    def _get_suffix(figures_stated_in):
        """Get appropriate suffix based on figures_stated_in parameter."""
        if isinstance(figures_stated_in, str):
            if 'thousand' in figures_stated_in.lower():
                return 'Th'
            elif 'billion' in figures_stated_in.lower():
                return 'B'
            elif 'million' in figures_stated_in.lower():
                return 'M'
        return 'M'  # Default to millions

    def _fmt_money_m(value, decimals=0):
        formatted = _fmt_money(value, decimals)
        suffix = _get_suffix(figures_stated_in)
        return formatted if formatted == "N/A" else f"{formatted} {suffix}"

    def _fmt_price(value):
        return _fmt_money(value, decimals=2)

    def _fmt_float(value, decimals=2, suffix=""):
        try:
            if value is None:
                return "N/A"
            num = float(value)
            if num != num:
                return "N/A"
            return f"{num:.{decimals}f}{suffix}"
        except (TypeError, ValueError):
            return "N/A"

    def _fmt_currency_with_suffix(value, decimals=0):
        """Format currency with dynamic suffix."""
        try:
            if value is None:
                return "N/A"
            num = float(value)
            if num != num:
                return "N/A"
            suffix = _get_suffix(figures_stated_in)
            return f"${num:,.{decimals}f} {suffix}"
        except (TypeError, ValueError):
            return "N/A"

    def _fmt_float_with_suffix(value, decimals=1):
        """Format float with dynamic suffix based on figures_stated_in."""
        if value is None:
            return "N/A"
        try:
            num = float(value)
            if num != num:
                return "N/A"
            suffix = _get_suffix(figures_stated_in)
            return f"{num:,.{decimals}f} {suffix}"
        except (TypeError, ValueError):
            return "N/A"
    
    # ============================================================================
    # Appendix B - Bear Case Heading
    # ============================================================================
    appendix_heading = Paragraph("Appendix B - Bear Case", appendix_heading_style)
    elements.append(appendix_heading)
    elements.append(Spacer(1, 0.2*inch))
    
    # ============================================================================
    # Intrinsic Valuation - FCFF Only
    # ============================================================================
    intrinsic_heading = Paragraph("Intrinsic Valuation", subheading_style)
    elements.append(intrinsic_heading)
    
    if forecast_years_data is None:
        forecast_years_data = []
    
    # Build FCF table with FCFF only
    years = [item.get("year", "N/A") for item in forecast_years_data]
    if not years:
        years = ["N/A"]
    
    fcf_table_data = [
        ['Metric'] + [f'Year {year}' for year in years]
    ]
    
    # FCFF row only - no intermediate calculations
    fcff_row = ['FCFF'] + [_fmt_money(i.get("fcff"), decimals=0) for i in forecast_years_data]
    fcf_table_data.append(fcff_row)
    
    # Create column widths dynamically
    years_count = len(years)
    col_width = 5.5 / (years_count + 1)
    fcf_col_widths = [1.5*inch] + [col_width*inch for _ in range(years_count)]
    
    fcf_table = Table(fcf_table_data, colWidths=fcf_col_widths)
    fcf_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#124734')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.HexColor('#4A4A4A')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#F2F2EB'), colors.white]),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#CCCCCC')),
    ]))
    elements.append(fcf_table)
    elements.append(Spacer(1, 0.25*inch))
    
    # ============================================================================
    # DCF Valuation - Terminal Growth Method
    # ============================================================================
    tgm_heading = Paragraph("DCF Valuation - Terminal Growth Method", subheading_style)
    elements.append(tgm_heading)
    
    terminal_growth_summary = terminal_growth_summary or {}
    
    # Valuation summary table for Terminal Growth Method
    valuation_tgm_data = [
        [Paragraph('Metric', table_heading_style), Paragraph('Value', table_heading_style)]
    ]
    valuation_tgm_data.append(['Terminal Value', _fmt_money_m(terminal_growth_summary.get('terminal_value'), decimals=0)])
    valuation_tgm_data.append(['Enterprise Value', _fmt_money_m(terminal_growth_summary.get('enterprise_value'), decimals=0)])
    valuation_tgm_data.append(['Less: Net Debt', _fmt_money_m(terminal_growth_summary.get('net_debt'), decimals=0)])
    valuation_tgm_data.append(['Total Equity Value', _fmt_money_m(terminal_growth_summary.get('total_equity_value'), decimals=0)])
    valuation_tgm_data.append(['Shares Outstanding', _fmt_float_with_suffix(terminal_growth_summary.get('shares_outstanding'), decimals=1)])
    valuation_tgm_data.append(['Implied Share Price', _fmt_price(terminal_growth_summary.get('implied_share_price'))])
    
    valuation_tgm_table = Table(valuation_tgm_data, colWidths=[3.5*inch, 2*inch])
    valuation_tgm_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#124734')),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('ALIGN', (1, 0), (1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.HexColor('#4A4A4A')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#F2F2EB'), colors.white]),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#CCCCCC')),
    ]))
    elements.append(valuation_tgm_table)
    elements.append(Spacer(1, 0.25*inch))
    
    # ============================================================================
    # DCF Valuation - Exit Multiple Method
    # ============================================================================
    em_heading = Paragraph("DCF Valuation - Exit Multiple Method", subheading_style)
    elements.append(em_heading)
    
    exit_multiple_summary = exit_multiple_summary or {}
    
    # Valuation summary table for Exit Multiple Method
    valuation_em_data = [
        [Paragraph('Metric', table_heading_style), Paragraph('Value', table_heading_style)]
    ]
    valuation_em_data.append(['Terminal Value', _fmt_money_m(exit_multiple_summary.get('terminal_value'), decimals=0)])
    valuation_em_data.append(['Enterprise Value', _fmt_money_m(exit_multiple_summary.get('enterprise_value'), decimals=0)])
    valuation_em_data.append(['Less: Net Debt', _fmt_money_m(exit_multiple_summary.get('net_debt'), decimals=0)])
    valuation_em_data.append(['Total Equity Value', _fmt_money_m(exit_multiple_summary.get('total_equity_value'), decimals=0)])
    valuation_em_data.append(['Shares Outstanding', _fmt_float_with_suffix(exit_multiple_summary.get('shares_outstanding'), decimals=1)])
    valuation_em_data.append(['Implied Share Price', _fmt_price(exit_multiple_summary.get('implied_share_price'))])
    
    valuation_em_table = Table(valuation_em_data, colWidths=[3.5*inch, 2*inch])
    valuation_em_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#124734')),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('ALIGN', (1, 0), (1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.HexColor('#4A4A4A')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#F2F2EB'), colors.white]),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#CCCCCC')),
    ]))
    elements.append(valuation_em_table)
    elements.append(Spacer(1, 0.25*inch))
    
    # ============================================================================
    # Relative Valuation - Comp Analysis
    # ============================================================================
    comp_heading = Paragraph("Relative Valuation", subheading_style)
    elements.append(comp_heading)
    
    # Create comp analysis table showing share price, equity value, and enterprise value
    comp_table_data = [
        [
            Paragraph('Method', table_heading_style),
            Paragraph('Share Price', table_heading_style),
            Paragraph('Equity Value', table_heading_style),
            Paragraph('Enterprise Value', table_heading_style)
        ]
    ]
    
    comp_table_data.append([
        'EV/EBITDA LTM',
        f'${implied_valuations["ev_ebitda_ltm"]["share_price"]:.2f}',
        _fmt_currency_with_suffix(implied_valuations["ev_ebitda_ltm"]["equity_value"], decimals=0),
        _fmt_currency_with_suffix(implied_valuations["ev_ebitda_ltm"]["enterprise_value"], decimals=0)
    ])
    
    comp_table_data.append([
        'EV/EBITDA FY',
        f'${implied_valuations["ev_ebitda_fy"]["share_price"]:.2f}',
        _fmt_currency_with_suffix(implied_valuations["ev_ebitda_fy"]["equity_value"], decimals=0),
        _fmt_currency_with_suffix(implied_valuations["ev_ebitda_fy"]["enterprise_value"], decimals=0)
    ])
    
    comp_table_data.append([
        'P/E LTM',
        f'${implied_valuations["pe_ltm"]["share_price"]:.2f}',
        _fmt_currency_with_suffix(implied_valuations["pe_ltm"]["equity_value"], decimals=0),
        _fmt_currency_with_suffix(implied_valuations["pe_ltm"]["enterprise_value"], decimals=0)
    ])
    
    comp_table_data.append([
        'P/E FY',
        f'${implied_valuations["pe_fy"]["share_price"]:.2f}',
        _fmt_currency_with_suffix(implied_valuations["pe_fy"]["equity_value"], decimals=0),
        _fmt_currency_with_suffix(implied_valuations["pe_fy"]["enterprise_value"], decimals=0)
    ])
    
    comp_table = Table(comp_table_data, colWidths=[1.8*inch, 1.5*inch, 1.8*inch, 1.8*inch])
    comp_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#124734')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.HexColor('#4A4A4A')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#F2F2EB'), colors.white]),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#CCCCCC')),
    ]))
    elements.append(comp_table)
    elements.append(Spacer(1, 0.25*inch))
    
    # ============================================================================
    # Summary - Football Field Chart and Average Price
    # ============================================================================
    summary_heading = Paragraph("Summary", subheading_style)
    elements.append(summary_heading)
    
    # Extract valuation data for football field
    dcf_tg_low = valuation_summary['dcf_terminal_growth']['low']
    dcf_tg_base = valuation_summary['dcf_terminal_growth']['base']
    dcf_tg_high = valuation_summary['dcf_terminal_growth']['high']
    
    dcf_em_low = valuation_summary['dcf_exit_multiple']['low']
    dcf_em_base = valuation_summary['dcf_exit_multiple']['base']
    dcf_em_high = valuation_summary['dcf_exit_multiple']['high']
    
    # Calculate EV/EBITDA range
    ev_ebitda_ltm = valuation_summary['ev_ebitda_ltm']
    ev_ebitda_fy = valuation_summary['ev_ebitda_fy']
    ev_ebitda_low = min(ev_ebitda_ltm, ev_ebitda_fy)
    ev_ebitda_high = max(ev_ebitda_ltm, ev_ebitda_fy)
    ev_ebitda_avg = (ev_ebitda_ltm + ev_ebitda_fy) / 2
    
    # Calculate P/E range
    pe_ltm = valuation_summary['pe_ltm']
    pe_fy = valuation_summary['pe_fy']
    pe_low = min(pe_ltm, pe_fy)
    pe_high = max(pe_ltm, pe_fy)
    pe_avg = (pe_ltm + pe_fy) / 2
    
    # Calculate average share price from all base case values
    all_base_values = [dcf_tg_base, dcf_em_base, ev_ebitda_avg, pe_avg]
    average_share_price = sum(all_base_values) / len(all_base_values)
    
    # Find min and max values for scaling
    all_values = [dcf_tg_low, dcf_tg_high, dcf_em_low, dcf_em_high, ev_ebitda_low, ev_ebitda_high, pe_low, pe_high, average_share_price]
    min_val = min(all_values) * 0.9
    max_val = max(all_values) * 1.1
    value_range = max_val - min_val
    
    # Create football field chart
    drawing_width = 6.5 * inch
    drawing_height = 3.0 * inch
    d = Drawing(drawing_width, drawing_height)
    
    # Chart parameters
    chart_left = 2.0 * inch
    chart_width = 4.0 * inch
    bar_height = 0.30 * inch
    bar_spacing = 0.55 * inch
    chart_top = drawing_height - 0.4 * inch
    
    # Helper function to convert value to x position
    def value_to_x(value):
        proportion = (value - min_val) / value_range
        return chart_left + (proportion * chart_width)
    
    # Draw vertical grid lines
    grid_values = [min_val, (min_val + max_val) / 2, max_val]
    for grid_val in grid_values:
        x_pos = value_to_x(grid_val)
        d.add(Line(x_pos, 0.25 * inch, x_pos, chart_top + 0.25 * inch,
                   strokeColor=colors.HexColor('#CCCCCC'), strokeWidth=0.5, strokeDashArray=[2, 2]))
        d.add(String(x_pos, 0.08 * inch, f'${grid_val:.0f}',
                    fontSize=8, fillColor=colors.HexColor('#4A4A4A'), textAnchor='middle'))
    
    # Draw average target price line
    avg_x = value_to_x(average_share_price)
    d.add(Line(avg_x, 0.25 * inch, avg_x, chart_top + 0.25 * inch,
               strokeColor=colors.HexColor('#FF6B6B'), strokeWidth=2))
    d.add(String(avg_x, chart_top + 0.35 * inch, f'Avg: ${average_share_price:.2f}',
                fontSize=8, fillColor=colors.HexColor('#FF6B6B'), textAnchor='middle', fontName='Helvetica-Bold'))
    
    # Method 1: DCF Terminal Growth
    y_pos = chart_top - (0 * bar_spacing)
    x_start = value_to_x(dcf_tg_low)
    x_end = value_to_x(dcf_tg_high)
    x_base = value_to_x(dcf_tg_base)
    
    d.add(Rect(x_start, y_pos - bar_height/2, x_end - x_start, bar_height,
               fillColor=colors.HexColor('#124734'), strokeColor=colors.HexColor('#0A0A0A'), strokeWidth=1))
    d.add(Line(x_base, y_pos - bar_height/2 - 0.04*inch, x_base, y_pos + bar_height/2 + 0.04*inch,
               strokeColor=colors.HexColor('#FFD700'), strokeWidth=2.5))
    d.add(String(0.2 * inch, y_pos, 'DCF-TG',
                fontSize=8, fillColor=colors.HexColor('#0A0A0A'), textAnchor='start'))
    
    # Method 2: DCF Exit Multiple
    y_pos = chart_top - (1 * bar_spacing)
    x_start = value_to_x(dcf_em_low)
    x_end = value_to_x(dcf_em_high)
    x_base = value_to_x(dcf_em_base)
    
    d.add(Rect(x_start, y_pos - bar_height/2, x_end - x_start, bar_height,
               fillColor=colors.HexColor('#124734'), strokeColor=colors.HexColor('#0A0A0A'), strokeWidth=1))
    d.add(Line(x_base, y_pos - bar_height/2 - 0.04*inch, x_base, y_pos + bar_height/2 + 0.04*inch,
               strokeColor=colors.HexColor('#FFD700'), strokeWidth=2.5))
    d.add(String(0.2 * inch, y_pos, 'DCF-EM',
                fontSize=8, fillColor=colors.HexColor('#0A0A0A'), textAnchor='start'))
    
    # Method 3: EV/EBITDA
    y_pos = chart_top - (2 * bar_spacing)
    x_start = value_to_x(ev_ebitda_low)
    x_end = value_to_x(ev_ebitda_high)
    x_avg = value_to_x(ev_ebitda_avg)
    
    d.add(Rect(x_start, y_pos - bar_height/2, x_end - x_start, bar_height,
               fillColor=colors.HexColor('#2E6F40'), strokeColor=colors.HexColor('#0A0A0A'), strokeWidth=1))
    d.add(Line(x_avg, y_pos - bar_height/2 - 0.04*inch, x_avg, y_pos + bar_height/2 + 0.04*inch,
               strokeColor=colors.HexColor('#FFD700'), strokeWidth=2.5))
    d.add(String(0.2 * inch, y_pos, 'EV/EBITDA',
                fontSize=8, fillColor=colors.HexColor('#0A0A0A'), textAnchor='start'))
    
    # Method 4: P/E
    y_pos = chart_top - (3 * bar_spacing)
    x_start = value_to_x(pe_low)
    x_end = value_to_x(pe_high)
    x_avg = value_to_x(pe_avg)
    
    d.add(Rect(x_start, y_pos - bar_height/2, x_end - x_start, bar_height,
               fillColor=colors.HexColor('#2E6F40'), strokeColor=colors.HexColor('#0A0A0A'), strokeWidth=1))
    d.add(Line(x_avg, y_pos - bar_height/2 - 0.04*inch, x_avg, y_pos + bar_height/2 + 0.04*inch,
               strokeColor=colors.HexColor('#FFD700'), strokeWidth=2.5))
    d.add(String(0.2 * inch, y_pos, 'P/E',
                fontSize=8, fillColor=colors.HexColor('#0A0A0A'), textAnchor='start'))
    
    elements.append(d)
    elements.append(Spacer(1, 0.2*inch))
    
    # Add summary text with average price
    summary_text = Paragraph(
        f"<b>Average Target Price: ${average_share_price:.2f}</b><br/>"
        f"Implied Upside / (Downside) vs Current Price: {((average_share_price - current_share_price) / current_share_price) * 100:+.2f}%",
        body_style
    )
    elements.append(summary_text)
    elements.append(Spacer(1, 0.3*inch))
    
    elements.append(PageBreak())


if __name__ == "__main__":
    # ============================================================================
    # EXAMPLE USAGE - Generate a Complete DCF Valuation Report PDF
    # ============================================================================
    # Customize these values for your analysis
    ticker_symbol = "GOOGL"  # Change this to any ticker symbol you want
    company_name = "Alphabet Inc."
    current_share_price = 145.67
    shares_outstanding = 2500  # in millions
    market_capitalization = 1820.5
    
    # Step 1: Create PDF document with dynamic ticker in title
    doc = SimpleDocTemplate(
        "data/reports/MyReport.pdf",
        pagesize=A4,
        rightMargin=0.75*inch,
        leftMargin=0.75*inch,
        topMargin=0.75*inch,
        bottomMargin=0.75*inch,
        title=f"{ticker_symbol} Valuation Report",
        author="DCF Model"
    )
    elements = []
    
    # Step 2: Add Executive Summary page
    add_executive_summary_page(
        doc=doc,
        elements=elements,
        company_name=company_name,
        ticker_symbol=ticker_symbol,
        current_share_price=current_share_price,
        shares_outstanding=shares_outstanding,
        market_capitalization=market_capitalization,
        intrinsic_valuation_data={
            'Terminal Growth Model': {
                'Base Case': '$168.50',
                'Bull Case': '$195.75',
                'Bear Case': '$145.25'
            },
            'EBITDA Exit Multiple': {
                'Base Case': '$162.30',
                'Bull Case': '$188.90',
                'Bear Case': '$140.75'
            }
        },
        relative_valuation_data={
            'EV/EBITDA LTM': {
                'Base Case': '$172.45',
                'Bull Case': '$195.30',
                'Bear Case': '$152.75'
            },
            'EV/EBITDA FY': {
                'Base Case': '$169.80',
                'Bull Case': '$192.50',
                'Bear Case': '$150.25'
            },
            'P/E LTM': {
                'Base Case': '$165.25',
                'Bull Case': '$188.90',
                'Bear Case': '$148.50'
            },
            'P/E FY': {
                'Base Case': '$167.90',
                'Bull Case': '$191.25',
                'Bear Case': '$151.00'
            }
        }
    )
    
    # Step 3: Add Model Assumptions & Parameters page
    add_model_assumptions_page(
        doc=doc,
        elements=elements,
        forecast_period=10,
        effective_tax_rate=0.20,        # 20%
        cost_of_equity=0.106,            # 10.6%
        beta=1.10,
        risk_free_rate=0.045,            # 4.5%
        market_return=0.10,              # 10%
        cost_of_debt=0.042,              # 4.2%
        wacc=0.092,                      # 9.2%
        terminal_growth_rate=0.03,       # 3%
        exit_multiple=12.0               # 12.0x
    )
    
    # Step 4: Add Intrinsic Valuation page with DCF analysis
    # You can create different scenarios by changing the forecast data
    
    # BASE CASE - Most likely scenario (10 years)
    base_case_forecast = [
        {'year': 1, 'ebit_1_minus_t': 120, 'da': 50, 'capex': 40, 'change_nwc': 10},
        {'year': 2, 'ebit_1_minus_t': 135, 'da': 52, 'capex': 42, 'change_nwc': 12},
        {'year': 3, 'ebit_1_minus_t': 150, 'da': 54, 'capex': 44, 'change_nwc': 14},
        {'year': 4, 'ebit_1_minus_t': 165, 'da': 56, 'capex': 46, 'change_nwc': 16},
        {'year': 5, 'ebit_1_minus_t': 180, 'da': 58, 'capex': 48, 'change_nwc': 18},
        {'year': 6, 'ebit_1_minus_t': 195, 'da': 60, 'capex': 50, 'change_nwc': 20},
        {'year': 7, 'ebit_1_minus_t': 210, 'da': 62, 'capex': 52, 'change_nwc': 22},
        {'year': 8, 'ebit_1_minus_t': 225, 'da': 64, 'capex': 54, 'change_nwc': 24},
        {'year': 9, 'ebit_1_minus_t': 240, 'da': 66, 'capex': 56, 'change_nwc': 26},
        {'year': 10, 'ebit_1_minus_t': 255, 'da': 68, 'capex': 58, 'change_nwc': 28},
    ]
    
    # BULL CASE - Optimistic scenario (higher cash flows)
    bull_case_forecast = [
        {'year': 1, 'ebit_1_minus_t': 140, 'da': 55, 'capex': 35, 'change_nwc': 8},
        {'year': 2, 'ebit_1_minus_t': 160, 'da': 58, 'capex': 37, 'change_nwc': 10},
        {'year': 3, 'ebit_1_minus_t': 180, 'da': 61, 'capex': 39, 'change_nwc': 12},
        {'year': 4, 'ebit_1_minus_t': 200, 'da': 64, 'capex': 41, 'change_nwc': 14},
        {'year': 5, 'ebit_1_minus_t': 220, 'da': 67, 'capex': 43, 'change_nwc': 16},
        {'year': 6, 'ebit_1_minus_t': 240, 'da': 70, 'capex': 45, 'change_nwc': 18},
        {'year': 7, 'ebit_1_minus_t': 260, 'da': 73, 'capex': 47, 'change_nwc': 20},
        {'year': 8, 'ebit_1_minus_t': 280, 'da': 76, 'capex': 49, 'change_nwc': 22},
        {'year': 9, 'ebit_1_minus_t': 300, 'da': 79, 'capex': 51, 'change_nwc': 24},
        {'year': 10, 'ebit_1_minus_t': 320, 'da': 82, 'capex': 53, 'change_nwc': 26},
    ]
    
    # BEAR CASE - Conservative scenario (lower cash flows)
    bear_case_forecast = [
        {'year': 1, 'ebit_1_minus_t': 100, 'da': 45, 'capex': 45, 'change_nwc': 12},
        {'year': 2, 'ebit_1_minus_t': 110, 'da': 47, 'capex': 47, 'change_nwc': 14},
        {'year': 3, 'ebit_1_minus_t': 120, 'da': 49, 'capex': 49, 'change_nwc': 16},
        {'year': 4, 'ebit_1_minus_t': 130, 'da': 51, 'capex': 51, 'change_nwc': 18},
        {'year': 5, 'ebit_1_minus_t': 140, 'da': 53, 'capex': 53, 'change_nwc': 20},
        {'year': 6, 'ebit_1_minus_t': 150, 'da': 55, 'capex': 55, 'change_nwc': 22},
        {'year': 7, 'ebit_1_minus_t': 160, 'da': 57, 'capex': 57, 'change_nwc': 24},
        {'year': 8, 'ebit_1_minus_t': 170, 'da': 59, 'capex': 59, 'change_nwc': 26},
        {'year': 9, 'ebit_1_minus_t': 180, 'da': 61, 'capex': 61, 'change_nwc': 28},
        {'year': 10, 'ebit_1_minus_t': 190, 'da': 63, 'capex': 63, 'change_nwc': 30},
    ]
    
    # Add Base Case Intrinsic Valuation (precomputed inputs)
    terminal_growth_summary = {
        'terminal_value': 5200,
        'enterprise_value': 8400,
        'net_debt': 100,
        'total_equity_value': 8300,
        'shares_outstanding': 2500,
        'implied_share_price': 3.32
    }
    exit_multiple_summary = {
        'terminal_ebitda': 1800,
        'exit_multiple': 12.0,
        'terminal_value': 21600,
        'enterprise_value': 9500,
        'net_debt': 100,
        'total_equity_value': 9400,
        'shares_outstanding': 2500,
        'implied_share_price': 3.76
    }
    sensitivity_tgm = [
        ['WACC / TGR', '2.0%', '3.0%', '4.0%'],
        ['7.0%', '$2.90', '$3.10', '$3.35'],
        ['8.0%', '$2.75', '$2.95', '$3.20'],
        ['9.0%', '$2.60', '$2.80', '$3.05'],
        ['10.0%', '$2.45', '$2.65', '$2.90']
    ]
    sensitivity_em = [
        ['WACC / Multiple', '10.0x', '11.0x', '12.0x', '13.0x', '14.0x'],
        ['7.0%', '$2.90', '$3.05', '$3.20', '$3.35', '$3.50'],
        ['8.0%', '$2.75', '$2.90', '$3.05', '$3.20', '$3.35'],
        ['9.0%', '$2.60', '$2.75', '$2.90', '$3.05', '$3.20'],
        ['10.0%', '$2.45', '$2.60', '$2.75', '$2.90', '$3.05']
    ]
    add_intrinsic_valuation_page(
        doc=doc,
        elements=elements,
        forecast_years_data=base_case_forecast,
        terminal_growth_summary=terminal_growth_summary,
        exit_multiple_summary=exit_multiple_summary,
        sensitivity_tgm=sensitivity_tgm,
        sensitivity_em=sensitivity_em
    )
    
    # Step 5: Add Relative Valuation page
    # Define peer group data
    peer_companies = [
        {
            'ticker': 'MSFT',
            'company_name': 'Microsoft Corporation',
            'ev_ebitda_ltm': 18.5,
            'ev_ebitda_fy': 17.2,
            'pe_ltm': 32.4,
            'pe_fy': 29.8
        },
        {
            'ticker': 'AAPL',
            'company_name': 'Apple Inc.',
            'ev_ebitda_ltm': 22.1,
            'ev_ebitda_fy': 20.5,
            'pe_ltm': 28.7,
            'pe_fy': 26.3
        },
        {
            'ticker': 'META',
            'company_name': 'Meta Platforms Inc.',
            'ev_ebitda_ltm': 14.8,
            'ev_ebitda_fy': 13.9,
            'pe_ltm': 24.2,
            'pe_fy': 22.1
        },
        {
            'ticker': 'AMZN',
            'company_name': 'Amazon.com Inc.',
            'ev_ebitda_ltm': 16.3,
            'ev_ebitda_fy': 15.1,
            'pe_ltm': 45.6,
            'pe_fy': 38.9
        }
    ]
    
    # Define target company metrics for valuation calculation
    target_metrics = {
        'ebitda_ltm': 850,              # $850M EBITDA LTM
        'ebitda_fy': 920,               # $920M EBITDA FY
        'net_income_ltm': 650,          # $650M Net Income LTM
        'net_income_fy': 710,           # $710M Net Income FY
        'net_debt': 100,                # $100M Net Debt
        'shares_outstanding': 2500      # 2,500M shares
    }
    
    # Define summary statistics (these would typically be calculated from peer_companies)
    summary_statistics = {
        'ev_ebitda_ltm': {
            'average': 17.93,
            'median': 17.40,
            'highest': 22.10,
            'lowest': 14.80
        },
        'ev_ebitda_fy': {
            'average': 16.68,
            'median': 16.15,
            'highest': 20.50,
            'lowest': 13.90
        },
        'pe_ltm': {
            'average': 32.73,
            'median': 30.55,
            'highest': 45.60,
            'lowest': 24.20
        },
        'pe_fy': {
            'average': 29.28,
            'median': 28.05,
            'highest': 38.90,
            'lowest': 22.10
        }
    }
    
    # Calculate implied valuations (normally done outside this function with your own logic)
    # Method 1: EV/EBITDA LTM
    ev_ebitda_ltm_multiple = summary_statistics['ev_ebitda_ltm']['average']
    enterprise_value_ltm = target_metrics['ebitda_ltm'] * ev_ebitda_ltm_multiple
    equity_value_ltm = enterprise_value_ltm - target_metrics['net_debt']
    share_price_ltm = equity_value_ltm / target_metrics['shares_outstanding']
    
    # Method 2: EV/EBITDA FY
    ev_ebitda_fy_multiple = summary_statistics['ev_ebitda_fy']['average']
    enterprise_value_fy = target_metrics['ebitda_fy'] * ev_ebitda_fy_multiple
    equity_value_fy = enterprise_value_fy - target_metrics['net_debt']
    share_price_fy = equity_value_fy / target_metrics['shares_outstanding']
    
    # Method 3: P/E LTM
    pe_ltm_multiple = summary_statistics['pe_ltm']['average']
    equity_value_pe_ltm = target_metrics['net_income_ltm'] * pe_ltm_multiple
    share_price_pe_ltm = equity_value_pe_ltm / target_metrics['shares_outstanding']
    enterprise_value_pe_ltm = equity_value_pe_ltm + target_metrics['net_debt']
    
    # Method 4: P/E FY
    pe_fy_multiple = summary_statistics['pe_fy']['average']
    equity_value_pe_fy = target_metrics['net_income_fy'] * pe_fy_multiple
    share_price_pe_fy = equity_value_pe_fy / target_metrics['shares_outstanding']
    enterprise_value_pe_fy = equity_value_pe_fy + target_metrics['net_debt']
    
    # Create implied valuations dictionary
    implied_vals = {
        'ev_ebitda_ltm': {
            'share_price': share_price_ltm,
            'equity_value': equity_value_ltm,
            'enterprise_value': enterprise_value_ltm
        },
        'ev_ebitda_fy': {
            'share_price': share_price_fy,
            'equity_value': equity_value_fy,
            'enterprise_value': enterprise_value_fy
        },
        'pe_ltm': {
            'share_price': share_price_pe_ltm,
            'equity_value': equity_value_pe_ltm,
            'enterprise_value': enterprise_value_pe_ltm
        },
        'pe_fy': {
            'share_price': share_price_pe_fy,
            'equity_value': equity_value_pe_fy,
            'enterprise_value': enterprise_value_pe_fy
        }
    }
    
    add_relative_valuation_page(
        doc=doc,
        elements=elements,
        peer_data=peer_companies,
        summary_stats=summary_statistics,
        implied_valuations=implied_vals
    )
    
    # Step 6: Add Conclusion page with valuation summary
    # Calculate EV/EBITDA valuations (LTM and FY)
    ev_ebitda_ltm_price = (target_metrics['ebitda_ltm'] * summary_statistics['ev_ebitda_ltm']['average'] - target_metrics['net_debt']) / target_metrics['shares_outstanding']
    ev_ebitda_fy_price = (target_metrics['ebitda_fy'] * summary_statistics['ev_ebitda_fy']['average'] - target_metrics['net_debt']) / target_metrics['shares_outstanding']
    
    # Calculate P/E valuations (LTM and FY)
    pe_ltm_price = (target_metrics['net_income_ltm'] * summary_statistics['pe_ltm']['average']) / target_metrics['shares_outstanding']
    pe_fy_price = (target_metrics['net_income_fy'] * summary_statistics['pe_fy']['average']) / target_metrics['shares_outstanding']
    
    # Prepare valuation summary (using example values - replace with actual calculated values)
    valuation_summary_data = {
        'dcf_terminal_growth': {
            'low': 135.50,    # Min from sensitivity table (WACC vs TGR)
            'base': 152.30,   # Base case
            'high': 168.75    # Max from sensitivity table (WACC vs TGR)
        },
        'dcf_exit_multiple': {
            'low': 138.20,    # Min from sensitivity table (WACC vs Exit Multiple)
            'base': 155.80,   # Base case
            'high': 172.40    # Max from sensitivity table (WACC vs Exit Multiple)
        },
        'ev_ebitda_ltm': ev_ebitda_ltm_price,
        'ev_ebitda_fy': ev_ebitda_fy_price,
        'pe_ltm': pe_ltm_price,
        'pe_fy': pe_fy_price
    }
    
    add_conclusion_page(
        doc=doc,
        elements=elements,
        valuation_summary=valuation_summary_data,
        current_share_price=current_share_price
    )
    
    # Step 7: Build the PDF
    doc.build(elements)
    print(f"✓ Complete DCF Valuation Report created successfully with all pages")
    print(f"  - Executive Summary page")
    print(f"  - Model Assumptions & Parameters page")
    print(f"  - Intrinsic Valuation (Base Case) page")
    print(f"  - Relative Valuation page")
    print(f"  - Conclusion & Valuation Summary page")
    print(f"\nTo generate Bull or Bear case scenarios, modify the forecast data:")
    print(f"  - Bull Case: Higher growth and lower capital requirements")
    print(f"  - Bear Case: Lower growth and higher capital requirements")
