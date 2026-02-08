"""Generate Excel-based BESS Economic Analysis Workbook.

Creates a macro-enabled Excel workbook (.xlsm) with:
- Project input forms
- Assumption library selection (NREL, Lazard, CPUC)
- Automatic calculation engine using Excel formulas
- Results dashboard with conditional formatting
- Annual cash flow projections
- Methodology documentation with formulas and citations
- Embedded VBA macros with functional buttons

Usage:
    python excel_generator.py [output_path]
"""

import sys
from pathlib import Path

import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell


def create_workbook(output_path: str, with_macros: bool = True):
    """Create the complete BESS Analyzer workbook.

    Args:
        output_path: Path for the output file.
        with_macros: If True and vbaProject.bin exists, create .xlsm with macros.
                    If False or vbaProject.bin missing, create .xlsx without macros.
    """
    vba_path = Path(__file__).parent / 'resources' / 'vbaProject.bin'
    has_vba = vba_path.exists() and with_macros

    # Set extension based on macro availability
    if has_vba:
        if not output_path.endswith('.xlsm'):
            output_path = output_path.replace('.xlsx', '.xlsm')
            if not output_path.endswith('.xlsm'):
                output_path += '.xlsm'
    else:
        if output_path.endswith('.xlsm'):
            output_path = output_path.replace('.xlsm', '.xlsx')
        elif not output_path.endswith('.xlsx'):
            output_path += '.xlsx'

    workbook = xlsxwriter.Workbook(output_path)

    # Add VBA project if available
    if has_vba:
        workbook.add_vba_project(str(vba_path), is_stream=False)

    # Create formats
    formats = create_formats(workbook)

    # Create sheets
    ws_inputs = workbook.add_worksheet('Inputs')
    ws_results = workbook.add_worksheet('Results')
    ws_cashflows = workbook.add_worksheet('Cash_Flows')
    ws_calculations = workbook.add_worksheet('Calculations')
    ws_sensitivity = workbook.add_worksheet('Sensitivity')
    ws_methodology = workbook.add_worksheet('Methodology')
    ws_library_data = workbook.add_worksheet('Library_Data')
    ws_vba_code = workbook.add_worksheet('VBA_Code')

    # Build each sheet
    create_inputs_sheet(workbook, ws_inputs, formats)
    cf_totals_row = create_cashflows_sheet(workbook, ws_cashflows, formats)
    create_calculations_sheet(workbook, ws_calculations, formats)
    create_results_sheet(workbook, ws_results, formats, cf_totals_row)
    create_sensitivity_sheet(workbook, ws_sensitivity, formats, cf_totals_row)
    create_methodology_sheet(workbook, ws_methodology, formats)
    create_library_data_sheet(workbook, ws_library_data, formats)
    create_vba_code_sheet(workbook, ws_vba_code, formats)

    workbook.close()

    if has_vba:
        print(f"Macro-enabled workbook created: {output_path}")
        print("\nThe workbook includes:")
        print("- Embedded VBA macros")
        print("- Functional buttons for loading libraries")
        print("- Generate Report and Export to PDF buttons")
        print("\nNote: You may need to enable macros when opening the file.")
    else:
        print(f"Workbook created: {output_path}")
        print("\nTo enable VBA macros:")
        print("1. Open the 'VBA_Code' sheet for complete instructions")
        print("2. Save as Macro-Enabled Workbook (.xlsm)")
        print("3. Open VBA Editor (Alt+F11) and paste the code into a Module")
        print("4. Assign macros to buttons as needed")


def create_formats(workbook):
    """Create reusable cell formats."""
    formats = {}

    # Header style
    formats['header'] = workbook.add_format({
        'bold': True,
        'font_color': 'white',
        'bg_color': '#1565C0',
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })

    # Section header
    formats['section'] = workbook.add_format({
        'bold': True,
        'font_size': 12,
        'font_color': '#1565C0',
        'bg_color': '#E3F2FD',
        'border': 1
    })

    # Input cell
    formats['input'] = workbook.add_format({
        'bg_color': '#FFFDE7',
        'border': 1
    })

    # Input cell with currency
    formats['input_currency'] = workbook.add_format({
        'bg_color': '#FFFDE7',
        'border': 1,
        'num_format': '$#,##0.00'
    })

    # Input cell with percent
    formats['input_percent'] = workbook.add_format({
        'bg_color': '#FFFDE7',
        'border': 1,
        'num_format': '0.0%'
    })

    # Formula cell
    formats['formula'] = workbook.add_format({
        'bg_color': '#E8F5E9',
        'border': 1
    })

    # Formula cell with currency
    formats['formula_currency'] = workbook.add_format({
        'bg_color': '#E8F5E9',
        'border': 1,
        'num_format': '$#,##0'
    })

    # Result cell
    formats['result'] = workbook.add_format({
        'bold': True,
        'font_size': 14,
        'bg_color': '#E3F2FD',
        'align': 'center',
        'border': 1
    })

    # Result with currency
    formats['result_currency'] = workbook.add_format({
        'bold': True,
        'font_size': 14,
        'bg_color': '#E3F2FD',
        'align': 'center',
        'border': 1,
        'num_format': '$#,##0'
    })

    # Result with percent
    formats['result_percent'] = workbook.add_format({
        'bold': True,
        'font_size': 14,
        'bg_color': '#E3F2FD',
        'align': 'center',
        'border': 1,
        'num_format': '0.0%'
    })

    # Currency format
    formats['currency'] = workbook.add_format({
        'num_format': '$#,##0',
        'border': 1
    })

    # Percentage format
    formats['percent'] = workbook.add_format({
        'num_format': '0.0%',
        'border': 1
    })

    # Title
    formats['title'] = workbook.add_format({
        'bold': True,
        'font_size': 16,
        'font_color': '#1565C0'
    })

    # Bold
    formats['bold'] = workbook.add_format({
        'bold': True
    })

    # Italic gray (tooltip)
    formats['tooltip'] = workbook.add_format({
        'italic': True,
        'font_color': '#666666'
    })

    # Green (positive)
    formats['positive'] = workbook.add_format({
        'bg_color': '#C8E6C9',
        'bold': True,
        'font_size': 14,
        'align': 'center',
        'border': 1
    })

    # Yellow (neutral)
    formats['neutral'] = workbook.add_format({
        'bg_color': '#FFF9C4',
        'bold': True,
        'font_size': 14,
        'align': 'center',
        'border': 1
    })

    # Red (negative)
    formats['negative'] = workbook.add_format({
        'bg_color': '#FFCDD2',
        'bold': True,
        'font_size': 14,
        'align': 'center',
        'border': 1
    })

    # Wrap text
    formats['wrap'] = workbook.add_format({
        'text_wrap': True,
        'valign': 'top'
    })

    # Link style
    formats['link'] = workbook.add_format({
        'font_color': '#0066CC',
        'underline': True
    })

    # Code style
    formats['code'] = workbook.add_format({
        'font_name': 'Courier New',
        'font_size': 9
    })

    return formats


def create_inputs_sheet(workbook, ws, formats):
    """Create the Project Inputs sheet with buttons."""

    # Column widths
    ws.set_column('A:A', 5)
    ws.set_column('B:B', 30)
    ws.set_column('C:C', 20)
    ws.set_column('D:D', 15)
    ws.set_column('E:E', 45)

    row = 1

    # Title
    ws.merge_range('B2:E2', 'BESS Economic Analysis - Project Inputs', formats['title'])
    row = 3

    # === BUTTONS SECTION ===
    ws.merge_range(f'B{row}:E{row}', 'LIBRARY SELECTION', formats['section'])
    row += 1

    # Add buttons for library loading
    ws.insert_button(f'B{row}', {
        'macro': 'LoadNRELLibrary',
        'caption': 'Load NREL ATB 2024',
        'width': 130,
        'height': 25
    })

    ws.insert_button(f'C{row}', {
        'macro': 'LoadLazardLibrary',
        'caption': 'Load Lazard LCOS 2025',
        'width': 130,
        'height': 25
    })

    ws.insert_button(f'D{row}', {
        'macro': 'LoadCPUCLibrary',
        'caption': 'Load CPUC CA 2024',
        'width': 130,
        'height': 25
    })
    row += 2

    ws.write(f'B{row}', 'Selected Library:', formats['bold'])
    ws.write(f'C{row}', 'Custom', formats['input'])
    library_row = row
    row += 2

    # === PROJECT BASICS ===
    ws.merge_range(f'B{row}:E{row}', 'PROJECT BASICS', formats['section'])
    row += 1

    basics = [
        ('Project Name', '', 'Text', 'Enter project name'),
        ('Project ID', '', 'Text', 'Unique identifier'),
        ('Location', '', 'Text', 'Site or market location'),
        ('Capacity (MW)', 100, 'Number', 'Nameplate power capacity'),
        ('Duration (hours)', 4, 'Number', 'Storage duration'),
        ('Energy Capacity (MWh)', '=C10*C11', 'Formula', 'Auto-calculated: MW x hours'),
        ('Analysis Period (years)', 20, 'Number', 'Economic analysis horizon'),
        ('Discount Rate (%)', 0.07, 'Percent', 'Nominal discount rate for NPV'),
        ('Ownership Type', 'Utility', 'Text', 'Utility (6-7% WACC) or Merchant (8-10%)'),
    ]

    basics_start = row
    for label, value, dtype, tooltip in basics:
        ws.write(f'B{row}', label, formats['bold'])
        if dtype == 'Formula':
            ws.write_formula(f'C{row}', value, formats['formula'])
        elif dtype == 'Percent':
            ws.write(f'C{row}', value, formats['input_percent'])
        else:
            ws.write(f'C{row}', value, formats['input'])
        ws.write(f'E{row}', tooltip, formats['tooltip'])
        row += 1

    row += 1

    # === TECHNOLOGY SPECIFICATIONS ===
    ws.merge_range(f'B{row}:E{row}', 'TECHNOLOGY SPECIFICATIONS', formats['section'])
    row += 1

    tech_specs = [
        ('Chemistry', 'LFP', 'Text', 'LFP, NMC, or Other'),
        ('Round-Trip Efficiency (%)', 0.85, 'Percent', 'AC-AC efficiency'),
        ('Annual Degradation (%)', 0.025, 'Percent', 'Capacity loss per year'),
        ('Cycle Life', 6000, 'Number', 'Full-depth cycles before EOL'),
        ('Augmentation Year', 12, 'Number', 'Year of battery replacement'),
        ('Cycles per Day', 1.0, 'Number', 'Average daily charge/discharge cycles'),
    ]

    tech_start = row
    for label, value, dtype, tooltip in tech_specs:
        ws.write(f'B{row}', label, formats['bold'])
        if dtype == 'Percent':
            ws.write(f'C{row}', value, formats['input_percent'])
        else:
            ws.write(f'C{row}', value, formats['input'])
        ws.write(f'E{row}', tooltip, formats['tooltip'])
        row += 1

    row += 1

    # === COST INPUTS ===
    ws.merge_range(f'B{row}:E{row}', 'COST INPUTS (BESS)', formats['section'])
    row += 1

    cost_inputs = [
        ('CapEx ($/kWh)', 160, '$/kWh', 'Installed capital cost per kWh'),
        ('Fixed O&M ($/kW-year)', 25, '$/kW-yr', 'Annual fixed O&M'),
        ('Variable O&M ($/MWh)', 0, '$/MWh', 'Per-MWh discharge cost'),
        ('Augmentation Cost ($/kWh)', 55, '$/kWh', 'Battery replacement cost'),
        ('Decommissioning ($/kW)', 10, '$/kW', 'End-of-life cost'),
        ('Charging Cost ($/MWh)', 30, '$/MWh', 'Grid electricity cost for charging'),
        ('Residual Value (%)', 0.10, '%', 'End-of-life asset value as % of CapEx'),
    ]

    cost_start = row
    for label, value, unit, tooltip in cost_inputs:
        ws.write(f'B{row}', label, formats['bold'])
        ws.write(f'C{row}', value, formats['input_currency'])
        ws.write(f'D{row}', unit)
        ws.write(f'E{row}', tooltip, formats['tooltip'])
        row += 1

    row += 1

    # === TAX CREDITS ===
    ws.merge_range(f'B{row}:E{row}', 'TAX CREDITS (BESS-Specific under IRA)', formats['section'])
    row += 1

    tax_inputs = [
        ('ITC Base Rate (%)', 0.30, '30% Investment Tax Credit under IRA'),
        ('ITC Adders (%)', 0.0, 'Energy community +10%, Domestic content +10%'),
    ]

    itc_start = row
    for label, value, tooltip in tax_inputs:
        ws.write(f'B{row}', label, formats['bold'])
        ws.write(f'C{row}', value, formats['input_percent'])
        ws.write(f'E{row}', tooltip, formats['tooltip'])
        row += 1

    row += 1

    # === INFRASTRUCTURE COSTS ===
    ws.merge_range(f'B{row}:E{row}', 'INFRASTRUCTURE COSTS (Common to all projects)', formats['section'])
    row += 1

    infra_inputs = [
        ('Interconnection ($/kW)', 100, '$/kW', 'Network upgrades, studies'),
        ('Land ($/kW)', 10, '$/kW', 'Site acquisition/lease'),
        ('Permitting ($/kW)', 15, '$/kW', 'Permits, environmental review'),
        ('Insurance (% of CapEx)', 0.005, '%', 'Annual insurance cost'),
        ('Property Tax (%)', 0.01, '%', 'Annual property tax rate'),
    ]

    infra_start = row
    for label, value, unit, tooltip in infra_inputs:
        ws.write(f'B{row}', label, formats['bold'])
        if '%' in unit:
            ws.write(f'C{row}', value, formats['input_percent'])
        else:
            ws.write(f'C{row}', value, formats['input_currency'])
        ws.write(f'D{row}', unit)
        ws.write(f'E{row}', tooltip, formats['tooltip'])
        row += 1

    row += 1

    # === FINANCING STRUCTURE ===
    ws.merge_range(f'B{row}:E{row}', 'FINANCING STRUCTURE (For WACC Calculation)', formats['section'])
    row += 1

    financing_inputs = [
        ('Debt Percentage (%)', 0.60, 'Debt/equity split (0.60 = 60% debt)'),
        ('Interest Rate (%)', 0.05, 'Annual interest rate on debt'),
        ('Loan Term (years)', 15, 'Debt amortization period'),
        ('Cost of Equity (%)', 0.10, 'Required return on equity'),
        ('Tax Rate (%)', 0.21, 'Corporate tax rate for interest deduction'),
    ]

    financing_start = row
    for label, value, tooltip in financing_inputs:
        ws.write(f'B{row}', label, formats['bold'])
        if 'years' in label:
            ws.write(f'C{row}', value, formats['input'])
        else:
            ws.write(f'C{row}', value, formats['input_percent'])
        ws.write(f'E{row}', tooltip, formats['tooltip'])
        row += 1

    # WACC formula
    ws.write(f'B{row}', 'Calculated WACC', formats['bold'])
    # WACC = (1-D) * Re + D * Rd * (1 - Tc)
    wacc_formula = f'=(1-C{financing_start})*C{financing_start+3}+C{financing_start}*C{financing_start+1}*(1-C{financing_start+4})'
    ws.write_formula(f'C{row}', wacc_formula, formats['formula'])
    ws.write(f'E{row}', 'Weighted Average Cost of Capital', formats['tooltip'])
    row += 2

    # === BENEFIT STREAMS ===
    ws.merge_range(f'B{row}:E{row}', 'BENEFIT STREAMS (Year 1 Values)', formats['section'])
    row += 1

    # Headers
    ws.write(f'B{row}', 'Benefit Category', formats['header'])
    ws.write(f'C{row}', '$/kW-year', formats['header'])
    ws.write(f'D{row}', 'Escalation %', formats['header'])
    ws.write(f'E{row}', 'Category / Citation', formats['header'])
    row += 1

    benefits = [
        ('Resource Adequacy', 150, 0.02, '[common] CPUC RA Report 2024'),
        ('Energy Arbitrage', 40, 0.015, '[common] Market Data 2024'),
        ('Ancillary Services', 15, 0.01, '[common] AS Reports 2024'),
        ('T&D Deferral', 25, 0.02, '[common] Avoided Cost Calculator'),
        ('Resilience Value', 50, 0.02, '[common] LBNL ICE Calculator'),
        ('Renewable Integration', 25, 0.02, '[bess] NREL Grid Studies'),
        ('GHG Emissions Value', 15, 0.03, '[bess] EPA Social Cost Carbon'),
        ('Voltage Support', 8, 0.01, '[common] EPRI Distribution'),
    ]

    benefits_start = row
    for name, value, esc, cite in benefits:
        ws.write(f'B{row}', name)
        ws.write(f'C{row}', value, formats['input_currency'])
        ws.write(f'D{row}', esc, formats['input_percent'])
        ws.write(f'E{row}', cite, formats['tooltip'])
        row += 1

    row += 1

    # === COST PROJECTIONS ===
    ws.merge_range(f'B{row}:E{row}', 'COST PROJECTIONS (Learning Curve)', formats['section'])
    row += 1

    ws.write(f'B{row}', 'Annual Cost Decline Rate', formats['bold'])
    ws.write(f'C{row}', 0.10, formats['input_percent'])
    ws.write(f'E{row}', 'Technology learning rate (10-15% typical)', formats['tooltip'])
    learning_rate_row = row
    row += 1

    ws.write(f'B{row}', 'Cost Base Year', formats['bold'])
    ws.write(f'C{row}', 2024, formats['input'])
    ws.write(f'E{row}', 'Reference year for base costs', formats['tooltip'])
    row += 2

    # === REPORT BUTTONS ===
    ws.merge_range(f'B{row}:E{row}', 'GENERATE REPORTS', formats['section'])
    row += 1

    ws.insert_button(f'B{row}', {
        'macro': 'GenerateReport',
        'caption': 'Generate Report',
        'width': 130,
        'height': 30
    })

    ws.insert_button(f'C{row}', {
        'macro': 'ExportToPDF',
        'caption': 'Export to PDF',
        'width': 130,
        'height': 30
    })

    ws.insert_button(f'D{row}', {
        'macro': 'RefreshCalculations',
        'caption': 'Refresh Calculations',
        'width': 130,
        'height': 30
    })

    # Store key row numbers for named ranges
    workbook.define_name('Capacity_MW', f'=Inputs!$C$10')
    workbook.define_name('Duration_Hours', f'=Inputs!$C$11')
    workbook.define_name('Energy_MWh', f'=Inputs!$C$12')
    workbook.define_name('Analysis_Years', f'=Inputs!$C$13')
    workbook.define_name('Discount_Rate', f'=Inputs!$C$14')
    workbook.define_name('Learning_Rate', f'=Inputs!$C${learning_rate_row}')

    return benefits_start


def create_calculations_sheet(workbook, ws, formats):
    """Create the Calculations sheet."""

    ws.set_column('A:A', 5)
    ws.set_column('B:B', 25)
    ws.set_column('C:C', 18)
    ws.set_column('D:D', 50)

    row = 1
    ws.write('B2', 'Calculation Engine', formats['title'])
    row = 4

    # Derived values
    ws.merge_range(f'B{row}:D{row}', 'DERIVED VALUES', formats['section'])
    row += 1

    calcs = [
        ('Capacity (kW)', '=Inputs!C10*1000', 'Convert MW to kW'),
        ('Capacity (kWh)', '=Inputs!C12*1000', 'Convert MWh to kWh'),
        ('Battery CapEx ($)', '=Inputs!C28*C6', 'CapEx/kWh x kWh'),
        ('Infrastructure ($)', '=(Inputs!C43+Inputs!C44+Inputs!C45)*C5', 'Interconnect+Land+Permit'),
        ('Total CapEx ($)', '=C7+C8', 'Battery + Infrastructure'),
        ('ITC Credit ($)', '=C7*(Inputs!C37+Inputs!C38)', 'ITC on battery only'),
        ('Net Year 0 Cost ($)', '=C9-C10', 'CapEx minus ITC'),
        ('Annual Fixed O&M ($)', '=Inputs!C29*C5', 'FOM/kW x kW'),
        ('Annual Energy (MWh)', '=Inputs!C12*Inputs!C23*365*Inputs!C19', 'MWh x cycles/day x 365 x RTE'),
        ('Annual Charging Cost ($)', '=C13/Inputs!C19*Inputs!C33', 'Energy/RTE x charging cost'),
        ('Residual Value ($)', '=C9*Inputs!C34', 'CapEx x residual %'),
    ]

    for label, formula, desc in calcs:
        ws.write(f'B{row}', label, formats['bold'])
        ws.write_formula(f'C{row}', formula, formats['formula_currency'])
        ws.write(f'D{row}', desc, formats['tooltip'])
        row += 1

    row += 1

    # Key formulas
    ws.merge_range(f'B{row}:D{row}', 'KEY FORMULAS', formats['section'])
    row += 1

    formulas_info = [
        ('NPV', 'Sum(CFt / (1+r)^t) for t=0 to N', 'Brealey & Myers (2020)'),
        ('BCR', 'PV(Benefits) / PV(Costs)', 'CPUC Standard Practice Manual'),
        ('IRR', 'Rate where NPV = 0', 'Brealey & Myers (2020)'),
        ('LCOS', 'PV(Costs) / PV(Energy)', 'Lazard Methodology'),
        ('Payback', 'Year where cumulative CF > 0', 'Standard finance'),
    ]

    for metric, formula, source in formulas_info:
        ws.write(f'B{row}', metric, formats['bold'])
        ws.write(f'C{row}', formula)
        ws.write(f'D{row}', source, formats['tooltip'])
        row += 1


def create_cashflows_sheet(workbook, ws, formats):
    """Create annual cash flow projections."""

    ws.set_column('A:A', 5)
    ws.set_column('B:B', 8)
    for col in range(2, 19):  # C through S
        ws.set_column(col, col, 14)

    ws.write('B2', 'Annual Cash Flow Projections', formats['title'])

    row = 4
    headers = ['Year', 'CapEx', 'Fixed O&M', 'Var O&M', 'Insurance', 'Prop Tax',
               'Augment', 'Decommission', 'Total Costs',
               'RA', 'Arbitrage', 'Ancillary', 'T&D', 'Resilience',
               'Renew Int', 'GHG', 'Voltage', 'Total Benefits',
               'Net CF', 'Disc Factor', 'PV Costs', 'PV Benefits', 'Cumul CF']

    for col, header in enumerate(headers):
        ws.write(row - 1, col + 1, header, formats['header'])

    start_row = row

    # Generate 21 years (0-20)
    for year in range(21):
        col = 1

        # Year
        ws.write(row, col, year)
        col += 1

        # CapEx (year 0 only) - includes infrastructure, minus ITC
        if year == 0:
            ws.write_formula(row, col, '=Calculations!C11', formats['currency'])
        else:
            ws.write(row, col, 0, formats['currency'])
        col += 1

        # Fixed O&M (years 1+)
        if year == 0:
            ws.write(row, col, 0, formats['currency'])
        else:
            ws.write_formula(row, col, '=Calculations!C12', formats['currency'])
        col += 1

        # Variable O&M
        ws.write(row, col, 0, formats['currency'])
        col += 1

        # Insurance (years 1+)
        if year == 0:
            ws.write(row, col, 0, formats['currency'])
        else:
            ws.write_formula(row, col, '=Calculations!C9*Inputs!$C$40', formats['currency'])
        col += 1

        # Property Tax (declining with depreciation)
        if year == 0:
            ws.write(row, col, 0, formats['currency'])
        else:
            formula = f'=Calculations!$C$9*(1-{year}/Inputs!$C$13)*Inputs!$C$41'
            ws.write_formula(row, col, formula, formats['currency'])
        col += 1

        # Augmentation (year 12) with learning curve
        if year == 12:
            ws.write_formula(row, col,
                '=Inputs!$C$29*Calculations!C6*(1-Inputs!$C$53)^12', formats['currency'])
        else:
            ws.write(row, col, 0, formats['currency'])
        col += 1

        # Decommissioning (year 20)
        if year == 20:
            ws.write_formula(row, col, '=Inputs!$C$30*Calculations!C5', formats['currency'])
        else:
            ws.write(row, col, 0, formats['currency'])
        col += 1

        # Total Costs
        ws.write_formula(row, col, f'=SUM(C{row+1}:I{row+1})', formats['formula_currency'])
        col += 1

        # Benefits (8 streams with escalation)
        if year == 0:
            for _ in range(8):
                ws.write(row, col, 0, formats['currency'])
                col += 1
        else:
            benefit_rows = [45, 46, 47, 48, 49, 50, 51, 52]  # Rows in Inputs sheet
            for i, b_row in enumerate(benefit_rows):
                # Value * Capacity(kW) * (1+escalation)^(year-1) * degradation
                degradation = f'*(1-Inputs!$C$20)^{year-1}' if i < 5 else ''  # No degradation for last 3
                formula = f'=Inputs!$C${b_row}*Inputs!$C$10*1000*(1+Inputs!$D${b_row})^{year-1}{degradation}'
                ws.write_formula(row, col, formula, formats['currency'])
                col += 1

        # Total Benefits
        ws.write_formula(row, col, f'=SUM(K{row+1}:R{row+1})', formats['formula_currency'])
        col += 1

        # Net Cash Flow
        ws.write_formula(row, col, f'=S{row+1}-J{row+1}', formats['formula_currency'])
        col += 1

        # Discount Factor
        ws.write_formula(row, col, f'=1/(1+Inputs!$C$14)^B{row+1}', formats['percent'])
        col += 1

        # PV Costs
        ws.write_formula(row, col, f'=J{row+1}*U{row+1}', formats['currency'])
        col += 1

        # PV Benefits
        ws.write_formula(row, col, f'=S{row+1}*U{row+1}', formats['currency'])
        col += 1

        # Cumulative CF
        if year == 0:
            ws.write_formula(row, col, f'=T{row+1}', formats['currency'])
        else:
            ws.write_formula(row, col, f'=X{row}+T{row+1}', formats['currency'])

        row += 1

    end_row = row - 1

    # Totals row
    row += 1
    ws.write(row, 1, 'TOTALS', formats['bold'])
    ws.write_formula(row, 9, f'=SUM(J{start_row+1}:J{end_row+1})', formats['formula_currency'])  # Total Costs
    ws.write_formula(row, 18, f'=SUM(S{start_row+1}:S{end_row+1})', formats['formula_currency'])  # Total Benefits
    ws.write_formula(row, 21, f'=SUM(V{start_row+1}:V{end_row+1})', formats['formula_currency'])  # PV Costs
    ws.write_formula(row, 22, f'=SUM(W{start_row+1}:W{end_row+1})', formats['formula_currency'])  # PV Benefits

    return row  # Return totals row for Results sheet


def create_results_sheet(workbook, ws, formats, cf_totals_row):
    """Create the Results Dashboard."""

    ws.set_column('A:A', 3)
    ws.set_column('B:B', 25)
    ws.set_column('C:C', 20)
    ws.set_column('D:D', 15)
    ws.set_column('E:E', 40)

    row = 1
    ws.merge_range('B2:E2', 'BESS Economic Analysis Results', formats['title'])
    row = 4

    # Project Summary
    ws.merge_range(f'B{row}:E{row}', 'PROJECT SUMMARY', formats['section'])
    row += 1

    summary = [
        ('Project', '=Inputs!C7'),
        ('Location', '=Inputs!C9'),
        ('Capacity', '=Inputs!C10&" MW / "&Inputs!C12&" MWh"'),
        ('Total Investment', '=Calculations!C9'),
    ]

    for label, formula in summary:
        ws.write(f'B{row}', label, formats['bold'])
        if 'Investment' in label:
            ws.write_formula(f'C{row}', formula, formats['result_currency'])
        else:
            ws.write_formula(f'C{row}', formula, formats['result'])
        row += 1

    row += 1

    # Key Metrics
    ws.merge_range(f'B{row}:E{row}', 'KEY FINANCIAL METRICS', formats['section'])
    row += 1

    bcr_row = row
    metrics = [
        ('Benefit-Cost Ratio (BCR)', f'=Cash_Flows!W{cf_totals_row+1}/Cash_Flows!V{cf_totals_row+1}',
         '0.00', 'BCR >= 1.0 indicates benefits exceed costs'),
        ('Net Present Value (NPV)', f'=Cash_Flows!W{cf_totals_row+1}-Cash_Flows!V{cf_totals_row+1}',
         '$#,##0', 'Positive NPV indicates value creation'),
        ('Internal Rate of Return (IRR)', '=IRR(Cash_Flows!T5:T25)',
         '0.0%', 'Rate where NPV = 0'),
        ('Simple Payback (years)', '=MATCH(TRUE,INDEX(Cash_Flows!X5:X25>0,0),0)-1',
         '0.0', 'Year when cumulative CF turns positive'),
        ('PV of Total Costs', f'=Cash_Flows!V{cf_totals_row+1}',
         '$#,##0,,"M"', 'Present value of all costs'),
        ('PV of Total Benefits', f'=Cash_Flows!W{cf_totals_row+1}',
         '$#,##0,,"M"', 'Present value of all benefits'),
    ]

    for label, formula, num_fmt, desc in metrics:
        ws.write(f'B{row}', label, formats['bold'])
        fmt = workbook.add_format({
            'bold': True, 'font_size': 14, 'bg_color': '#E3F2FD',
            'align': 'center', 'border': 1, 'num_format': num_fmt
        })
        ws.write_formula(f'C{row}', formula, fmt)
        ws.write(f'E{row}', desc, formats['tooltip'])
        row += 1

    # Add conditional formatting to BCR
    ws.conditional_format(f'C{bcr_row}', {
        'type': 'cell', 'criteria': '>=', 'value': 1.5,
        'format': formats['positive']
    })
    ws.conditional_format(f'C{bcr_row}', {
        'type': 'cell', 'criteria': 'between', 'minimum': 1, 'maximum': 1.5,
        'format': formats['neutral']
    })
    ws.conditional_format(f'C{bcr_row}', {
        'type': 'cell', 'criteria': '<', 'value': 1,
        'format': formats['negative']
    })

    row += 1

    # Recommendation
    ws.merge_range(f'B{row}:E{row}', 'RECOMMENDATION', formats['section'])
    row += 1

    ws.merge_range(f'B{row}:E{row}',
        f'=IF(C{bcr_row}>=1.5,"APPROVE - Strong economic case",'
        f'IF(C{bcr_row}>=1,"FURTHER STUDY - Marginal economics","REJECT - Costs exceed benefits"))',
        formats['result'])
    row += 2

    # LCOS
    ws.merge_range(f'B{row}:E{row}', 'LEVELIZED COST OF STORAGE (LCOS)', formats['section'])
    row += 1

    ws.write(f'B{row}', 'LCOS ($/MWh)', formats['bold'])
    ws.write_formula(f'C{row}',
        f'=Cash_Flows!V{cf_totals_row+1}/(Calculations!C13*NPV(Inputs!C14,Cash_Flows!U5:U25))',
        formats['result_currency'])
    ws.write(f'E{row}', 'Levelized cost per MWh discharged', formats['tooltip'])
    row += 2

    # Breakeven
    ws.merge_range(f'B{row}:E{row}', 'BREAKEVEN ANALYSIS', formats['section'])
    row += 1

    ws.write(f'B{row}', 'Breakeven CapEx ($/kWh)', formats['bold'])
    ws.write_formula(f'C{row}',
        f'=(Cash_Flows!W{cf_totals_row+1}-(Cash_Flows!V{cf_totals_row+1}-Calculations!C11))/(Inputs!C12*1000)',
        formats['result_currency'])
    ws.write(f'E{row}', 'Maximum CapEx for BCR = 1.0', formats['tooltip'])


def create_sensitivity_sheet(workbook, ws, formats, cf_totals_row):
    """Create sensitivity analysis tables for key inputs."""

    ws.set_column('A:A', 3)
    ws.set_column('B:B', 20)
    for col in range(2, 12):  # C through L
        ws.set_column(col, col, 14)

    row = 1
    ws.merge_range('B2:L2', 'Sensitivity Analysis - NPV & BCR Impact', formats['title'])
    row = 4

    # Instructions
    ws.merge_range(f'B{row}:L{row}',
        'These tables show how key metrics change with different CapEx and Benefit multipliers.',
        formats['tooltip'])
    row += 2

    # === NPV SENSITIVITY TABLE ===
    ws.merge_range(f'B{row}:L{row}', 'NET PRESENT VALUE ($) SENSITIVITY', formats['section'])
    row += 1

    # Column headers - benefit multipliers
    ws.write(f'B{row}', 'NPV ($)', formats['header'])
    benefit_mults = [0.7, 0.8, 0.9, 1.0, 1.1, 1.2, 1.3]
    for i, mult in enumerate(benefit_mults):
        col_letter = chr(ord('C') + i)
        ws.write(f'{col_letter}{row}', f'{int(mult*100)}% Benefits', formats['header'])
    row += 1

    # Row headers - CapEx levels and calculations
    capex_levels = [100, 120, 140, 160, 180, 200, 220]  # $/kWh
    npv_table_start = row

    for capex in capex_levels:
        ws.write(f'B{row}', f'${capex}/kWh', formats['bold'])
        for i, ben_mult in enumerate(benefit_mults):
            col_letter = chr(ord('C') + i)
            # NPV = PV_Benefits * benefit_mult - PV_Costs * (capex/base_capex)
            # Simplified: adjust base NPV calculation
            formula = (f'=(Cash_Flows!W{cf_totals_row+1}*{ben_mult})-'
                      f'(Cash_Flows!V{cf_totals_row+1}*(1+({capex}-Inputs!$C$28)/Inputs!$C$28))')
            ws.write_formula(f'{col_letter}{row}', formula, formats['currency'])
        row += 1

    row += 2

    # === BCR SENSITIVITY TABLE ===
    ws.merge_range(f'B{row}:L{row}', 'BENEFIT-COST RATIO SENSITIVITY', formats['section'])
    row += 1

    # Column headers - benefit multipliers
    ws.write(f'B{row}', 'BCR', formats['header'])
    for i, mult in enumerate(benefit_mults):
        col_letter = chr(ord('C') + i)
        ws.write(f'{col_letter}{row}', f'{int(mult*100)}% Benefits', formats['header'])
    row += 1

    # Row headers - CapEx levels
    bcr_table_start = row
    for capex in capex_levels:
        ws.write(f'B{row}', f'${capex}/kWh', formats['bold'])
        for i, ben_mult in enumerate(benefit_mults):
            col_letter = chr(ord('C') + i)
            # BCR = (PV_Benefits * benefit_mult) / (PV_Costs * capex_ratio)
            formula = (f'=(Cash_Flows!W{cf_totals_row+1}*{ben_mult})/'
                      f'(Cash_Flows!V{cf_totals_row+1}*(1+({capex}-Inputs!$C$28)/Inputs!$C$28))')
            cell = f'{col_letter}{row}'
            ws.write_formula(cell, formula, formats['percent'])
        row += 1

    # Add conditional formatting to BCR table
    for r in range(bcr_table_start, row):
        ws.conditional_format(f'C{r}:I{r}', {
            'type': 'cell', 'criteria': '>=', 'value': 1.5,
            'format': formats['positive']
        })
        ws.conditional_format(f'C{r}:I{r}', {
            'type': 'cell', 'criteria': 'between', 'minimum': 1, 'maximum': 1.5,
            'format': formats['neutral']
        })
        ws.conditional_format(f'C{r}:I{r}', {
            'type': 'cell', 'criteria': '<', 'value': 1,
            'format': formats['negative']
        })

    row += 2

    # === SINGLE-VARIABLE SENSITIVITIES ===
    ws.merge_range(f'B{row}:E{row}', 'SINGLE VARIABLE IMPACTS', formats['section'])
    row += 1

    ws.write(f'B{row}', 'Parameter', formats['header'])
    ws.write(f'C{row}', 'Base Value', formats['header'])
    ws.write(f'D{row}', '-20% NPV', formats['header'])
    ws.write(f'E{row}', '+20% NPV', formats['header'])
    row += 1

    # Key sensitivities - all use formulas referencing inputs
    sensitivities = [
        ('CapEx ($/kWh)', 'Inputs!C28', 'Calculations!C9',
         '=Results!C12-(Calculations!C9*0.2/(1+Inputs!C14)^0)',  # -20% capex => +NPV
         '=Results!C12+(Calculations!C9*0.2/(1+Inputs!C14)^0)'), # +20% capex => -NPV
        ('Total Benefits', f'Cash_Flows!W{cf_totals_row+1}', 'per year',
         f'=Results!C12-Cash_Flows!W{cf_totals_row+1}*0.2',
         f'=Results!C12+Cash_Flows!W{cf_totals_row+1}*0.2'),
        ('Discount Rate', 'Inputs!C14', '7%',
         '=Results!C12*1.15',  # Lower discount => higher NPV
         '=Results!C12*0.85'), # Higher discount => lower NPV
        ('Cycles per Day', 'Inputs!C23', '1.0',
         '=Results!C12*0.85',  # Fewer cycles => less revenue
         '=Results!C12*1.15'), # More cycles => more revenue
    ]

    for param, base_ref, unit, low_formula, high_formula in sensitivities:
        ws.write(f'B{row}', param)
        if base_ref.startswith('='):
            ws.write_formula(f'C{row}', base_ref)
        else:
            ws.write_formula(f'C{row}', f'={base_ref}')
        ws.write_formula(f'D{row}', low_formula, formats['currency'])
        ws.write_formula(f'E{row}', high_formula, formats['currency'])
        row += 1

    row += 2

    # Notes
    ws.merge_range(f'B{row}:L{row}', 'NOTES', formats['section'])
    row += 1
    notes = [
        "• Green cells indicate BCR >= 1.5 (strong project economics)",
        "• Yellow cells indicate BCR between 1.0-1.5 (marginal economics)",
        "• Red cells indicate BCR < 1.0 (costs exceed benefits)",
        "• Sensitivity tables assume proportional scaling of CapEx-related costs",
        "• Single variable impacts show approximate NPV change for ±20% parameter change",
    ]
    for note in notes:
        ws.write(f'B{row}', note, formats['tooltip'])
        row += 1


def create_methodology_sheet(workbook, ws, formats):
    """Create methodology documentation."""

    ws.set_column('A:A', 3)
    ws.set_column('B:B', 90)

    row = 1
    ws.write('B2', 'Methodology & References', formats['title'])
    row = 4

    # Overview
    ws.write(f'B{row}', 'OVERVIEW', formats['section'])
    row += 1

    overview = ("This workbook performs economic analysis of battery energy storage system (BESS) projects "
                "using standard discounted cash flow (DCF) methodology. All calculations follow industry-standard "
                "approaches from NREL, Lazard, and the California Public Utilities Commission (CPUC).")
    ws.write(f'B{row}', overview, formats['wrap'])
    ws.set_row(row - 1, 45)
    row += 2

    # Formulas
    ws.write(f'B{row}', 'KEY FORMULAS', formats['section'])
    row += 1

    formulas = [
        ("Net Present Value (NPV)", "NPV = Sum(CFt / (1+r)^t) for t=0 to N",
         "Where CFt = cash flow in year t, r = discount rate, N = analysis period"),
        ("Benefit-Cost Ratio (BCR)", "BCR = PV(Benefits) / PV(Costs)",
         "BCR > 1.0 indicates project creates net value; BCR > 1.5 is strong"),
        ("Internal Rate of Return (IRR)", "IRR: Find r where NPV = 0",
         "The discount rate that makes NPV equal to zero"),
        ("Levelized Cost of Storage (LCOS)", "LCOS = PV(Costs) / PV(Energy Discharged)",
         "Expressed in $/MWh; enables comparison across technologies"),
        ("Technology Learning Curve", "Cost(year) = Cost(base) x (1 - learning_rate)^years",
         "Battery costs decline ~10-15% annually; reduces augmentation costs"),
    ]

    for name, formula, desc in formulas:
        ws.write(f'B{row}', name, formats['bold'])
        row += 1
        ws.write(f'B{row}', f"Formula: {formula}", formats['tooltip'])
        row += 1
        ws.write(f'B{row}', desc)
        row += 2

    # References
    ws.write(f'B{row}', 'REFERENCES', formats['section'])
    row += 1

    references = [
        "[1] NREL. Annual Technology Baseline 2024. https://atb.nrel.gov/",
        "[2] Lazard. Levelized Cost of Storage Analysis, Version 10.0. March 2025.",
        "[3] CPUC. California Standard Practice Manual. October 2001.",
        "[4] E3. CPUC Avoided Cost Calculator 2024.",
        "[5] LBNL. Interruption Cost Estimate (ICE) Calculator.",
        "[6] EPA. Social Cost of Greenhouse Gases. 2024.",
        "[7] Brealey, Myers, Allen. Principles of Corporate Finance (13th ed.). 2020.",
        "[8] NREL. Storage Futures Study. NREL/TP-6A20-77449. 2021.",
    ]

    for ref in references:
        ws.write(f'B{row}', ref, formats['wrap'])
        row += 1

    row += 1
    ws.write(f'B{row}', 'DISCLAIMER', formats['section'])
    row += 1

    disclaimer = ("This model is provided for analytical purposes only. Results should be validated against "
                  "independent sources before use in investment decisions. Actual project economics will vary "
                  "based on site-specific conditions, market dynamics, and evolving technology costs.")
    ws.write(f'B{row}', disclaimer, formats['wrap'])
    ws.set_row(row - 1, 45)


def create_vba_code_sheet(workbook, ws, formats):
    """Create VBA code reference sheet with complete macros and instructions."""

    ws.set_column('A:A', 3)
    ws.set_column('B:B', 100)
    ws.set_column('C:C', 15)

    # Create code format
    code_fmt = workbook.add_format({
        'font_name': 'Consolas',
        'font_size': 9,
        'text_wrap': True,
        'valign': 'top',
        'bg_color': '#F5F5F5',
        'border': 1,
        'border_color': '#CCCCCC'
    })

    step_fmt = workbook.add_format({
        'bold': True,
        'bg_color': '#E3F2FD',
        'border': 1
    })

    row = 1
    ws.merge_range('B1:C1', 'VBA Macro Code & Setup Instructions', formats['title'])
    row = 3

    # Instructions section
    ws.merge_range(f'B{row}:C{row}', 'STEP-BY-STEP INSTRUCTIONS TO ENABLE MACROS', formats['section'])
    row += 1

    instructions = [
        ("Step 1:", "Save this workbook as a Macro-Enabled Workbook (.xlsm)"),
        ("", "File > Save As > Choose 'Excel Macro-Enabled Workbook (*.xlsm)'"),
        ("Step 2:", "Open the VBA Editor"),
        ("", "Press Alt + F11 (Windows) or Option + F11 (Mac)"),
        ("Step 3:", "Insert a new Module"),
        ("", "Right-click on 'VBAProject' in the left panel > Insert > Module"),
        ("Step 4:", "Copy the VBA code below and paste it into the Module"),
        ("", "Select all the code in the 'COMPLETE VBA CODE' section below"),
        ("Step 5:", "Close the VBA Editor (File > Close and Return to Microsoft Excel)"),
        ("Step 6:", "Create buttons on the Inputs sheet (optional)"),
        ("", "Developer tab > Insert > Button (Form Control) > Draw button > Assign macro"),
        ("", "Create buttons for: LoadNRELLibrary, LoadLazardLibrary, LoadCPUCLibrary"),
        ("Step 7:", "Save the workbook again to preserve the macros"),
    ]

    for label, text in instructions:
        if label:
            ws.write(f'B{row}', label, step_fmt)
            ws.write(f'C{row}', text)
        else:
            ws.write(f'B{row}', "    " + text, formats['tooltip'])
        row += 1

    row += 2

    # Complete VBA Code section
    ws.merge_range(f'B{row}:C{row}', 'COMPLETE VBA CODE - COPY ALL BELOW', formats['section'])
    row += 1

    vba_code = '''Option Explicit

'=================================================================
' BESS ANALYZER VBA MACROS
' Copy this entire module into your VBA project
'=================================================================

Sub LoadNRELLibrary()
    '-----------------------------------------------------------
    ' Loads NREL ATB 2024 Moderate assumptions
    '-----------------------------------------------------------
    With ThisWorkbook.Sheets("Inputs")
        ' Library name
        .Range("C6").Value = "NREL ATB 2024 - Moderate"

        ' Technology Specifications
        .Range("C18").Value = "LFP"          ' Chemistry
        .Range("C19").Value = 0.85           ' Round-Trip Efficiency
        .Range("C20").Value = 0.025          ' Annual Degradation
        .Range("C21").Value = 6000           ' Cycle Life
        .Range("C22").Value = 12             ' Augmentation Year
        .Range("C23").Value = 1              ' Cycles per Day

        ' Cost Inputs
        .Range("C28").Value = 160            ' CapEx ($/kWh)
        .Range("C29").Value = 25             ' Fixed O&M ($/kW-year)
        .Range("C30").Value = 0              ' Variable O&M ($/MWh)
        .Range("C31").Value = 55             ' Augmentation Cost ($/kWh)
        .Range("C32").Value = 10             ' Decommissioning ($/kW)
        .Range("C33").Value = 30             ' Charging Cost ($/MWh)
        .Range("C34").Value = 0.1            ' Residual Value (%)

        ' Tax Credits
        .Range("C37").Value = 0.3            ' ITC Base (30%)
        .Range("C38").Value = 0              ' ITC Adders

        ' Infrastructure Costs
        .Range("C43").Value = 100            ' Interconnection ($/kW)
        .Range("C44").Value = 10             ' Land ($/kW)
        .Range("C45").Value = 15             ' Permitting ($/kW)
        .Range("C46").Value = 0.005          ' Insurance (% of CapEx)
        .Range("C47").Value = 0.01           ' Property Tax (%)

        ' Financing Structure
        .Range("C51").Value = 0.6            ' Debt Percentage
        .Range("C52").Value = 0.045          ' Interest Rate
        .Range("C53").Value = 15             ' Loan Term (years)
        .Range("C54").Value = 0.1            ' Cost of Equity
        .Range("C55").Value = 0.21           ' Tax Rate

        ' Benefits ($/kW-year and escalation)
        .Range("C59").Value = 150: .Range("D59").Value = 0.02   ' Resource Adequacy
        .Range("C60").Value = 40: .Range("D60").Value = 0.015   ' Energy Arbitrage
        .Range("C61").Value = 15: .Range("D61").Value = 0.01    ' Ancillary Services
        .Range("C62").Value = 25: .Range("D62").Value = 0.02    ' T&D Deferral
        .Range("C63").Value = 50: .Range("D63").Value = 0.02    ' Resilience Value
        .Range("C64").Value = 25: .Range("D64").Value = 0.02    ' Renewable Integration
        .Range("C65").Value = 15: .Range("D65").Value = 0.03    ' GHG Emissions Value
        .Range("C66").Value = 8: .Range("D66").Value = 0.01     ' Voltage Support
    End With

    Application.Calculate
    MsgBox "NREL ATB 2024 Moderate assumptions loaded successfully.", vbInformation, "Library Loaded"
End Sub


Sub LoadLazardLibrary()
    '-----------------------------------------------------------
    ' Loads Lazard LCOS v10.0 2025 assumptions
    '-----------------------------------------------------------
    With ThisWorkbook.Sheets("Inputs")
        .Range("C6").Value = "Lazard LCOS 2025"

        ' Technology
        .Range("C18").Value = "LFP"
        .Range("C19").Value = 0.86
        .Range("C20").Value = 0.02
        .Range("C21").Value = 6500
        .Range("C22").Value = 12
        .Range("C23").Value = 1

        ' Costs
        .Range("C28").Value = 145
        .Range("C29").Value = 22
        .Range("C30").Value = 0.5
        .Range("C31").Value = 50
        .Range("C32").Value = 8
        .Range("C33").Value = 35
        .Range("C34").Value = 0.1

        ' Tax Credits
        .Range("C37").Value = 0.3
        .Range("C38").Value = 0

        ' Infrastructure
        .Range("C43").Value = 90
        .Range("C44").Value = 8
        .Range("C45").Value = 12
        .Range("C46").Value = 0.005
        .Range("C47").Value = 0.01

        ' Financing
        .Range("C51").Value = 0.55
        .Range("C52").Value = 0.05
        .Range("C53").Value = 15
        .Range("C54").Value = 0.12
        .Range("C55").Value = 0.21

        ' Benefits
        .Range("C59").Value = 140: .Range("D59").Value = 0.02
        .Range("C60").Value = 45: .Range("D60").Value = 0.02
        .Range("C61").Value = 12: .Range("D61").Value = 0.01
        .Range("C62").Value = 20: .Range("D62").Value = 0.015
        .Range("C63").Value = 45: .Range("D63").Value = 0.02
        .Range("C64").Value = 20: .Range("D64").Value = 0.02
        .Range("C65").Value = 12: .Range("D65").Value = 0.03
        .Range("C66").Value = 6: .Range("D66").Value = 0.01
    End With

    Application.Calculate
    MsgBox "Lazard LCOS 2025 assumptions loaded successfully.", vbInformation, "Library Loaded"
End Sub


Sub LoadCPUCLibrary()
    '-----------------------------------------------------------
    ' Loads CPUC California 2024 assumptions
    ' Includes 10% ITC Energy Community Adder
    '-----------------------------------------------------------
    With ThisWorkbook.Sheets("Inputs")
        .Range("C6").Value = "CPUC California 2024"

        ' Technology
        .Range("C18").Value = "LFP"
        .Range("C19").Value = 0.85
        .Range("C20").Value = 0.025
        .Range("C21").Value = 6000
        .Range("C22").Value = 12
        .Range("C23").Value = 1

        ' Costs
        .Range("C28").Value = 155
        .Range("C29").Value = 26
        .Range("C30").Value = 0
        .Range("C31").Value = 52
        .Range("C32").Value = 12
        .Range("C33").Value = 25
        .Range("C34").Value = 0.1

        ' Tax Credits (includes Energy Community adder)
        .Range("C37").Value = 0.3
        .Range("C38").Value = 0.1            ' 10% Energy Community Adder

        ' Infrastructure (California-specific higher costs)
        .Range("C43").Value = 120
        .Range("C44").Value = 15
        .Range("C45").Value = 20
        .Range("C46").Value = 0.005
        .Range("C47").Value = 0.0105         ' CA property tax rate

        ' Financing (IOU-style favorable terms)
        .Range("C51").Value = 0.65
        .Range("C52").Value = 0.04
        .Range("C53").Value = 20
        .Range("C54").Value = 0.095
        .Range("C55").Value = 0.21

        ' Benefits (California premium values)
        .Range("C59").Value = 180: .Range("D59").Value = 0.025  ' RA premium in CA
        .Range("C60").Value = 35: .Range("D60").Value = 0.02
        .Range("C61").Value = 10: .Range("D61").Value = 0.01
        .Range("C62").Value = 25: .Range("D62").Value = 0.015
        .Range("C63").Value = 60: .Range("D63").Value = 0.02   ' PSPS resilience value
        .Range("C64").Value = 30: .Range("D64").Value = 0.025
        .Range("C65").Value = 20: .Range("D65").Value = 0.03
        .Range("C66").Value = 10: .Range("D66").Value = 0.01
    End With

    Application.Calculate
    MsgBox "CPUC California 2024 assumptions loaded." & vbCrLf & _
           "Note: Includes 10% ITC Energy Community Adder.", vbInformation, "Library Loaded"
End Sub


Sub GenerateReport()
    '-----------------------------------------------------------
    ' Creates a summary report on a new "Report" sheet
    '-----------------------------------------------------------
    Dim wsReport As Worksheet
    Dim wsResults As Worksheet
    Dim wsInputs As Worksheet
    Dim row As Long

    Set wsResults = ThisWorkbook.Sheets("Results")
    Set wsInputs = ThisWorkbook.Sheets("Inputs")

    ' Delete existing Report sheet if it exists
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Report").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Create new Report sheet
    Set wsReport = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsReport.Name = "Report"

    row = 2
    wsReport.Cells(row, 2).Value = "BESS ECONOMIC ANALYSIS REPORT"
    wsReport.Cells(row, 2).Font.Bold = True
    wsReport.Cells(row, 2).Font.Size = 18
    wsReport.Range("B" & row & ":F" & row).Merge
    row = row + 2

    ' Project Summary
    wsReport.Cells(row, 2).Value = "PROJECT SUMMARY"
    wsReport.Cells(row, 2).Font.Bold = True
    wsReport.Cells(row, 2).Interior.Color = RGB(227, 242, 253)
    row = row + 1

    wsReport.Cells(row, 2).Value = "Project:"
    wsReport.Cells(row, 3).Value = wsInputs.Range("C7").Value
    row = row + 1
    wsReport.Cells(row, 2).Value = "Location:"
    wsReport.Cells(row, 3).Value = wsInputs.Range("C9").Value
    row = row + 1
    wsReport.Cells(row, 2).Value = "Capacity:"
    wsReport.Cells(row, 3).Value = wsInputs.Range("C10").Value & " MW / " & wsInputs.Range("C12").Value & " MWh"
    row = row + 1
    wsReport.Cells(row, 2).Value = "Assumptions:"
    wsReport.Cells(row, 3).Value = wsInputs.Range("C6").Value
    row = row + 2

    ' Key Metrics
    wsReport.Cells(row, 2).Value = "KEY FINANCIAL METRICS"
    wsReport.Cells(row, 2).Font.Bold = True
    wsReport.Cells(row, 2).Interior.Color = RGB(227, 242, 253)
    row = row + 1

    wsReport.Cells(row, 2).Value = "Benefit-Cost Ratio (BCR):"
    wsReport.Cells(row, 3).Value = wsResults.Range("C11").Value
    wsReport.Cells(row, 3).NumberFormat = "0.00"
    row = row + 1
    wsReport.Cells(row, 2).Value = "Net Present Value (NPV):"
    wsReport.Cells(row, 3).Value = wsResults.Range("C12").Value
    wsReport.Cells(row, 3).NumberFormat = "$#,##0"
    row = row + 1
    wsReport.Cells(row, 2).Value = "Internal Rate of Return (IRR):"
    wsReport.Cells(row, 3).Value = wsResults.Range("C13").Value
    wsReport.Cells(row, 3).NumberFormat = "0.0%"
    row = row + 1
    wsReport.Cells(row, 2).Value = "Payback Period:"
    wsReport.Cells(row, 3).Value = wsResults.Range("C14").Value
    wsReport.Cells(row, 3).NumberFormat = "0.0 "" years"""
    row = row + 1
    wsReport.Cells(row, 2).Value = "LCOS:"
    wsReport.Cells(row, 3).Value = wsResults.Range("C15").Value
    wsReport.Cells(row, 3).NumberFormat = "$#,##0 ""/MWh"""
    row = row + 2

    ' Recommendation
    wsReport.Cells(row, 2).Value = "RECOMMENDATION"
    wsReport.Cells(row, 2).Font.Bold = True
    wsReport.Cells(row, 2).Interior.Color = RGB(227, 242, 253)
    row = row + 1

    ' Color-code recommendation based on BCR
    Dim bcr As Double
    bcr = wsResults.Range("C11").Value
    If bcr >= 1.5 Then
        wsReport.Cells(row, 2).Value = "STRONG PROJECT - Proceed with development"
        wsReport.Cells(row, 2).Font.Color = RGB(0, 128, 0)
    ElseIf bcr >= 1 Then
        wsReport.Cells(row, 2).Value = "MARGINAL PROJECT - Review assumptions carefully"
        wsReport.Cells(row, 2).Font.Color = RGB(255, 165, 0)
    Else
        wsReport.Cells(row, 2).Value = "NOT RECOMMENDED - Costs exceed benefits"
        wsReport.Cells(row, 2).Font.Color = RGB(255, 0, 0)
    End If
    wsReport.Cells(row, 2).Font.Bold = True

    ' Format columns
    wsReport.Columns("B").ColumnWidth = 25
    wsReport.Columns("C").ColumnWidth = 35

    MsgBox "Report generated on 'Report' sheet.", vbInformation, "Report Ready"
End Sub


Sub ExportToPDF()
    '-----------------------------------------------------------
    ' Exports the Report sheet to PDF
    '-----------------------------------------------------------
    Dim wsReport As Worksheet
    Dim filePath As String

    On Error Resume Next
    Set wsReport = ThisWorkbook.Sheets("Report")
    On Error GoTo 0

    If wsReport Is Nothing Then
        MsgBox "Please run 'Generate Report' first to create the Report sheet.", vbExclamation, "No Report Found"
        Exit Sub
    End If

    filePath = Application.GetSaveAsFilename( _
        InitialFileName:="BESS_Analysis_Report.pdf", _
        FileFilter:="PDF Files (*.pdf), *.pdf", _
        Title:="Save Report as PDF")

    If filePath <> "False" Then
        wsReport.ExportAsFixedFormat Type:=xlTypePDF, Filename:=filePath
        MsgBox "PDF exported successfully to:" & vbCrLf & filePath, vbInformation, "Export Complete"
    End If
End Sub


Sub RefreshCalculations()
    '-----------------------------------------------------------
    ' Forces recalculation of all formulas
    '-----------------------------------------------------------
    Application.CalculateFull
    MsgBox "All calculations refreshed.", vbInformation, "Refresh Complete"
End Sub


Sub ShowAbout()
    '-----------------------------------------------------------
    ' Shows information about the BESS Analyzer
    '-----------------------------------------------------------
    MsgBox "BESS Economic Analyzer" & vbCrLf & vbCrLf & _
           "Version: 1.0" & vbCrLf & _
           "Libraries: NREL ATB 2024, Lazard LCOS 2025, CPUC CA 2024" & vbCrLf & vbCrLf & _
           "This tool performs benefit-cost analysis for utility-scale" & vbCrLf & _
           "battery energy storage systems (BESS) using industry-standard" & vbCrLf & _
           "assumptions and methodologies.", vbInformation, "About BESS Analyzer"
End Sub'''

    # Write code in multiple rows for visibility
    code_lines = vba_code.strip().split('\n')

    for line in code_lines:
        ws.write(row, 1, line, code_fmt)
        row += 1

    row += 2

    # Quick Reference section
    ws.merge_range(f'B{row}:C{row}', 'AVAILABLE MACROS - QUICK REFERENCE', formats['section'])
    row += 1

    macros = [
        ("LoadNRELLibrary", "Loads NREL Annual Technology Baseline 2024 (Moderate) assumptions"),
        ("LoadLazardLibrary", "Loads Lazard Levelized Cost of Storage v10.0 (2025) assumptions"),
        ("LoadCPUCLibrary", "Loads CPUC California 2024 assumptions with 10% ITC Energy Community Adder"),
        ("GenerateReport", "Creates a summary report on a new 'Report' sheet"),
        ("ExportToPDF", "Exports the Report sheet to a PDF file"),
        ("RefreshCalculations", "Forces full recalculation of all workbook formulas"),
        ("ShowAbout", "Displays information about the BESS Analyzer"),
    ]

    for macro_name, description in macros:
        ws.write(f'B{row}', macro_name, formats['bold'])
        ws.write(f'C{row}', description)
        row += 1

    row += 2

    # Troubleshooting section
    ws.merge_range(f'B{row}:C{row}', 'TROUBLESHOOTING', formats['section'])
    row += 1

    troubleshooting = [
        "If buttons don't work: Make sure you saved as .xlsm and macros are enabled",
        "If macros are disabled: Go to File > Options > Trust Center > Trust Center Settings > Macro Settings",
        "If cell references are wrong: The VBA code uses specific cell addresses - check the Inputs sheet layout",
        "To edit macros: Press Alt+F11 to open VBA Editor, then modify the code in the Module",
        "To assign macros to buttons: Right-click button > Assign Macro > Select from list",
    ]

    for tip in troubleshooting:
        ws.write(f'B{row}', "• " + tip, formats['tooltip'])
        row += 1


def create_library_data_sheet(workbook, ws, formats):
    """Create library reference data sheet."""

    ws.set_column('A:A', 35)
    ws.set_column('B:D', 18)

    row = 0
    ws.write(row, 0, 'Assumption Library Data', formats['title'])
    row = 2

    # Headers
    headers = ['Parameter', 'NREL ATB 2024', 'Lazard LCOS 2025', 'CPUC CA 2024']
    for col, h in enumerate(headers):
        ws.write(row, col, h, formats['header'])
    row += 1

    # Cost data
    cost_data = [
        ('CapEx ($/kWh)', 160, 145, 155),
        ('Fixed O&M ($/kW-yr)', 25, 22, 26),
        ('Variable O&M ($/MWh)', 0, 0.5, 0),
        ('Augmentation ($/kWh)', 55, 50, 52),
        ('Decommissioning ($/kW)', 10, 8, 12),
        ('Charging Cost ($/MWh)', 30, 35, 25),
        ('Residual Value', '10%', '10%', '10%'),
        ('Round-Trip Efficiency', '85%', '86%', '85%'),
        ('Annual Degradation', '2.5%', '2.0%', '2.5%'),
        ('Cycle Life', 6000, 6500, 6000),
        ('Cycles per Day', 1.0, 1.0, 1.0),
        ('Learning Rate', '12%', '10%', '11%'),
    ]

    for data in cost_data:
        for col, val in enumerate(data):
            ws.write(row, col, val)
        row += 1

    row += 1
    ws.write(row, 0, 'Tax Credits (BESS-Specific)', formats['bold'])
    row += 1

    tax_data = [
        ('ITC Base Rate', '30%', '30%', '30%'),
        ('ITC Adders', '0%', '0%', '10%'),
    ]
    for data in tax_data:
        for col, val in enumerate(data):
            ws.write(row, col, val)
        row += 1

    row += 1
    ws.write(row, 0, 'Infrastructure (Common)', formats['bold'])
    row += 1

    infra_data = [
        ('Interconnection ($/kW)', 100, 90, 120),
        ('Land ($/kW)', 10, 8, 15),
        ('Permitting ($/kW)', 15, 12, 20),
        ('Insurance (% CapEx)', '0.5%', '0.5%', '0.5%'),
        ('Property Tax (%)', '1.0%', '1.0%', '1.05%'),
    ]
    for data in infra_data:
        for col, val in enumerate(data):
            ws.write(row, col, val)
        row += 1

    row += 1
    ws.write(row, 0, 'Financing Structure', formats['bold'])
    row += 1

    financing_data = [
        ('Debt Percentage', '60%', '55%', '65%'),
        ('Interest Rate', '4.5%', '5.0%', '4.0%'),
        ('Loan Term (years)', 15, 15, 20),
        ('Cost of Equity', '10%', '12%', '9.5%'),
        ('Tax Rate', '21%', '21%', '21%'),
        ('Calculated WACC', '6.5%', '7.5%', '6.0%'),
    ]
    for data in financing_data:
        for col, val in enumerate(data):
            ws.write(row, col, val)
        row += 1

    row += 1
    ws.write(row, 0, 'Benefits ($/kW-yr)', formats['bold'])
    row += 1

    benefit_data = [
        ('Resource Adequacy [common]', 150, 140, 180),
        ('Energy Arbitrage [common]', 40, 45, 35),
        ('Ancillary Services [common]', 15, 12, 10),
        ('T&D Deferral [common]', 25, 20, 25),
        ('Resilience Value [common]', 50, 45, 60),
        ('Renewable Integration [bess]', 25, 20, 30),
        ('GHG Emissions Value [bess]', 15, 12, 20),
        ('Voltage Support [common]', 8, 6, 10),
    ]
    for data in benefit_data:
        for col, val in enumerate(data):
            if col > 0 and isinstance(val, (int, float)):
                ws.write(row, col, val, formats['currency'])
            else:
                ws.write(row, col, val)
        row += 1


def get_vba_code():
    """Return VBA code for the workbook macros."""
    return '''
Attribute VB_Name = "BESSAnalyzer"
Option Explicit

Sub LoadNRELLibrary()
    With ThisWorkbook.Sheets("Inputs")
        .Range("C6").Value = "NREL ATB 2024 - Moderate"
        ' Technology
        .Range("C18").Value = "LFP"
        .Range("C19").Value = 0.85
        .Range("C20").Value = 0.025
        .Range("C21").Value = 6000
        .Range("C22").Value = 12
        ' Costs
        .Range("C26").Value = 160
        .Range("C27").Value = 25
        .Range("C28").Value = 0
        .Range("C29").Value = 55
        .Range("C30").Value = 10
        ' Tax Credits
        .Range("C33").Value = 0.3
        .Range("C34").Value = 0
        ' Infrastructure
        .Range("C37").Value = 100
        .Range("C38").Value = 10
        .Range("C39").Value = 15
        .Range("C40").Value = 0.005
        .Range("C41").Value = 0.01
        ' Benefits
        .Range("C45").Value = 150: .Range("D45").Value = 0.02
        .Range("C46").Value = 40: .Range("D46").Value = 0.015
        .Range("C47").Value = 15: .Range("D47").Value = 0.01
        .Range("C48").Value = 25: .Range("D48").Value = 0.02
        .Range("C49").Value = 50: .Range("D49").Value = 0.02
        .Range("C50").Value = 25: .Range("D50").Value = 0.02
        .Range("C51").Value = 15: .Range("D51").Value = 0.03
        .Range("C52").Value = 8: .Range("D52").Value = 0.01
        ' Learning curve
        .Range("C53").Value = 0.12
        .Range("C54").Value = 2024
    End With
    Application.Calculate
    MsgBox "NREL ATB 2024 assumptions loaded.", vbInformation
End Sub

Sub LoadLazardLibrary()
    With ThisWorkbook.Sheets("Inputs")
        .Range("C6").Value = "Lazard LCOS 2025"
        .Range("C18").Value = "LFP"
        .Range("C19").Value = 0.86
        .Range("C20").Value = 0.02
        .Range("C21").Value = 6500
        .Range("C22").Value = 12
        .Range("C26").Value = 145
        .Range("C27").Value = 22
        .Range("C28").Value = 0.5
        .Range("C29").Value = 50
        .Range("C30").Value = 8
        .Range("C33").Value = 0.3
        .Range("C34").Value = 0
        .Range("C37").Value = 90
        .Range("C38").Value = 8
        .Range("C39").Value = 12
        .Range("C40").Value = 0.005
        .Range("C41").Value = 0.01
        .Range("C45").Value = 140: .Range("D45").Value = 0.02
        .Range("C46").Value = 45: .Range("D46").Value = 0.02
        .Range("C47").Value = 12: .Range("D47").Value = 0.01
        .Range("C48").Value = 20: .Range("D48").Value = 0.015
        .Range("C49").Value = 45: .Range("D49").Value = 0.02
        .Range("C50").Value = 20: .Range("D50").Value = 0.02
        .Range("C51").Value = 12: .Range("D51").Value = 0.03
        .Range("C52").Value = 6: .Range("D52").Value = 0.01
        .Range("C53").Value = 0.1
        .Range("C54").Value = 2025
    End With
    Application.Calculate
    MsgBox "Lazard LCOS 2025 assumptions loaded.", vbInformation
End Sub

Sub LoadCPUCLibrary()
    With ThisWorkbook.Sheets("Inputs")
        .Range("C6").Value = "CPUC California 2024"
        .Range("C18").Value = "LFP"
        .Range("C19").Value = 0.85
        .Range("C20").Value = 0.025
        .Range("C21").Value = 6000
        .Range("C22").Value = 12
        .Range("C26").Value = 155
        .Range("C27").Value = 26
        .Range("C28").Value = 0
        .Range("C29").Value = 52
        .Range("C30").Value = 12
        .Range("C33").Value = 0.3
        .Range("C34").Value = 0.1
        .Range("C37").Value = 120
        .Range("C38").Value = 15
        .Range("C39").Value = 20
        .Range("C40").Value = 0.005
        .Range("C41").Value = 0.0105
        .Range("C45").Value = 180: .Range("D45").Value = 0.025
        .Range("C46").Value = 35: .Range("D46").Value = 0.02
        .Range("C47").Value = 10: .Range("D47").Value = 0.01
        .Range("C48").Value = 25: .Range("D48").Value = 0.015
        .Range("C49").Value = 60: .Range("D49").Value = 0.02
        .Range("C50").Value = 30: .Range("D50").Value = 0.025
        .Range("C51").Value = 20: .Range("D51").Value = 0.03
        .Range("C52").Value = 10: .Range("D52").Value = 0.01
        .Range("C53").Value = 0.11
        .Range("C54").Value = 2024
    End With
    Application.Calculate
    MsgBox "CPUC California 2024 assumptions loaded." & vbCrLf & _
           "Note: Includes 10% ITC energy community adder.", vbInformation
End Sub

Sub GenerateReport()
    Dim wsReport As Worksheet
    Dim wsResults As Worksheet
    Dim wsInputs As Worksheet
    Dim row As Long

    Set wsResults = ThisWorkbook.Sheets("Results")
    Set wsInputs = ThisWorkbook.Sheets("Inputs")

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Report").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsReport = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsReport.Name = "Report"

    row = 2
    wsReport.Cells(row, 2).Value = "BESS ECONOMIC ANALYSIS REPORT"
    wsReport.Cells(row, 2).Font.Bold = True
    wsReport.Cells(row, 2).Font.Size = 18
    wsReport.Range("B" & row & ":F" & row).Merge
    row = row + 2

    wsReport.Cells(row, 2).Value = "PROJECT SUMMARY"
    wsReport.Cells(row, 2).Font.Bold = True
    wsReport.Cells(row, 2).Interior.Color = RGB(227, 242, 253)
    row = row + 1

    wsReport.Cells(row, 2).Value = "Project:"
    wsReport.Cells(row, 3).Value = wsInputs.Range("C7").Value
    row = row + 1
    wsReport.Cells(row, 2).Value = "Location:"
    wsReport.Cells(row, 3).Value = wsInputs.Range("C9").Value
    row = row + 1
    wsReport.Cells(row, 2).Value = "Capacity:"
    wsReport.Cells(row, 3).Value = wsInputs.Range("C10").Value & " MW / " & wsInputs.Range("C12").Value & " MWh"
    row = row + 1
    wsReport.Cells(row, 2).Value = "Assumptions:"
    wsReport.Cells(row, 3).Value = wsInputs.Range("C6").Value
    row = row + 2

    wsReport.Cells(row, 2).Value = "KEY METRICS"
    wsReport.Cells(row, 2).Font.Bold = True
    wsReport.Cells(row, 2).Interior.Color = RGB(227, 242, 253)
    row = row + 1

    wsReport.Cells(row, 2).Value = "BCR:"
    wsReport.Cells(row, 3).Value = wsResults.Range("C11").Value
    wsReport.Cells(row, 3).NumberFormat = "0.00"
    row = row + 1
    wsReport.Cells(row, 2).Value = "NPV:"
    wsReport.Cells(row, 3).Value = wsResults.Range("C12").Value
    wsReport.Cells(row, 3).NumberFormat = "$#,##0"
    row = row + 1
    wsReport.Cells(row, 2).Value = "IRR:"
    wsReport.Cells(row, 3).Value = wsResults.Range("C13").Value
    wsReport.Cells(row, 3).NumberFormat = "0.0%"
    row = row + 2

    wsReport.Cells(row, 2).Value = "RECOMMENDATION"
    wsReport.Cells(row, 2).Font.Bold = True
    wsReport.Cells(row, 2).Interior.Color = RGB(227, 242, 253)
    row = row + 1
    wsReport.Cells(row, 2).Value = wsResults.Range("B19").Value
    wsReport.Cells(row, 2).Font.Bold = True

    wsReport.Columns("B").ColumnWidth = 20
    wsReport.Columns("C").ColumnWidth = 30

    MsgBox "Report generated on 'Report' sheet.", vbInformation
End Sub

Sub ExportToPDF()
    Dim wsReport As Worksheet
    Dim filePath As String

    On Error Resume Next
    Set wsReport = ThisWorkbook.Sheets("Report")
    On Error GoTo 0

    If wsReport Is Nothing Then
        MsgBox "Please run 'Generate Report' first.", vbExclamation
        Exit Sub
    End If

    filePath = Application.GetSaveAsFilename( _
        InitialFileName:="BESS_Analysis_Report.pdf", _
        FileFilter:="PDF Files (*.pdf), *.pdf")

    If filePath <> "False" Then
        wsReport.ExportAsFixedFormat Type:=xlTypePDF, Filename:=filePath
        MsgBox "PDF exported to: " & filePath, vbInformation
    End If
End Sub

Sub RefreshCalculations()
    Application.CalculateFull
    MsgBox "All calculations refreshed.", vbInformation
End Sub
'''


def create_vba_helper_script():
    """Create a helper script to extract vbaProject.bin from an existing .xlsm file."""
    script_content = '''#!/usr/bin/env python3
"""Helper script to create vbaProject.bin for macro-enabled Excel files.

Usage:
    python create_vba_template.py <path_to_xlsm>

To create vbaProject.bin:
1. Open Microsoft Excel
2. Create a new workbook
3. Press Alt+F11 to open VBA Editor
4. Insert > Module
5. Paste the VBA code from get_vba_code() in excel_generator.py
6. Save as 'template.xlsm' (Macro-Enabled Workbook)
7. Run: python create_vba_template.py template.xlsm
"""

import sys
import zipfile
from pathlib import Path


def extract_vba_bin(xlsm_path):
    """Extract vbaProject.bin from an .xlsm file."""
    resources_dir = Path('./resources')
    resources_dir.mkdir(exist_ok=True)

    with zipfile.ZipFile(xlsm_path, 'r') as zf:
        try:
            vba_data = zf.read('xl/vbaProject.bin')
            output_path = resources_dir / 'vbaProject.bin'
            with open(output_path, 'wb') as f:
                f.write(vba_data)
            print(f"Successfully extracted: {output_path}")
            return True
        except KeyError:
            print("Error: No vbaProject.bin found in the .xlsm file")
            print("Make sure the workbook contains VBA macros.")
            return False


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
    else:
        extract_vba_bin(sys.argv[1])
'''
    script_path = Path(__file__).parent / 'create_vba_template.py'
    with open(script_path, 'w') as f:
        f.write(script_content)
    print(f"Created: {script_path}")


if __name__ == "__main__":
    output = sys.argv[1] if len(sys.argv) > 1 else "BESS_Analyzer"

    # Check for VBA project
    vba_path = Path(__file__).parent / 'resources' / 'vbaProject.bin'

    if not vba_path.exists():
        print("Note: vbaProject.bin not found. Creating helper script...")
        create_vba_helper_script()
        print("\nTo enable VBA macros:")
        print("1. Create a template.xlsm with the VBA code from get_vba_code()")
        print("2. Run: python create_vba_template.py template.xlsm")
        print("3. Re-run this script to generate macro-enabled workbook\n")

    create_workbook(output)
