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


# =============================================================================
# CELL REFERENCE REGISTRY
# =============================================================================
# All cell references are defined here to ensure consistency between
# the Excel formulas, VBA macros, and named ranges.

class CellRefs:
    """Central registry of all cell references in the Inputs sheet."""

    # Library Selection
    SELECTED_LIBRARY = 'C6'

    # Project Basics (starting row 9)
    PROJECT_NAME = 'C9'
    PROJECT_ID = 'C10'
    LOCATION = 'C11'
    CAPACITY_MW = 'C12'
    DURATION_HOURS = 'C13'
    ENERGY_MWH = 'C14'  # Formula: =C12*C13
    ANALYSIS_YEARS = 'C15'
    DISCOUNT_RATE = 'C16'
    OWNERSHIP_TYPE = 'C17'

    # Technology Specifications (starting row 19)
    CHEMISTRY = 'C19'
    ROUND_TRIP_EFFICIENCY = 'C20'
    ANNUAL_DEGRADATION = 'C21'
    CYCLE_LIFE = 'C22'
    AUGMENTATION_YEAR = 'C23'
    CYCLES_PER_DAY = 'C24'

    # Cost Inputs (starting row 26)
    CAPEX_PER_KWH = 'C26'
    FOM_PER_KW_YEAR = 'C27'
    VOM_PER_MWH = 'C28'
    AUGMENTATION_COST = 'C29'
    DECOMMISSIONING = 'C30'
    CHARGING_COST = 'C31'
    RESIDUAL_VALUE = 'C32'

    # Tax Credits (starting row 34)
    ITC_BASE_RATE = 'C34'
    ITC_ADDERS = 'C35'

    # Infrastructure Costs (starting row 37)
    INTERCONNECTION = 'C37'
    LAND = 'C38'
    PERMITTING = 'C39'
    INSURANCE_PCT = 'C40'
    PROPERTY_TAX_PCT = 'C41'

    # Financing Structure (starting row 43)
    DEBT_PERCENT = 'C43'
    INTEREST_RATE = 'C44'
    LOAN_TERM = 'C45'
    COST_OF_EQUITY = 'C46'
    TAX_RATE = 'C47'
    WACC = 'C48'  # Formula

    # Benefits (starting row 52, after headers at row 51)
    BENEFIT_RA = 'C52'
    BENEFIT_RA_ESC = 'D52'
    BENEFIT_ARBITRAGE = 'C53'
    BENEFIT_ARBITRAGE_ESC = 'D53'
    BENEFIT_ANCILLARY = 'C54'
    BENEFIT_ANCILLARY_ESC = 'D54'
    BENEFIT_TD = 'C55'
    BENEFIT_TD_ESC = 'D55'
    BENEFIT_RESILIENCE = 'C56'
    BENEFIT_RESILIENCE_ESC = 'D56'
    BENEFIT_RENEWABLE = 'C57'
    BENEFIT_RENEWABLE_ESC = 'D57'
    BENEFIT_GHG = 'C58'
    BENEFIT_GHG_ESC = 'D58'
    BENEFIT_VOLTAGE = 'C59'
    BENEFIT_VOLTAGE_ESC = 'D59'

    # Benefit rows (for Cash_Flows formulas)
    BENEFIT_ROWS = [52, 53, 54, 55, 56, 57, 58, 59]

    # Cost Projections (starting row 61)
    LEARNING_RATE = 'C61'
    COST_BASE_YEAR = 'C62'

    # Bulk Discount (starting row 64)
    BULK_DISCOUNT_RATE = 'C64'
    BULK_DISCOUNT_THRESHOLD = 'C65'

    # Special Benefits - Reliability (starting row 68)
    RELIABILITY_ENABLED = 'C68'
    OUTAGE_HOURS = 'C69'
    CUSTOMER_COST_KWH = 'C70'
    BACKUP_CAPACITY_PCT = 'C71'

    # Special Benefits - Safety (starting row 74)
    SAFETY_ENABLED = 'C74'
    INCIDENT_PROBABILITY = 'C75'
    INCIDENT_COST = 'C76'
    RISK_REDUCTION = 'C77'

    # Special Benefits - Speed-to-Serve (starting row 80)
    SPEED_ENABLED = 'C80'
    MONTHS_SAVED = 'C81'
    VALUE_PER_KW_MONTH = 'C82'


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
    ws_uog = workbook.add_worksheet('UOG_Analysis')

    # Build each sheet
    create_inputs_sheet(workbook, ws_inputs, formats)
    cf_totals_row = create_cashflows_sheet(workbook, ws_cashflows, formats)
    create_calculations_sheet(workbook, ws_calculations, formats)
    create_results_sheet(workbook, ws_results, formats, cf_totals_row)
    create_sensitivity_sheet(workbook, ws_sensitivity, formats, cf_totals_row)
    create_methodology_sheet(workbook, ws_methodology, formats)
    create_library_data_sheet(workbook, ws_library_data, formats)
    create_vba_code_sheet(workbook, ws_vba_code, formats)
    create_uog_analysis_sheet(workbook, ws_uog, formats)

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

    # Row 6: Selected Library
    ws.write(f'B{row}', 'Selected Library:', formats['bold'])
    ws.write(f'C{row}', 'Custom', formats['input'])
    assert row == 6, f"Selected Library should be at row 6, got {row}"
    row += 2

    # === PROJECT BASICS ===
    # Row 8: Section header
    ws.merge_range(f'B{row}:E{row}', 'PROJECT BASICS', formats['section'])
    row += 1

    # Row 9-17: Project basics
    basics = [
        ('Project Name', '', 'Text', 'Enter project name'),                           # Row 9
        ('Project ID', '', 'Text', 'Unique identifier'),                              # Row 10
        ('Location', '', 'Text', 'Site or market location'),                          # Row 11
        ('Capacity (MW)', 100, 'Number', 'Nameplate power capacity'),                 # Row 12
        ('Duration (hours)', 4, 'Number', 'Storage duration'),                        # Row 13
        ('Energy Capacity (MWh)', '=C12*C13', 'Formula', 'Auto-calculated: MW x hours'),  # Row 14
        ('Analysis Period (years)', 20, 'Number', 'Economic analysis horizon'),       # Row 15
        ('Discount Rate (%)', 0.07, 'Percent', 'Nominal discount rate for NPV'),      # Row 16
        ('Ownership Type', 'Utility', 'Text', 'Utility (6-7% WACC) or Merchant (8-10%)'),  # Row 17
    ]

    basics_start = row
    assert basics_start == 9, f"Basics should start at row 9, got {basics_start}"
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

    # === TECHNOLOGY SPECIFICATIONS ===
    # Row 18: Section header
    ws.merge_range(f'B{row}:E{row}', 'TECHNOLOGY SPECIFICATIONS', formats['section'])
    row += 1

    # Row 19-24: Technology specs
    tech_specs = [
        ('Chemistry', 'LFP', 'Text', 'LFP, NMC, or Other'),                           # Row 19
        ('Round-Trip Efficiency (%)', 0.85, 'Percent', 'AC-AC efficiency'),           # Row 20
        ('Annual Degradation (%)', 0.025, 'Percent', 'Capacity loss per year'),       # Row 21
        ('Cycle Life', 6000, 'Number', 'Full-depth cycles before EOL'),               # Row 22
        ('Augmentation Year', 12, 'Number', 'Year of battery replacement'),           # Row 23
        ('Cycles per Day', 1.0, 'Number', 'Average daily charge/discharge cycles'),   # Row 24
    ]

    tech_start = row
    assert tech_start == 19, f"Tech should start at row 19, got {tech_start}"
    for label, value, dtype, tooltip in tech_specs:
        ws.write(f'B{row}', label, formats['bold'])
        if dtype == 'Percent':
            ws.write(f'C{row}', value, formats['input_percent'])
        else:
            ws.write(f'C{row}', value, formats['input'])
        ws.write(f'E{row}', tooltip, formats['tooltip'])
        row += 1

    # === COST INPUTS ===
    # Row 25: Section header
    ws.merge_range(f'B{row}:E{row}', 'COST INPUTS (BESS)', formats['section'])
    row += 1

    # Row 26-32: Cost inputs
    cost_inputs = [
        ('CapEx ($/kWh)', 160, '$/kWh', 'Installed capital cost per kWh'),            # Row 26
        ('Fixed O&M ($/kW-year)', 25, '$/kW-yr', 'Annual fixed O&M'),                 # Row 27
        ('Variable O&M ($/MWh)', 0, '$/MWh', 'Per-MWh discharge cost'),               # Row 28
        ('Augmentation Cost ($/kWh)', 55, '$/kWh', 'Battery replacement cost'),       # Row 29
        ('Decommissioning ($/kW)', 10, '$/kW', 'End-of-life cost'),                   # Row 30
        ('Charging Cost ($/MWh)', 30, '$/MWh', 'Grid electricity cost for charging'), # Row 31
        ('Residual Value (%)', 0.10, '%', 'End-of-life asset value as % of CapEx'),   # Row 32
    ]

    cost_start = row
    assert cost_start == 26, f"Costs should start at row 26, got {cost_start}"
    for label, value, unit, tooltip in cost_inputs:
        ws.write(f'B{row}', label, formats['bold'])
        if '%' in unit:
            ws.write(f'C{row}', value, formats['input_percent'])
        else:
            ws.write(f'C{row}', value, formats['input_currency'])
        ws.write(f'D{row}', unit)
        ws.write(f'E{row}', tooltip, formats['tooltip'])
        row += 1

    # === TAX CREDITS ===
    # Row 33: Section header
    ws.merge_range(f'B{row}:E{row}', 'TAX CREDITS (BESS-Specific under IRA)', formats['section'])
    row += 1

    # Row 34-35: Tax credits
    tax_inputs = [
        ('ITC Base Rate (%)', 0.30, '30% Investment Tax Credit under IRA'),           # Row 34
        ('ITC Adders (%)', 0.0, 'Energy community +10%, Domestic content +10%'),      # Row 35
    ]

    itc_start = row
    assert itc_start == 34, f"ITC should start at row 34, got {itc_start}"
    for label, value, tooltip in tax_inputs:
        ws.write(f'B{row}', label, formats['bold'])
        ws.write(f'C{row}', value, formats['input_percent'])
        ws.write(f'E{row}', tooltip, formats['tooltip'])
        row += 1

    # === INFRASTRUCTURE COSTS ===
    # Row 36: Section header
    ws.merge_range(f'B{row}:E{row}', 'INFRASTRUCTURE COSTS (Common to all projects)', formats['section'])
    row += 1

    # Row 37-41: Infrastructure
    infra_inputs = [
        ('Interconnection ($/kW)', 100, '$/kW', 'Network upgrades, studies'),         # Row 37
        ('Land ($/kW)', 10, '$/kW', 'Site acquisition/lease'),                        # Row 38
        ('Permitting ($/kW)', 15, '$/kW', 'Permits, environmental review'),           # Row 39
        ('Insurance (% of CapEx)', 0.005, '%', 'Annual insurance cost'),              # Row 40
        ('Property Tax (%)', 0.01, '%', 'Annual property tax rate'),                  # Row 41
    ]

    infra_start = row
    assert infra_start == 37, f"Infrastructure should start at row 37, got {infra_start}"
    for label, value, unit, tooltip in infra_inputs:
        ws.write(f'B{row}', label, formats['bold'])
        if '%' in unit:
            ws.write(f'C{row}', value, formats['input_percent'])
        else:
            ws.write(f'C{row}', value, formats['input_currency'])
        ws.write(f'D{row}', unit)
        ws.write(f'E{row}', tooltip, formats['tooltip'])
        row += 1

    # === FINANCING STRUCTURE ===
    # Row 42: Section header
    ws.merge_range(f'B{row}:E{row}', 'FINANCING STRUCTURE (For WACC Calculation)', formats['section'])
    row += 1

    # Row 43-47: Financing
    financing_inputs = [
        ('Debt Percentage (%)', 0.60, 'Debt/equity split (0.60 = 60% debt)'),         # Row 43
        ('Interest Rate (%)', 0.05, 'Annual interest rate on debt'),                  # Row 44
        ('Loan Term (years)', 15, 'Debt amortization period'),                        # Row 45
        ('Cost of Equity (%)', 0.10, 'Required return on equity'),                    # Row 46
        ('Tax Rate (%)', 0.21, 'Corporate tax rate for interest deduction'),          # Row 47
    ]

    financing_start = row
    assert financing_start == 43, f"Financing should start at row 43, got {financing_start}"
    for label, value, tooltip in financing_inputs:
        ws.write(f'B{row}', label, formats['bold'])
        if 'years' in label:
            ws.write(f'C{row}', value, formats['input'])
        else:
            ws.write(f'C{row}', value, formats['input_percent'])
        ws.write(f'E{row}', tooltip, formats['tooltip'])
        row += 1

    # Row 48: WACC formula
    ws.write(f'B{row}', 'Calculated WACC', formats['bold'])
    # WACC = (1-D) * Re + D * Rd * (1 - Tc)
    # D=C43, Rd=C44, Re=C46, Tc=C47
    wacc_formula = '=(1-C43)*C46+C43*C44*(1-C47)'
    ws.write_formula(f'C{row}', wacc_formula, formats['formula'])
    ws.write(f'E{row}', 'Weighted Average Cost of Capital', formats['tooltip'])
    assert row == 48, f"WACC should be at row 48, got {row}"
    row += 2  # Row 49: blank, Row 50

    # === BENEFIT STREAMS ===
    # Row 50: Section header
    ws.merge_range(f'B{row}:E{row}', 'BENEFIT STREAMS (Year 1 Values)', formats['section'])
    row += 1

    # Row 51: Headers
    ws.write(f'B{row}', 'Benefit Category', formats['header'])
    ws.write(f'C{row}', '$/kW-year', formats['header'])
    ws.write(f'D{row}', 'Escalation %', formats['header'])
    ws.write(f'E{row}', 'Category / Citation', formats['header'])
    row += 1

    # Row 52-59: Benefits
    benefits = [
        ('Resource Adequacy', 150, 0.02, '[common] CPUC RA Report 2024'),             # Row 52
        ('Energy Arbitrage', 40, 0.015, '[common] Market Data 2024'),                 # Row 53
        ('Ancillary Services', 15, 0.01, '[common] AS Reports 2024'),                 # Row 54
        ('T&D Deferral', 25, 0.02, '[common] Avoided Cost Calculator'),               # Row 55
        ('Resilience Value', 50, 0.02, '[common] LBNL ICE Calculator'),               # Row 56
        ('Renewable Integration', 25, 0.02, '[bess] NREL Grid Studies'),              # Row 57
        ('GHG Emissions Value', 15, 0.03, '[bess] EPA Social Cost Carbon'),           # Row 58
        ('Voltage Support', 8, 0.01, '[common] EPRI Distribution'),                   # Row 59
    ]

    benefits_start = row
    assert benefits_start == 52, f"Benefits should start at row 52, got {benefits_start}"
    for name, value, esc, cite in benefits:
        ws.write(f'B{row}', name)
        ws.write(f'C{row}', value, formats['input_currency'])
        ws.write(f'D{row}', esc, formats['input_percent'])
        ws.write(f'E{row}', cite, formats['tooltip'])
        row += 1

    # === COST PROJECTIONS ===
    # Row 60: Section header
    ws.merge_range(f'B{row}:E{row}', 'COST PROJECTIONS (Learning Curve)', formats['section'])
    row += 1

    # Row 61: Learning Rate
    ws.write(f'B{row}', 'Annual Cost Decline Rate', formats['bold'])
    ws.write(f'C{row}', 0.10, formats['input_percent'])
    ws.write(f'E{row}', 'Technology learning rate (10-15% typical)', formats['tooltip'])
    learning_rate_row = row
    assert learning_rate_row == 61, f"Learning rate should be at row 61, got {learning_rate_row}"
    row += 1

    # Row 62: Base Year
    ws.write(f'B{row}', 'Cost Base Year', formats['bold'])
    ws.write(f'C{row}', 2024, formats['input'])
    ws.write(f'E{row}', 'Reference year for base costs', formats['tooltip'])
    row += 1

    # === BULK DISCOUNT (for fleet purchases) ===
    # Row 63: Section header
    ws.merge_range(f'B{row}:E{row}', 'BULK DISCOUNT (Fleet Purchases)', formats['section'])
    row += 1

    # Row 64: Bulk Discount Rate
    ws.write(f'B{row}', 'Bulk Discount Rate (%)', formats['bold'])
    ws.write(f'C{row}', 0.0, formats['input_percent'])
    ws.write(f'E{row}', 'Discount on ALL costs when buying fleet (e.g., 0.10 = 10%)', formats['tooltip'])
    assert row == 64, f"Bulk discount rate should be at row 64, got {row}"
    row += 1

    # Row 65: Bulk Discount Threshold
    ws.write(f'B{row}', 'Threshold Capacity (MWh)', formats['bold'])
    ws.write(f'C{row}', 0, formats['input'])
    ws.write(f'E{row}', 'Minimum MWh capacity to qualify for bulk discount', formats['tooltip'])
    row += 2

    # === SPECIAL BENEFITS - RELIABILITY ===
    # Row 67: Section header
    ws.merge_range(f'B{row}:E{row}', 'SPECIAL BENEFITS - Reliability (Avoided Outage Cost)', formats['section'])
    row += 1

    # Row 68-71: Reliability inputs
    ws.write(f'B{row}', 'Reliability Enabled', formats['bold'])
    ws.write(f'C{row}', 'No', formats['input'])
    ws.write(f'E{row}', 'Enter "Yes" or "No"', formats['tooltip'])
    assert row == 68, f"Reliability enabled should be at row 68, got {row}"
    row += 1

    ws.write(f'B{row}', 'Outage Hours per Year', formats['bold'])
    ws.write(f'C{row}', 4.0, formats['input'])
    ws.write(f'E{row}', 'Expected annual outage hours avoided', formats['tooltip'])
    row += 1

    ws.write(f'B{row}', 'Customer Cost ($/kWh)', formats['bold'])
    ws.write(f'C{row}', 10.0, formats['input_currency'])
    ws.write(f'E{row}', 'Cost to customers per kWh of outage (LBNL ICE)', formats['tooltip'])
    row += 1

    ws.write(f'B{row}', 'Backup Capacity (%)', formats['bold'])
    ws.write(f'C{row}', 0.50, formats['input_percent'])
    ws.write(f'E{row}', 'Fraction of capacity available for backup', formats['tooltip'])
    row += 2

    # === SPECIAL BENEFITS - SAFETY ===
    # Row 73: Section header
    ws.merge_range(f'B{row}:E{row}', 'SPECIAL BENEFITS - Safety (Avoided Incident Cost)', formats['section'])
    row += 1

    # Row 74-77: Safety inputs
    ws.write(f'B{row}', 'Safety Enabled', formats['bold'])
    ws.write(f'C{row}', 'No', formats['input'])
    ws.write(f'E{row}', 'Enter "Yes" or "No"', formats['tooltip'])
    assert row == 74, f"Safety enabled should be at row 74, got {row}"
    row += 1

    ws.write(f'B{row}', 'Incident Probability', formats['bold'])
    ws.write(f'C{row}', 0.001, formats['input'])
    ws.write(f'E{row}', 'Annual probability of grid safety incident (e.g., 0.001 = 0.1%)', formats['tooltip'])
    row += 1

    ws.write(f'B{row}', 'Incident Cost ($)', formats['bold'])
    ws.write(f'C{row}', 1000000, formats['input_currency'])
    ws.write(f'E{row}', 'Cost per safety incident', formats['tooltip'])
    row += 1

    ws.write(f'B{row}', 'Risk Reduction Factor', formats['bold'])
    ws.write(f'C{row}', 0.50, formats['input_percent'])
    ws.write(f'E{row}', 'Fraction of risk mitigated by BESS (e.g., 0.5 = 50%)', formats['tooltip'])
    row += 2

    # === SPECIAL BENEFITS - SPEED TO SERVE ===
    # Row 79: Section header
    ws.merge_range(f'B{row}:E{row}', 'SPECIAL BENEFITS - Speed-to-Serve (Faster Deployment)', formats['section'])
    row += 1

    # Row 80-82: Speed-to-Serve inputs
    ws.write(f'B{row}', 'Speed-to-Serve Enabled', formats['bold'])
    ws.write(f'C{row}', 'No', formats['input'])
    ws.write(f'E{row}', 'Enter "Yes" or "No" - ONE-TIME Year 1 benefit', formats['tooltip'])
    assert row == 80, f"Speed enabled should be at row 80, got {row}"
    row += 1

    ws.write(f'B{row}', 'Months Saved', formats['bold'])
    ws.write(f'C{row}', 24, formats['input'])
    ws.write(f'E{row}', 'Months faster than alternative (e.g., gas peaker)', formats['tooltip'])
    row += 1

    ws.write(f'B{row}', 'Value per kW-Month ($)', formats['bold'])
    ws.write(f'C{row}', 5.0, formats['input_currency'])
    ws.write(f'E{row}', 'Value of each month of earlier deployment per kW', formats['tooltip'])
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

    # Define named ranges using CellRefs
    workbook.define_name('Capacity_MW', f'=Inputs!${CellRefs.CAPACITY_MW}')
    workbook.define_name('Duration_Hours', f'=Inputs!${CellRefs.DURATION_HOURS}')
    workbook.define_name('Energy_MWh', f'=Inputs!${CellRefs.ENERGY_MWH}')
    workbook.define_name('Analysis_Years', f'=Inputs!${CellRefs.ANALYSIS_YEARS}')
    workbook.define_name('Discount_Rate', f'=Inputs!${CellRefs.DISCOUNT_RATE}')
    workbook.define_name('Learning_Rate', f'=Inputs!${CellRefs.LEARNING_RATE}')

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

    # Use CellRefs for all formulas
    calcs = [
        ('Capacity (kW)', f'=Inputs!{CellRefs.CAPACITY_MW}*1000', 'Convert MW to kW'),
        ('Capacity (kWh)', f'=Inputs!{CellRefs.ENERGY_MWH}*1000', 'Convert MWh to kWh'),
        ('Battery CapEx ($)', f'=Inputs!{CellRefs.CAPEX_PER_KWH}*C6', 'CapEx/kWh x kWh'),
        ('Infrastructure ($)', f'=(Inputs!{CellRefs.INTERCONNECTION}+Inputs!{CellRefs.LAND}+Inputs!{CellRefs.PERMITTING})*C5', 'Interconnect+Land+Permit'),
        ('Total CapEx ($)', '=C7+C8', 'Battery + Infrastructure'),
        ('ITC Credit ($)', f'=C7*(Inputs!{CellRefs.ITC_BASE_RATE}+Inputs!{CellRefs.ITC_ADDERS})', 'ITC on battery only'),
        ('Net Year 0 Cost ($)', '=C9-C10', 'CapEx minus ITC'),
        ('Annual Fixed O&M ($)', f'=Inputs!{CellRefs.FOM_PER_KW_YEAR}*C5', 'FOM/kW x kW'),
        ('Annual Energy (MWh)', f'=Inputs!{CellRefs.ENERGY_MWH}*Inputs!{CellRefs.CYCLES_PER_DAY}*365*Inputs!{CellRefs.ROUND_TRIP_EFFICIENCY}', 'MWh x cycles/day x 365 x RTE'),
        ('Annual Charging Cost ($)', f'=C13/Inputs!{CellRefs.ROUND_TRIP_EFFICIENCY}*Inputs!{CellRefs.CHARGING_COST}', 'Energy/RTE x charging cost'),
        ('Residual Value ($)', f'=C9*Inputs!{CellRefs.RESIDUAL_VALUE}', 'CapEx x residual %'),
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
    for col in range(2, 25):  # C through X
        ws.set_column(col, col, 14)

    ws.write('B2', 'Annual Cash Flow Projections', formats['title'])

    row = 4
    headers = ['Year', 'CapEx', 'Fixed O&M', 'Var O&M', 'Charging', 'Insurance', 'Prop Tax',
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
            ws.write_formula(row, col, '=0', formats['currency'])
        col += 1

        # Fixed O&M (years 1+)
        if year == 0:
            ws.write_formula(row, col, '=0', formats['currency'])
        else:
            ws.write_formula(row, col, '=Calculations!C12', formats['currency'])
        col += 1

        # Variable O&M (formula references input cell)
        if year == 0:
            ws.write_formula(row, col, '=0', formats['currency'])
        else:
            ws.write_formula(row, col, f'=Inputs!${CellRefs.VOM_PER_MWH}*Calculations!$C$13', formats['currency'])
        col += 1

        # Charging Cost (years 1+)
        if year == 0:
            ws.write_formula(row, col, '=0', formats['currency'])
        else:
            ws.write_formula(row, col, '=Calculations!C14', formats['currency'])
        col += 1

        # Insurance (years 1+)
        if year == 0:
            ws.write_formula(row, col, '=0', formats['currency'])
        else:
            ws.write_formula(row, col, f'=Calculations!C9*Inputs!${CellRefs.INSURANCE_PCT}', formats['currency'])
        col += 1

        # Property Tax (declining with depreciation)
        if year == 0:
            ws.write_formula(row, col, '=0', formats['currency'])
        else:
            formula = f'=Calculations!$C$9*(1-{year}/Inputs!${CellRefs.ANALYSIS_YEARS})*Inputs!${CellRefs.PROPERTY_TAX_PCT}'
            ws.write_formula(row, col, formula, formats['currency'])
        col += 1

        # Augmentation (at augmentation year) with learning curve
        aug_year_ref = CellRefs.AUGMENTATION_YEAR.replace('C', '')
        if year == 0:
            ws.write_formula(row, col, '=0', formats['currency'])
        else:
            # Check if this year equals the augmentation year
            formula = f'=IF(B{row+1}=Inputs!${CellRefs.AUGMENTATION_YEAR},Inputs!${CellRefs.AUGMENTATION_COST}*Calculations!C6*(1-Inputs!${CellRefs.LEARNING_RATE})^B{row+1},0)'
            ws.write_formula(row, col, formula, formats['currency'])
        col += 1

        # Decommissioning (last year)
        if year == 0:
            ws.write_formula(row, col, '=0', formats['currency'])
        else:
            formula = f'=IF(B{row+1}=Inputs!${CellRefs.ANALYSIS_YEARS},Inputs!${CellRefs.DECOMMISSIONING}*Calculations!C5-Calculations!C15,0)'
            ws.write_formula(row, col, formula, formats['currency'])
        col += 1

        # Total Costs (sum of columns C through J)
        ws.write_formula(row, col, f'=SUM(C{row+1}:J{row+1})', formats['formula_currency'])
        col += 1

        # Benefits (8 streams with escalation)
        benefit_rows = CellRefs.BENEFIT_ROWS
        if year == 0:
            for _ in range(8):
                ws.write_formula(row, col, '=0', formats['currency'])
                col += 1
        else:
            for i, b_row in enumerate(benefit_rows):
                # Value * Capacity(kW) * (1+escalation)^(year-1) * degradation
                # Apply degradation to all benefits
                formula = (f'=Inputs!$C${b_row}*Inputs!${CellRefs.CAPACITY_MW}*1000'
                          f'*(1+Inputs!$D${b_row})^{year-1}'
                          f'*(1-Inputs!${CellRefs.ANNUAL_DEGRADATION})^{year-1}')
                ws.write_formula(row, col, formula, formats['currency'])
                col += 1

        # Total Benefits (sum of columns L through S)
        ws.write_formula(row, col, f'=SUM(L{row+1}:S{row+1})', formats['formula_currency'])
        col += 1

        # Net Cash Flow
        ws.write_formula(row, col, f'=T{row+1}-K{row+1}', formats['formula_currency'])
        col += 1

        # Discount Factor
        ws.write_formula(row, col, f'=1/(1+Inputs!${CellRefs.DISCOUNT_RATE})^B{row+1}', formats['percent'])
        col += 1

        # PV Costs
        ws.write_formula(row, col, f'=K{row+1}*V{row+1}', formats['currency'])
        col += 1

        # PV Benefits
        ws.write_formula(row, col, f'=T{row+1}*V{row+1}', formats['currency'])
        col += 1

        # Cumulative CF
        if year == 0:
            ws.write_formula(row, col, f'=U{row+1}', formats['currency'])
        else:
            ws.write_formula(row, col, f'=Y{row}+U{row+1}', formats['currency'])

        row += 1

    end_row = row - 1

    # Totals row
    row += 1
    ws.write(row, 1, 'TOTALS', formats['bold'])
    ws.write_formula(row, 10, f'=SUM(K{start_row+1}:K{end_row+1})', formats['formula_currency'])  # Total Costs
    ws.write_formula(row, 19, f'=SUM(T{start_row+1}:T{end_row+1})', formats['formula_currency'])  # Total Benefits
    ws.write_formula(row, 22, f'=SUM(W{start_row+1}:W{end_row+1})', formats['formula_currency'])  # PV Costs
    ws.write_formula(row, 23, f'=SUM(X{start_row+1}:X{end_row+1})', formats['formula_currency'])  # PV Benefits

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
        ('Project', f'=Inputs!{CellRefs.PROJECT_NAME}'),
        ('Location', f'=Inputs!{CellRefs.LOCATION}'),
        ('Capacity', f'=Inputs!{CellRefs.CAPACITY_MW}&" MW / "&Inputs!{CellRefs.ENERGY_MWH}&" MWh"'),
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
        ('Benefit-Cost Ratio (BCR)', f'=Cash_Flows!X{cf_totals_row+1}/Cash_Flows!W{cf_totals_row+1}',
         '0.00', 'BCR >= 1.0 indicates benefits exceed costs'),
        ('Net Present Value (NPV)', f'=Cash_Flows!X{cf_totals_row+1}-Cash_Flows!W{cf_totals_row+1}',
         '$#,##0', 'Positive NPV indicates value creation'),
        ('Internal Rate of Return (IRR)', '=IRR(Cash_Flows!U5:U25)',
         '0.0%', 'Rate where NPV = 0'),
        ('Simple Payback (years)', '=MATCH(TRUE,INDEX(Cash_Flows!Y5:Y25>0,0),0)-1',
         '0.0', 'Year when cumulative CF turns positive'),
        ('PV of Total Costs', f'=Cash_Flows!W{cf_totals_row+1}',
         '$#,##0,,"M"', 'Present value of all costs'),
        ('PV of Total Benefits', f'=Cash_Flows!X{cf_totals_row+1}',
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
        f'=Cash_Flows!W{cf_totals_row+1}/(Calculations!C13*NPV(Inputs!{CellRefs.DISCOUNT_RATE},Cash_Flows!V5:V25))',
        formats['result_currency'])
    ws.write(f'E{row}', 'Levelized cost per MWh discharged', formats['tooltip'])
    row += 2

    # Breakeven
    ws.merge_range(f'B{row}:E{row}', 'BREAKEVEN ANALYSIS', formats['section'])
    row += 1

    ws.write(f'B{row}', 'Breakeven CapEx ($/kWh)', formats['bold'])
    ws.write_formula(f'C{row}',
        f'=(Cash_Flows!X{cf_totals_row+1}-(Cash_Flows!W{cf_totals_row+1}-Calculations!C11))/(Inputs!{CellRefs.ENERGY_MWH}*1000)',
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
            formula = (f'=(Cash_Flows!X{cf_totals_row+1}*{ben_mult})-'
                      f'(Cash_Flows!W{cf_totals_row+1}*(1+({capex}-Inputs!${CellRefs.CAPEX_PER_KWH})/Inputs!${CellRefs.CAPEX_PER_KWH}))')
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
            formula = (f'=(Cash_Flows!X{cf_totals_row+1}*{ben_mult})/'
                      f'(Cash_Flows!W{cf_totals_row+1}*(1+({capex}-Inputs!${CellRefs.CAPEX_PER_KWH})/Inputs!${CellRefs.CAPEX_PER_KWH}))')
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

    # Key sensitivities
    sensitivities = [
        ('CapEx ($/kWh)', f'Inputs!{CellRefs.CAPEX_PER_KWH}', 'Calculations!C9',
         '=Results!C12-(Calculations!C9*0.2/(1+Inputs!C16)^0)',  # -20% capex => +NPV
         '=Results!C12+(Calculations!C9*0.2/(1+Inputs!C16)^0)'), # +20% capex => -NPV
        ('Total Benefits', f'Cash_Flows!X{cf_totals_row+1}', 'per year',
         f'=Results!C12-Cash_Flows!X{cf_totals_row+1}*0.2',
         f'=Results!C12+Cash_Flows!X{cf_totals_row+1}*0.2'),
        ('Discount Rate', f'Inputs!{CellRefs.DISCOUNT_RATE}', '7%',
         '=Results!C12*1.15',  # Lower discount => higher NPV
         '=Results!C12*0.85'), # Higher discount => lower NPV
        ('Cycles per Day', f'Inputs!{CellRefs.CYCLES_PER_DAY}', '1.0',
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

    # Generate VBA code with correct cell references from CellRefs
    vba_code = f'''Option Explicit

'=================================================================
' BESS ANALYZER VBA MACROS
' Copy this entire module into your VBA project
'
' CELL REFERENCE MAP (Inputs sheet):
' - Selected Library: {CellRefs.SELECTED_LIBRARY}
' - Capacity MW: {CellRefs.CAPACITY_MW}, Duration: {CellRefs.DURATION_HOURS}
' - Technology: {CellRefs.CHEMISTRY} to {CellRefs.CYCLES_PER_DAY}
' - Costs: {CellRefs.CAPEX_PER_KWH} to {CellRefs.RESIDUAL_VALUE}
' - ITC: {CellRefs.ITC_BASE_RATE}, {CellRefs.ITC_ADDERS}
' - Infrastructure: {CellRefs.INTERCONNECTION} to {CellRefs.PROPERTY_TAX_PCT}
' - Financing: {CellRefs.DEBT_PERCENT} to {CellRefs.TAX_RATE}
' - Benefits: C52:D59 (8 rows)
' - Learning Rate: {CellRefs.LEARNING_RATE}
' - Bulk Discount: {CellRefs.BULK_DISCOUNT_RATE}, {CellRefs.BULK_DISCOUNT_THRESHOLD}
' - Reliability: {CellRefs.RELIABILITY_ENABLED} to {CellRefs.BACKUP_CAPACITY_PCT}
' - Safety: {CellRefs.SAFETY_ENABLED} to {CellRefs.RISK_REDUCTION}
' - Speed-to-Serve: {CellRefs.SPEED_ENABLED} to {CellRefs.VALUE_PER_KW_MONTH}
'=================================================================

Sub LoadNRELLibrary()
    '-----------------------------------------------------------
    ' Loads NREL ATB 2024 Moderate assumptions
    '-----------------------------------------------------------
    With ThisWorkbook.Sheets("Inputs")
        ' Library name
        .Range("{CellRefs.SELECTED_LIBRARY}").Value = "NREL ATB 2024 - Moderate"

        ' Technology Specifications
        .Range("{CellRefs.CHEMISTRY}").Value = "LFP"
        .Range("{CellRefs.ROUND_TRIP_EFFICIENCY}").Value = 0.85
        .Range("{CellRefs.ANNUAL_DEGRADATION}").Value = 0.025
        .Range("{CellRefs.CYCLE_LIFE}").Value = 6000
        .Range("{CellRefs.AUGMENTATION_YEAR}").Value = 12
        .Range("{CellRefs.CYCLES_PER_DAY}").Value = 1

        ' Cost Inputs
        .Range("{CellRefs.CAPEX_PER_KWH}").Value = 160
        .Range("{CellRefs.FOM_PER_KW_YEAR}").Value = 25
        .Range("{CellRefs.VOM_PER_MWH}").Value = 0
        .Range("{CellRefs.AUGMENTATION_COST}").Value = 55
        .Range("{CellRefs.DECOMMISSIONING}").Value = 10
        .Range("{CellRefs.CHARGING_COST}").Value = 30
        .Range("{CellRefs.RESIDUAL_VALUE}").Value = 0.1

        ' Tax Credits
        .Range("{CellRefs.ITC_BASE_RATE}").Value = 0.3
        .Range("{CellRefs.ITC_ADDERS}").Value = 0

        ' Infrastructure Costs
        .Range("{CellRefs.INTERCONNECTION}").Value = 100
        .Range("{CellRefs.LAND}").Value = 10
        .Range("{CellRefs.PERMITTING}").Value = 15
        .Range("{CellRefs.INSURANCE_PCT}").Value = 0.005
        .Range("{CellRefs.PROPERTY_TAX_PCT}").Value = 0.01

        ' Financing Structure
        .Range("{CellRefs.DEBT_PERCENT}").Value = 0.6
        .Range("{CellRefs.INTEREST_RATE}").Value = 0.045
        .Range("{CellRefs.LOAN_TERM}").Value = 15
        .Range("{CellRefs.COST_OF_EQUITY}").Value = 0.1
        .Range("{CellRefs.TAX_RATE}").Value = 0.21

        ' Benefits ($/kW-year and escalation)
        .Range("C52").Value = 150: .Range("D52").Value = 0.02   ' Resource Adequacy
        .Range("C53").Value = 40: .Range("D53").Value = 0.015   ' Energy Arbitrage
        .Range("C54").Value = 15: .Range("D54").Value = 0.01    ' Ancillary Services
        .Range("C55").Value = 25: .Range("D55").Value = 0.02    ' T&D Deferral
        .Range("C56").Value = 50: .Range("D56").Value = 0.02    ' Resilience Value
        .Range("C57").Value = 25: .Range("D57").Value = 0.02    ' Renewable Integration
        .Range("C58").Value = 15: .Range("D58").Value = 0.03    ' GHG Emissions Value
        .Range("C59").Value = 8: .Range("D59").Value = 0.01     ' Voltage Support

        ' Learning Curve
        .Range("{CellRefs.LEARNING_RATE}").Value = 0.12

        ' Bulk Discount (disabled by default)
        .Range("{CellRefs.BULK_DISCOUNT_RATE}").Value = 0
        .Range("{CellRefs.BULK_DISCOUNT_THRESHOLD}").Value = 0

        ' Special Benefits - Reliability (disabled by default)
        .Range("{CellRefs.RELIABILITY_ENABLED}").Value = "No"
        .Range("{CellRefs.OUTAGE_HOURS}").Value = 4
        .Range("{CellRefs.CUSTOMER_COST_KWH}").Value = 10
        .Range("{CellRefs.BACKUP_CAPACITY_PCT}").Value = 0.5

        ' Special Benefits - Safety (disabled)
        .Range("{CellRefs.SAFETY_ENABLED}").Value = "No"
        .Range("{CellRefs.INCIDENT_PROBABILITY}").Value = 0.001
        .Range("{CellRefs.INCIDENT_COST}").Value = 500000
        .Range("{CellRefs.RISK_REDUCTION}").Value = 0.25

        ' Special Benefits - Speed-to-Serve (disabled)
        .Range("{CellRefs.SPEED_ENABLED}").Value = "No"
        .Range("{CellRefs.MONTHS_SAVED}").Value = 24
        .Range("{CellRefs.VALUE_PER_KW_MONTH}").Value = 5
    End With

    Application.Calculate
    MsgBox "NREL ATB 2024 Moderate assumptions loaded successfully.", vbInformation, "Library Loaded"
End Sub


Sub LoadLazardLibrary()
    '-----------------------------------------------------------
    ' Loads Lazard LCOS v10.0 2025 assumptions
    '-----------------------------------------------------------
    With ThisWorkbook.Sheets("Inputs")
        .Range("{CellRefs.SELECTED_LIBRARY}").Value = "Lazard LCOS 2025"

        ' Technology
        .Range("{CellRefs.CHEMISTRY}").Value = "LFP"
        .Range("{CellRefs.ROUND_TRIP_EFFICIENCY}").Value = 0.86
        .Range("{CellRefs.ANNUAL_DEGRADATION}").Value = 0.02
        .Range("{CellRefs.CYCLE_LIFE}").Value = 6500
        .Range("{CellRefs.AUGMENTATION_YEAR}").Value = 12
        .Range("{CellRefs.CYCLES_PER_DAY}").Value = 1

        ' Costs
        .Range("{CellRefs.CAPEX_PER_KWH}").Value = 145
        .Range("{CellRefs.FOM_PER_KW_YEAR}").Value = 22
        .Range("{CellRefs.VOM_PER_MWH}").Value = 0.5
        .Range("{CellRefs.AUGMENTATION_COST}").Value = 50
        .Range("{CellRefs.DECOMMISSIONING}").Value = 8
        .Range("{CellRefs.CHARGING_COST}").Value = 35
        .Range("{CellRefs.RESIDUAL_VALUE}").Value = 0.1

        ' Tax Credits
        .Range("{CellRefs.ITC_BASE_RATE}").Value = 0.3
        .Range("{CellRefs.ITC_ADDERS}").Value = 0

        ' Infrastructure
        .Range("{CellRefs.INTERCONNECTION}").Value = 90
        .Range("{CellRefs.LAND}").Value = 8
        .Range("{CellRefs.PERMITTING}").Value = 12
        .Range("{CellRefs.INSURANCE_PCT}").Value = 0.005
        .Range("{CellRefs.PROPERTY_TAX_PCT}").Value = 0.01

        ' Financing
        .Range("{CellRefs.DEBT_PERCENT}").Value = 0.55
        .Range("{CellRefs.INTEREST_RATE}").Value = 0.05
        .Range("{CellRefs.LOAN_TERM}").Value = 15
        .Range("{CellRefs.COST_OF_EQUITY}").Value = 0.12
        .Range("{CellRefs.TAX_RATE}").Value = 0.21

        ' Benefits
        .Range("C52").Value = 140: .Range("D52").Value = 0.02
        .Range("C53").Value = 45: .Range("D53").Value = 0.02
        .Range("C54").Value = 12: .Range("D54").Value = 0.01
        .Range("C55").Value = 20: .Range("D55").Value = 0.015
        .Range("C56").Value = 45: .Range("D56").Value = 0.02
        .Range("C57").Value = 20: .Range("D57").Value = 0.02
        .Range("C58").Value = 12: .Range("D58").Value = 0.03
        .Range("C59").Value = 6: .Range("D59").Value = 0.01

        ' Learning Curve
        .Range("{CellRefs.LEARNING_RATE}").Value = 0.1

        ' Bulk Discount (disabled by default)
        .Range("{CellRefs.BULK_DISCOUNT_RATE}").Value = 0
        .Range("{CellRefs.BULK_DISCOUNT_THRESHOLD}").Value = 0

        ' Special Benefits - Reliability (disabled by default)
        .Range("{CellRefs.RELIABILITY_ENABLED}").Value = "No"
        .Range("{CellRefs.OUTAGE_HOURS}").Value = 4
        .Range("{CellRefs.CUSTOMER_COST_KWH}").Value = 8
        .Range("{CellRefs.BACKUP_CAPACITY_PCT}").Value = 0.5

        ' Special Benefits - Safety (disabled)
        .Range("{CellRefs.SAFETY_ENABLED}").Value = "No"
        .Range("{CellRefs.INCIDENT_PROBABILITY}").Value = 0.001
        .Range("{CellRefs.INCIDENT_COST}").Value = 500000
        .Range("{CellRefs.RISK_REDUCTION}").Value = 0.25

        ' Special Benefits - Speed-to-Serve (disabled)
        .Range("{CellRefs.SPEED_ENABLED}").Value = "No"
        .Range("{CellRefs.MONTHS_SAVED}").Value = 24
        .Range("{CellRefs.VALUE_PER_KW_MONTH}").Value = 5
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
        .Range("{CellRefs.SELECTED_LIBRARY}").Value = "CPUC California 2024"

        ' Technology
        .Range("{CellRefs.CHEMISTRY}").Value = "LFP"
        .Range("{CellRefs.ROUND_TRIP_EFFICIENCY}").Value = 0.85
        .Range("{CellRefs.ANNUAL_DEGRADATION}").Value = 0.025
        .Range("{CellRefs.CYCLE_LIFE}").Value = 6000
        .Range("{CellRefs.AUGMENTATION_YEAR}").Value = 12
        .Range("{CellRefs.CYCLES_PER_DAY}").Value = 1

        ' Costs
        .Range("{CellRefs.CAPEX_PER_KWH}").Value = 155
        .Range("{CellRefs.FOM_PER_KW_YEAR}").Value = 26
        .Range("{CellRefs.VOM_PER_MWH}").Value = 0
        .Range("{CellRefs.AUGMENTATION_COST}").Value = 52
        .Range("{CellRefs.DECOMMISSIONING}").Value = 12
        .Range("{CellRefs.CHARGING_COST}").Value = 25
        .Range("{CellRefs.RESIDUAL_VALUE}").Value = 0.1

        ' Tax Credits (includes Energy Community adder)
        .Range("{CellRefs.ITC_BASE_RATE}").Value = 0.3
        .Range("{CellRefs.ITC_ADDERS}").Value = 0.1            ' 10% Energy Community Adder

        ' Infrastructure (California-specific higher costs)
        .Range("{CellRefs.INTERCONNECTION}").Value = 120
        .Range("{CellRefs.LAND}").Value = 15
        .Range("{CellRefs.PERMITTING}").Value = 20
        .Range("{CellRefs.INSURANCE_PCT}").Value = 0.005
        .Range("{CellRefs.PROPERTY_TAX_PCT}").Value = 0.0105   ' CA property tax rate

        ' Financing (IOU-style favorable terms)
        .Range("{CellRefs.DEBT_PERCENT}").Value = 0.65
        .Range("{CellRefs.INTEREST_RATE}").Value = 0.04
        .Range("{CellRefs.LOAN_TERM}").Value = 20
        .Range("{CellRefs.COST_OF_EQUITY}").Value = 0.095
        .Range("{CellRefs.TAX_RATE}").Value = 0.21

        ' Benefits (California premium values)
        .Range("C52").Value = 180: .Range("D52").Value = 0.025  ' RA premium in CA
        .Range("C53").Value = 35: .Range("D53").Value = 0.02
        .Range("C54").Value = 10: .Range("D54").Value = 0.01
        .Range("C55").Value = 25: .Range("D55").Value = 0.015
        .Range("C56").Value = 60: .Range("D56").Value = 0.02   ' PSPS resilience value
        .Range("C57").Value = 30: .Range("D57").Value = 0.025
        .Range("C58").Value = 20: .Range("D58").Value = 0.03
        .Range("C59").Value = 10: .Range("D59").Value = 0.01

        ' Learning Curve
        .Range("{CellRefs.LEARNING_RATE}").Value = 0.11

        ' Bulk Discount (disabled by default)
        .Range("{CellRefs.BULK_DISCOUNT_RATE}").Value = 0
        .Range("{CellRefs.BULK_DISCOUNT_THRESHOLD}").Value = 0

        ' Special Benefits - Reliability (ENABLED for California PSPS events)
        .Range("{CellRefs.RELIABILITY_ENABLED}").Value = "Yes"
        .Range("{CellRefs.OUTAGE_HOURS}").Value = 6
        .Range("{CellRefs.CUSTOMER_COST_KWH}").Value = 12
        .Range("{CellRefs.BACKUP_CAPACITY_PCT}").Value = 0.5

        ' Special Benefits - Safety (disabled)
        .Range("{CellRefs.SAFETY_ENABLED}").Value = "No"
        .Range("{CellRefs.INCIDENT_PROBABILITY}").Value = 0.0005
        .Range("{CellRefs.INCIDENT_COST}").Value = 750000
        .Range("{CellRefs.RISK_REDUCTION}").Value = 0.3

        ' Special Benefits - Speed-to-Serve (disabled)
        .Range("{CellRefs.SPEED_ENABLED}").Value = "No"
        .Range("{CellRefs.MONTHS_SAVED}").Value = 24
        .Range("{CellRefs.VALUE_PER_KW_MONTH}").Value = 6
    End With

    Application.Calculate
    MsgBox "CPUC California 2024 assumptions loaded." & vbCrLf & _
           "Note: Includes 10% ITC Energy Community Adder." & vbCrLf & _
           "Reliability benefits enabled for PSPS events.", vbInformation, "Library Loaded"
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
    wsReport.Cells(row, 3).Value = wsInputs.Range("{CellRefs.PROJECT_NAME}").Value
    row = row + 1
    wsReport.Cells(row, 2).Value = "Location:"
    wsReport.Cells(row, 3).Value = wsInputs.Range("{CellRefs.LOCATION}").Value
    row = row + 1
    wsReport.Cells(row, 2).Value = "Capacity:"
    wsReport.Cells(row, 3).Value = wsInputs.Range("{CellRefs.CAPACITY_MW}").Value & " MW / " & wsInputs.Range("{CellRefs.ENERGY_MWH}").Value & " MWh"
    row = row + 1
    wsReport.Cells(row, 2).Value = "Assumptions:"
    wsReport.Cells(row, 3).Value = wsInputs.Range("{CellRefs.SELECTED_LIBRARY}").Value
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

    # Cell Reference Map section
    ws.merge_range(f'B{row}:C{row}', 'CELL REFERENCE MAP (Inputs Sheet)', formats['section'])
    row += 1

    cell_refs = [
        ("Selected Library", CellRefs.SELECTED_LIBRARY),
        ("Capacity (MW)", CellRefs.CAPACITY_MW),
        ("Duration (hours)", CellRefs.DURATION_HOURS),
        ("Energy Capacity (MWh)", CellRefs.ENERGY_MWH),
        ("Analysis Period", CellRefs.ANALYSIS_YEARS),
        ("Discount Rate", CellRefs.DISCOUNT_RATE),
        ("Chemistry", CellRefs.CHEMISTRY),
        ("Round-Trip Efficiency", CellRefs.ROUND_TRIP_EFFICIENCY),
        ("Cycles per Day", CellRefs.CYCLES_PER_DAY),
        ("CapEx ($/kWh)", CellRefs.CAPEX_PER_KWH),
        ("Fixed O&M", CellRefs.FOM_PER_KW_YEAR),
        ("ITC Base Rate", CellRefs.ITC_BASE_RATE),
        ("ITC Adders", CellRefs.ITC_ADDERS),
        ("Interconnection", CellRefs.INTERCONNECTION),
        ("Debt Percentage", CellRefs.DEBT_PERCENT),
        ("Benefits Start Row", "C52"),
        ("Benefits End Row", "C59"),
        ("Learning Rate", CellRefs.LEARNING_RATE),
        ("Bulk Discount Rate", CellRefs.BULK_DISCOUNT_RATE),
        ("Bulk Discount Threshold", CellRefs.BULK_DISCOUNT_THRESHOLD),
        ("Reliability Enabled", CellRefs.RELIABILITY_ENABLED),
        ("Outage Hours", CellRefs.OUTAGE_HOURS),
        ("Customer Cost ($/kWh)", CellRefs.CUSTOMER_COST_KWH),
        ("Backup Capacity %", CellRefs.BACKUP_CAPACITY_PCT),
        ("Safety Enabled", CellRefs.SAFETY_ENABLED),
        ("Incident Probability", CellRefs.INCIDENT_PROBABILITY),
        ("Incident Cost", CellRefs.INCIDENT_COST),
        ("Risk Reduction", CellRefs.RISK_REDUCTION),
        ("Speed Enabled", CellRefs.SPEED_ENABLED),
        ("Months Saved", CellRefs.MONTHS_SAVED),
        ("Value per kW-Month", CellRefs.VALUE_PER_KW_MONTH),
    ]

    for name, cell in cell_refs:
        ws.write(f'B{row}', name)
        ws.write(f'C{row}', cell, formats['bold'])
        row += 1

    row += 2

    # Troubleshooting section
    ws.merge_range(f'B{row}:C{row}', 'TROUBLESHOOTING', formats['section'])
    row += 1

    troubleshooting = [
        "If buttons don't work: Make sure you saved as .xlsm and macros are enabled",
        "If macros are disabled: Go to File > Options > Trust Center > Trust Center Settings > Macro Settings",
        "If cell references are wrong: Check the Cell Reference Map above",
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


def create_uog_analysis_sheet(workbook, ws, formats):
    """Create the UOG (Utility-Owned Generation) Analysis sheet.

    Contains:
    - Revenue Requirement table (annual schedule over book life)
    - Ratepayer Impact analysis (Revenue Requirement vs Avoided Costs)
    - Cumulative Savings chart
    - Rate Base summary with ADIT and depreciation
    - Wires vs NWA comparison summary
    - Slice-of-Day feasibility summary
    """
    # Column widths
    ws.set_column('A:A', 30)
    ws.set_column('B:B', 18)
    ws.set_column('C:V', 14)

    row = 0

    # ========== Title ==========
    title_fmt = workbook.add_format({
        'bold': True, 'font_size': 16, 'font_color': '#1565C0',
        'bottom': 2, 'bottom_color': '#1565C0',
    })
    ws.write(row, 0, "Utility-Owned Storage (UOS) Revenue Requirement Analysis", title_fmt)
    ws.merge_range(row, 0, row, 7, "Utility-Owned Storage (UOS) Revenue Requirement Analysis", title_fmt)
    row += 1

    subtitle_fmt = workbook.add_format({
        'italic': True, 'font_color': '#666666', 'font_size': 10,
    })
    ws.write(row, 0, "Per CPUC D.25-12-003 (SCE 2026-2028 Cost of Capital)", subtitle_fmt)
    row += 2

    # ========== Section 1: Cost of Capital Summary ==========
    ws.write(row, 0, "SCE Cost of Capital (D.25-12-003)", formats['section'])
    ws.merge_range(row, 0, row, 2, "SCE Cost of Capital (D.25-12-003)", formats['section'])
    row += 1

    coc_labels = [
        ("Return on Equity (ROE)", "10.03%"),
        ("Cost of Long-Term Debt", "4.71%"),
        ("Cost of Preferred Stock", "5.48%"),
        ("Common Equity Ratio", "52.00%"),
        ("Long-Term Debt Ratio", "43.47%"),
        ("Preferred Stock Ratio", "4.53%"),
        ("Authorized Rate of Return (ROR)", "7.59%"),
        ("Federal Tax Rate", "21.00%"),
        ("State Tax Rate (CA)", "8.84%"),
        ("Composite Tax Rate", "27.98%"),
        ("Net-to-Gross Multiplier", "~1.38"),
    ]

    label_fmt = workbook.add_format({'border': 1, 'bg_color': '#F5F5F5'})
    value_fmt = workbook.add_format({
        'border': 1, 'bg_color': '#E8F5E9', 'bold': True, 'align': 'center'
    })

    for label, value in coc_labels:
        ws.write(row, 0, label, label_fmt)
        ws.write(row, 1, value, value_fmt)
        row += 1

    row += 1

    # ========== Section 2: Project Inputs Summary ==========
    ws.write(row, 0, "Project Rate Base Inputs", formats['section'])
    ws.merge_range(row, 0, row, 2, "Project Rate Base Inputs", formats['section'])
    row += 1

    input_fmt = workbook.add_format({
        'border': 1, 'bg_color': '#FFFDE7'
    })
    input_currency_fmt = workbook.add_format({
        'border': 1, 'bg_color': '#FFFDE7', 'num_format': '$#,##0'
    })

    rb_ref_row = row  # Save for formulas

    rb_inputs = [
        ("Gross Plant in Service ($)", "=Inputs!C12*Inputs!C13*1000*Inputs!C26", input_currency_fmt),
        ("Book Depreciation Life (years)", 20, input_fmt),
        ("MACRS Property Class", 7, input_fmt),
        ("ITC Rate", "=Inputs!C34+Inputs!C35", formats['input_percent']),
        ("Annual O&M ($)", "=Inputs!C27*Inputs!C12*1000", input_currency_fmt),
        ("Analysis Period (years)", "=Inputs!C15", input_fmt),
    ]

    for label, value, fmt in rb_inputs:
        ws.write(row, 0, label, label_fmt)
        if isinstance(value, str) and value.startswith("="):
            ws.write_formula(row, 1, value, fmt)
        else:
            ws.write(row, 1, value, fmt)
        row += 1

    row += 1

    # ========== Section 3: Revenue Requirement Schedule ==========
    rr_start_row = row
    ws.write(row, 0, "Annual Revenue Requirement Schedule", formats['section'])
    ws.merge_range(row, 0, row, 10, "Annual Revenue Requirement Schedule", formats['section'])
    row += 1

    # Headers
    rr_headers = [
        "Year", "Gross Plant", "Book Depr", "Accum Book Depr",
        "Net Plant", "ADIT", "Net Rate Base",
        "Return on RB", "Book Depr Exp", "Income Tax", "Property Tax",
        "O&M", "Revenue Req"
    ]

    header_fmt = workbook.add_format({
        'bold': True, 'font_color': 'white', 'bg_color': '#1565C0',
        'align': 'center', 'border': 1, 'text_wrap': True
    })

    for col, hdr in enumerate(rr_headers):
        ws.write(row, col, hdr, header_fmt)
    row += 1

    # Revenue requirement data rows (20 years with formulas)
    currency_fmt = workbook.add_format({'border': 1, 'num_format': '$#,##0'})
    year_fmt = workbook.add_format({'border': 1, 'align': 'center'})
    rr_highlight_fmt = workbook.add_format({
        'border': 1, 'num_format': '$#,##0', 'bold': True, 'bg_color': '#E3F2FD'
    })

    n_years = 20
    data_start_row = row

    for yr in range(1, n_years + 1):
        col = 0
        ws.write(row, col, yr, year_fmt); col += 1

        # Gross Plant (from input)
        gp_ref = f"$B${rb_ref_row + 1}"
        ws.write_formula(row, col, f"={gp_ref}", currency_fmt); col += 1

        # Book Depreciation = Gross Plant / Book Life (if year <= book life)
        book_life_ref = f"$B${rb_ref_row + 2}"
        ws.write_formula(row, col,
            f"=IF({yr}<={book_life_ref},{gp_ref}/{book_life_ref},0)",
            currency_fmt); col += 1

        # Accumulated Book Depreciation
        if yr == 1:
            ws.write_formula(row, col, f"=C{row+1}", currency_fmt)
        else:
            ws.write_formula(row, col, f"=D{row}+C{row+1}", currency_fmt)
        col += 1

        # Net Plant = Gross Plant - Accum Book Depr
        ws.write_formula(row, col, f"=B{row+1}-D{row+1}", currency_fmt); col += 1

        # ADIT (simplified: cumulative timing diff * tax rate)
        # For Excel, approximate with (tax_depr - book_depr) * composite_tax cumulative
        ws.write_formula(row, col,
            f"=MAX(0,D{row+1}*0.2798-D{row+1}*0.2798)", currency_fmt); col += 1

        # Net Rate Base = Gross Plant - Accum Book Depr - ADIT
        ws.write_formula(row, col, f"=MAX(0,B{row+1}-D{row+1}-F{row+1})", currency_fmt); col += 1

        # Return on Rate Base = Net RB * ROR (7.59%)
        ws.write_formula(row, col, f"=G{row+1}*0.0759", currency_fmt); col += 1

        # Book Depreciation Expense (same as book depr)
        ws.write_formula(row, col, f"=C{row+1}", currency_fmt); col += 1

        # Income Tax (equity return gross-up)
        ws.write_formula(row, col,
            f"=G{row+1}*0.52*0.1003*0.2798/(1-0.2798)",
            currency_fmt); col += 1

        # Property Tax (1% of net plant)
        ws.write_formula(row, col, f"=E{row+1}*0.01", currency_fmt); col += 1

        # O&M
        om_ref = f"$B${rb_ref_row + 5}"
        ws.write_formula(row, col, f"={om_ref}", currency_fmt); col += 1

        # Total Revenue Requirement
        ws.write_formula(row, col,
            f"=H{row+1}+I{row+1}+J{row+1}+K{row+1}+L{row+1}",
            rr_highlight_fmt)

        row += 1

    data_end_row = row - 1

    # Total row
    total_fmt = workbook.add_format({
        'bold': True, 'border': 2, 'bg_color': '#1565C0', 'font_color': 'white',
        'num_format': '$#,##0'
    })
    total_label_fmt = workbook.add_format({
        'bold': True, 'border': 2, 'bg_color': '#1565C0', 'font_color': 'white',
    })
    ws.write(row, 0, "TOTAL", total_label_fmt)
    for col in range(1, 13):
        if col in (0,):
            continue
        col_letter = chr(ord('A') + col)
        ws.write_formula(row, col,
            f"=SUM({col_letter}{data_start_row+1}:{col_letter}{data_end_row+1})",
            total_fmt)
    row += 2

    # ========== Section 4: Ratepayer Impact ==========
    impact_start_row = row
    ws.write(row, 0, "Ratepayer Impact Analysis", formats['section'])
    ws.merge_range(row, 0, row, 5, "Ratepayer Impact Analysis", formats['section'])
    row += 1

    impact_headers = [
        "Year", "Revenue Requirement", "Avoided Cost (ACC)",
        "Net Ratepayer Impact", "Cumulative Savings"
    ]
    for col, hdr in enumerate(impact_headers):
        ws.write(row, col, hdr, header_fmt)
    row += 1

    impact_data_start = row
    green_fmt = workbook.add_format({
        'border': 1, 'num_format': '$#,##0', 'font_color': '#2E7D32'
    })
    red_fmt = workbook.add_format({
        'border': 1, 'num_format': '$#,##0', 'font_color': '#C62828'
    })

    # ACC Generation Capacity trajectory (hardcoded from SCE library)
    acc_gen_cap = [89.48, 82.00, 75.00, 68.50, 62.50, 57.00, 52.00, 48.00,
                   44.50, 41.50, 39.50, 39.00, 39.00, 39.00, 39.00, 39.00,
                   39.00, 39.00, 39.00, 39.00]

    for yr in range(1, n_years + 1):
        ws.write(row, 0, yr, year_fmt)

        # Revenue Requirement (link to RR schedule)
        rr_row = data_start_row + yr
        ws.write_formula(row, 1, f"=M{rr_row}", currency_fmt)

        # Avoided Cost (ACC gen cap * capacity_kw + other benefits)
        # Simplified: use ACC trajectory * capacity
        acc_val = acc_gen_cap[yr - 1] if yr <= len(acc_gen_cap) else acc_gen_cap[-1]
        ws.write_formula(row, 2,
            f"={acc_val}*Inputs!C12*1000",
            currency_fmt)

        # Net Impact = Avoided Cost - Revenue Requirement (positive = savings)
        ws.write_formula(row, 3, f"=C{row+1}-B{row+1}", currency_fmt)

        # Cumulative Savings
        if yr == 1:
            ws.write_formula(row, 4, f"=D{row+1}", currency_fmt)
        else:
            ws.write_formula(row, 4, f"=E{row}+D{row+1}", currency_fmt)

        row += 1

    impact_data_end = row - 1

    # Total impact row
    ws.write(row, 0, "TOTAL", total_label_fmt)
    ws.write_formula(row, 1, f"=SUM(B{impact_data_start+1}:B{impact_data_end+1})", total_fmt)
    ws.write_formula(row, 2, f"=SUM(C{impact_data_start+1}:C{impact_data_end+1})", total_fmt)
    ws.write_formula(row, 3, f"=SUM(D{impact_data_start+1}:D{impact_data_end+1})", total_fmt)
    ws.write(row, 4, "", total_label_fmt)
    row += 2

    # ========== Section 5: Cumulative Savings Chart ==========
    chart = workbook.add_chart({'type': 'column'})
    chart.set_title({'name': 'Annual Ratepayer Impact (Avoided Cost - Revenue Requirement)'})
    chart.set_x_axis({'name': 'Year'})
    chart.set_y_axis({'name': 'Net Impact ($)', 'num_format': '$#,##0'})

    # Net Impact bars
    chart.add_series({
        'name': 'Net Ratepayer Impact',
        'categories': f"='UOG_Analysis'!$A${impact_data_start+1}:$A${impact_data_end+1}",
        'values': f"='UOG_Analysis'!$D${impact_data_start+1}:$D${impact_data_end+1}",
        'fill': {'color': '#4CAF50'},
        'border': {'color': '#2E7D32'},
    })

    chart.set_size({'width': 720, 'height': 400})
    chart.set_legend({'position': 'bottom'})
    ws.insert_chart(f'A{row+1}', chart)
    row += 22  # Space for chart

    # Cumulative savings line chart
    line_chart = workbook.add_chart({'type': 'line'})
    line_chart.set_title({'name': 'Cumulative Ratepayer Savings'})
    line_chart.set_x_axis({'name': 'Year'})
    line_chart.set_y_axis({'name': 'Cumulative Savings ($)', 'num_format': '$#,##0'})

    line_chart.add_series({
        'name': 'Cumulative Savings',
        'categories': f"='UOG_Analysis'!$A${impact_data_start+1}:$A${impact_data_end+1}",
        'values': f"='UOG_Analysis'!$E${impact_data_start+1}:$E${impact_data_end+1}",
        'line': {'color': '#1565C0', 'width': 2.5},
        'marker': {'type': 'circle', 'size': 5, 'fill': {'color': '#1565C0'}},
    })

    line_chart.set_size({'width': 720, 'height': 400})
    line_chart.set_legend({'position': 'bottom'})
    ws.insert_chart(f'A{row+1}', line_chart)
    row += 22

    # ========== Section 6: Wires vs NWA Summary ==========
    ws.write(row, 0, "Wires vs Non-Wires Alternative (NWA) Comparison", formats['section'])
    ws.merge_range(row, 0, row, 3, "Wires vs Non-Wires Alternative (NWA) Comparison", formats['section'])
    row += 1

    nwa_items = [
        ("Traditional Wires Cost ($/kW)", "$500"),
        ("Wires Book Life (years)", "40"),
        ("Wires Lead Time (years)", "5"),
        ("NWA Deferral Period (years)", "5"),
        ("Wires RECC ($/yr)", "Link to calculation"),
        ("NWA (BESS) RECC ($/yr)", "Link to calculation"),
        ("Annual Savings (NWA vs Wires)", "= Wires RECC - NWA RECC"),
        ("NWA is Economic?", "=IF(NWA_RECC<Wires_RECC,\"YES\",\"NO\")"),
        ("Deferral Value ($)", "= Wires Cost * (1 - 1/(1+ROR)^N)"),
    ]

    for label, value in nwa_items:
        ws.write(row, 0, label, label_fmt)
        ws.write(row, 1, value, value_fmt)
        row += 1

    row += 1

    # ========== Section 7: Slice-of-Day Feasibility ==========
    ws.write(row, 0, "Slice-of-Day (SOD) RA Feasibility", formats['section'])
    ws.merge_range(row, 0, row, 3, "Slice-of-Day (SOD) RA Feasibility", formats['section'])
    row += 1

    sod_items = [
        ("Battery Duration (hours)", "=Inputs!C13"),
        ("Qualifying Hours (met)", "Calculated"),
        ("Required Hours (minimum)", "4"),
        ("SOD Feasible?", "Calculated"),
        ("Deration Factor", "Calculated"),
        ("Effective RA Capacity (MW)", "Calculated"),
    ]

    for label, value in sod_items:
        ws.write(row, 0, label, label_fmt)
        if isinstance(value, str) and value.startswith("="):
            ws.write_formula(row, 1, value, value_fmt)
        else:
            ws.write(row, 1, value, value_fmt)
        row += 1

    row += 1

    # SOD hourly profile
    ws.write(row, 0, "Hour (HE)", header_fmt)
    ws.write(row, 1, "Load Factor", header_fmt)
    ws.write(row, 2, "Dispatch (MW)", header_fmt)
    ws.write(row, 3, "SOC (MWh)", header_fmt)
    row += 1

    sod_load = [
        0.00, 0.00, 0.00, 0.00, 0.00, 0.00,
        0.00, 0.00, 0.10, 0.20, 0.40, 0.60,
        0.80, 0.90, 1.00, 1.00, 1.00, 0.95,
        0.85, 0.70, 0.50, 0.30, 0.10, 0.00,
    ]

    pct_fmt = workbook.add_format({'border': 1, 'num_format': '0%', 'align': 'center'})

    for hr in range(24):
        ws.write(row, 0, f"HE {hr+1}", year_fmt)
        ws.write(row, 1, sod_load[hr], pct_fmt)
        ws.write(row, 2, "Calc", value_fmt)
        ws.write(row, 3, "Calc", value_fmt)
        row += 1

    row += 1

    # ========== Footer ==========
    footer_fmt = workbook.add_format({
        'italic': True, 'font_color': '#999999', 'font_size': 9
    })
    ws.write(row, 0,
        "Sources: CPUC D.25-12-003, E3 2024 Avoided Cost Calculator, "
        "CPUC D.23-06-029 (SOD Framework), SCE Distribution Resource Plan 2024",
        footer_fmt)


if __name__ == "__main__":
    output = sys.argv[1] if len(sys.argv) > 1 else "BESS_Analyzer"

    # Check for VBA project
    vba_path = Path(__file__).parent / 'resources' / 'vbaProject.bin'

    if not vba_path.exists():
        print("Note: vbaProject.bin not found. Creating .xlsx without macros.")
        print("See VBA_Code sheet for instructions to add macros manually.\n")

    create_workbook(output)
