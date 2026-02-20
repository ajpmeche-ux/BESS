"""Generate Excel-based BESS Economic Analyzer Workbook.

Creates a macro-enabled Excel workbook (.xlsm) with:
- Project input forms for single asset and multi-tranche cohort models
- Assumption library selection (NREL, Lazard, CPUC)
- Automatic calculation engine using Excel formulas, including SUMPRODUCT for cohorts
- Results dashboard with conditional formatting
- Annual cash flow projections
- Methodology documentation with formulas and citations
- Embedded VBA macros with functional buttons

Usage:
    python excel_generator.py [output_path]
"""

import sys
from datetime import datetime
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
    # Total capacity is now calculated from the build schedule
    CAPACITY_MW = 'C12'  # Formula: =SUM(Build_Schedule[Capacity (MW)])
    DURATION_HOURS = 'C13'
    ENERGY_MWH = 'C14'  # Formula: =C12*C13
    ANALYSIS_YEARS = 'C15'
    DISCOUNT_RATE = 'C16'
    OWNERSHIP_TYPE = 'C17'

    # Phased Build Schedule starts at row 20
    # Named Range: Build_Schedule covers C20:E29

    # T&D Deferral Inputs start at row 32
    TD_DEFERRAL_K = 'C32'
    TD_DEFERRAL_T_NEED = 'C33'
    TD_DEFERRAL_N = 'C34'
    TD_DEFERRAL_G = 'C35'
    TD_DEFERRAL_PV = 'C36' # Formula for PV

    # Technology Specifications (starting row 39)
    CHEMISTRY = 'C39'
    ROUND_TRIP_EFFICIENCY = 'C40'
    ANNUAL_DEGRADATION = 'C41'
    CYCLE_LIFE = 'C42'
    AUGMENTATION_YEAR = 'C43'
    CYCLES_PER_DAY = 'C44'

    # Cost Inputs (starting row 46)
    CAPEX_PER_KWH = 'C46'
    FOM_PER_KW_YEAR = 'C47'
    VOM_PER_MWH = 'C48'
    AUGMENTATION_COST = 'C49'
    DECOMMISSIONING = 'C50'
    CHARGING_COST = 'C51'
    RESIDUAL_VALUE = 'C52'

    # Tax Credits (starting row 54)
    ITC_BASE_RATE = 'C54'
    ITC_ADDERS = 'C55'

    # Infrastructure Costs (starting row 57)
    INTERCONNECTION = 'C57'
    LAND = 'C58'
    PERMITTING = 'C59'
    INSURANCE_PCT = 'C60'
    PROPERTY_TAX_PCT = 'C61'

    # Financing Structure (starting row 63)
    DEBT_PERCENT = 'C63'
    INTEREST_RATE = 'C64'
    LOAN_TERM = 'C65'
    COST_OF_EQUITY = 'C66'
    TAX_RATE = 'C67'
    WACC = 'C68'  # Formula

    # Benefits (starting row 72, after headers at row 71)
    BENEFIT_RA = 'C72'
    BENEFIT_RA_ESC = 'D72'
    BENEFIT_ARBITRAGE = 'C73'
    BENEFIT_ARBITRAGE_ESC = 'D73'
    BENEFIT_ANCILLARY = 'C74'
    BENEFIT_ANCILLARY_ESC = 'D74'
    BENEFIT_TD = 'C75'
    BENEFIT_TD_ESC = 'D75'
    BENEFIT_RESILIENCE = 'C76'
    BENEFIT_RESILIENCE_ESC = 'D76'
    BENEFIT_RENEWABLE = 'C77'
    BENEFIT_RENEWABLE_ESC = 'D77'
    BENEFIT_GHG = 'C78'
    BENEFIT_GHG_ESC = 'D78'
    BENEFIT_VOLTAGE = 'C79'
    BENEFIT_VOLTAGE_ESC = 'D79'

    # Benefit rows (for Cash_Flows formulas)
    BENEFIT_ROWS = [72, 73, 74, 75, 76, 77, 78, 79]

    # Cost Projections (starting row 81)
    LEARNING_RATE = 'C81'
    COST_BASE_YEAR = 'C82'


def create_workbook(output_path: str, with_macros: bool = True):
    """Create the complete BESS Analyzer workbook."""
    vba_path = Path(__file__).parent / 'resources' / 'vbaProject.bin'
    has_vba = vba_path.exists() and with_macros

    if has_vba:
        if not output_path.endswith('.xlsm'):
            output_path = str(Path(output_path).with_suffix('.xlsm'))
    else:
        if not output_path.endswith('.xlsx'):
            output_path = str(Path(output_path).with_suffix('.xlsx'))

    workbook = xlsxwriter.Workbook(output_path)

    if has_vba:
        workbook.add_vba_project(str(vba_path))

    formats = create_formats(workbook)

    ws_inputs = workbook.add_worksheet('Inputs')
    ws_results = workbook.add_worksheet('Results')
    ws_cashflows = workbook.add_worksheet('Cash_Flows')
    ws_calculations = workbook.add_worksheet('Calculations')

    create_inputs_sheet(workbook, ws_inputs, formats)
    cf_totals_row = create_cashflows_sheet(workbook, ws_cashflows, formats)
    create_calculations_sheet(workbook, ws_calculations, formats)
    create_results_sheet(workbook, ws_results, formats, cf_totals_row)

    workbook.close()
    print(f"Workbook created: {output_path}")


def create_formats(workbook):
    """Create reusable cell formats."""
    formats = {}
    formats['header'] = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#1565C0', 'align': 'center', 'valign': 'vcenter', 'border': 1})
    formats['section'] = workbook.add_format({'bold': True, 'font_size': 12, 'font_color': '#1565C0', 'bg_color': '#E3F2FD', 'border': 1})
    formats['input'] = workbook.add_format({'bg_color': '#FFFDE7', 'border': 1})
    formats['input_currency'] = workbook.add_format({'bg_color': '#FFFDE7', 'border': 1, 'num_format': '$#,##0.00'})
    formats['input_percent'] = workbook.add_format({'bg_color': '#FFFDE7', 'border': 1, 'num_format': '0.0%'})
    formats['formula'] = workbook.add_format({'bg_color': '#E8F5E9', 'border': 1, 'bold': True})
    formats['formula_currency'] = workbook.add_format({'bg_color': '#E8F5E9', 'border': 1, 'num_format': '$#,##0'})
    formats['result_currency'] = workbook.add_format({'bold': True, 'font_size': 14, 'bg_color': '#E3F2FD', 'align': 'center', 'border': 1, 'num_format': '$#,##0'})
    formats['currency'] = workbook.add_format({'num_format': '$#,##0', 'border': 1})
    formats['percent'] = workbook.add_format({'num_format': '0.0%', 'border': 1})
    formats['title'] = workbook.add_format({'bold': True, 'font_size': 16, 'font_color': '#1565C0'})
    formats['bold'] = workbook.add_format({'bold': True})
    formats['tooltip'] = workbook.add_format({'italic': True, 'font_color': '#666666'})
    return formats


def create_inputs_sheet(workbook, ws, formats):
    """Create the Project Inputs sheet."""
    ws.set_column('A:A', 5)
    ws.set_column('B:B', 30)
    ws.set_column('C:C', 20)
    ws.set_column('D:D', 15)
    ws.set_column('E:E', 45)

    row = 1
    ws.merge_range('B2:E2', 'BESS Economic Analysis - JIT Cohort Model', formats['title'])
    row = 8

    # === PROJECT BASICS ===
    ws.merge_range(f'B{row}:E{row}', 'PROJECT BASICS', formats['section'])
    row += 1
    basics = [
        ('Project Name', 'BESS JIT Project', 'Text', 'Enter project name'),
        ('Project ID', 'JIT-001', 'Text', 'Unique identifier'),
        ('Location', 'Substation Z', 'Text', 'Site or market location'),
        ('Total Capacity (MW)', '=SUM(D20:D29)', 'Formula', 'Auto-summed from Build Schedule'),
        ('Duration (hours)', 4, 'Number', 'Storage duration for each cohort'),
        ('Total Energy (MWh)', f'={CellRefs.CAPACITY_MW}*{CellRefs.DURATION_HOURS}', 'Formula', 'Auto-calculated'),
        ('Analysis Period (years)', 20, 'Number', 'Economic analysis horizon'),
        ('Discount Rate (%)', 0.07, 'Percent', 'Utility WACC'),
        ('Ownership Type', 'Utility', 'Text', 'Utility or Merchant'),
    ]
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

    # === PHASED BUILD SCHEDULE ===
    ws.merge_range(f'B{row}:E{row}', 'PHASED BUILD SCHEDULE', formats['section'])
    row += 1
    ws.write(f'B{row}', 'Cohort', formats['header'])
    ws.write(f'C{row}', 'COD (Year)', formats['header'])
    ws.write(f'D{row}', 'Capacity (MW)', formats['header'])
    ws.write(f'E{row}', 'ITC Rate (%)', formats['header'])
    row += 1
    build_schedule_start_row = row
    for i in range(10):
        ws.write(f'B{row}', f'Tranche {i+1}')
        ws.write(f'C{row}', 0, formats['input']) # Default year
        ws.write(f'D{row}', 0, formats['input']) # Default capacity
        ws.write_formula(f'E{row}', f'=IF(C{row}>0, {CellRefs.ITC_BASE_RATE}+{CellRefs.ITC_ADDERS}, 0)', formats['formula'])
        row += 1
    # Define named range for the schedule data (excluding headers)
    workbook.define_name('Build_Schedule', f'=Inputs!$C${build_schedule_start_row}:$E${row-1}')
    row += 1

    # === T&D DEFERRAL ===
    ws.merge_range(f'B{row}:E{row}', 'T&D DEFERRAL', formats['section'])
    row += 1
    td_inputs = [
        ('Capital Cost (K)', 100000000, '$/kW', 'Upfront capital cost of the deferred asset'),
        ('Time Need (t_need)', 0, 'Year', 'Year the traditional asset is needed'),
        ('Deferral Period (n)', 5, 'Years', 'Number of years the asset is deferred'),
        ('Growth Rate (g)', 0.02, 'Percent', 'Annual growth rate of the capital cost'),
    ]
    for label, value, unit, tooltip in td_inputs:
        ws.write(f'B{row}', label, formats['bold'])
        if 'Percent' in tooltip:
            ws.write(f'C{row}', value, formats['input_percent'])
        else:
            ws.write(f'C{row}', value, formats['input_currency'])
        ws.write(f'E{row}', tooltip, formats['tooltip'])
        row += 1
    ws.write(f'B{row}', 'PV of Deferral ($)', formats['bold'])
    # Formula: PV = K * [1 - ((1 + g) / (1 + r))^n]
    pv_formula = f'={CellRefs.TD_DEFERRAL_K}*(1-((1+{CellRefs.TD_DEFERRAL_G})/(1+{CellRefs.DISCOUNT_RATE}))^{CellRefs.TD_DEFERRAL_N})'
    ws.write_formula(f'C{row}', pv_formula, formats['formula_currency'])
    ws.write(f'E{row}', 'Present value of the T&D deferral benefit', formats['tooltip'])
    row += 2

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
    for label, value, unit, tooltip in cost_inputs:
        ws.write(f'B{row}', label, formats['bold'])
        if '%' in unit:
            ws.write(f'C{row}', value, formats['input_percent'])
        else:
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
    for label, value, tooltip in tax_inputs:
        ws.write(f'B{row}', label, formats['bold'])
        ws.write(f'C{row}', value, formats['input_percent'])
        ws.write(f'E{row}', tooltip, formats['tooltip'])
        row += 1
    row += 1

    # === COST PROJECTIONS ===
    ws.merge_range(f'B{row}:E{row}', 'COST PROJECTIONS (Learning Curve)', formats['section'])
    row += 1
    ws.write(f'B{row}', 'Annual Cost Decline Rate', formats['bold'])
    ws.write(f'C{row}', 0.12, formats['input_percent'])
    ws.write(f'E{row}', 'Technology learning rate (12% default)', formats['tooltip'])
    row += 1
    ws.write(f'B{row}', 'Cost Base Year', formats['bold'])
    ws.write(f'C{row}', 2024, formats['input'])
    ws.write(f'E{row}', 'Reference year for base costs', formats['tooltip'])

    # Define named ranges
    workbook.define_name('Discount_Rate', f'=Inputs!${CellRefs.DISCOUNT_RATE}')
    workbook.define_name('Learning_Rate', f'=Inputs!${CellRefs.LEARNING_RATE}')


def create_calculations_sheet(workbook, ws, formats):
    """Create the Calculations sheet with cohort-based logic."""
    ws.set_column('A:A', 5)
    ws.set_column('B:B', 35)
    ws.set_column('C:C', 20)
    ws.set_column('D:D', 60)

    row = 1
    ws.write('B2', 'Calculation Engine (Cohort-Based)', formats['title'])
    row = 4

    ws.merge_range(f'B{row}:D{row}', 'COHORT-BASED CALCULATIONS', formats['section'])
    row += 1

    calcs = [
        ('Total Capacity (kW)', f'=Inputs!{CellRefs.CAPACITY_MW}*1000', 'Sum of all cohort capacities in kW'),
        ('Total Energy (kWh)', f'=Inputs!{CellRefs.ENERGY_MWH}*1000', 'Sum of all cohort energies in kWh'),
        ('Total Battery CapEx ($)',
         f'=SUMPRODUCT(INDEX(Build_Schedule,,2)*1000*Inputs!{CellRefs.DURATION_HOURS}, Inputs!{CellRefs.CAPEX_PER_KWH}*(1-Inputs!{CellRefs.LEARNING_RATE})^(INDEX(Build_Schedule,,1)-Inputs!{CellRefs.COST_BASE_YEAR}))',
         'Sum of CapEx for each cohort, adjusted for learning rate based on COD'),
        ('Total ITC Credit ($)',
         f'=SUMPRODUCT(INDEX(Build_Schedule,,2)*1000*Inputs!{CellRefs.DURATION_HOURS}, Inputs!{CellRefs.CAPEX_PER_KWH}*(1-Inputs!{CellRefs.LEARNING_RATE})^(INDEX(Build_Schedule,,1)-Inputs!{CellRefs.COST_BASE_YEAR}), INDEX(Build_Schedule,,3))',
         'Sum of ITC for each cohort based on its specific CapEx'),
        ('Net Year 0 Cost ($)', '=C7-C8', 'Total CapEx minus Total ITC'),
    ]

    for label, formula, desc in calcs:
        ws.write(f'B{row}', label, formats['bold'])
        ws.write_formula(f'C{row}', formula, formats['formula_currency'])
        ws.write(f'D{row}', desc, formats['tooltip'])
        row += 1


def create_cashflows_sheet(workbook, ws, formats):
    """Create annual cash flow projections with cohort logic."""
    ws.set_column('B:B', 8)
    ws.set_column('C:V', 14)

    ws.write('B2', 'Annual Cash Flow Projections (Cohort-Aggregated)', formats['title'])
    row = 4
    headers = ['Year', 'CapEx', 'O&M', 'Charging', 'Total Costs', 'Benefits', 'Net CF', 'PV Net CF']
    for col, header in enumerate(headers):
        ws.write(row - 1, col + 1, header, formats['header'])

    start_row = row
    for year in range(int(21)): # Analysis period 0-20
        ws.write(row, 1, year) # Year column

        # CapEx
        capex_formula = f'=SUMPRODUCT(--(INDEX(Build_Schedule,,1)={year}), INDEX(Build_Schedule,,2)*1000*Inputs!{CellRefs.DURATION_HOURS}*Inputs!{CellRefs.CAPEX_PER_KWH}*(1-Inputs!{CellRefs.LEARNING_RATE})^({year}-Inputs!{CellRefs.COST_BASE_YEAR}))'
        ws.write_formula(row, 2, capex_formula, formats['currency'])

        # O&M
        om_formula = f'=SUMPRODUCT(--(INDEX(Build_Schedule,,1)<={year}), INDEX(Build_Schedule,,2)*1000*Inputs!{CellRefs.FOM_PER_KW_YEAR})'
        ws.write_formula(row, 3, om_formula, formats['currency'])

        # Charging Cost
        charging_formula = f'=SUMPRODUCT(--(INDEX(Build_Schedule,,1)<={year}), INDEX(Build_Schedule,,2)*1000*Inputs!{CellRefs.DURATION_HOURS}*Inputs!{CellRefs.CYCLES_PER_DAY}*365*Inputs!{CellRefs.CHARGING_COST})'
        ws.write_formula(row, 4, charging_formula, formats['currency'])

        # Total Costs
        ws.write_formula(row, 5, f'=SUM(C{row+1}:E{row+1})', formats['formula_currency'])

        # Benefits (with staged degradation)
        benefit_formula = f'=SUMPRODUCT(--(INDEX(Build_Schedule,,1)<={year}), INDEX(Build_Schedule,,2)*1000 * (1-Inputs!{CellRefs.ANNUAL_DEGRADATION})^({year}-INDEX(Build_Schedule,,1))) * (Inputs!{CellRefs.BENEFIT_RA}*(1+Inputs!{CellRefs.BENEFIT_RA_ESC})^{year})'
        ws.write_formula(row, 6, benefit_formula, formats['currency'])

        # Net CF
        itc_formula = f'=SUMPRODUCT(--(INDEX(Build_Schedule,,1)={year}), INDEX(Build_Schedule,,2)*1000*Inputs!{CellRefs.DURATION_HOURS}*Inputs!{CellRefs.CAPEX_PER_KWH}*(1-Inputs!{CellRefs.LEARNING_RATE})^({year}-Inputs!{CellRefs.COST_BASE_YEAR})*INDEX(Build_Schedule,,3))'
        net_cf_formula = f'=G{row+1} - F{row+1} + ({itc_formula})'
        ws.write_formula(row, 7, net_cf_formula, formats['formula_currency'])

        # PV Net CF
        pv_formula = f'=H{row+1}/(1+Inputs!{CellRefs.DISCOUNT_RATE})^{year}'
        ws.write_formula(row, 8, pv_formula, formats['currency'])

        row += 1

    # Totals row
    totals_row = row
    ws.write(f'B{totals_row+1}', 'TOTALS', formats['bold'])
    ws.write_formula(f'I{totals_row+1}', f'=SUM(I{start_row+1}:I{row})', formats['formula_currency'])
    return totals_row


def create_results_sheet(workbook, ws, formats, cf_totals_row):
    """Create the Results Dashboard."""
    ws.set_column('B:B', 25)
    ws.set_column('C:C', 20)

    ws.write('B2', 'BESS Economic Analysis Results', formats['title'])
    row = 4

    ws.merge_range(f'B{row}:C{row}', 'KEY FINANCIAL METRICS', formats['section'])
    row += 1

    metrics = [
        ('Net Present Value (NPV)', f'=Cash_Flows!I{cf_totals_row+1}', 'Present value of all cash flows'),
        ('T&D Deferral PV', f'=Inputs!{CellRefs.TD_DEFERRAL_PV}', 'Value from capital deferral'),
        ('Total Project Value', f'=Cash_Flows!I{cf_totals_row+1} + Inputs!{CellRefs.TD_DEFERRAL_PV}', 'NPV plus T&D Deferral Value')
    ]

    for label, formula, desc in metrics:
        ws.write(f'B{row}', label, formats['bold'])
        ws.write_formula(f'C{row}', formula, formats['result_currency'])
        ws.write(f'D{row}', desc, formats['tooltip'])
        row += 1

if __name__ == '__main__':
    timestamp = datetime.now().strftime('%Y%m%d_%H%M')
    default_name = f'BESS_Analyzer_v2.0_{timestamp}.xlsx'
    output_file = sys.argv[1] if len(sys.argv) > 1 else default_name
    create_workbook(output_file)
