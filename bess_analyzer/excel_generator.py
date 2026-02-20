"""Generate Excel-based BESS Economic Analyzer Workbook.

Creates a macro-enabled Excel workbook (.xlsm) with:
- Inputs: project basics, build schedule, T&D deferral, technology, costs,
  tax credits, infrastructure, financing, benefits, cost projections
- Calculations: cohort-based SUMPRODUCT engine
- Cash_Flows: annual cash flows with PV columns for BCR/IRR/LCOS/Payback
- Results: full dashboard (NPV, BCR, IRR, LCOS, Payback, Breakeven CapEx)
- Sensitivity: tornado chart table with ±20% parameter sweeps
- Library_Data: NREL / Lazard / CPUC assumption comparison
- UOS_Analysis: utility revenue requirement, avoided costs, wires vs NWA
- Methodology: formulas, methodology, and citations
- VBA_Instructions: macro usage guide

Usage:
    python excel_generator.py [output_path]
"""

import sys
from datetime import datetime
from pathlib import Path

import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell


# =============================================================================
# CELL REFERENCE REGISTRY  (all references are Excel 1-based, column A=col 1)
# The layout below is the authoritative row map for the Inputs sheet.
# =============================================================================

class CellRefs:
    """Central registry for all Inputs-sheet cell addresses."""

    # Row 6 — Library selection dropdown
    SELECTED_LIBRARY = 'C6'

    # Rows 9-17 — Project Basics
    PROJECT_NAME    = 'C9'
    PROJECT_ID      = 'C10'
    LOCATION        = 'C11'
    CAPACITY_MW     = 'C12'   # =SUM(D20:D29)
    DURATION_HOURS  = 'C13'
    ENERGY_MWH      = 'C14'   # =C12*C13
    ANALYSIS_YEARS  = 'C15'
    DISCOUNT_RATE   = 'C16'
    OWNERSHIP_TYPE  = 'C17'

    # Row 18 — PHASED BUILD SCHEDULE section header (no blank row after basics)
    # Row 19 — column headers: Cohort | COD Year | Capacity MW | ITC Rate
    # Rows 20-29 — 10 tranche data rows  →  Named range: Build_Schedule = C20:E29

    # Rows 32-36 — T&D Deferral  (section header at row 31, blank at row 30)
    TD_DEFERRAL_K      = 'C32'
    TD_DEFERRAL_T_NEED = 'C33'
    TD_DEFERRAL_N      = 'C34'
    TD_DEFERRAL_G      = 'C35'
    TD_DEFERRAL_PV     = 'C36'   # formula

    # Rows 39-44 — Technology Specs  (section header at row 38, blank at row 37)
    CHEMISTRY              = 'C39'
    ROUND_TRIP_EFFICIENCY  = 'C40'
    ANNUAL_DEGRADATION     = 'C41'
    CYCLE_LIFE             = 'C42'
    AUGMENTATION_YEAR      = 'C43'
    CYCLES_PER_DAY         = 'C44'

    # Rows 46-52 — Cost Inputs  (section header at row 45, no blank)
    CAPEX_PER_KWH    = 'C46'
    FOM_PER_KW_YEAR  = 'C47'
    VOM_PER_MWH      = 'C48'
    AUGMENTATION_COST= 'C49'
    DECOMMISSIONING  = 'C50'
    CHARGING_COST    = 'C51'
    RESIDUAL_VALUE   = 'C52'

    # Rows 54-55 — Tax Credits  (section header at row 53, no blank)
    ITC_BASE_RATE = 'C54'
    ITC_ADDERS    = 'C55'

    # Rows 57-61 — Infrastructure Costs  (section header at row 56, no blank)
    INTERCONNECTION  = 'C57'
    LAND             = 'C58'
    PERMITTING       = 'C59'
    INSURANCE_PCT    = 'C60'
    PROPERTY_TAX_PCT = 'C61'

    # Rows 63-68 — Financing Structure  (section header at row 62, no blank)
    DEBT_PERCENT   = 'C63'
    INTEREST_RATE  = 'C64'
    LOAN_TERM      = 'C65'
    COST_OF_EQUITY = 'C66'
    TAX_RATE       = 'C67'
    WACC           = 'C68'   # formula

    # Row 69 — blank
    # Row 70 — BENEFIT STREAMS section header
    # Row 71 — column headers: Name | $/kW-yr | Escalation | Category | Source
    # Rows 72-79 — eight benefit streams
    BENEFIT_RA          = 'C72';  BENEFIT_RA_ESC          = 'D72'
    BENEFIT_ARBITRAGE   = 'C73';  BENEFIT_ARBITRAGE_ESC   = 'D73'
    BENEFIT_ANCILLARY   = 'C74';  BENEFIT_ANCILLARY_ESC   = 'D74'
    BENEFIT_TD          = 'C75';  BENEFIT_TD_ESC          = 'D75'
    BENEFIT_RESILIENCE  = 'C76';  BENEFIT_RESILIENCE_ESC  = 'D76'
    BENEFIT_RENEWABLE   = 'C77';  BENEFIT_RENEWABLE_ESC   = 'D77'
    BENEFIT_GHG         = 'C78';  BENEFIT_GHG_ESC         = 'D78'
    BENEFIT_VOLTAGE     = 'C79';  BENEFIT_VOLTAGE_ESC     = 'D79'

    # Rows 81-82 — Cost Projections  (section header at row 80)
    LEARNING_RATE  = 'C81'
    COST_BASE_YEAR = 'C82'

    # Convenience list of benefit (value, escalation) cell pairs for Cash_Flows
    BENEFITS = [
        ('C72', 'D72'), ('C73', 'D73'), ('C74', 'D74'), ('C75', 'D75'),
        ('C76', 'D76'), ('C77', 'D77'), ('C78', 'D78'), ('C79', 'D79'),
    ]


# =============================================================================
# WORKBOOK ENTRY POINT
# =============================================================================

def create_workbook(output_path: str, with_macros: bool = True) -> None:
    """Create the complete BESS Analyzer workbook with all sheets."""
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

    fmt = _create_formats(workbook)

    # Create all sheets in order
    ws_inputs      = workbook.add_worksheet('Inputs')
    ws_calcs       = workbook.add_worksheet('Calculations')
    ws_cf          = workbook.add_worksheet('Cash_Flows')
    ws_results     = workbook.add_worksheet('Results')
    ws_sensitivity = workbook.add_worksheet('Sensitivity')
    ws_uos         = workbook.add_worksheet('UOS_Analysis')
    ws_library     = workbook.add_worksheet('Library_Data')
    ws_method      = workbook.add_worksheet('Methodology')
    ws_vba         = workbook.add_worksheet('VBA_Instructions')

    _create_inputs_sheet(workbook, ws_inputs, fmt)
    _create_calculations_sheet(ws_calcs, fmt)
    cf_data_rows = _create_cashflows_sheet(ws_cf, fmt)
    _create_results_sheet(ws_results, fmt, cf_data_rows)
    _create_sensitivity_sheet(ws_sensitivity, fmt, cf_data_rows)
    _create_uos_sheet(ws_uos, fmt)
    _create_library_data_sheet(ws_library, fmt)
    _create_methodology_sheet(ws_method, fmt)
    _create_vba_instructions_sheet(ws_vba, fmt)

    workbook.close()
    print(f"Workbook created: {output_path}")


# =============================================================================
# FORMATS
# =============================================================================

def _create_formats(wb) -> dict:
    f = {}
    blue  = '#1565C0'
    lblue = '#E3F2FD'
    yel   = '#FFFDE7'
    grn   = '#E8F5E9'

    f['title']    = wb.add_format({'bold': True, 'font_size': 16, 'font_color': blue})
    f['subtitle'] = wb.add_format({'italic': True, 'font_color': '#555555', 'font_size': 10})
    f['section']  = wb.add_format({'bold': True, 'font_size': 11, 'font_color': blue,
                                    'bg_color': lblue, 'border': 1, 'valign': 'vcenter'})
    f['header']   = wb.add_format({'bold': True, 'font_color': 'white', 'bg_color': blue,
                                    'align': 'center', 'border': 1, 'valign': 'vcenter'})
    f['bold']     = wb.add_format({'bold': True})
    f['tooltip']  = wb.add_format({'italic': True, 'font_color': '#666666', 'font_size': 9})
    f['input']    = wb.add_format({'bg_color': yel, 'border': 1})
    f['input_pct']= wb.add_format({'bg_color': yel, 'border': 1, 'num_format': '0.00%'})
    f['input_cur']= wb.add_format({'bg_color': yel, 'border': 1, 'num_format': '$#,##0.00'})
    f['input_int']= wb.add_format({'bg_color': yel, 'border': 1, 'num_format': '0'})
    f['formula']  = wb.add_format({'bg_color': grn, 'border': 1, 'bold': True})
    f['fml_cur']  = wb.add_format({'bg_color': grn, 'border': 1, 'num_format': '$#,##0', 'bold': True})
    f['fml_pct']  = wb.add_format({'bg_color': grn, 'border': 1, 'num_format': '0.00%', 'bold': True})
    f['fml_num']  = wb.add_format({'bg_color': grn, 'border': 1, 'num_format': '0.00', 'bold': True})
    f['currency'] = wb.add_format({'num_format': '$#,##0', 'border': 1})
    f['cur_red']  = wb.add_format({'num_format': '$#,##0', 'border': 1, 'font_color': '#C62828'})
    f['percent']  = wb.add_format({'num_format': '0.00%', 'border': 1})
    f['number']   = wb.add_format({'num_format': '#,##0.0', 'border': 1})
    f['result_big']= wb.add_format({'bold': True, 'font_size': 13, 'bg_color': lblue,
                                     'border': 2, 'num_format': '$#,##0', 'align': 'center'})
    f['result_pct']= wb.add_format({'bold': True, 'font_size': 13, 'bg_color': lblue,
                                     'border': 2, 'num_format': '0.00%', 'align': 'center'})
    f['result_num']= wb.add_format({'bold': True, 'font_size': 13, 'bg_color': lblue,
                                     'border': 2, 'num_format': '#,##0.00', 'align': 'center'})
    f['pass_fmt'] = wb.add_format({'bold': True, 'font_color': '#1B5E20', 'bg_color': '#C8E6C9', 'border': 1})
    f['fail_fmt'] = wb.add_format({'bold': True, 'font_color': '#B71C1C', 'bg_color': '#FFCDD2', 'border': 1})
    f['warn_fmt'] = wb.add_format({'bold': True, 'font_color': '#E65100', 'bg_color': '#FFE0B2', 'border': 1})
    f['center']   = wb.add_format({'align': 'center', 'border': 1})
    return f


# =============================================================================
# INPUTS SHEET
# =============================================================================

def _create_inputs_sheet(wb, ws, f) -> None:
    ws.set_column('A:A', 3)
    ws.set_column('B:B', 32)
    ws.set_column('C:C', 20)
    ws.set_column('D:D', 14)
    ws.set_column('E:E', 12)
    ws.set_column('F:F', 42)

    R = CellRefs

    # ── Title ──────────────────────────────────────────────────────────────
    ws.merge_range('B2:F2', 'BESS Economic Analyzer — JIT Cohort Model', f['title'])
    ws.merge_range('B3:F3', 'Version 2.0  |  February 2026  |  For Electric Utility Engineering & Finance',
                   f['subtitle'])

    # ── Library Selection (row 6) ──────────────────────────────────────────
    ws.write('B5', 'ASSUMPTION LIBRARY', f['bold'])
    ws.write('B6', 'Selected Library', f['bold'])
    ws.write(R.SELECTED_LIBRARY, 'NREL ATB 2024 - Moderate Scenario', f['input'])
    ws.write('D6', '↑ Overrides defaults below', f['tooltip'])
    ws.write('F6', 'Options: NREL ATB 2024 - Moderate Scenario | Lazard LCOS v10.0 - 2025 | CPUC California 2024',
             f['tooltip'])

    # ── Project Basics (rows 8-17) ─────────────────────────────────────────
    ws.merge_range('B8:F8', 'PROJECT BASICS', f['section'])
    basics = [
        # (row, label, cell, value, format, tooltip)
        (9,  'Project Name',            'C9',  'My BESS Project',         f['input'],     'Descriptive project name'),
        (10, 'Project ID',              'C10', 'BESS-001',                f['input'],     'Unique identifier for tracking'),
        (11, 'Location / Market',       'C11', 'CAISO NP15',              f['input'],     'ISO zone or substation name'),
        (12, 'Total Capacity (MW)',     'C12', '=SUM(D20:D29)',           f['fml_num'],   'Auto-summed from Build Schedule'),
        (13, 'Duration (hours)',        'C13', 4,                         f['input'],     '1–8 hrs typical utility-scale'),
        (14, 'Total Energy (MWh)',      'C14', f'={R.CAPACITY_MW}*{R.DURATION_HOURS}',
                                                                           f['formula'],   'Auto-calculated'),
        (15, 'Analysis Period (years)', 'C15', 20,                        f['input_int'], '15–25 yrs typical'),
        (16, 'Discount Rate (WACC)',    'C16', 0.07,                      f['input_pct'], '6–10% utility; 8–12% merchant'),
        (17, 'Ownership Type',          'C17', 'Utility',                 f['input'],     '"Utility" or "Merchant"'),
    ]
    for row, label, cell, value, fmt, tip in basics:
        ws.write(f'B{row}', label, f['bold'])
        if isinstance(value, str) and value.startswith('='):
            ws.write_formula(cell, value, fmt)
        else:
            ws.write(cell, value, fmt)
        ws.write(f'F{row}', tip, f['tooltip'])

    # ── Phased Build Schedule (rows 18-29) ────────────────────────────────
    ws.merge_range('B18:F18', 'PHASED BUILD SCHEDULE  (JIT Cohort Model)', f['section'])
    for col, hdr in enumerate(['Cohort', 'COD Year', 'Capacity (MW)', 'ITC Rate (%)']):
        ws.write(18, 1 + col, hdr, f['header'])   # row 19 → 0-indexed 18

    default_years = [2027, 2028, 2029, 2030, 2031, 2032, 0, 0, 0, 0]
    default_mws   = [20,   20,   20,   20,   20,   0,    0, 0, 0, 0]
    for i in range(10):
        r = 19 + i   # 0-indexed → Excel rows 20-29
        ws.write(r, 1, f'Tranche {i+1}')
        ws.write(r, 2, default_years[i], f['input_int'])
        ws.write(r, 3, default_mws[i],   f['input'])
        ws.write_formula(r, 4,
            f'=IF(D{r+1}>0,Inputs!{R.ITC_BASE_RATE}+Inputs!{R.ITC_ADDERS},0)',
            f['fml_pct'])

    wb.define_name('Build_Schedule', '=Inputs!$C$20:$E$29')
    wb.define_name('Discount_Rate',  f'=Inputs!${R.DISCOUNT_RATE}')
    wb.define_name('Learning_Rate',  f'=Inputs!${R.LEARNING_RATE}')

    # ── T&D Deferral (rows 31-36, blank at 30) ────────────────────────────
    ws.merge_range('B31:F31', 'T&D CAPITAL DEFERRAL  |  PV = K × [1 − ((1+g)/(1+r))^n]',
                   f['section'])
    td_rows = [
        (32, 'Deferred Capital Cost (K)', R.TD_DEFERRAL_K,      100_000_000, f['input_cur'],
             'Total traditional wires capital cost deferred ($)'),
        (33, 'Year Asset Needed (t_need)', R.TD_DEFERRAL_T_NEED, 2032,       f['input_int'],
             'Calendar year the T&D asset would otherwise be needed'),
        (34, 'Deferral Period (n years)',  R.TD_DEFERRAL_N,      5,           f['input_int'],
             'Number of years the traditional investment is deferred'),
        (35, 'Load Growth Rate (g)',       R.TD_DEFERRAL_G,      0.02,        f['input_pct'],
             'Annual growth / inflation of the deferred capital cost'),
    ]
    for row, label, cell, value, fmt, tip in td_rows:
        ws.write(f'B{row}', label, f['bold'])
        ws.write(cell, value, fmt)
        ws.write(f'F{row}', tip, f['tooltip'])
    ws.write('B36', 'PV of Deferral Benefit ($)', f['bold'])
    ws.write_formula(R.TD_DEFERRAL_PV,
        f'={R.TD_DEFERRAL_K}*(1-((1+{R.TD_DEFERRAL_G})/(1+{R.DISCOUNT_RATE}))^{R.TD_DEFERRAL_N})',
        f['fml_cur'])
    ws.write('F36', 'Present value of the T&D deferral benefit — added to Results', f['tooltip'])

    # ── Technology Specs (rows 38-44, blank at 37) ────────────────────────
    ws.merge_range('B38:F38', 'TECHNOLOGY SPECIFICATIONS', f['section'])
    tech = [
        (39, 'Battery Chemistry',           R.CHEMISTRY,             'LFP',   f['input'],
             'LFP (default) | NMC | Other'),
        (40, 'Round-Trip Efficiency (RTE)', R.ROUND_TRIP_EFFICIENCY,  0.85,   f['input_pct'],
             'AC-AC efficiency including inverter losses (70–95%)'),
        (41, 'Annual Degradation Rate',     R.ANNUAL_DEGRADATION,     0.025,  f['input_pct'],
             'Capacity loss per year; 2.5% typical for LFP'),
        (42, 'Cycle Life (full cycles)',    R.CYCLE_LIFE,             6000,   f['input_int'],
             'Number of full-depth cycles before end-of-life'),
        (43, 'Augmentation Year',           R.AUGMENTATION_YEAR,      12,     f['input_int'],
             'Year battery modules are replaced; typically 10–12'),
        (44, 'Cycles per Day',              R.CYCLES_PER_DAY,         1.0,    f['input'],
             'Average daily full charge-discharge cycles (0.1–3.0)'),
    ]
    for row, label, cell, value, fmt, tip in tech:
        ws.write(f'B{row}', label, f['bold'])
        ws.write(cell, value, fmt)
        ws.write(f'F{row}', tip, f['tooltip'])

    # ── Cost Inputs (rows 45-52, no blank) ───────────────────────────────
    ws.merge_range('B45:F45', 'COST INPUTS — BESS', f['section'])
    costs = [
        (46, 'CapEx ($/kWh)',                R.CAPEX_PER_KWH,    160,  f['input_cur'], '$130–200/kWh utility-scale LFP'),
        (47, 'Fixed O&M ($/kW-year)',        R.FOM_PER_KW_YEAR,   25,  f['input_cur'], 'Site maintenance, monitoring, security'),
        (48, 'Variable O&M ($/MWh)',         R.VOM_PER_MWH,        0,  f['input_cur'], 'Per-MWh discharge cost; often $0 for LFP'),
        (49, 'Augmentation Cost ($/kWh)',    R.AUGMENTATION_COST,  55,  f['input_cur'], 'Battery module replacement; learning-curve adjusted'),
        (50, 'Decommissioning ($/kW)',       R.DECOMMISSIONING,    10,  f['input_cur'], 'End-of-life removal and site restoration'),
        (51, 'Charging Cost ($/MWh)',        R.CHARGING_COST,      30,  f['input_cur'], 'Grid electricity cost for charging cycles'),
        (52, 'Residual Value (% of CapEx)',  R.RESIDUAL_VALUE,    0.10, f['input_pct'], 'Salvage value at end of analysis period'),
    ]
    for row, label, cell, value, fmt, tip in costs:
        ws.write(f'B{row}', label, f['bold'])
        ws.write(cell, value, fmt)
        ws.write(f'F{row}', tip, f['tooltip'])

    # ── Tax Credits (rows 53-55, no blank) ────────────────────────────────
    ws.merge_range('B53:F53', 'INVESTMENT TAX CREDIT — BESS-Specific (IRA 2022)', f['section'])
    ws.write('B54', 'ITC Base Rate',   f['bold']); ws.write(R.ITC_BASE_RATE, 0.30, f['input_pct'])
    ws.write('F54', '30% base rate for standalone storage under IRA', f['tooltip'])
    ws.write('B55', 'ITC Adders',      f['bold']); ws.write(R.ITC_ADDERS, 0.00, f['input_pct'])
    ws.write('F55', 'Energy community +10%, Domestic content +10%, Low-income +10–20%', f['tooltip'])

    # ── Infrastructure Costs (rows 56-61) ─────────────────────────────────
    ws.merge_range('B56:F56', 'INFRASTRUCTURE COSTS — Common to All Utility Projects', f['section'])
    infra = [
        (57, 'Interconnection ($/kW)',  R.INTERCONNECTION,  100,   f['input_cur'], 'Network upgrades, studies, metering'),
        (58, 'Land ($/kW)',             R.LAND,              10,   f['input_cur'], 'Site acquisition or capitalized lease'),
        (59, 'Permitting ($/kW)',       R.PERMITTING,        15,   f['input_cur'], 'Environmental review, permits, legal'),
        (60, 'Insurance (% of CapEx)',  R.INSURANCE_PCT,    0.005, f['input_pct'], 'Annual property & liability insurance'),
        (61, 'Property Tax (% of book)',R.PROPERTY_TAX_PCT, 0.010, f['input_pct'], 'Annual property tax on net book value'),
    ]
    for row, label, cell, value, fmt, tip in infra:
        ws.write(f'B{row}', label, f['bold'])
        ws.write(cell, value, fmt)
        ws.write(f'F{row}', tip, f['tooltip'])

    # ── Financing Structure (rows 62-68) ─────────────────────────────────
    ws.merge_range('B62:F62', 'FINANCING STRUCTURE  (overrides Discount Rate for WACC)', f['section'])
    fin = [
        (63, 'Debt Financing (%)',    R.DEBT_PERCENT,   0.60, f['input_pct'], '60% debt / 40% equity typical utility'),
        (64, 'Interest Rate on Debt', R.INTEREST_RATE,  0.045,f['input_pct'], 'Annual debt interest rate'),
        (65, 'Loan Term (years)',     R.LOAN_TERM,      15,   f['input_int'], 'Amortization period'),
        (66, 'Cost of Equity',        R.COST_OF_EQUITY, 0.10, f['input_pct'], 'Required equity return'),
        (67, 'Corporate Tax Rate',    R.TAX_RATE,       0.21, f['input_pct'], '21% federal; add state if applicable'),
    ]
    for row, label, cell, value, fmt, tip in fin:
        ws.write(f'B{row}', label, f['bold'])
        ws.write(cell, value, fmt)
        ws.write(f'F{row}', tip, f['tooltip'])
    ws.write('B68', 'Calculated WACC', f['bold'])
    ws.write_formula(R.WACC,
        f'=(1-{R.DEBT_PERCENT})*{R.COST_OF_EQUITY}+{R.DEBT_PERCENT}*{R.INTEREST_RATE}*(1-{R.TAX_RATE})',
        f['fml_pct'])
    ws.write('F68', 'WACC = (E/V)×Re + (D/V)×Rd×(1-Tc)  — use this as Discount Rate', f['tooltip'])

    # ── Benefit Streams (rows 70-79, blank at 69) ─────────────────────────
    ws.merge_range('B70:F70', 'ANNUAL BENEFIT STREAMS  ($/kW-year, Year 1)', f['section'])
    for col, hdr in enumerate(['Benefit Name', '$/kW-yr (Y1)', 'Escalation/yr',
                                'Category', 'Source']):
        ws.write(70, 1 + col, hdr, f['header'])   # row 71 → 0-indexed 70

    benefit_data = [
        # (row, name, $/kW, esc, category, source)
        (72, 'Resource Adequacy',       150, 0.020, 'Common',        'CPUC RA Program D.24-06-050'),
        (73, 'Energy Arbitrage',         40, 0.020, 'Common',        'CAISO OASIS historical LMPs'),
        (74, 'Ancillary Services',       15, 0.010, 'Common',        'CAISO AS market data'),
        (75, 'T&D Deferral',             25, 0.015, 'Common',        'CPUC Avoided Cost Calculator'),
        (76, 'Resilience Value',         50, 0.020, 'Common',        'LBNL ICE Calculator'),
        (77, 'Renewable Integration',    25, 0.025, 'BESS-Specific', 'CAISO Curtailment Reports'),
        (78, 'GHG Emissions Value',      15, 0.030, 'BESS-Specific', 'EPA Social Cost of Carbon'),
        (79, 'Voltage Support',          10, 0.010, 'Common',        'EPRI Distribution Studies'),
    ]
    val_cell  = ['C72','C73','C74','C75','C76','C77','C78','C79']
    esc_cell  = ['D72','D73','D74','D75','D76','D77','D78','D79']
    for i, (row, name, val, esc, cat, src) in enumerate(benefit_data):
        r = row - 1   # 0-indexed
        ws.write(r, 1, name)
        ws.write(val_cell[i], val,  f['input_cur'])
        ws.write(esc_cell[i], esc,  f['input_pct'])
        ws.write(r, 4, cat)
        ws.write(r, 5, src, f['tooltip'])

    # ── Cost Projections (rows 80-82) ─────────────────────────────────────
    ws.merge_range('B80:F80', 'COST PROJECTIONS — Technology Learning Curve', f['section'])
    ws.write('B81', 'Annual Cost Decline Rate', f['bold'])
    ws.write(R.LEARNING_RATE, 0.12, f['input_pct'])
    ws.write('F81', '12% per NREL ATB 2024 — applied to augmentation & multi-tranche CapEx', f['tooltip'])
    ws.write('B82', 'Cost Base Year', f['bold'])
    ws.write(R.COST_BASE_YEAR, 2024, f['input_int'])
    ws.write('F82', 'Reference year for base CapEx and augmentation costs', f['tooltip'])


# =============================================================================
# CALCULATIONS SHEET
# =============================================================================

def _create_calculations_sheet(ws, f) -> None:
    ws.set_column('A:A', 3)
    ws.set_column('B:B', 38)
    ws.set_column('C:C', 20)
    ws.set_column('D:D', 58)

    R = CellRefs
    ws.merge_range('B2:D2', 'CALCULATION ENGINE — Cohort-Based SUMPRODUCT Formulas', f['title'])
    ws.merge_range('B4:D4', 'COHORT-BASED COST AGGREGATION', f['section'])

    calcs = [
        ('Total Capacity (kW)',
         f'=Inputs!{R.CAPACITY_MW}*1000',
         'Sum of all online cohort capacities in kW'),
        ('Total Energy (kWh)',
         f'=Inputs!{R.ENERGY_MWH}*1000',
         'Total energy capacity kWh'),
        ('Total Battery CapEx ($)',
         f'=SUMPRODUCT((INDEX(Build_Schedule,,2)*1000*Inputs!{R.DURATION_HOURS})'
         f'*(Inputs!{R.CAPEX_PER_KWH}*(1-Inputs!{R.LEARNING_RATE})'
         f'^MAX(0,INDEX(Build_Schedule,,1)-Inputs!{R.COST_BASE_YEAR})))',
         'Cohort CapEx with learning-curve adjustment by COD year'),
        ('Total ITC Credit ($)',
         f'=SUMPRODUCT((INDEX(Build_Schedule,,2)*1000*Inputs!{R.DURATION_HOURS})'
         f'*(Inputs!{R.CAPEX_PER_KWH}*(1-Inputs!{R.LEARNING_RATE})'
         f'^MAX(0,INDEX(Build_Schedule,,1)-Inputs!{R.COST_BASE_YEAR}))'
         f'*INDEX(Build_Schedule,,3))',
         'ITC credit per cohort based on its learning-curve-adjusted CapEx'),
        ('Net Year-0 Battery Cost ($)',
         '=C7-C8',
         'Total CapEx minus ITC (before infrastructure costs)'),
        ('Infrastructure Costs ($)',
         f'=Inputs!{R.CAPACITY_MW}*1000'
         f'*(Inputs!{R.INTERCONNECTION}+Inputs!{R.LAND}+Inputs!{R.PERMITTING})',
         'Interconnection + land + permitting (one-time, Year 0)'),
        ('Total Net Year-0 Cost ($)',
         '=C9+C10',
         'Battery net cost + infrastructure'),
        ('Annual Fixed O&M ($)',
         f'=Inputs!{R.CAPACITY_MW}*1000*Inputs!{R.FOM_PER_KW_YEAR}',
         'Fixed O&M when all tranches are online'),
        ('Annual Charging Cost ($)',
         f'=Inputs!{R.ENERGY_MWH}*Inputs!{R.CYCLES_PER_DAY}*365'
         f'*Inputs!{R.CHARGING_COST}',
         'Annual electricity cost for charging (all tranches online)'),
        ('Annual Insurance ($)',
         f'=C7*Inputs!{R.INSURANCE_PCT}',
         'Insurance = total CapEx × insurance rate'),
        ('Augmentation Cost (Year {aug_yr}) ($)',
         f'=Inputs!{R.ENERGY_MWH}*1000'
         f'*Inputs!{R.AUGMENTATION_COST}'
         f'*(1-Inputs!{R.LEARNING_RATE})^Inputs!{R.AUGMENTATION_YEAR}',
         'Augmentation adjusted for learning curve'),
    ]

    for i, (label, formula, desc) in enumerate(calcs):
        row = 4 + i   # 0-indexed; Excel rows 5, 6, 7…
        ws.write(row, 1, label, f['bold'])
        ws.write_formula(row, 2, formula, f['fml_cur'])
        ws.write(row, 3, desc, f['tooltip'])


# =============================================================================
# CASH FLOWS SHEET
# =============================================================================
# Columns (0-indexed): B=1 Year | C=2 CapEx(net ITC) | D=3 O&M | E=4 Charging
#   F=5 Aug | G=6 Infra | H=7 Total Costs | I=8 PV Costs
#   J=9 Benefits | K=10 PV Benefits | L=11 Net CF | M=12 PV Net CF
#   N=13 Cumulative CF | O=14 Energy (MWh)

CF_COL = {
    'year':   1,   # B
    'capex':  2,   # C
    'om':     3,   # D
    'charge': 4,   # E
    'aug':    5,   # F
    'infra':  6,   # G
    'costs':  7,   # H
    'pv_c':   8,   # I
    'ben':    9,   # J
    'pv_b':  10,   # K
    'net':   11,   # L
    'pv_n':  12,   # M
    'cum':   13,   # N
    'energy':14,   # O
}

def _create_cashflows_sheet(ws, f) -> dict:
    """Build annual cash flow sheet; return {'data_start': row, 'data_end': row}."""
    ws.set_column('A:A', 3)
    ws.set_column('B:B', 6)
    ws.set_column('C:N', 14)
    ws.set_column('O:O', 14)

    R   = CellRefs
    COL = CF_COL

    ws.merge_range('B2:O2', 'Annual Cash Flow Projections — Cohort-Aggregated', f['title'])

    headers = [
        'Year', 'CapEx\n(net ITC)', 'Fixed O&M', 'Charging',
        'Augment', 'Infra', 'Total\nCosts', 'PV Costs',
        'Total\nBenefits', 'PV\nBenefits', 'Net CF',
        'PV Net CF', 'Cumul CF', 'Energy\n(MWh)',
    ]
    ws.set_row(3, 30)
    for c, hdr in enumerate(headers):
        ws.write(3, 1 + c, hdr, f['header'])

    DATA_START = 4      # 0-indexed → Excel row 5
    N_YEARS    = 21     # years 0 – 20

    for y in range(N_YEARS):
        row   = DATA_START + y       # 0-indexed
        erow  = row + 1              # Excel 1-based row for A1 refs

        ws.write(row, COL['year'], y, f['center'])

        # ── CapEx (net of ITC) ────────────────────────────────────────────
        # For each cohort online in year y: capex_kwh × capacity_kwh × learning − ITC
        capex_f = (
            f'=SUMPRODUCT(--(INDEX(Build_Schedule,,1)={y}),'
            f'(INDEX(Build_Schedule,,2)*1000*Inputs!{R.DURATION_HOURS})'
            f'*(Inputs!{R.CAPEX_PER_KWH}*(1-Inputs!{R.LEARNING_RATE})'
            f'^MAX(0,{y}-Inputs!{R.COST_BASE_YEAR}))'
            f'*(1-INDEX(Build_Schedule,,3)))'
        )
        ws.write_formula(row, COL['capex'], capex_f, f['currency'])

        # ── Fixed O&M (cohorts online) ────────────────────────────────────
        om_f = (
            f'=SUMPRODUCT(--(INDEX(Build_Schedule,,1)<{y}),'
            f'INDEX(Build_Schedule,,2)*1000*Inputs!{R.FOM_PER_KW_YEAR})'
            if y > 0 else '=0'
        )
        ws.write_formula(row, COL['om'], om_f, f['currency'])

        # ── Charging Cost (online cohorts, with degradation) ──────────────
        charge_f = (
            f'=SUMPRODUCT(--(INDEX(Build_Schedule,,1)<{y}),'
            f'INDEX(Build_Schedule,,2)*1000*Inputs!{R.DURATION_HOURS}'
            f'*Inputs!{R.CYCLES_PER_DAY}*365*Inputs!{R.ROUND_TRIP_EFFICIENCY}'
            f'*(1-Inputs!{R.ANNUAL_DEGRADATION})^({y}-INDEX(Build_Schedule,,1))'
            f'*Inputs!{R.CHARGING_COST})'
            if y > 0 else '=0'
        )
        ws.write_formula(row, COL['charge'], charge_f, f['currency'])

        # ── Augmentation ──────────────────────────────────────────────────
        aug_f = (
            f'=SUMPRODUCT(--(INDEX(Build_Schedule,,1)+Inputs!{R.AUGMENTATION_YEAR}={y}),'
            f'INDEX(Build_Schedule,,2)*1000*Inputs!{R.DURATION_HOURS}'
            f'*Inputs!{R.AUGMENTATION_COST}'
            f'*(1-Inputs!{R.LEARNING_RATE})^MAX(0,Inputs!{R.AUGMENTATION_YEAR}))'
            if y > 0 else '=0'
        )
        ws.write_formula(row, COL['aug'], aug_f, f['currency'])

        # ── Infrastructure (one-time at COD year of each cohort) ──────────
        infra_f = (
            f'=SUMPRODUCT(--(INDEX(Build_Schedule,,1)={y}),'
            f'INDEX(Build_Schedule,,2)*1000'
            f'*(Inputs!{R.INTERCONNECTION}+Inputs!{R.LAND}+Inputs!{R.PERMITTING}))'
        )
        ws.write_formula(row, COL['infra'], infra_f, f['currency'])

        # ── Total Costs ───────────────────────────────────────────────────
        ws.write_formula(row, COL['costs'],
            f'=SUM(C{erow}:G{erow})', f['currency'])

        # ── PV Costs ─────────────────────────────────────────────────────
        ws.write_formula(row, COL['pv_c'],
            f'=H{erow}/(1+Inputs!{R.DISCOUNT_RATE})^{y}', f['currency'])

        # ── Total Benefits (all 8 streams, escalated, capacity-weighted) ──
        # Sum: $/kW × capacity_kW × (1+esc)^(year-1) × degradation_weighted capacity
        ben_parts = []
        for val_c, esc_c in R.BENEFITS:
            if y > 0:
                ben_parts.append(
                    f'SUMPRODUCT(--(INDEX(Build_Schedule,,1)<{y}),'
                    f'INDEX(Build_Schedule,,2)*1000'
                    f'*(1-Inputs!{R.ANNUAL_DEGRADATION})^({y}-INDEX(Build_Schedule,,1))'
                    f')*Inputs!{val_c}*(1+Inputs!{esc_c})^{y-1}'
                    f'/IFERROR(Inputs!{R.CAPACITY_MW}*1000,1)'
                )
        if ben_parts:
            ben_f = '=' + '+'.join(ben_parts)
        else:
            ben_f = '=0'
        ws.write_formula(row, COL['ben'], ben_f, f['currency'])

        # ── PV Benefits ──────────────────────────────────────────────────
        ws.write_formula(row, COL['pv_b'],
            f'=J{erow}/(1+Inputs!{R.DISCOUNT_RATE})^{y}', f['currency'])

        # ── Net CF ───────────────────────────────────────────────────────
        ws.write_formula(row, COL['net'],
            f'=J{erow}-H{erow}', f['currency'])

        # ── PV Net CF ────────────────────────────────────────────────────
        ws.write_formula(row, COL['pv_n'],
            f'=L{erow}/(1+Inputs!{R.DISCOUNT_RATE})^{y}', f['currency'])

        # ── Cumulative CF ─────────────────────────────────────────────────
        cum_f = (f'=N{erow-1}+L{erow}' if y > 0 else f'=L{erow}')
        ws.write_formula(row, COL['cum'], cum_f, f['currency'])

        # ── Energy Discharged (MWh) ───────────────────────────────────────
        energy_f = (
            f'=SUMPRODUCT(--(INDEX(Build_Schedule,,1)<{y}),'
            f'INDEX(Build_Schedule,,2)*Inputs!{R.DURATION_HOURS}'
            f'*Inputs!{R.CYCLES_PER_DAY}*365*Inputs!{R.ROUND_TRIP_EFFICIENCY}'
            f'*(1-Inputs!{R.ANNUAL_DEGRADATION})^({y}-INDEX(Build_Schedule,,1)))'
            if y > 0 else '=0'
        )
        ws.write_formula(row, COL['energy'], energy_f, f['number'])

    # ── Totals row ────────────────────────────────────────────────────────
    tot = DATA_START + N_YEARS    # 0-indexed
    etot = tot + 1
    d0   = DATA_START + 1         # Excel row of first data row
    ws.write(tot, 1, 'TOTALS / PV SUM', f['bold'])
    for col_idx in [COL['pv_c'], COL['pv_b'], COL['pv_n'], COL['energy']]:
        col_ltr = xl_rowcol_to_cell(DATA_START, col_idx)[0]  # get letter
        ws.write_formula(tot, col_idx,
            f'=SUM({col_ltr}{d0}:{col_ltr}{etot-1})', f['fml_cur'])

    return {'data_start': DATA_START, 'data_end': DATA_START + N_YEARS - 1,
            'tot_row': tot, 'n_years': N_YEARS}


# =============================================================================
# RESULTS SHEET
# =============================================================================

def _create_results_sheet(ws, f, cf) -> None:
    ws.set_column('A:A', 3)
    ws.set_column('B:B', 32)
    ws.set_column('C:C', 22)
    ws.set_column('D:D', 14)
    ws.set_column('E:E', 40)

    R   = CellRefs
    COL = CF_COL

    d0  = cf['data_start'] + 1          # Excel first data row
    dn  = cf['data_end']   + 1          # Excel last data row

    # Helper: Excel column letter from 0-indexed column index
    def col_ltr(idx):
        return xl_rowcol_to_cell(0, idx)[0]

    pv_c_col  = col_ltr(COL['pv_c'])
    pv_b_col  = col_ltr(COL['pv_b'])
    pv_n_col  = col_ltr(COL['pv_n'])
    net_col   = col_ltr(COL['net'])
    cum_col   = col_ltr(COL['cum'])
    en_col    = col_ltr(COL['energy'])

    ws.merge_range('B2:E2', 'BESS Economic Analysis — Results Dashboard', f['title'])
    ws.merge_range('B3:E3',
        f'Project: [auto from Inputs!{R.PROJECT_NAME}]  |  Analysis Date: {datetime.now():%B %d, %Y}',
        f['subtitle'])

    # ── Key Financial Metrics ─────────────────────────────────────────────
    ws.merge_range('B5:E5', 'KEY FINANCIAL METRICS', f['section'])
    ws.write('B6', 'Metric',                  f['header'])
    ws.write('C6', 'Value',                   f['header'])
    ws.write('D6', 'Benchmark',               f['header'])
    ws.write('E6', 'Interpretation',          f['header'])

    metrics = [
        # (label, formula, fmt_key, benchmark, interpretation)
        ('Net Present Value (NPV)',
         f"=SUM(Cash_Flows!{pv_n_col}{d0}:Cash_Flows!{pv_n_col}{dn})"
         f"+Inputs!{R.TD_DEFERRAL_PV}",
         'result_big', '> $0', 'Positive NPV creates shareholder/ratepayer value'),

        ('T&D Deferral PV',
         f"=Inputs!{R.TD_DEFERRAL_PV}",
         'result_big', '—', 'Present value of deferred T&D capital investment'),

        ('Total Project Value',
         f"=SUM(Cash_Flows!{pv_n_col}{d0}:Cash_Flows!{pv_n_col}{dn})"
         f"+Inputs!{R.TD_DEFERRAL_PV}",
         'result_big', '> $0', 'NPV inclusive of T&D deferral benefit'),

        ('Benefit-Cost Ratio (BCR)',
         f"=IFERROR(SUM(Cash_Flows!{pv_b_col}{d0}:Cash_Flows!{pv_b_col}{dn})"
         f"/SUM(Cash_Flows!{pv_c_col}{d0}:Cash_Flows!{pv_c_col}{dn}),0)",
         'result_num', '≥ 1.0 approve', '≥1.5 strong | 1.0–1.5 marginal | <1.0 reject (CPUC SPM)'),

        ('Internal Rate of Return (IRR)',
         f"=IFERROR(IRR(Cash_Flows!{net_col}{d0}:Cash_Flows!{net_col}{dn}),\"N/A\")",
         'result_pct', '> WACC', 'IRR > WACC creates value above cost of capital'),

        ('Simple Payback (years)',
         f"=IFERROR(MATCH(TRUE,Cash_Flows!{cum_col}{d0}:Cash_Flows!{cum_col}{dn}>0,0)-1,\"Never\")",
         'result_num', '< 10 yrs', '< 7 yrs strong; 7–10 acceptable; > 10 yrs challenging'),

        ('LCOS ($/MWh)',
         f"=IFERROR(SUM(Cash_Flows!{pv_c_col}{d0}:Cash_Flows!{pv_c_col}{dn})"
         f"/SUM(Cash_Flows!{en_col}{d0}:Cash_Flows!{en_col}{dn}),0)",
         'result_num', '$100–200/MWh', 'Levelized cost per MWh discharged (Lazard methodology)'),

        ('Breakeven CapEx ($/kWh)',
         f"=IFERROR((SUM(Cash_Flows!{pv_b_col}{d0}:Cash_Flows!{pv_b_col}{dn})"
         f"-(SUM(Cash_Flows!{pv_c_col}{d0}:Cash_Flows!{pv_c_col}{dn})"
         f"-Cash_Flows!{pv_c_col}{d0}))"
         f"/(Inputs!{R.CAPACITY_MW}*Inputs!{R.DURATION_HOURS}*1000),0)",
         'result_num', '> CapEx input', 'Max CapEx where BCR = 1.0 (margin of safety)'),
    ]

    for i, (label, formula, fmt_key, bench, interp) in enumerate(metrics):
        row = 6 + i   # 0-indexed → Excel rows 7–14
        ws.write(row, 1, label, f['bold'])
        ws.write_formula(row, 2, formula, f[fmt_key])
        ws.write(row, 3, bench, f['center'])
        ws.write(row, 4, interp, f['tooltip'])

    # ── Project Summary ───────────────────────────────────────────────────
    ws.merge_range('B16:E16', 'PROJECT SUMMARY', f['section'])
    summary = [
        ('Project Name',     f"=Inputs!{R.PROJECT_NAME}"),
        ('Capacity (MW)',    f"=Inputs!{R.CAPACITY_MW}"),
        ('Duration (hrs)',   f"=Inputs!{R.DURATION_HOURS}"),
        ('Analysis Period',  f"=Inputs!{R.ANALYSIS_YEARS}&\" years\""),
        ('Discount Rate',    f"=Inputs!{R.DISCOUNT_RATE}"),
        ('WACC (Financing)', f"=Inputs!{R.WACC}"),
        ('Ownership Type',   f"=Inputs!{R.OWNERSHIP_TYPE}"),
        ('CapEx ($/kWh)',    f"=Inputs!{R.CAPEX_PER_KWH}"),
        ('ITC Rate',         f"=Inputs!{R.ITC_BASE_RATE}+Inputs!{R.ITC_ADDERS}"),
    ]
    for i, (label, formula) in enumerate(summary):
        row = 16 + i
        ws.write(row, 1, label, f['bold'])
        ws.write_formula(row, 2, formula, f['formula'])


# =============================================================================
# SENSITIVITY SHEET
# =============================================================================

def _create_sensitivity_sheet(ws, f, cf) -> None:
    ws.set_column('A:A', 3)
    ws.set_column('B:B', 30)
    ws.set_column('C:G', 16)
    ws.set_column('H:H', 36)

    R   = CellRefs
    COL = CF_COL

    d0 = cf['data_start'] + 1
    dn = cf['data_end']   + 1

    def col_ltr(idx):
        return xl_rowcol_to_cell(0, idx)[0]

    pv_n_col = col_ltr(COL['pv_n'])
    base_npv = (f"SUM(Cash_Flows!{pv_n_col}{d0}:Cash_Flows!{pv_n_col}{dn})"
                f"+Inputs!{R.TD_DEFERRAL_PV}")

    ws.merge_range('B2:H2', 'SENSITIVITY ANALYSIS — Tornado Chart Parameters', f['title'])
    ws.merge_range('B3:H3',
        'Each row shows the base case and ±20% deviation for one input parameter. '
        'Recalculate NPV in Python GUI for exact values; Excel shows analytical estimates.',
        f['subtitle'])

    ws.merge_range('B5:H5', 'PARAMETER SENSITIVITY TABLE', f['section'])
    for c, hdr in enumerate(['Parameter', 'Low Case', 'Base Case', 'High Case',
                               'Low NPV (est.)', 'Base NPV', 'High NPV (est.)']):
        ws.write(5, 1 + c, hdr, f['header'])

    # Base NPV reference
    base_npv_ref = f'={base_npv}'

    sensitivity_params = [
        # (label, input_cell, low_mult, high_mult, low_label, high_label)
        ('CapEx ($/kWh)',       R.CAPEX_PER_KWH,    0.80, 1.20, '−20% CapEx', '+20% CapEx'),
        ('Discount Rate',       R.DISCOUNT_RATE,    0.80, 1.20, '−20% rate',  '+20% rate'),
        ('Resource Adequacy ($/kW-yr)', R.BENEFIT_RA, 0.75, 1.25, '−25% RA', '+25% RA'),
        ('Learning Rate',       R.LEARNING_RATE,    0.417, 1.25, '5% rate', '15% rate'),
        ('Round-Trip Efficiency',R.ROUND_TRIP_EFFICIENCY,0.90,1.05,'−10% RTE','+5% RTE'),
        ('Analysis Period (yrs)',R.ANALYSIS_YEARS,   0.80, 1.20, '−20% years', '+20% years'),
        ('ITC Rate',            R.ITC_BASE_RATE,    0.67, 1.00, 'No adders', '+10% adder'),
        ('Fixed O&M ($/kW-yr)', R.FOM_PER_KW_YEAR,  0.80, 1.20, '−20% O&M', '+20% O&M'),
    ]

    for i, (label, cell, low_m, high_m, low_lbl, high_lbl) in enumerate(sensitivity_params):
        row = 6 + i
        ws.write(row, 1, label, f['bold'])
        ws.write(row, 2, f'Inputs!{cell} × {low_m:.2f}  ({low_lbl})',  f['tooltip'])
        ws.write_formula(row, 3, f'=Inputs!{cell}', f['formula'])
        ws.write(row, 4, f'Inputs!{cell} × {high_m:.2f}  ({high_lbl})', f['tooltip'])
        # Estimated NPV for low/high (analytical delta for key drivers)
        ws.write(row, 5, '← run Python GUI for exact', f['tooltip'])
        ws.write_formula(row, 6, base_npv_ref, f['fml_cur'])
        ws.write(row, 7, '← run Python GUI for exact', f['tooltip'])

    # ── Interpretation guide ──────────────────────────────────────────────
    ws.merge_range('B16:H16', 'HOW TO READ A TORNADO CHART', f['section'])
    guide = [
        'Parameters are sorted by impact (longest bar = highest sensitivity)',
        'Bars extending left (negative) show parameters where increase hurts NPV',
        'CapEx and Discount Rate typically dominate BESS sensitivity',
        'RA value is the largest single benefit driver in capacity-constrained markets',
        'Use the Python GUI → Sensitivity tab for exact recalculations and chart export',
    ]
    for i, line in enumerate(guide):
        ws.write(16 + i, 1, f'• {line}', f['tooltip'])

    # ── Recommended test ranges ───────────────────────────────────────────
    ws.merge_range('B23:H23', 'RECOMMENDED SCENARIO RANGES', f['section'])
    headers2 = ['Parameter', 'Low', 'Base', 'High', 'Source']
    for c, hdr in enumerate(headers2):
        ws.write(23, 1 + c, hdr, f['header'])

    ranges = [
        ('CapEx ($/kWh)',          '$130',  '$160', '$200',  'NREL ATB 2024'),
        ('Discount Rate',          '6.0%',  '7.0%', '9.0%',  'CPUC authorized WACC range'),
        ('Resource Adequacy',      '$120/kW','$150/kW','$200/kW','CPUC RA decisions'),
        ('Learning Rate',          '5%',    '12%',  '15%',   'BNEF / NREL ATB'),
        ('Degradation Rate',       '1.5%',  '2.5%', '3.5%',  'Manufacturer data'),
        ('Analysis Period (yrs)',  '15',    '20',   '25',    'Utility planning horizon'),
    ]
    for i, (param, lo, base, hi, src) in enumerate(ranges):
        row = 24 + i
        ws.write(row, 1, param, f['bold'])
        ws.write(row, 2, lo,   f['center'])
        ws.write(row, 3, base, f['formula'])
        ws.write(row, 4, hi,   f['center'])
        ws.write(row, 5, src,  f['tooltip'])


# =============================================================================
# UOS ANALYSIS SHEET
# =============================================================================

def _create_uos_sheet(ws, f) -> None:
    ws.set_column('A:A', 3)
    ws.set_column('B:B', 34)
    ws.set_column('C:C', 18)
    ws.set_column('D:D', 18)
    ws.set_column('E:E', 40)

    R = CellRefs

    ws.merge_range('B2:E2', 'UTILITY-OWNED STORAGE (UOS) — Revenue Requirement Analysis', f['title'])
    ws.merge_range('B3:E3',
        'Per CPUC D.25-12-003 (SCE 2026-2028 Cost of Capital) | '
        'Populate from Python GUI or override below.',
        f['subtitle'])

    # ── SCE Cost of Capital ───────────────────────────────────────────────
    ws.merge_range('B5:E5', 'COST OF CAPITAL  (D.25-12-003 SCE Defaults)', f['section'])
    coc_data = [
        ('Return on Equity (ROE)',       0.1003, '10.03%'),
        ('Cost of Debt',                 0.0471, '4.71%'),
        ('Cost of Preferred',            0.0548, '5.48%'),
        ('Equity Ratio',                 0.5200, '52.00%'),
        ('Debt Ratio',                   0.4347, '43.47%'),
        ('Preferred Ratio',              0.0453, '4.53%'),
        ('Authorized ROR',               0.0759, '7.59%'),
        ('Federal Tax Rate',             0.2100, '21.00%'),
        ('State Tax Rate (CA)',           0.0884, '8.84%'),
        ('Property Tax Rate',            0.0100, '1.00%'),
    ]
    ws.write('B6', 'Parameter', f['header']); ws.write('C6', 'Value', f['header'])
    ws.write('D6', 'Default',   f['header']); ws.write('E6', 'Notes', f['header'])
    for i, (label, value, default) in enumerate(coc_data):
        r = 6 + i
        ws.write(r, 1, label, f['bold'])
        ws.write(r, 2, value, f['input_pct'])
        ws.write(r, 3, default, f['center'])
    ws.write('E7', 'Per D.25-12-003', f['tooltip'])
    ws.write('E13', 'Composite: T_state + T_fed×(1-T_state)', f['tooltip'])

    # ── Rate Base Inputs ──────────────────────────────────────────────────
    ws.merge_range('B18:E18', 'RATE BASE INPUTS', f['section'])
    rb_data = [
        ('Book Life (years)',              20,   'Straight-line depreciation life'),
        ('MACRS Property Class',            7,   '7-year MACRS for utility storage'),
        ('Bonus Depreciation (%)',         0.0,  'IRA bonus depreciation if applicable'),
        ('Gross Plant (from Inputs)',
         f'=Inputs!{R.CAPACITY_MW}*Inputs!{R.DURATION_HOURS}*1000*Inputs!{R.CAPEX_PER_KWH}'
         f'+Inputs!{R.CAPACITY_MW}*1000*(Inputs!{R.INTERCONNECTION}+Inputs!{R.LAND}+Inputs!{R.PERMITTING})',
         'Battery CapEx + infrastructure'),
        ('ITC Amount',
         f'=Inputs!{R.CAPACITY_MW}*Inputs!{R.DURATION_HOURS}*1000*Inputs!{R.CAPEX_PER_KWH}'
         f'*(Inputs!{R.ITC_BASE_RATE}+Inputs!{R.ITC_ADDERS})',
         'ITC applied to battery CapEx only'),
        ('Annual O&M',
         f'=Inputs!{R.CAPACITY_MW}*1000*Inputs!{R.FOM_PER_KW_YEAR}',
         'Fixed O&M from Inputs sheet'),
    ]
    ws.write('B19', 'Parameter', f['header']); ws.write('C19', 'Value', f['header'])
    ws.write('E19', 'Notes',     f['header'])
    for i, row_data in enumerate(rb_data):
        r = 19 + i
        label = row_data[0]; value = row_data[1]; tip = row_data[2]
        ws.write(r, 1, label, f['bold'])
        if isinstance(value, str) and value.startswith('='):
            ws.write_formula(r, 2, value, f['fml_cur'])
        else:
            ws.write(r, 2, value, f['input'])
        ws.write(r, 4, tip, f['tooltip'])

    # ── Revenue Requirement Schedule Template ─────────────────────────────
    ws.merge_range('B28:E28', 'REVENUE REQUIREMENT SCHEDULE  (populate via Python GUI)', f['section'])
    rr_headers = ['Year', 'Gross Plant', 'Book Depr', 'Tax Depr (MACRS)',
                  'ADIT', 'Net Rate Base', 'Return on RB', 'Income Tax',
                  'Property Tax', 'O&M', 'Revenue Req.']
    ws.set_row(28, 28)
    for c, hdr in enumerate(rr_headers):
        ws.write(28, 1 + c, hdr, f['header'])
    for yr in range(1, 21):
        ws.write(28 + yr, 1, yr, f['center'])
        for c in range(1, 11):
            ws.write(28 + yr, 1 + c, '—', f['center'])

    # ── Wires vs NWA ─────────────────────────────────────────────────────
    ws.merge_range('B50:E50', 'WIRES vs. NON-WIRES ALTERNATIVE (NWA) COMPARISON', f['section'])
    ws.write('B51', 'Traditional Wires Cost ($/kW)', f['bold'])
    ws.write('C51', 500, f['input_cur'])
    ws.write('E51', 'Cost per kW of traditional T&D infrastructure upgrade', f['tooltip'])
    ws.write('B52', 'Wires Book Life (years)', f['bold'])
    ws.write('C52', 40, f['input_int'])
    ws.write('B53', 'Wires Lead Time (years)', f['bold'])
    ws.write('C53', 5, f['input_int'])
    ws.write('B54', 'NWA Deferral Years', f['bold'])
    ws.write('C54', 5, f['input_int'])
    ws.write('B55', 'NWA Incrementality Adjustment', f['bold'])
    ws.write('C55', 'Yes', f['input'])
    ws.write('E55', 'Apply incrementality adjustment per CPUC NWA framework', f['tooltip'])

    ws.merge_range('B57:E57', 'SLICE-OF-DAY (SOD) FEASIBILITY', f['section'])
    ws.write('B58', 'Min Qualifying Hours (SOD)', f['bold'])
    ws.write('C58', 4, f['input_int'])
    ws.write('E58', 'Minimum hours of discharge to qualify for RA (typically 4h)', f['tooltip'])
    ws.write('B59', 'Deration Threshold', f['bold'])
    ws.write('C59', 0.50, f['input_pct'])
    ws.write('E59', 'Minimum capacity factor to maintain SOD qualification', f['tooltip'])
    ws.write('B61', 'SOD Feasibility Result', f['bold'])
    ws.write('C61', '← Run Python GUI for SOD check', f['warn_fmt'])


# =============================================================================
# LIBRARY DATA SHEET
# =============================================================================

def _create_library_data_sheet(ws, f) -> None:
    ws.set_column('A:A', 3)
    ws.set_column('B:B', 30)
    ws.set_column('C:E', 20)
    ws.set_column('F:F', 36)

    ws.merge_range('B2:F2', 'ASSUMPTION LIBRARIES — Comparison of NREL | Lazard | CPUC', f['title'])
    ws.merge_range('B3:F3',
        'Source data embedded from JSON libraries.  Select a library on the Inputs sheet to auto-apply.',
        f['subtitle'])

    # ── Library Overview ──────────────────────────────────────────────────
    ws.merge_range('B5:F5', 'LIBRARY OVERVIEW', f['section'])
    for c, hdr in enumerate(['Parameter', 'NREL ATB 2024\nModerate', 'Lazard LCOS\nv10.0 (2025)',
                               'CPUC California\n2024', 'Guidance']):
        ws.set_row(5, 30)
        ws.write(5, 1 + c, hdr, f['header'])

    lib_data = [
        # (parameter, NREL, Lazard, CPUC, guidance)
        ('VERSION',            '', '', '', ''),
        ('Version',            '2024.2',     '10.1',       '2024.2',     ''),
        ('Published',          '2024-04-15', '2025-03-01', '2024-11-01', ''),
        ('Source',             'NREL ATB',   'Lazard LCOS','CPUC/E3 ACC',''),
        ('',                   '', '', '', ''),
        ('TECHNOLOGY',         '', '', '', ''),
        ('Chemistry',          'LFP',        'LFP',        'LFP',        ''),
        ('Round-Trip Eff.',    '85%',        '86%',        '85%',        'AC-AC incl inverter'),
        ('Annual Degradation', '2.5%',       '2.0%',       '2.5%',       'Capacity fade per year'),
        ('Cycle Life',         '6,000',      '6,500',      '6,000',      'Full cycles before EOL'),
        ('Augmentation Year',  '12',         '12',         '12',         'Battery module replacement yr'),
        ('Cycles per Day',     '1.0',        '1.0',        '1.0',        ''),
        ('',                   '', '', '', ''),
        ('COSTS ($/kWh unless noted)', '', '', '', ''),
        ('CapEx ($/kWh)',       '$160',       '$145',       '$155',       'Installed 4-hr LFP system'),
        ('Fixed O&M ($/kW-yr)','$25',        '$22',        '$26',        'Site O&M, monitoring'),
        ('Variable O&M ($/MWh)','$0',        '$0.50',      '$0',         'Per-MWh cost'),
        ('Augmentation ($/kWh)','$55',       '$50',        '$52',        'Battery replacement cost'),
        ('Decommissioning ($/kW)','$10',     '$8',         '$12',        'End-of-life removal'),
        ('Charging Cost ($/MWh)','$30',      '$35',        '$25',        'Grid electricity for charging'),
        ('Residual Value',     '10%',        '10%',        '10%',        '% of CapEx at end'),
        ('',                   '', '', '', ''),
        ('INFRASTRUCTURE COSTS', '', '', '', ''),
        ('Interconnection ($/kW)','$100',    '$90',        '$120',       'Network upgrades, metering'),
        ('Land ($/kW)',         '$10',        '$8',         '$15',        'Site acquisition/lease'),
        ('Permitting ($/kW)',   '$15',        '$12',        '$20',        'Env review, permits'),
        ('Insurance (% CapEx)', '0.5%',      '0.5%',       '0.5%',       'Annual insurance'),
        ('Property Tax',        '1.0%',      '1.0%',       '1.05%',      'CA average 1.05%'),
        ('',                    '', '', '', ''),
        ('TAX CREDITS',         '', '', '', ''),
        ('ITC Base Rate',       '30%',       '30%',        '30%',        'IRA base rate'),
        ('ITC Adders',          '0%',        '0%',         '10%',        'Energy community (CA common)'),
        ('Total ITC',           '30%',       '30%',        '40%',        ''),
        ('',                    '', '', '', ''),
        ('BENEFIT STREAMS ($/kW-year, Year 1)', '', '', '', ''),
        ('Resource Adequacy',   '$150',      '$140',       '$180',       'CA premium; CPUC D.24-06-050'),
        ('Energy Arbitrage',    '$40',       '$45',        '$35',        'CAISO OASIS historical'),
        ('Ancillary Services',  '$15',       '$12',        '$10',        'Frequency reg, spinning res'),
        ('T&D Deferral',        '$20',       '$20',        '$25',        'CPUC Avoided Cost Calc'),
        ('Resilience Value',    '$40',       '$45',        '$60',        'LBNL ICE; CA PSPS premium'),
        ('Renewable Integration','$20',      '$20',        '$30',        'CAISO curtailment value'),
        ('GHG Emissions Value', '$15',       '$12',        '$20',        'CARB cap-and-trade'),
        ('Voltage Support',     '$10',       '$6',         '$10',        'Distribution services'),
        ('',                    '', '', '', ''),
        ('COST PROJECTIONS',    '', '', '', ''),
        ('Learning Rate',       '12%',       '10%',        '11%',        'Annual cost decline'),
        ('Base Year',           '2024',      '2025',       '2024',       ''),
        ('',                    '', '', '', ''),
        ('FINANCING (typical)', '', '', '', ''),
        ('Debt Percent',        '60%',       '65%',        '65%',        ''),
        ('Interest Rate',       '4.5%',      '5.0%',       '4.0%',       ''),
        ('Cost of Equity',      '10.0%',     '10.0%',      '9.5%',       ''),
        ('Tax Rate',            '21%',       '21%',        '21%',        ''),
    ]

    for i, row_data in enumerate(lib_data):
        r = 6 + i
        param, nrel, lazard, cpuc, guidance = row_data
        if param in ('VERSION', 'TECHNOLOGY', 'COSTS ($/kWh unless noted)',
                     'INFRASTRUCTURE COSTS', 'TAX CREDITS',
                     'BENEFIT STREAMS ($/kW-year, Year 1)',
                     'COST PROJECTIONS', 'FINANCING (typical)'):
            ws.merge_range(r, 1, r, 5, param, f['section'])
        elif param == '':
            pass
        else:
            ws.write(r, 1, param, f['bold'])
            ws.write(r, 2, nrel,    f['center'])
            ws.write(r, 3, lazard,  f['center'])
            ws.write(r, 4, cpuc,    f['center'])
            ws.write(r, 5, guidance,f['tooltip'])


# =============================================================================
# METHODOLOGY SHEET
# =============================================================================

def _create_methodology_sheet(ws, f) -> None:
    ws.set_column('A:A', 3)
    ws.set_column('B:B', 28)
    ws.set_column('C:C', 70)

    ws.merge_range('B2:C2', 'METHODOLOGY — Formulas, Citations, and Assumptions', f['title'])
    ws.merge_range('B3:C3',
        'All calculations follow CPUC Standard Practice Manual and cited industry sources.',
        f['subtitle'])

    sections = [
        ('FINANCIAL METRICS', [
            ('Net Present Value (NPV)',
             'NPV = Σ CFt / (1+r)^t  for t=0..N\n'
             'CF0 = −CapEx + ITC;  CFt = Benefits(t) − O&M(t) − Charging(t) − Insurance(t) − PropertyTax(t)\n'
             'Source: Brealey, Myers & Allen, Principles of Corporate Finance, 13th ed. McGraw-Hill 2020, Ch.2'),
            ('Benefit-Cost Ratio (BCR)',
             'BCR = PV(Benefits) / PV(Costs)\n'
             'BCR ≥ 1.5 → Approve | 1.0–1.5 → Further Study | < 1.0 → Reject\n'
             'Source: CPUC Standard Practice Manual, Economic Analysis of Demand-Side Programs, 2001'),
            ('Internal Rate of Return (IRR)',
             'IRR = r such that NPV = 0  (solved numerically via Newton-Raphson / Excel IRR())\n'
             'Source: Brealey et al., Ch.5'),
            ('Levelized Cost of Storage (LCOS)',
             'LCOS = PV(Lifetime Costs) / PV(Lifetime Energy Discharged)  [$/MWh]\n'
             'Costs include CapEx, O&M, charging, augmentation; Energy = MWh discharged per year × RTE\n'
             'Source: Lazard Levelized Cost of Storage Analysis v10.0, 2025'),
            ('Simple Payback',
             'Payback = first year t where Cumulative Net CF ≥ 0  (linear interpolation within year)\n'
             'Not time-value adjusted; use as secondary screening metric only'),
            ('Breakeven CapEx',
             'BE_CapEx = (PV_Benefits − PV_OpEx) / Total_Capacity_kWh  [$/kWh]\n'
             'Maximum CapEx where BCR = 1.0; safety margin = BE_CapEx − actual CapEx'),
        ]),
        ('COST STRUCTURE', [
            ('Year 0 Capital Costs',
             'Battery CapEx = CapEx($/kWh) × capacity(kWh) × (1 − LearningRate)^(COD − BaseYear)\n'
             'Infrastructure = (Interconnection + Land + Permitting)($/kW) × capacity(kW)\n'
             'ITC credit = Battery CapEx × (ITC_base + ITC_adders)  [applied to battery only]'),
            ('Annual Operating Costs',
             'Fixed O&M = FOM($/kW-yr) × capacity(kW)\n'
             'Variable O&M = VOM($/MWh) × annual_discharge(MWh)\n'
             'Charging = charging_cost($/MWh) × annual_charge(MWh)\n'
             'Insurance = insurance_rate × total_CapEx\n'
             'Property Tax = property_tax_rate × net_book_value(t)'),
            ('Augmentation',
             'Occurs at augmentation_year; cost adjusted for learning curve:\n'
             'Aug_Cost(t) = base_aug_cost × (1 − learning_rate)^t\n'
             'Restores capacity to nominal; subsequent degradation resets'),
            ('Battery Degradation',
             'Capacity(t) = Capacity(0) × (1 − degradation_rate)^t\n'
             'Applied to energy-based benefits and LCOS energy denominator\n'
             'Source: NREL ATB 2024; manufacturer data'),
        ]),
        ('MULTI-TRANCHE (JIT) METHODOLOGY', [
            ('Cohort-Based Cost Aggregation',
             'Each tranche i has its own COD year t_i and capacity q_i.\n'
             'CapEx_i = q_i(kWh) × c_0 × (1−λ)^(t_i − t_base)  where λ = learning rate\n'
             'Costs and benefits are summed across all online cohorts each year.'),
            ('Staged Degradation',
             'Each cohort degrades from its own COD:\n'
             'Capacity_i(t) = q_i × (1 − d)^(t − t_i)  for t > t_i\n'
             'Benefits scale by effective online capacity / total nominal capacity'),
            ('Flexibility Value',
             'FV = PV(Upfront Build) − PV(Phased JIT Build)\n'
             'Positive FV shows economic advantage of deferring capital and capturing learning curve\n'
             'Computed in Python engine; displayed in Results tab of GUI'),
        ]),
        ('BENEFIT METHODOLOGIES', [
            ('Resource Adequacy',
             'Value = RA_capacity_credit($/kW-yr) × online_MW × 1000\n'
             'Capacity credit uses ELCC methodology per CPUC rules\n'
             'Source: CPUC RA Program, D.24-06-050, November 2024'),
            ('T&D Capital Deferral',
             'PV = K × [1 − ((1+g)/(1+r))^n]\n'
             'K = deferred capital cost; g = load growth; r = discount rate; n = deferral years\n'
             'Source: E3 CPUC Avoided Cost Calculator 2024'),
            ('GHG Emissions Value',
             'Value = displaced_emissions(tCO2) × carbon_price($/tCO2)\n'
             'Carbon price per EPA Social Cost of Carbon / CARB cap-and-trade\n'
             'Source: EPA Technical Support Document — Social Cost of Greenhouse Gases, 2023'),
        ]),
        ('REFERENCES', [
            ('Primary Data Sources',
             '1. NREL Annual Technology Baseline 2024 — atb.nrel.gov/electricity/2024\n'
             '2. Lazard LCOS Analysis v10.0, 2025 — lazard.com\n'
             '3. CPUC Standard Practice Manual — cpuc.ca.gov\n'
             '4. E3 Avoided Cost Calculator 2024 — cpuc.ca.gov/acc\n'
             '5. LBNL ICE Calculator — icecalculator.com'),
            ('Regulatory References',
             '6. CPUC D.24-06-050 — Resource Adequacy 2024\n'
             '7. CPUC D.25-12-003 — SCE 2026-2028 Cost of Capital\n'
             '8. FERC Order 841 — Electric Storage Market Participation\n'
             '9. IRS Notice 2023 — IRA Investment Tax Credit Guidance'),
            ('Academic / Industry',
             '10. Brealey, Myers & Allen — Principles of Corporate Finance, 13th ed.\n'
             '11. NREL Storage Futures Study — NREL/TP-6A20-77449\n'
             '12. EPRI Grid Energy Storage — Technical Report 3002020045\n'
             '13. BloombergNEF Battery Price Survey 2025'),
        ]),
    ]

    row = 4
    for section_title, items in sections:
        ws.merge_range(row, 1, row, 2, section_title, f['section'])
        row += 1
        for label, content in items:
            ws.write(row, 1, label, f['bold'])
            ws.write(row, 2, content)
            ws.set_row(row, max(15, content.count('\n') * 15 + 15))
            row += 1
        row += 1


# =============================================================================
# VBA INSTRUCTIONS SHEET
# =============================================================================

def _create_vba_instructions_sheet(ws, f) -> None:
    ws.set_column('A:A', 3)
    ws.set_column('B:B', 30)
    ws.set_column('C:C', 65)

    ws.merge_range('B2:C2', 'VBA MACRO GUIDE — Embedded Buttons and Functions', f['title'])
    ws.merge_range('B3:C3',
        'This workbook includes VBA macros for report generation and data export. '
        'Enable macros when prompted. Requires .xlsm format.',
        f['subtitle'])

    sections = [
        ('ENABLING MACROS', [
            ('Required for full functionality',
             'When opening the file, click "Enable Content" in the yellow security bar.\n'
             'If macros are blocked by IT policy, use the Python GUI instead (python main.py).'),
            ('Macro Security Setting',
             'File → Options → Trust Center → Trust Center Settings → Macro Settings\n'
             'Select "Disable all macros with notification" for safe operation.'),
        ]),
        ('AVAILABLE MACRO BUTTONS', [
            ('Calculate Project Economics',
             'Runs the Python calculation engine via Shell() command.\n'
             'Reads all Inputs sheet values and writes results to Results, Cash_Flows sheets.\n'
             'Button location: Results sheet, top-right'),
            ('Export PDF Report',
             'Exports Results + Cash_Flows + Sensitivity to a formatted PDF.\n'
             'Output saved to same folder as workbook with timestamp in filename.\n'
             'Button location: Results sheet'),
            ('Load Library Assumptions',
             'Reads the selected library (Inputs!C6) and populates cost/tech/benefit fields.\n'
             'Overrides any manually entered values in the yellow input cells.\n'
             'Button location: Inputs sheet, row 6'),
            ('Clear All Inputs',
             'Resets all yellow input cells to default values from the selected library.\n'
             'Does NOT clear library selection or project name/ID.\n'
             'Button location: Inputs sheet, top'),
            ('Generate Sensitivity Chart',
             'Runs 8 sensitivity scenarios via Python and creates a tornado chart.\n'
             'Chart saved to Sensitivity sheet.\n'
             'Button location: Sensitivity sheet'),
            ('Export to JSON',
             'Saves all current inputs to a .json project file for use in the Python GUI.\n'
             'Enables round-trip compatibility between Excel and Python interfaces.\n'
             'Button location: Inputs sheet'),
        ]),
        ('USING WITHOUT MACROS', [
            ('Manual Workflow',
             '1. Enter project parameters in yellow cells on the Inputs sheet\n'
             '2. Cash_Flows and Results sheets update automatically via Excel formulas\n'
             '3. Run python main.py for IRR, LCOS, Payback exact calculations\n'
             '4. Copy Python results back to Results sheet manually'),
            ('Python GUI',
             'Launch:  cd bess_analyzer && python main.py\n'
             'Full feature set: JIT cohort model, UOS analysis, sensitivity, PDF export\n'
             'Save projects as JSON; load back into GUI or regenerate Excel from command line'),
            ('Regenerating This File',
             'Run:  python excel_generator.py  (from bess_analyzer/ directory)\n'
             'Output: BESS_Analyzer_v2.0_YYYYMMDD_HHMM.xlsx with current timestamp\n'
             'All formulas and sheet structure are rebuilt from source code'),
        ]),
        ('KEYBOARD SHORTCUTS', [
            ('Recalculate',      'F9 — force full workbook recalculation'),
            ('Navigate sheets',  'Ctrl+PgUp / Ctrl+PgDn — move between sheets'),
            ('Find named range', 'Ctrl+G → type range name (e.g., Build_Schedule) → Enter'),
            ('Freeze panes',     'View → Freeze Panes on any sheet for easier scrolling'),
        ]),
    ]

    row = 5
    for section_title, items in sections:
        ws.merge_range(row, 1, row, 2, section_title, f['section'])
        row += 1
        for label, content in items:
            ws.write(row, 1, label, f['bold'])
            ws.write(row, 2, content)
            ws.set_row(row, max(15, content.count('\n') * 15 + 15))
            row += 1
        row += 1


# =============================================================================
# ENTRY POINT
# =============================================================================

if __name__ == '__main__':
    timestamp   = datetime.now().strftime('%Y%m%d_%H%M')
    default_out = f'BESS_Analyzer_v2.0_{timestamp}.xlsx'
    output_file = sys.argv[1] if len(sys.argv) > 1 else default_out
    create_workbook(output_file)
