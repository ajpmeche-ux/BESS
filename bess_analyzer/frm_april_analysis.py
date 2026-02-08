#!/usr/bin/env python3
"""
FRM April Analysis - BESS Financial Performance Model
======================================================

Generates detailed multi-sheet Excel export for April Financial Review Meeting.
Models three technical scenarios across two cost variations (Turnkey vs Equipment Only).

Author: RESS Analysis Team
Date: April 2026
"""

import numpy as np
import numpy_financial as npf
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Tuple
from pathlib import Path
import xlsxwriter
from datetime import datetime


# =============================================================================
# PART 1: CONFIGURATION & COST VARIABLES
# =============================================================================

# Cost Variations ($/kWh) - Placeholders for actual data
VAR_TURNKEY_COST = 350  # $/kWh - Full EPC included (installation, commissioning, etc.)
VAR_EQUIPMENT_ONLY = 220  # $/kWh - Battery modules + inverters only

# Additional Cost Components for Equipment Only scenario
VAR_INSTALLATION_PER_KW = 80  # $/kW - Installation labor
VAR_COMMISSIONING_PER_KW = 25  # $/kW - Commissioning & testing
VAR_ENGINEERING_PCT = 0.08  # % of equipment cost - Engineering & design

# Common Financial Parameters
DISCOUNT_RATE = 0.07  # 7% WACC
ANALYSIS_PERIOD = 20  # years
DEGRADATION_RATE = 0.025  # 2.5% annual capacity degradation
ROUND_TRIP_EFFICIENCY = 0.85  # 85% AC-AC
CYCLES_PER_DAY = 1.0  # Average daily cycles
INFLATION_RATE = 0.025  # 2.5% annual escalation

# ITC Parameters (IRA)
ITC_BASE_RATE = 0.30  # 30% base ITC
ITC_ADDER_DOMESTIC = 0.10  # 10% domestic content adder (if applicable)

# Operating Costs
FOM_PER_KW_YEAR = 22  # $/kW-year Fixed O&M
VOM_PER_MWH = 0.50  # $/MWh Variable O&M
INSURANCE_PCT = 0.005  # 0.5% of CapEx annually
PROPERTY_TAX_PCT = 0.01  # 1% of asset value

# Value Stack Components ($/kW-year unless noted)
VALUE_STACK = {
    "deferral_high": 45,  # $/kW-year - High deferral value (constrained areas)
    "deferral_medium": 30,  # $/kW-year - Medium deferral value
    "deferral_low": 15,  # $/kW-year - Low deferral value
    "arbitrage_high": 55,  # $/kW-year - High arbitrage (volatile markets)
    "arbitrage_medium": 35,  # $/kW-year - Medium arbitrage
    "arbitrage_low": 20,  # $/kW-year - Low arbitrage
    "ra_per_kw_month": 12,  # $/kW-month - Resource Adequacy
    "resiliency_premium": 25,  # $/kW-year - Resiliency premium (critical loads)
    "ancillary_services": 12,  # $/kW-year - Frequency regulation, spinning reserve
    "voltage_support": 8,  # $/kW-year - VAR support, power quality
}


# =============================================================================
# PART 2: SCENARIO DEFINITIONS
# =============================================================================

@dataclass
class BESSScenario:
    """Technical scenario definition for BESS deployment."""

    name: str
    code: str  # A, B, or C
    voltage_kv: float
    power_mw: float
    energy_mwh: float
    primary_value: str
    description: str

    # Value stack configuration
    deferral_tier: str = "medium"  # high, medium, low
    arbitrage_tier: str = "medium"  # high, medium, low
    has_resiliency_premium: bool = False

    # Derived properties
    @property
    def power_kw(self) -> float:
        return self.power_mw * 1000

    @property
    def energy_kwh(self) -> float:
        return self.energy_mwh * 1000

    @property
    def duration_hours(self) -> float:
        return self.energy_mwh / self.power_mw


def create_scenarios() -> Dict[str, BESSScenario]:
    """Create the three FRM analysis scenarios."""

    scenarios = {
        "A": BESSScenario(
            name="The Urban Modular",
            code="A",
            voltage_kv=4.0,
            power_mw=0.5,
            energy_mwh=2.0,
            primary_value="Phase Balancing + Resiliency",
            description="Small-scale urban deployment for critical load support and distribution phase balancing",
            deferral_tier="high",
            arbitrage_tier="low",
            has_resiliency_premium=True,
        ),
        "B": BESSScenario(
            name="The Feeder Relief",
            code="B",
            voltage_kv=12.0,
            power_mw=2.5,
            energy_mwh=10.0,
            primary_value="Peak Shaving + Deferral",
            description="Mid-scale feeder deployment for peak load management and T&D deferral",
            deferral_tier="high",
            arbitrage_tier="medium",
            has_resiliency_premium=False,
        ),
        "C": BESSScenario(
            name="The Sub-Transmission",
            code="C",
            voltage_kv=34.5,
            power_mw=10.0,
            energy_mwh=40.0,
            primary_value="Capacity Injection + Arbitrage",
            description="Large-scale sub-transmission deployment for capacity and energy market participation",
            deferral_tier="low",
            arbitrage_tier="high",
            has_resiliency_premium=False,
        ),
    }

    return scenarios


# =============================================================================
# PART 3: COST VARIATION MODELS
# =============================================================================

@dataclass
class CostVariation:
    """Cost structure for a specific procurement approach."""

    name: str
    code: str  # "TURNKEY" or "EQUIP_ONLY"
    description: str

    # Cost components
    equipment_per_kwh: float  # $/kWh
    installation_per_kw: float = 0.0  # $/kW (only for equipment-only)
    commissioning_per_kw: float = 0.0  # $/kW
    engineering_pct: float = 0.0  # % of equipment cost

    # Flags
    includes_epc: bool = True
    itc_eligible: bool = True

    def calculate_total_capex(self, energy_kwh: float, power_kw: float) -> Dict[str, float]:
        """Calculate total CapEx breakdown for a given system size."""

        equipment_cost = self.equipment_per_kwh * energy_kwh
        installation_cost = self.installation_per_kw * power_kw
        commissioning_cost = self.commissioning_per_kw * power_kw
        engineering_cost = equipment_cost * self.engineering_pct

        total_capex = equipment_cost + installation_cost + commissioning_cost + engineering_cost

        return {
            "equipment": equipment_cost,
            "installation": installation_cost,
            "commissioning": commissioning_cost,
            "engineering": engineering_cost,
            "total": total_capex,
            "per_kwh": total_capex / energy_kwh if energy_kwh > 0 else 0,
            "per_kw": total_capex / power_kw if power_kw > 0 else 0,
        }


def create_cost_variations() -> Dict[str, CostVariation]:
    """Create the two cost variation models."""

    variations = {
        "TURNKEY": CostVariation(
            name="Turnkey (EPC Included)",
            code="TURNKEY",
            description="Full turnkey solution including engineering, procurement, construction, installation, and commissioning",
            equipment_per_kwh=VAR_TURNKEY_COST,
            installation_per_kw=0,  # Included in turnkey price
            commissioning_per_kw=0,  # Included in turnkey price
            engineering_pct=0,  # Included in turnkey price
            includes_epc=True,
            itc_eligible=True,
        ),
        "EQUIP_ONLY": CostVariation(
            name="Equipment Only",
            code="EQUIP_ONLY",
            description="Battery modules and inverters only; installation, commissioning, and engineering separate",
            equipment_per_kwh=VAR_EQUIPMENT_ONLY,
            installation_per_kw=VAR_INSTALLATION_PER_KW,
            commissioning_per_kw=VAR_COMMISSIONING_PER_KW,
            engineering_pct=VAR_ENGINEERING_PCT,
            includes_epc=False,
            itc_eligible=True,
        ),
    }

    return variations


# =============================================================================
# PART 4: VALUE STACK CALCULATION
# =============================================================================

def calculate_annual_value_stack(
    scenario: BESSScenario,
    year: int = 1,
    degradation_factor: float = 1.0
) -> Dict[str, float]:
    """
    Calculate annual value stack for a scenario.

    Returns dict with each value component and total.
    All values in $/year (not $/kW-year).
    """

    capacity_kw = scenario.power_kw * degradation_factor

    # Deferral value
    deferral_key = f"deferral_{scenario.deferral_tier}"
    deferral_value = VALUE_STACK[deferral_key] * capacity_kw

    # Arbitrage value
    arbitrage_key = f"arbitrage_{scenario.arbitrage_tier}"
    arbitrage_value = VALUE_STACK[arbitrage_key] * capacity_kw

    # Resource Adequacy ($/kW-month * 12 months * kW)
    ra_value = VALUE_STACK["ra_per_kw_month"] * 12 * capacity_kw

    # Ancillary Services
    ancillary_value = VALUE_STACK["ancillary_services"] * capacity_kw

    # Voltage Support
    voltage_value = VALUE_STACK["voltage_support"] * capacity_kw

    # Resiliency Premium (only for applicable scenarios)
    resiliency_value = 0.0
    if scenario.has_resiliency_premium:
        resiliency_value = VALUE_STACK["resiliency_premium"] * capacity_kw

    # Apply inflation escalation for years > 1
    escalation = (1 + INFLATION_RATE) ** (year - 1)

    values = {
        "deferral": deferral_value * escalation,
        "arbitrage": arbitrage_value * escalation,
        "resource_adequacy": ra_value * escalation,
        "ancillary_services": ancillary_value * escalation,
        "voltage_support": voltage_value * escalation,
        "resiliency": resiliency_value * escalation,
    }

    values["total"] = sum(values.values())

    return values


# =============================================================================
# PART 5: FINANCIAL MODEL
# =============================================================================

@dataclass
class FinancialResults:
    """Complete financial analysis results."""

    scenario: BESSScenario
    cost_variation: CostVariation

    # CapEx breakdown
    capex_breakdown: Dict[str, float] = field(default_factory=dict)
    itc_credit: float = 0.0
    net_capex: float = 0.0

    # Annual streams (lists indexed by year, 0 = Year 0)
    annual_costs: List[float] = field(default_factory=list)
    annual_benefits: List[float] = field(default_factory=list)
    annual_net: List[float] = field(default_factory=list)

    # Value stack breakdown by year
    value_stack_detail: List[Dict[str, float]] = field(default_factory=list)

    # Key metrics
    npv: float = 0.0
    bcr: float = 0.0
    irr: Optional[float] = None
    payback_years: Optional[float] = None
    lcos_per_mwh: float = 0.0

    # PV totals
    pv_costs: float = 0.0
    pv_benefits: float = 0.0


def run_financial_analysis(
    scenario: BESSScenario,
    cost_variation: CostVariation,
    discount_rate: float = DISCOUNT_RATE,
    analysis_years: int = ANALYSIS_PERIOD,
    itc_rate: float = ITC_BASE_RATE,
) -> FinancialResults:
    """
    Run complete financial analysis for a scenario/cost combination.
    """

    results = FinancialResults(scenario=scenario, cost_variation=cost_variation)

    n = analysis_years
    capacity_kw = scenario.power_kw
    capacity_kwh = scenario.energy_kwh

    # --- Calculate CapEx ---
    capex = cost_variation.calculate_total_capex(capacity_kwh, capacity_kw)
    results.capex_breakdown = capex

    # ITC credit (on ITC-eligible portion)
    if cost_variation.itc_eligible:
        results.itc_credit = capex["total"] * itc_rate

    results.net_capex = capex["total"] - results.itc_credit

    # --- Build annual cost stream ---
    annual_costs = [0.0] * (n + 1)

    # Year 0: Net CapEx
    annual_costs[0] = results.net_capex

    # Years 1-N: Operating costs
    for t in range(1, n + 1):
        degradation_factor = (1 - DEGRADATION_RATE) ** (t - 1)

        # Fixed O&M
        fom = FOM_PER_KW_YEAR * capacity_kw

        # Variable O&M (based on discharged energy)
        annual_discharge_mwh = (
            scenario.energy_mwh * CYCLES_PER_DAY * 365 * ROUND_TRIP_EFFICIENCY * degradation_factor
        )
        vom = VOM_PER_MWH * annual_discharge_mwh

        # Insurance
        insurance = capex["total"] * INSURANCE_PCT

        # Property tax (on depreciating value)
        remaining_value = capex["total"] * (1 - t / n)
        property_tax = remaining_value * PROPERTY_TAX_PCT

        annual_costs[t] = fom + vom + insurance + property_tax

    results.annual_costs = annual_costs

    # --- Build annual benefit stream ---
    annual_benefits = [0.0] * (n + 1)
    value_stack_detail = [{} for _ in range(n + 1)]

    for t in range(1, n + 1):
        degradation_factor = (1 - DEGRADATION_RATE) ** (t - 1)
        values = calculate_annual_value_stack(scenario, t, degradation_factor)
        value_stack_detail[t] = values
        annual_benefits[t] = values["total"]

    results.annual_benefits = annual_benefits
    results.value_stack_detail = value_stack_detail

    # --- Calculate net cash flows ---
    results.annual_net = [b - c for b, c in zip(annual_benefits, annual_costs)]

    # --- Calculate financial metrics ---

    # NPV
    results.npv = sum(cf / (1 + discount_rate) ** t for t, cf in enumerate(results.annual_net))

    # PV of costs and benefits
    results.pv_costs = sum(c / (1 + discount_rate) ** t for t, c in enumerate(annual_costs))
    results.pv_benefits = sum(b / (1 + discount_rate) ** t for t, b in enumerate(annual_benefits))

    # BCR
    if results.pv_costs > 0:
        results.bcr = results.pv_benefits / results.pv_costs

    # IRR
    try:
        irr = npf.irr(results.annual_net)
        if not np.isnan(irr) and -1 < irr < 1:
            results.irr = irr
    except Exception:
        pass

    # Payback period
    cumulative = 0.0
    for t, cf in enumerate(results.annual_net):
        cumulative += cf
        if cumulative >= 0 and t > 0:
            # Interpolate
            prev_cumulative = cumulative - cf
            if cf != 0:
                fraction = -prev_cumulative / cf
                results.payback_years = t - 1 + fraction
            else:
                results.payback_years = float(t)
            break

    # LCOS ($/MWh)
    total_energy_mwh = 0.0
    for t in range(1, n + 1):
        degradation_factor = (1 - DEGRADATION_RATE) ** (t - 1)
        annual_discharge = scenario.energy_mwh * CYCLES_PER_DAY * 365 * ROUND_TRIP_EFFICIENCY * degradation_factor
        total_energy_mwh += annual_discharge / (1 + discount_rate) ** t

    if total_energy_mwh > 0:
        results.lcos_per_mwh = results.pv_costs / total_energy_mwh

    return results


# =============================================================================
# PART 6: EXCEL EXPORT
# =============================================================================

def create_excel_formats(workbook: xlsxwriter.Workbook) -> Dict[str, xlsxwriter.workbook.Format]:
    """Create standard formatting for Excel output."""

    formats = {
        "title": workbook.add_format({
            "bold": True, "font_size": 16, "font_color": "#1F4E79",
            "bottom": 2, "bottom_color": "#1F4E79"
        }),
        "subtitle": workbook.add_format({
            "bold": True, "font_size": 12, "font_color": "#2E75B6"
        }),
        "section": workbook.add_format({
            "bold": True, "font_size": 11, "bg_color": "#D6DCE4",
            "border": 1, "border_color": "#8EA9DB"
        }),
        "header": workbook.add_format({
            "bold": True, "bg_color": "#4472C4", "font_color": "white",
            "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True
        }),
        "header_green": workbook.add_format({
            "bold": True, "bg_color": "#70AD47", "font_color": "white",
            "border": 1, "align": "center", "valign": "vcenter"
        }),
        "header_orange": workbook.add_format({
            "bold": True, "bg_color": "#ED7D31", "font_color": "white",
            "border": 1, "align": "center", "valign": "vcenter"
        }),
        "bold": workbook.add_format({"bold": True}),
        "currency": workbook.add_format({"num_format": "$#,##0", "align": "right"}),
        "currency_k": workbook.add_format({"num_format": "$#,##0,K", "align": "right"}),
        "currency_m": workbook.add_format({"num_format": '$#,##0.0,,"M"', "align": "right"}),
        "percent": workbook.add_format({"num_format": "0.0%", "align": "right"}),
        "percent_2": workbook.add_format({"num_format": "0.00%", "align": "right"}),
        "number": workbook.add_format({"num_format": "#,##0", "align": "right"}),
        "number_1": workbook.add_format({"num_format": "#,##0.0", "align": "right"}),
        "number_2": workbook.add_format({"num_format": "#,##0.00", "align": "right"}),
        "positive": workbook.add_format({
            "num_format": "$#,##0", "bg_color": "#C6EFCE", "font_color": "#006100"
        }),
        "negative": workbook.add_format({
            "num_format": "$#,##0", "bg_color": "#FFC7CE", "font_color": "#9C0006"
        }),
        "neutral": workbook.add_format({
            "num_format": "$#,##0", "bg_color": "#FFEB9C", "font_color": "#9C5700"
        }),
        "bcr_good": workbook.add_format({
            "num_format": "0.00", "bg_color": "#C6EFCE", "font_color": "#006100", "bold": True
        }),
        "bcr_marginal": workbook.add_format({
            "num_format": "0.00", "bg_color": "#FFEB9C", "font_color": "#9C5700", "bold": True
        }),
        "bcr_poor": workbook.add_format({
            "num_format": "0.00", "bg_color": "#FFC7CE", "font_color": "#9C0006", "bold": True
        }),
        "border": workbook.add_format({"border": 1}),
        "right_border": workbook.add_format({"right": 1, "right_color": "#8EA9DB"}),
    }

    return formats


def create_summary_sheet(
    workbook: xlsxwriter.Workbook,
    formats: Dict,
    all_results: Dict[str, Dict[str, FinancialResults]]
):
    """Create executive summary sheet."""

    ws = workbook.add_worksheet("Executive_Summary")
    ws.set_column("A:A", 3)
    ws.set_column("B:B", 28)
    ws.set_column("C:H", 16)

    row = 1

    # Title
    ws.merge_range(f"B{row}:H{row}", "FRM April 2026 - BESS Financial Analysis Summary", formats["title"])
    row += 2

    # Generation timestamp
    ws.write(f"B{row}", f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}", formats["subtitle"])
    row += 2

    # --- Scenario Overview ---
    ws.merge_range(f"B{row}:H{row}", "SCENARIO OVERVIEW", formats["section"])
    row += 1

    headers = ["Scenario", "Voltage", "Power", "Energy", "Duration", "Primary Value"]
    for col, header in enumerate(headers):
        ws.write(row, col + 1, header, formats["header"])
    row += 1

    scenarios = create_scenarios()
    for code, scenario in scenarios.items():
        ws.write(row, 1, f"{code}: {scenario.name}", formats["bold"])
        ws.write(row, 2, f"{scenario.voltage_kv} kV")
        ws.write(row, 3, f"{scenario.power_mw} MW")
        ws.write(row, 4, f"{scenario.energy_mwh} MWh")
        ws.write(row, 5, f"{scenario.duration_hours} hr")
        ws.write(row, 6, scenario.primary_value)
        row += 1

    row += 2

    # --- Key Metrics Comparison ---
    ws.merge_range(f"B{row}:H{row}", "KEY FINANCIAL METRICS BY SCENARIO & COST VARIATION", formats["section"])
    row += 1

    # Headers
    ws.write(row, 1, "Metric", formats["header"])
    col = 2
    for scenario_code in ["A", "B", "C"]:
        ws.merge_range(row, col, row, col + 1, f"Scenario {scenario_code}", formats["header"])
        col += 2
    row += 1

    ws.write(row, 1, "", formats["header"])
    col = 2
    for _ in range(3):
        ws.write(row, col, "Turnkey", formats["header_green"])
        ws.write(row, col + 1, "Equip Only", formats["header_orange"])
        col += 2
    row += 1

    # Metrics rows
    metrics = [
        ("Total CapEx", lambda r: r.capex_breakdown["total"], formats["currency"]),
        ("Net CapEx (after ITC)", lambda r: r.net_capex, formats["currency"]),
        ("NPV", lambda r: r.npv, formats["currency"]),
        ("BCR", lambda r: r.bcr, None),  # Special formatting
        ("IRR", lambda r: r.irr if r.irr else 0, formats["percent_2"]),
        ("Payback (years)", lambda r: r.payback_years if r.payback_years else 0, formats["number_1"]),
        ("LCOS ($/MWh)", lambda r: r.lcos_per_mwh, formats["currency"]),
        ("Year 1 Revenue", lambda r: r.annual_benefits[1] if len(r.annual_benefits) > 1 else 0, formats["currency"]),
        ("Year 1 Costs", lambda r: r.annual_costs[1] if len(r.annual_costs) > 1 else 0, formats["currency"]),
    ]

    for metric_name, getter, fmt in metrics:
        ws.write(row, 1, metric_name, formats["bold"])
        col = 2
        for scenario_code in ["A", "B", "C"]:
            for var_code in ["TURNKEY", "EQUIP_ONLY"]:
                result = all_results[scenario_code][var_code]
                value = getter(result)

                # Special BCR formatting
                if metric_name == "BCR":
                    if value >= 1.5:
                        ws.write(row, col, value, formats["bcr_good"])
                    elif value >= 1.0:
                        ws.write(row, col, value, formats["bcr_marginal"])
                    else:
                        ws.write(row, col, value, formats["bcr_poor"])
                elif fmt:
                    ws.write(row, col, value, fmt)
                else:
                    ws.write(row, col, value)
                col += 1
        row += 1

    row += 2

    # --- Assumptions Box ---
    ws.merge_range(f"B{row}:E{row}", "KEY ASSUMPTIONS", formats["section"])
    row += 1

    assumptions = [
        ("Discount Rate (WACC)", f"{DISCOUNT_RATE:.1%}"),
        ("Analysis Period", f"{ANALYSIS_PERIOD} years"),
        ("ITC Rate", f"{ITC_BASE_RATE:.0%}"),
        ("Round-Trip Efficiency", f"{ROUND_TRIP_EFFICIENCY:.0%}"),
        ("Annual Degradation", f"{DEGRADATION_RATE:.1%}"),
        ("Cycles per Day", f"{CYCLES_PER_DAY:.1f}"),
        ("Turnkey Cost", f"${VAR_TURNKEY_COST}/kWh"),
        ("Equipment Only Cost", f"${VAR_EQUIPMENT_ONLY}/kWh"),
    ]

    for label, value in assumptions:
        ws.write(row, 1, label, formats["bold"])
        ws.write(row, 2, value)
        row += 1


def create_scenario_detail_sheet(
    workbook: xlsxwriter.Workbook,
    formats: Dict,
    scenario: BESSScenario,
    results_turnkey: FinancialResults,
    results_equip: FinancialResults
):
    """Create detailed sheet for a single scenario."""

    ws = workbook.add_worksheet(f"Scenario_{scenario.code}")
    ws.set_column("A:A", 3)
    ws.set_column("B:B", 30)
    ws.set_column("C:F", 18)

    row = 1

    # Title
    ws.merge_range(f"B{row}:F{row}", f"Scenario {scenario.code}: {scenario.name}", formats["title"])
    row += 2

    # Description
    ws.write(f"B{row}", scenario.description)
    row += 2

    # --- Technical Specifications ---
    ws.merge_range(f"B{row}:D{row}", "TECHNICAL SPECIFICATIONS", formats["section"])
    row += 1

    specs = [
        ("Voltage Class", f"{scenario.voltage_kv} kV"),
        ("Power Rating", f"{scenario.power_mw} MW ({scenario.power_kw:,.0f} kW)"),
        ("Energy Capacity", f"{scenario.energy_mwh} MWh ({scenario.energy_kwh:,.0f} kWh)"),
        ("Duration", f"{scenario.duration_hours} hours"),
        ("Primary Value Stream", scenario.primary_value),
    ]

    for label, value in specs:
        ws.write(row, 1, label, formats["bold"])
        ws.write(row, 2, value)
        row += 1

    row += 1

    # --- CapEx Comparison ---
    ws.merge_range(f"B{row}:D{row}", "CAPITAL COST COMPARISON", formats["section"])
    row += 1

    ws.write(row, 1, "Component", formats["header"])
    ws.write(row, 2, "Turnkey", formats["header_green"])
    ws.write(row, 3, "Equipment Only", formats["header_orange"])
    row += 1

    capex_items = [
        ("Equipment/Base Cost", "equipment", "equipment"),
        ("Installation", "installation", "installation"),
        ("Commissioning", "commissioning", "commissioning"),
        ("Engineering", "engineering", "engineering"),
        ("Total CapEx", "total", "total"),
        ("ITC Credit", None, None),  # Special handling
        ("Net CapEx", None, None),  # Special handling
    ]

    for label, key_t, key_e in capex_items:
        ws.write(row, 1, label, formats["bold"])
        if label == "ITC Credit":
            ws.write(row, 2, -results_turnkey.itc_credit, formats["currency"])
            ws.write(row, 3, -results_equip.itc_credit, formats["currency"])
        elif label == "Net CapEx":
            ws.write(row, 2, results_turnkey.net_capex, formats["currency"])
            ws.write(row, 3, results_equip.net_capex, formats["currency"])
        else:
            ws.write(row, 2, results_turnkey.capex_breakdown.get(key_t, 0), formats["currency"])
            ws.write(row, 3, results_equip.capex_breakdown.get(key_e, 0), formats["currency"])
        row += 1

    # Per-unit costs
    row += 1
    ws.write(row, 1, "Cost per kWh", formats["bold"])
    ws.write(row, 2, results_turnkey.capex_breakdown["per_kwh"], formats["currency"])
    ws.write(row, 3, results_equip.capex_breakdown["per_kwh"], formats["currency"])
    row += 1
    ws.write(row, 1, "Cost per kW", formats["bold"])
    ws.write(row, 2, results_turnkey.capex_breakdown["per_kw"], formats["currency"])
    ws.write(row, 3, results_equip.capex_breakdown["per_kw"], formats["currency"])
    row += 2

    # --- Value Stack ---
    ws.merge_range(f"B{row}:D{row}", "ANNUAL VALUE STACK (YEAR 1)", formats["section"])
    row += 1

    ws.write(row, 1, "Revenue Stream", formats["header"])
    ws.write(row, 2, "$/kW-year", formats["header"])
    ws.write(row, 3, "Annual Value", formats["header"])
    row += 1

    year1_values = results_turnkey.value_stack_detail[1]
    value_items = [
        ("T&D Deferral", "deferral"),
        ("Energy Arbitrage", "arbitrage"),
        ("Resource Adequacy", "resource_adequacy"),
        ("Ancillary Services", "ancillary_services"),
        ("Voltage Support", "voltage_support"),
        ("Resiliency Premium", "resiliency"),
    ]

    for label, key in value_items:
        annual_value = year1_values.get(key, 0)
        per_kw = annual_value / scenario.power_kw if scenario.power_kw > 0 else 0
        ws.write(row, 1, label)
        ws.write(row, 2, per_kw, formats["currency"])
        ws.write(row, 3, annual_value, formats["currency"])
        row += 1

    # Total
    ws.write(row, 1, "TOTAL", formats["bold"])
    total_per_kw = year1_values["total"] / scenario.power_kw if scenario.power_kw > 0 else 0
    ws.write(row, 2, total_per_kw, formats["currency"])
    ws.write(row, 3, year1_values["total"], formats["currency"])
    row += 2

    # --- Financial Metrics ---
    ws.merge_range(f"B{row}:D{row}", "FINANCIAL METRICS COMPARISON", formats["section"])
    row += 1

    ws.write(row, 1, "Metric", formats["header"])
    ws.write(row, 2, "Turnkey", formats["header_green"])
    ws.write(row, 3, "Equipment Only", formats["header_orange"])
    row += 1

    fin_metrics = [
        ("Net Present Value", results_turnkey.npv, results_equip.npv, formats["currency"]),
        ("Benefit-Cost Ratio", results_turnkey.bcr, results_equip.bcr, formats["number_2"]),
        ("Internal Rate of Return", results_turnkey.irr or 0, results_equip.irr or 0, formats["percent_2"]),
        ("Payback Period (years)", results_turnkey.payback_years or 0, results_equip.payback_years or 0, formats["number_1"]),
        ("LCOS ($/MWh)", results_turnkey.lcos_per_mwh, results_equip.lcos_per_mwh, formats["currency"]),
        ("PV of Benefits", results_turnkey.pv_benefits, results_equip.pv_benefits, formats["currency"]),
        ("PV of Costs", results_turnkey.pv_costs, results_equip.pv_costs, formats["currency"]),
    ]

    for label, val_t, val_e, fmt in fin_metrics:
        ws.write(row, 1, label, formats["bold"])
        ws.write(row, 2, val_t, fmt)
        ws.write(row, 3, val_e, fmt)
        row += 1


def create_cashflow_sheet(
    workbook: xlsxwriter.Workbook,
    formats: Dict,
    all_results: Dict[str, Dict[str, FinancialResults]]
):
    """Create detailed cash flow sheet."""

    ws = workbook.add_worksheet("Cash_Flows")
    ws.set_column("A:A", 3)
    ws.set_column("B:B", 12)
    ws.set_column("C:Z", 14)

    row = 1

    ws.merge_range("B1:Z1", "Annual Cash Flow Analysis - All Scenarios", formats["title"])
    row = 3

    for scenario_code in ["A", "B", "C"]:
        for var_code in ["TURNKEY", "EQUIP_ONLY"]:
            result = all_results[scenario_code][var_code]
            scenario = result.scenario
            variation = result.cost_variation

            # Section header
            label = f"Scenario {scenario_code} - {variation.name}"
            ws.merge_range(f"B{row}:L{row}", label, formats["section"])
            row += 1

            # Headers
            ws.write(row, 1, "Year", formats["header"])
            ws.write(row, 2, "Costs", formats["header"])
            ws.write(row, 3, "Benefits", formats["header"])
            ws.write(row, 4, "Net CF", formats["header"])
            ws.write(row, 5, "Cumulative", formats["header"])
            row += 1

            # Data rows (show first 10 years + final year)
            cumulative = 0.0
            years_to_show = list(range(11)) + [ANALYSIS_PERIOD]

            for t in years_to_show:
                if t > ANALYSIS_PERIOD:
                    continue
                ws.write(row, 1, t)
                ws.write(row, 2, result.annual_costs[t], formats["currency"])
                ws.write(row, 3, result.annual_benefits[t], formats["currency"])
                ws.write(row, 4, result.annual_net[t], formats["currency"])
                cumulative += result.annual_net[t]

                # Conditional formatting for cumulative
                if cumulative >= 0:
                    ws.write(row, 5, cumulative, formats["positive"])
                else:
                    ws.write(row, 5, cumulative, formats["negative"])
                row += 1

            row += 2


def create_value_stack_sheet(
    workbook: xlsxwriter.Workbook,
    formats: Dict,
    scenarios: Dict[str, BESSScenario]
):
    """Create value stack assumptions sheet."""

    ws = workbook.add_worksheet("Value_Stack")
    ws.set_column("A:A", 3)
    ws.set_column("B:B", 25)
    ws.set_column("C:F", 16)

    row = 1

    ws.merge_range("B1:F1", "Value Stack Components & Assumptions", formats["title"])
    row = 3

    # Value stack parameters
    ws.merge_range(f"B{row}:D{row}", "BASE VALUE ASSUMPTIONS ($/kW-year)", formats["section"])
    row += 1

    ws.write(row, 1, "Component", formats["header"])
    ws.write(row, 2, "Low", formats["header"])
    ws.write(row, 3, "Medium", formats["header"])
    ws.write(row, 4, "High", formats["header"])
    row += 1

    ws.write(row, 1, "T&D Deferral", formats["bold"])
    ws.write(row, 2, VALUE_STACK["deferral_low"], formats["currency"])
    ws.write(row, 3, VALUE_STACK["deferral_medium"], formats["currency"])
    ws.write(row, 4, VALUE_STACK["deferral_high"], formats["currency"])
    row += 1

    ws.write(row, 1, "Energy Arbitrage", formats["bold"])
    ws.write(row, 2, VALUE_STACK["arbitrage_low"], formats["currency"])
    ws.write(row, 3, VALUE_STACK["arbitrage_medium"], formats["currency"])
    ws.write(row, 4, VALUE_STACK["arbitrage_high"], formats["currency"])
    row += 1

    ws.write(row, 1, "Resource Adequacy", formats["bold"])
    ws.write(row, 2, f"${VALUE_STACK['ra_per_kw_month']}/kW-mo")
    ws.write(row, 3, f"= ${VALUE_STACK['ra_per_kw_month'] * 12}/kW-yr")
    row += 1

    ws.write(row, 1, "Ancillary Services", formats["bold"])
    ws.write(row, 2, VALUE_STACK["ancillary_services"], formats["currency"])
    row += 1

    ws.write(row, 1, "Voltage Support", formats["bold"])
    ws.write(row, 2, VALUE_STACK["voltage_support"], formats["currency"])
    row += 1

    ws.write(row, 1, "Resiliency Premium", formats["bold"])
    ws.write(row, 2, VALUE_STACK["resiliency_premium"], formats["currency"])
    ws.write(row, 3, "(Scenario A only)")
    row += 2

    # Scenario value assignments
    ws.merge_range(f"B{row}:E{row}", "SCENARIO VALUE TIER ASSIGNMENTS", formats["section"])
    row += 1

    ws.write(row, 1, "Scenario", formats["header"])
    ws.write(row, 2, "Deferral Tier", formats["header"])
    ws.write(row, 3, "Arbitrage Tier", formats["header"])
    ws.write(row, 4, "Resiliency", formats["header"])
    row += 1

    for code, scenario in scenarios.items():
        ws.write(row, 1, f"{code}: {scenario.name}", formats["bold"])
        ws.write(row, 2, scenario.deferral_tier.title())
        ws.write(row, 3, scenario.arbitrage_tier.title())
        ws.write(row, 4, "Yes" if scenario.has_resiliency_premium else "No")
        row += 1


def create_cost_inputs_sheet(
    workbook: xlsxwriter.Workbook,
    formats: Dict,
    variations: Dict[str, CostVariation]
):
    """Create cost inputs reference sheet."""

    ws = workbook.add_worksheet("Cost_Inputs")
    ws.set_column("A:A", 3)
    ws.set_column("B:B", 30)
    ws.set_column("C:D", 18)

    row = 1

    ws.merge_range("B1:D1", "Cost Input Assumptions", formats["title"])
    row = 3

    # Cost variation comparison
    ws.merge_range(f"B{row}:D{row}", "COST VARIATION COMPARISON", formats["section"])
    row += 1

    ws.write(row, 1, "Parameter", formats["header"])
    ws.write(row, 2, "Turnkey", formats["header_green"])
    ws.write(row, 3, "Equipment Only", formats["header_orange"])
    row += 1

    turnkey = variations["TURNKEY"]
    equip = variations["EQUIP_ONLY"]

    cost_params = [
        ("Base Equipment ($/kWh)", turnkey.equipment_per_kwh, equip.equipment_per_kwh),
        ("Installation ($/kW)", "Included", equip.installation_per_kw),
        ("Commissioning ($/kW)", "Included", equip.commissioning_per_kw),
        ("Engineering (%)", "Included", f"{equip.engineering_pct:.0%}"),
        ("Includes EPC", "Yes", "No"),
        ("ITC Eligible", "Yes", "Yes"),
    ]

    for label, val_t, val_e in cost_params:
        ws.write(row, 1, label, formats["bold"])
        if isinstance(val_t, (int, float)):
            ws.write(row, 2, val_t, formats["currency"])
        else:
            ws.write(row, 2, val_t)
        if isinstance(val_e, (int, float)):
            ws.write(row, 3, val_e, formats["currency"])
        else:
            ws.write(row, 3, val_e)
        row += 1

    row += 2

    # Operating cost assumptions
    ws.merge_range(f"B{row}:C{row}", "OPERATING COST ASSUMPTIONS", formats["section"])
    row += 1

    op_costs = [
        ("Fixed O&M", f"${FOM_PER_KW_YEAR}/kW-year"),
        ("Variable O&M", f"${VOM_PER_MWH}/MWh"),
        ("Insurance", f"{INSURANCE_PCT:.1%} of CapEx"),
        ("Property Tax", f"{PROPERTY_TAX_PCT:.1%} of asset value"),
    ]

    for label, value in op_costs:
        ws.write(row, 1, label, formats["bold"])
        ws.write(row, 2, value)
        row += 1

    row += 2

    # Financial parameters
    ws.merge_range(f"B{row}:C{row}", "FINANCIAL PARAMETERS", formats["section"])
    row += 1

    fin_params = [
        ("Discount Rate (WACC)", f"{DISCOUNT_RATE:.1%}"),
        ("Analysis Period", f"{ANALYSIS_PERIOD} years"),
        ("ITC Base Rate", f"{ITC_BASE_RATE:.0%}"),
        ("Domestic Content Adder", f"{ITC_ADDER_DOMESTIC:.0%} (if applicable)"),
        ("Inflation Rate", f"{INFLATION_RATE:.1%}"),
    ]

    for label, value in fin_params:
        ws.write(row, 1, label, formats["bold"])
        ws.write(row, 2, value)
        row += 1

    row += 2

    # Technical parameters
    ws.merge_range(f"B{row}:C{row}", "TECHNICAL PARAMETERS", formats["section"])
    row += 1

    tech_params = [
        ("Round-Trip Efficiency", f"{ROUND_TRIP_EFFICIENCY:.0%}"),
        ("Annual Degradation", f"{DEGRADATION_RATE:.1%}"),
        ("Cycles per Day", f"{CYCLES_PER_DAY:.1f}"),
        ("Duration", "4 hours (all scenarios)"),
    ]

    for label, value in tech_params:
        ws.write(row, 1, label, formats["bold"])
        ws.write(row, 2, value)
        row += 1


def generate_frm_excel(output_path: str = "FRM_April_Analysis.xlsx"):
    """
    Main function to generate the complete FRM Excel analysis.
    """

    print("=" * 60)
    print("FRM April 2026 - BESS Financial Analysis")
    print("=" * 60)

    # Create scenarios and cost variations
    scenarios = create_scenarios()
    variations = create_cost_variations()

    print(f"\nAnalyzing {len(scenarios)} scenarios x {len(variations)} cost variations...")

    # Run all analyses
    all_results: Dict[str, Dict[str, FinancialResults]] = {}

    for scenario_code, scenario in scenarios.items():
        all_results[scenario_code] = {}

        for var_code, variation in variations.items():
            print(f"  Running Scenario {scenario_code} ({scenario.name}) - {variation.name}...")

            result = run_financial_analysis(scenario, variation)
            all_results[scenario_code][var_code] = result

            print(f"    NPV: ${result.npv:,.0f} | BCR: {result.bcr:.2f} | IRR: {result.irr:.1%}" if result.irr else f"    NPV: ${result.npv:,.0f} | BCR: {result.bcr:.2f}")

    # Create Excel workbook
    print(f"\nGenerating Excel workbook: {output_path}")

    workbook = xlsxwriter.Workbook(output_path)
    formats = create_excel_formats(workbook)

    # Create all sheets
    create_summary_sheet(workbook, formats, all_results)

    for scenario_code, scenario in scenarios.items():
        create_scenario_detail_sheet(
            workbook, formats, scenario,
            all_results[scenario_code]["TURNKEY"],
            all_results[scenario_code]["EQUIP_ONLY"]
        )

    create_cashflow_sheet(workbook, formats, all_results)
    create_value_stack_sheet(workbook, formats, scenarios)
    create_cost_inputs_sheet(workbook, formats, variations)

    workbook.close()

    print(f"\nExcel file created: {output_path}")
    print(f"  - Executive_Summary: High-level comparison")
    print(f"  - Scenario_A/B/C: Detailed analysis per scenario")
    print(f"  - Cash_Flows: Annual cash flow projections")
    print(f"  - Value_Stack: Revenue assumptions")
    print(f"  - Cost_Inputs: Cost parameter reference")
    print("\nDone!")

    return all_results


# =============================================================================
# MAIN ENTRY POINT
# =============================================================================

if __name__ == "__main__":
    import sys

    # Default output path
    output_file = "FRM_April_Analysis.xlsx"

    # Allow custom output path via command line
    if len(sys.argv) > 1:
        output_file = sys.argv[1]

    # Run analysis
    results = generate_frm_excel(output_file)

    # Print summary table
    print("\n" + "=" * 80)
    print("SUMMARY: NPV Comparison ($ millions)")
    print("=" * 80)
    print(f"{'Scenario':<35} {'Turnkey':>18} {'Equipment Only':>18}")
    print("-" * 80)

    for code in ["A", "B", "C"]:
        scenario = results[code]["TURNKEY"].scenario
        npv_t = results[code]["TURNKEY"].npv / 1_000_000
        npv_e = results[code]["EQUIP_ONLY"].npv / 1_000_000
        print(f"{code}: {scenario.name:<30} ${npv_t:>14,.2f}M  ${npv_e:>14,.2f}M")

    print("=" * 80)
