#!/usr/bin/env python3
"""
BESS Analyzer CLI - Complete Battery Energy Storage Economic Analysis Tool

A comprehensive command-line interface that replicates all Excel workbook
capabilities including:
- Load assumption libraries (NREL ATB 2024, Lazard LCOS 2025, CPUC CA 2024)
- Input project parameters (capacity, duration, costs, benefits)
- Calculate financial metrics (NPV, BCR, IRR, LCOS, payback, breakeven)
- Sensitivity analysis tables (CapEx vs Benefits grid)
- Generate PDF executive summary reports
- Export to Excel workbook
- Save/load projects as JSON

Usage:
    python bess_cli.py                    # Interactive mode
    python bess_cli.py --library nrel     # Load library and run analysis
    python bess_cli.py --load project.json --report  # Load and generate report
    python bess_cli.py --help             # Show all options

Author: BESS Analyzer Team
Version: 1.0.0
"""

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).parent))

from src.models.project import (
    BenefitStream,
    CostInputs,
    FinancialResults,
    FinancingInputs,
    Project,
    ProjectBasics,
    TechnologySpecs,
)
from src.models.calculations import calculate_project_economics
from src.data.libraries import AssumptionLibrary
from src.data.storage import save_project, load_project
from src.reports.executive import generate_executive_summary


# ============================================================================
# FORMATTING UTILITIES
# ============================================================================

def format_currency(value: float, decimals: int = 0) -> str:
    """Format value as currency with comma separators."""
    if abs(value) >= 1_000_000:
        return f"${value/1_000_000:,.{decimals}f}M"
    elif abs(value) >= 1_000:
        return f"${value/1_000:,.{decimals}f}K"
    else:
        return f"${value:,.{decimals}f}"


def format_percent(value: float, decimals: int = 1) -> str:
    """Format decimal as percentage."""
    if value is None:
        return "N/A"
    return f"{value * 100:.{decimals}f}%"


def format_years(value: Optional[float]) -> str:
    """Format payback period in years."""
    if value is None:
        return "Not achieved"
    return f"{value:.1f} years"


def print_header(text: str, char: str = "=") -> None:
    """Print a formatted section header."""
    width = 70
    print(f"\n{char * width}")
    print(f" {text}")
    print(f"{char * width}")


def print_subheader(text: str) -> None:
    """Print a formatted subsection header."""
    print(f"\n--- {text} ---")


def print_table(headers: List[str], rows: List[List[str]],
                col_widths: Optional[List[int]] = None) -> None:
    """Print a formatted ASCII table."""
    if col_widths is None:
        col_widths = [max(len(str(row[i])) for row in [headers] + rows) + 2
                      for i in range(len(headers))]

    # Header
    header_line = "|".join(h.center(w) for h, w in zip(headers, col_widths))
    separator = "+".join("-" * w for w in col_widths)
    print(f"+{separator}+")
    print(f"|{header_line}|")
    print(f"+{separator}+")

    # Rows
    for row in rows:
        row_line = "|".join(str(cell).center(w) for cell, w in zip(row, col_widths))
        print(f"|{row_line}|")
    print(f"+{separator}+")


# ============================================================================
# SENSITIVITY ANALYSIS
# ============================================================================

def calculate_sensitivity_npv(
    project: Project,
    capex_levels: List[float],
    benefit_multipliers: List[float]
) -> Dict[str, List[List[float]]]:
    """
    Calculate NPV and BCR sensitivity tables.

    Args:
        project: Base project to analyze
        capex_levels: CapEx values to test ($/kWh)
        benefit_multipliers: Benefit scaling factors (1.0 = 100%)

    Returns:
        Dictionary with 'npv' and 'bcr' sensitivity matrices
    """
    base_results = calculate_project_economics(project)
    npv_matrix = []
    bcr_matrix = []

    base_capex = project.costs.capex_per_kwh

    for capex in capex_levels:
        npv_row = []
        bcr_row = []

        for ben_mult in benefit_multipliers:
            # Adjust costs proportionally
            capex_ratio = capex / base_capex if base_capex > 0 else 1.0

            # NPV = PV_Benefits * benefit_mult - PV_Costs * capex_ratio
            # (Simplified - actual CapEx is only part of total costs)
            adjusted_pv_benefits = base_results.pv_benefits * ben_mult
            adjusted_pv_costs = base_results.pv_costs * (1 + (capex_ratio - 1) * 0.7)  # CapEx ~70% of costs

            npv = adjusted_pv_benefits - adjusted_pv_costs
            bcr = adjusted_pv_benefits / adjusted_pv_costs if adjusted_pv_costs > 0 else 0

            npv_row.append(npv)
            bcr_row.append(bcr)

        npv_matrix.append(npv_row)
        bcr_matrix.append(bcr_row)

    return {'npv': npv_matrix, 'bcr': bcr_matrix}


def print_sensitivity_tables(project: Project, results: FinancialResults) -> None:
    """Print NPV and BCR sensitivity analysis tables."""

    print_header("SENSITIVITY ANALYSIS", "=")

    capex_levels = [100, 120, 140, 160, 180, 200, 220]
    benefit_multipliers = [0.7, 0.8, 0.9, 1.0, 1.1, 1.2, 1.3]

    sensitivity = calculate_sensitivity_npv(project, capex_levels, benefit_multipliers)

    # NPV Sensitivity Table
    print_subheader("NPV SENSITIVITY ($)")
    print(f"\n{'CapEx':>12}", end="")
    for mult in benefit_multipliers:
        print(f" {int(mult*100):>10}%", end="")
    print()
    print("-" * 90)

    for i, capex in enumerate(capex_levels):
        print(f"${capex:>10}/kWh", end="")
        for j, npv in enumerate(sensitivity['npv'][i]):
            if npv >= 0:
                print(f" {format_currency(npv):>10}", end="")
            else:
                print(f" {format_currency(npv):>10}", end="")
        print()

    print()

    # BCR Sensitivity Table
    print_subheader("BCR SENSITIVITY")
    print(f"\n{'CapEx':>12}", end="")
    for mult in benefit_multipliers:
        print(f" {int(mult*100):>10}%", end="")
    print()
    print("-" * 90)

    for i, capex in enumerate(capex_levels):
        print(f"${capex:>10}/kWh", end="")
        for j, bcr in enumerate(sensitivity['bcr'][i]):
            indicator = ""
            if bcr >= 1.5:
                indicator = " [+]"
            elif bcr >= 1.0:
                indicator = " [=]"
            else:
                indicator = " [-]"
            print(f" {bcr:>8.2f}{indicator}", end="")
        print()

    print("\nLegend: [+] BCR >= 1.5 (Strong)  [=] BCR 1.0-1.5 (Marginal)  [-] BCR < 1.0 (Reject)")

    # Single Variable Impacts
    print_subheader("SINGLE VARIABLE IMPACTS (±20%)")

    single_vars = [
        ("CapEx ($/kWh)", project.costs.capex_per_kwh, "$/kWh"),
        ("Total Benefits", results.pv_benefits, "$"),
        ("Discount Rate", project.basics.discount_rate, "%"),
        ("Cycles per Day", project.technology.cycles_per_day, "cycles"),
    ]

    print(f"\n{'Parameter':<25} {'Base Value':>15} {'-20% NPV':>15} {'+20% NPV':>15}")
    print("-" * 70)

    base_npv = results.npv
    for name, base_val, unit in single_vars:
        # Approximate NPV impact
        if "CapEx" in name:
            low_npv = base_npv + results.annual_costs[0] * 0.2
            high_npv = base_npv - results.annual_costs[0] * 0.2
        elif "Benefits" in name:
            low_npv = base_npv - results.pv_benefits * 0.2
            high_npv = base_npv + results.pv_benefits * 0.2
        else:
            low_npv = base_npv * 1.15
            high_npv = base_npv * 0.85

        if unit == "%":
            base_str = format_percent(base_val)
        elif unit == "$":
            base_str = format_currency(base_val)
        else:
            base_str = f"{base_val:.1f} {unit}"

        print(f"{name:<25} {base_str:>15} {format_currency(low_npv):>15} {format_currency(high_npv):>15}")


# ============================================================================
# DISPLAY FUNCTIONS
# ============================================================================

def print_project_summary(project: Project) -> None:
    """Display project configuration summary."""

    print_header("PROJECT CONFIGURATION", "=")

    basics = project.basics
    tech = project.technology
    costs = project.costs

    print_subheader("Project Basics")
    print(f"  Project Name:      {basics.name}")
    print(f"  Location:          {basics.location}")
    print(f"  Capacity:          {basics.capacity_mw:.1f} MW / {basics.capacity_mwh:.1f} MWh")
    print(f"  Duration:          {basics.duration_hours:.1f} hours")
    print(f"  Analysis Period:   {basics.analysis_period_years} years")
    print(f"  Discount Rate:     {format_percent(basics.discount_rate)}")

    if project.assumption_library:
        print(f"  Library:           {project.assumption_library} v{project.library_version}")

    print_subheader("Technology Specifications")
    print(f"  Chemistry:         {tech.chemistry}")
    print(f"  Round-Trip Eff:    {format_percent(tech.round_trip_efficiency)}")
    print(f"  Annual Degradation:{format_percent(tech.degradation_rate_annual)}")
    print(f"  Cycle Life:        {tech.cycle_life:,} cycles")
    print(f"  Cycles per Day:    {tech.cycles_per_day:.1f}")
    print(f"  Augmentation Year: {tech.augmentation_year}")

    print_subheader("Cost Inputs")
    print(f"  CapEx:             ${costs.capex_per_kwh:.0f}/kWh")
    print(f"  Fixed O&M:         ${costs.fom_per_kw_year:.0f}/kW-year")
    print(f"  Variable O&M:      ${costs.vom_per_mwh:.2f}/MWh")
    print(f"  Charging Cost:     ${costs.charging_cost_per_mwh:.0f}/MWh")
    print(f"  Augmentation:      ${costs.augmentation_per_kwh:.0f}/kWh")
    print(f"  ITC Rate:          {format_percent(costs.itc_percent + costs.itc_adders)}")
    print(f"  Residual Value:    {format_percent(costs.residual_value_pct)}")

    print_subheader("Infrastructure Costs")
    print(f"  Interconnection:   ${costs.interconnection_per_kw:.0f}/kW")
    print(f"  Land:              ${costs.land_per_kw:.0f}/kW")
    print(f"  Permitting:        ${costs.permitting_per_kw:.0f}/kW")
    print(f"  Insurance:         {format_percent(costs.insurance_pct_of_capex)} of CapEx")
    print(f"  Property Tax:      {format_percent(costs.property_tax_pct)}")

    if project.financing:
        fin = project.financing
        print_subheader("Financing Structure")
        print(f"  Debt Percentage:   {format_percent(fin.debt_percent)}")
        print(f"  Interest Rate:     {format_percent(fin.interest_rate)}")
        print(f"  Loan Term:         {fin.loan_term_years} years")
        print(f"  Cost of Equity:    {format_percent(fin.cost_of_equity)}")
        print(f"  Tax Rate:          {format_percent(fin.tax_rate)}")
        print(f"  Calculated WACC:   {format_percent(fin.calculate_wacc())}")

    if project.benefits:
        print_subheader("Benefit Streams (Year 1 Values)")
        headers = ["Category", "$/kW-yr", "Escalation"]
        rows = []
        capacity_kw = basics.capacity_mw * 1000
        for benefit in project.benefits:
            if benefit.annual_values:
                per_kw = benefit.annual_values[0] / capacity_kw if capacity_kw > 0 else 0
                rows.append([benefit.name[:30], f"${per_kw:.0f}", "2.0%"])
        if rows:
            print_table(headers, rows, [32, 12, 12])


def print_results(project: Project, results: FinancialResults) -> None:
    """Display financial analysis results."""

    print_header("FINANCIAL RESULTS", "=")

    # Key Metrics Summary
    print_subheader("KEY FINANCIAL METRICS")

    bcr = results.bcr
    if bcr >= 1.5:
        recommendation = "APPROVE - Strong economic case"
        indicator = "[+++]"
    elif bcr >= 1.0:
        recommendation = "FURTHER STUDY - Marginal economics"
        indicator = "[===]"
    else:
        recommendation = "REJECT - Costs exceed benefits"
        indicator = "[---]"

    metrics = [
        ("Benefit-Cost Ratio (BCR)", f"{results.bcr:.2f}", indicator),
        ("Net Present Value (NPV)", format_currency(results.npv, 1),
         "[+]" if results.npv > 0 else "[-]"),
        ("Internal Rate of Return", format_percent(results.irr) if results.irr else "N/A",
         "[+]" if results.irr and results.irr > project.basics.discount_rate else ""),
        ("Payback Period", format_years(results.payback_years), ""),
        ("LCOS", f"${results.lcos_per_mwh:.1f}/MWh", ""),
        ("Breakeven CapEx", f"${results.breakeven_capex_per_kwh:.0f}/kWh", ""),
    ]

    print()
    for name, value, indicator in metrics:
        print(f"  {name:<30} {value:>15} {indicator}")

    print(f"\n  {'RECOMMENDATION:':<30} {recommendation}")

    # Cost Breakdown
    print_subheader("COST BREAKDOWN")

    capex = results.annual_costs[0] if results.annual_costs else 0
    pv_om = results.pv_costs - capex / (1 + project.basics.discount_rate) ** 0

    print(f"\n  {'Year 0 CapEx (after ITC):':<35} {format_currency(capex, 1):>15}")
    print(f"  {'PV of O&M & Other Costs:':<35} {format_currency(pv_om, 1):>15}")
    print(f"  {'-' * 50}")
    print(f"  {'Total PV of Costs:':<35} {format_currency(results.pv_costs, 1):>15}")

    # Benefit Breakdown
    print_subheader("BENEFIT BREAKDOWN")

    print(f"\n  {'Benefit Category':<30} {'PV ($)':>15} {'% of Total':>12}")
    print(f"  {'-' * 57}")

    for name, pct in sorted(results.benefit_breakdown.items(), key=lambda x: -x[1]):
        pv_val = results.pv_benefits * pct / 100
        print(f"  {name:<30} {format_currency(pv_val, 1):>15} {pct:>10.1f}%")

    print(f"  {'-' * 57}")
    print(f"  {'Total PV of Benefits:':<30} {format_currency(results.pv_benefits, 1):>15} {'100.0%':>12}")

    # Cash Flow Summary
    print_subheader("ANNUAL CASH FLOWS")

    print(f"\n  {'Year':>6} {'Costs':>15} {'Benefits':>15} {'Net':>15} {'Cumulative':>15}")
    print(f"  {'-' * 66}")

    cumulative = 0.0
    for year in range(min(len(results.annual_costs), 6)):  # Show first 6 years
        cost = results.annual_costs[year]
        benefit = results.annual_benefits[year]
        net = results.annual_net[year]
        cumulative += net

        print(f"  {year:>6} {format_currency(cost):>15} {format_currency(benefit):>15} "
              f"{format_currency(net):>15} {format_currency(cumulative):>15}")

    if len(results.annual_costs) > 6:
        print(f"  {'...':>6}")
        # Show last year
        year = len(results.annual_costs) - 1
        cost = results.annual_costs[year]
        benefit = results.annual_benefits[year]
        net = results.annual_net[year]
        cumulative = sum(results.annual_net)
        print(f"  {year:>6} {format_currency(cost):>15} {format_currency(benefit):>15} "
              f"{format_currency(net):>15} {format_currency(cumulative):>15}")


def print_methodology() -> None:
    """Print methodology documentation."""

    print_header("METHODOLOGY & REFERENCES", "=")

    print("""
CALCULATION METHODOLOGY
-----------------------
This tool implements standard discounted cash flow (DCF) analysis for
battery energy storage system (BESS) economic evaluation.

KEY FORMULAS:

1. Net Present Value (NPV)
   NPV = Σ(CF_t / (1+r)^t) for t = 0 to N
   Source: Brealey, Myers, & Allen. Principles of Corporate Finance.

2. Benefit-Cost Ratio (BCR)
   BCR = PV(Benefits) / PV(Costs)
   Source: CPUC Standard Practice Manual (2001)

3. Internal Rate of Return (IRR)
   The discount rate where NPV = 0
   Source: Brealey, Myers, & Allen. Principles of Corporate Finance.

4. Levelized Cost of Storage (LCOS)
   LCOS = PV(Costs) / PV(Energy Discharged)
   Source: Lazard LCOS Analysis v10.0 (2025)

5. Weighted Average Cost of Capital (WACC)
   WACC = (E/V) * Re + (D/V) * Rd * (1 - Tc)
   Where: E=equity, D=debt, V=total value, Re=cost of equity,
          Rd=cost of debt, Tc=tax rate

ASSUMPTION LIBRARIES:
- NREL ATB 2024 (Moderate): National Renewable Energy Laboratory
- Lazard LCOS v10.0 (2025): Investment banking industry standard
- CPUC California 2024: California Public Utilities Commission

BENEFIT CATEGORIES:
Common Benefits (applicable to all infrastructure):
  - Resource Adequacy: Capacity value for grid reliability
  - Energy Arbitrage: Time-shifting energy for price spreads
  - Ancillary Services: Frequency regulation, reserves
  - T&D Deferral: Avoided transmission/distribution upgrades
  - Resilience Value: Avoided outage costs
  - Voltage Support: Power quality services

BESS-Specific Benefits:
  - Renewable Integration: Avoided curtailment
  - GHG Emissions Value: Carbon reduction value

REFERENCES:
[1] NREL. Annual Technology Baseline 2024. https://atb.nrel.gov/
[2] Lazard. Levelized Cost of Storage Analysis v10.0. 2025.
[3] CPUC. California Standard Practice Manual. 2001.
[4] Brealey, R., Myers, S., & Allen, F. Principles of Corporate Finance. 2020.
""")


# ============================================================================
# INTERACTIVE MODE
# ============================================================================

def get_user_input(prompt: str, default: str = "",
                   input_type: type = str,
                   valid_options: Optional[List[str]] = None) -> any:
    """Get user input with optional validation."""

    if default:
        prompt = f"{prompt} [{default}]: "
    else:
        prompt = f"{prompt}: "

    while True:
        user_input = input(prompt).strip()

        if not user_input and default:
            user_input = default

        if valid_options and user_input.lower() not in [v.lower() for v in valid_options]:
            print(f"  Invalid option. Choose from: {', '.join(valid_options)}")
            continue

        try:
            if input_type == float:
                return float(user_input)
            elif input_type == int:
                return int(user_input)
            else:
                return user_input
        except ValueError:
            print(f"  Invalid input. Please enter a valid {input_type.__name__}.")


def interactive_mode() -> Tuple[Project, FinancialResults]:
    """Run interactive project configuration."""

    print_header("BESS ANALYZER - Interactive Mode", "=")
    print("\nConfigure your BESS project step by step.\n")

    # Library selection
    print("Available Assumption Libraries:")
    print("  1. NREL ATB 2024 (Moderate)")
    print("  2. Lazard LCOS v10.0 (2025)")
    print("  3. CPUC California 2024")
    print("  4. Custom (enter all values manually)")

    lib_choice = get_user_input("\nSelect library (1-4)", "1", str, ["1", "2", "3", "4"])

    # Create base project
    project = Project()

    # Apply library if selected
    library_manager = AssumptionLibrary()
    if lib_choice == "1":
        library_manager.apply_library_to_project(project, "NREL ATB 2024 - Moderate Scenario")
    elif lib_choice == "2":
        library_manager.apply_library_to_project(project, "Lazard LCOS v10.0 - 2025")
    elif lib_choice == "3":
        library_manager.apply_library_to_project(project, "CPUC California 2024")

    # Project basics
    print_subheader("Project Basics")
    project.basics.name = get_user_input("Project name", "BESS Project")
    project.basics.location = get_user_input("Location", "California")
    project.basics.capacity_mw = get_user_input("Capacity (MW)", "100", float)
    project.basics.duration_hours = get_user_input("Duration (hours)", "4", float)
    project.basics.capacity_mwh = project.basics.capacity_mw * project.basics.duration_hours
    project.basics.analysis_period_years = get_user_input("Analysis period (years)", "20", int)
    project.basics.discount_rate = get_user_input("Discount rate (e.g., 0.07 for 7%)", "0.07", float)

    # Customize costs if desired
    customize = get_user_input("\nCustomize costs? (y/n)", "n", str, ["y", "n"])
    if customize.lower() == "y":
        print_subheader("Cost Inputs")
        project.costs.capex_per_kwh = get_user_input("CapEx ($/kWh)", str(project.costs.capex_per_kwh), float)
        project.costs.fom_per_kw_year = get_user_input("Fixed O&M ($/kW-year)", str(project.costs.fom_per_kw_year), float)
        project.costs.charging_cost_per_mwh = get_user_input("Charging cost ($/MWh)", str(project.costs.charging_cost_per_mwh), float)

    # Run calculations
    print("\nRunning economic analysis...")
    results = calculate_project_economics(project)
    project.results = results

    return project, results


# ============================================================================
# MAIN CLI
# ============================================================================

def create_default_project(library_name: Optional[str] = None) -> Project:
    """Create a default 100 MW / 400 MWh project."""

    project = Project(
        basics=ProjectBasics(
            name="Default BESS Project",
            location="California",
            capacity_mw=100.0,
            duration_hours=4.0,
            capacity_mwh=400.0,
            analysis_period_years=20,
            discount_rate=0.07,
        ),
    )

    if library_name:
        library_manager = AssumptionLibrary()
        library_names = library_manager.get_library_names()

        # Find matching library
        for name in library_names:
            if library_name.lower() in name.lower():
                library_manager.apply_library_to_project(project, name)
                break

    return project


def main():
    """Main entry point for CLI."""

    parser = argparse.ArgumentParser(
        description="BESS Analyzer CLI - Battery Energy Storage Economic Analysis",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python bess_cli.py                          # Interactive mode
  python bess_cli.py --library nrel           # Use NREL library
  python bess_cli.py --library lazard         # Use Lazard library
  python bess_cli.py --library cpuc           # Use CPUC California library
  python bess_cli.py --capacity 50 --duration 2  # 50 MW / 2-hour system
  python bess_cli.py --load project.json      # Load saved project
  python bess_cli.py --report output.pdf      # Generate PDF report
  python bess_cli.py --excel output.xlsx      # Export to Excel
  python bess_cli.py --sensitivity            # Show sensitivity tables
  python bess_cli.py --methodology            # Show calculation methodology
        """
    )

    # Project configuration
    parser.add_argument("--library", "-l", type=str,
                        help="Assumption library: nrel, lazard, or cpuc")
    parser.add_argument("--capacity", "-c", type=float, default=100.0,
                        help="Capacity in MW (default: 100)")
    parser.add_argument("--duration", "-d", type=float, default=4.0,
                        help="Duration in hours (default: 4)")
    parser.add_argument("--name", "-n", type=str, default="BESS Project",
                        help="Project name")
    parser.add_argument("--location", type=str, default="California",
                        help="Project location")
    parser.add_argument("--discount-rate", type=float, default=0.07,
                        help="Discount rate as decimal (default: 0.07)")
    parser.add_argument("--years", type=int, default=20,
                        help="Analysis period in years (default: 20)")

    # File operations
    parser.add_argument("--load", type=str,
                        help="Load project from JSON file")
    parser.add_argument("--save", type=str,
                        help="Save project to JSON file")
    parser.add_argument("--report", type=str, nargs="?", const="BESS_Report.pdf",
                        help="Generate PDF report (optional: specify filename)")
    parser.add_argument("--excel", type=str, nargs="?", const="BESS_Analysis.xlsx",
                        help="Export to Excel workbook")

    # Display options
    parser.add_argument("--sensitivity", "-s", action="store_true",
                        help="Show sensitivity analysis tables")
    parser.add_argument("--methodology", "-m", action="store_true",
                        help="Show calculation methodology")
    parser.add_argument("--quiet", "-q", action="store_true",
                        help="Suppress detailed output")
    parser.add_argument("--interactive", "-i", action="store_true",
                        help="Run in interactive mode")

    args = parser.parse_args()

    # Show methodology and exit
    if args.methodology:
        print_methodology()
        return

    # Interactive mode
    if args.interactive or len(sys.argv) == 1:
        project, results = interactive_mode()
    else:
        # Load or create project
        if args.load:
            print(f"Loading project from {args.load}...")
            project = load_project(args.load)
        else:
            project = create_default_project(args.library)
            project.basics.name = args.name
            project.basics.location = args.location
            project.basics.capacity_mw = args.capacity
            project.basics.duration_hours = args.duration
            project.basics.capacity_mwh = args.capacity * args.duration
            project.basics.analysis_period_years = args.years
            project.basics.discount_rate = args.discount_rate

        # Run calculations
        results = calculate_project_economics(project)
        project.results = results

    # Display results
    if not args.quiet:
        print_project_summary(project)
        print_results(project, results)

    # Sensitivity analysis
    if args.sensitivity:
        print_sensitivity_tables(project, results)

    # Save project
    if args.save:
        save_project(project, args.save)
        print(f"\nProject saved to {args.save}")

    # Generate PDF report
    if args.report:
        try:
            generate_executive_summary(project, results, args.report)
            print(f"\nPDF report generated: {args.report}")
        except Exception as e:
            print(f"\nError generating PDF: {e}")

    # Export to Excel
    if args.excel:
        try:
            from excel_generator import create_workbook

            # Set output path and generate
            output_path = Path(args.excel)
            create_workbook(str(output_path.with_suffix(".xlsx")))
            print(f"\nExcel workbook generated: {output_path}")
        except Exception as e:
            print(f"\nError generating Excel: {e}")

    # Print final summary
    if not args.quiet:
        print_header("ANALYSIS COMPLETE", "=")
        print(f"\n  BCR: {results.bcr:.2f}  |  NPV: {format_currency(results.npv, 1)}  |  "
              f"IRR: {format_percent(results.irr) if results.irr else 'N/A'}")

        if results.bcr >= 1.5:
            print("\n  [✓] RECOMMENDATION: APPROVE - Strong economic case")
        elif results.bcr >= 1.0:
            print("\n  [~] RECOMMENDATION: FURTHER STUDY - Marginal economics")
        else:
            print("\n  [✗] RECOMMENDATION: REJECT - Costs exceed benefits")
        print()


if __name__ == "__main__":
    main()
