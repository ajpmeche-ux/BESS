"""Financial calculation engine for BESS Analyzer.

Implements standard energy storage economic metrics: NPV, BCR, IRR,
LCOS, payback period, and breakeven analysis. All formulas are cited
to authoritative sources and documented with LaTeX notation.
"""

from typing import List, Optional

import numpy as np
import numpy_financial as npf

from src.models.project import (
    BenefitStream,
    CostInputs,
    FinancialResults,
    FinancingInputs,
    Project,
    ProjectBasics,
    SpecialBenefitInputs,
    TechnologySpecs,
)


def calculate_npv(cash_flows: List[float], discount_rate: float) -> float:
    r"""Calculate net present value of a cash flow series.

    Formula:
        NPV = \sum_{t=0}^{N} \frac{CF_t}{(1+r)^t}

    Args:
        cash_flows: List of cash flows starting at year 0.
            Positive values are inflows (benefits), negative are outflows (costs).
        discount_rate: Annual discount rate as decimal (e.g., 0.07 for 7%).

    Returns:
        Net present value in the same currency units as cash_flows.

    Source:
        Brealey, R., Myers, S., & Allen, F. (2020). Principles of Corporate
        Finance (13th ed.). McGraw-Hill. Chapter 2.

    Example:
        >>> calculate_npv([-1000, 300, 300, 300, 300], 0.10)
        -49.04...
    """
    pv = 0.0
    for t, cf in enumerate(cash_flows):
        pv += cf / (1 + discount_rate) ** t
    return pv


def calculate_bcr(pv_benefits: float, pv_costs: float) -> float:
    r"""Calculate benefit-cost ratio.

    Formula:
        BCR = \frac{PV(Benefits)}{PV(Costs)}

    A BCR > 1.0 indicates benefits exceed costs. CPUC uses BCR as a
    primary screening metric for demand-side resources.

    Args:
        pv_benefits: Present value of all benefit streams ($).
        pv_costs: Present value of all cost streams ($), must be positive.

    Returns:
        Benefit-cost ratio (dimensionless).

    Raises:
        ValueError: If pv_costs is zero or negative.

    Source:
        California Public Utilities Commission. (2001). California Standard
        Practice Manual: Economic Analysis of Demand-Side Programs and Projects.
    """
    if pv_costs <= 0:
        raise ValueError(f"pv_costs must be > 0, got {pv_costs}")
    return pv_benefits / pv_costs


def calculate_irr(cash_flows: List[float]) -> Optional[float]:
    r"""Calculate internal rate of return for a cash flow series.

    The IRR is the discount rate r that makes NPV = 0:
        0 = \sum_{t=0}^{N} \frac{CF_t}{(1+IRR)^t}

    Args:
        cash_flows: List of cash flows starting at year 0.
            Typically year 0 is negative (investment) and subsequent years positive.

    Returns:
        IRR as a decimal (e.g., 0.12 for 12%), or None if no real solution exists.

    Source:
        Brealey, R., Myers, S., & Allen, F. (2020). Principles of Corporate
        Finance (13th ed.). McGraw-Hill. Chapter 5.
    """
    try:
        result = npf.irr(cash_flows)
        if np.isnan(result) or np.isinf(result):
            return None
        return float(result)
    except Exception:
        return None


def calculate_lcos(
    annual_costs: List[float],
    annual_energy_discharged_mwh: List[float],
    discount_rate: float,
) -> float:
    r"""Calculate levelized cost of storage.

    Formula:
        LCOS = \frac{\sum_{t=0}^{N} \frac{Cost_t}{(1+r)^t}}
                     {\sum_{t=1}^{N} \frac{Energy_t}{(1+r)^t}}

    Args:
        annual_costs: List of annual costs starting at year 0 (CapEx at year 0).
        annual_energy_discharged_mwh: List of annual energy discharged (MWh),
            starting at year 0 (typically 0 for year 0).
        discount_rate: Annual discount rate as decimal.

    Returns:
        LCOS in $/MWh.

    Source:
        Lazard. (2025). Lazard's Levelized Cost of Storage Analysis, Version 10.0.
    """
    pv_costs = sum(c / (1 + discount_rate) ** t for t, c in enumerate(annual_costs))
    pv_energy = sum(
        e / (1 + discount_rate) ** t
        for t, e in enumerate(annual_energy_discharged_mwh)
    )
    if pv_energy <= 0:
        return 0.0
    return pv_costs / pv_energy


def _calculate_payback(annual_net: List[float]) -> Optional[float]:
    """Calculate simple payback period from net annual cash flows.

    Finds the first year where cumulative net cash flow turns positive,
    with linear interpolation within that year.

    Args:
        annual_net: Net cash flows starting at year 0.

    Returns:
        Payback period in years, or None if never achieved.
    """
    cumulative = 0.0
    for t, cf in enumerate(annual_net):
        prev_cumulative = cumulative
        cumulative += cf
        if cumulative >= 0 and t > 0:
            # Linear interpolation within the year
            fraction = -prev_cumulative / cf if cf != 0 else 0
            return t - 1 + fraction
    return None


def _get_bulk_discount_factor(project: Project) -> float:
    """Calculate bulk discount multiplier for fleet purchases.

    Returns a multiplier to apply to all costs when project capacity
    meets or exceeds the bulk discount threshold.

    Args:
        project: Project with costs containing bulk discount parameters.

    Returns:
        Discount multiplier (e.g., 0.90 for 10% discount), or 1.0 if no discount.
    """
    costs = project.costs
    capacity_mwh = project.basics.capacity_mwh

    if (costs.bulk_discount_rate > 0 and
        costs.bulk_discount_threshold_mwh > 0 and
        capacity_mwh >= costs.bulk_discount_threshold_mwh):
        return 1.0 - costs.bulk_discount_rate
    return 1.0


def calculate_project_economics(project: Project) -> FinancialResults:
    """Run complete economic analysis on a BESS project.

    Calculation workflow:
        1. Build annual cost stream (Year 0: CapEx; Years 1-N: O&M + charging;
           augmentation year: battery replacement; Year N: decommissioning - residual).
        2. Aggregate annual benefit streams from all BenefitStream objects.
        3. Apply degradation to energy-based calculations.
        4. Discount all cash flows to present value using WACC or discount rate.
        5. Calculate NPV, BCR, IRR, payback, LCOS, and breakeven CapEx.

    Args:
        project: Complete Project object with all inputs populated.

    Returns:
        FinancialResults with all calculated metrics.

    Source:
        Methodology follows NREL's Storage Futures Study (2021) and
        CPUC Standard Practice Manual (2001) frameworks.
    """
    basics = project.basics
    tech = project.technology
    costs = project.costs
    n = basics.analysis_period_years
    # Use WACC if financing structure provided, otherwise use discount_rate
    r = project.get_discount_rate()

    capacity_kw = basics.capacity_mw * 1000
    capacity_kwh = basics.capacity_mwh * 1000

    # Calculate bulk discount factor for fleet purchases
    bulk_discount = _get_bulk_discount_factor(project)

    # --- Build annual cost stream ---
    annual_costs = [0.0] * (n + 1)

    # Year 0: Capital expenditure (battery system) with bulk discount
    battery_capex = costs.capex_per_kwh * capacity_kwh * bulk_discount

    # Year 0: Infrastructure costs (common to all utility projects) with bulk discount
    interconnection_cost = costs.interconnection_per_kw * capacity_kw * bulk_discount
    land_cost = costs.land_per_kw * capacity_kw * bulk_discount
    permitting_cost = costs.permitting_per_kw * capacity_kw * bulk_discount
    infrastructure_costs = interconnection_cost + land_cost + permitting_cost

    # Total pre-ITC CapEx
    total_capex = battery_capex + infrastructure_costs

    # Apply Investment Tax Credit (BESS-specific under IRA)
    # ITC applies only to the battery system, not infrastructure
    total_itc_rate = costs.itc_percent + costs.itc_adders
    itc_credit = battery_capex * total_itc_rate

    # Year 0 net capital cost (after ITC)
    annual_costs[0] = total_capex - itc_credit

    # Years 1-N: Fixed O&M with bulk discount
    for t in range(1, n + 1):
        annual_costs[t] += costs.fom_per_kw_year * capacity_kw * bulk_discount

    # Years 1-N: Variable O&M + Charging costs (based on annual discharge energy)
    # Uses cycles_per_day instead of hardcoded 1 cycle
    for t in range(1, n + 1):
        degradation_factor = (1 - tech.degradation_rate_annual) ** (t - 1)
        annual_discharge_mwh = (
            basics.capacity_mwh * tech.cycles_per_day * 365 * tech.round_trip_efficiency * degradation_factor
        )
        annual_costs[t] += costs.vom_per_mwh * annual_discharge_mwh

        # Charging cost: energy needed to charge = discharge / RTE
        annual_charge_mwh = annual_discharge_mwh / tech.round_trip_efficiency
        annual_costs[t] += costs.charging_cost_per_mwh * annual_charge_mwh

    # Years 1-N: Insurance (common to all utility projects)
    # Based on percentage of total CapEx
    annual_insurance = total_capex * costs.insurance_pct_of_capex
    for t in range(1, n + 1):
        annual_costs[t] += annual_insurance

    # Years 1-N: Property taxes (common to all utility projects)
    # Based on percentage of depreciating asset value (straight-line)
    for t in range(1, n + 1):
        # Simplified: property tax on remaining book value (straight-line depreciation)
        remaining_value = total_capex * (1 - t / n)
        annual_costs[t] += remaining_value * costs.property_tax_pct

    # Augmentation year: battery replacement cost (adjusted for learning curve + bulk discount)
    # Cost declines at learning_rate annually from base year
    aug_year = tech.augmentation_year
    if 1 <= aug_year <= n:
        # Get augmentation cost adjusted for technology cost decline and bulk discount
        adjusted_aug_cost = costs.get_augmentation_cost(aug_year) * bulk_discount
        annual_costs[aug_year] += adjusted_aug_cost * capacity_kwh

    # Final year: decommissioning minus residual value
    decommissioning_cost = costs.decommissioning_per_kw * capacity_kw
    residual_value = total_capex * costs.residual_value_pct
    annual_costs[n] += decommissioning_cost - residual_value

    # --- Build annual benefit stream ---
    annual_benefits = [0.0] * (n + 1)
    benefit_pvs = {}

    for benefit in project.benefits:
        for t in range(1, n + 1):
            if t - 1 < len(benefit.annual_values):
                val = benefit.annual_values[t - 1]
            else:
                val = 0.0
            annual_benefits[t] += val

        # Calculate PV for this benefit category
        pv_this = sum(
            (benefit.annual_values[t - 1] if t - 1 < len(benefit.annual_values) else 0.0)
            / (1 + r) ** t
            for t in range(1, n + 1)
        )
        benefit_pvs[benefit.name] = pv_this

    # --- Process special benefits (formula-based) ---
    special = project.special_benefits
    if special:
        # Reliability Benefits (annual, with degradation)
        if special.reliability_enabled:
            reliability_base = special.calculate_reliability_annual(basics.capacity_mwh)
            for t in range(1, n + 1):
                # Apply battery degradation to reliability benefit
                degradation_factor = (1 - tech.degradation_rate_annual) ** (t - 1)
                annual_benefits[t] += reliability_base * degradation_factor

            # Calculate PV for benefit breakdown
            pv_reliability = sum(
                reliability_base * (1 - tech.degradation_rate_annual) ** (t - 1) / (1 + r) ** t
                for t in range(1, n + 1)
            )
            benefit_pvs["Reliability (Avoided Outage)"] = pv_reliability

        # Safety Benefits (annual, constant - no degradation)
        if special.safety_enabled:
            safety_annual = special.calculate_safety_annual(basics.capacity_mw)
            for t in range(1, n + 1):
                annual_benefits[t] += safety_annual

            # Calculate PV for benefit breakdown
            pv_safety = sum(safety_annual / (1 + r) ** t for t in range(1, n + 1))
            benefit_pvs["Safety (Avoided Incident)"] = pv_safety

        # Speed-to-Serve Benefits (ONE-TIME in Year 1 only)
        if special.speed_enabled:
            speed_onetime = special.calculate_speed_onetime(capacity_kw)
            annual_benefits[1] += speed_onetime

            # PV is just the discounted Year 1 value
            pv_speed = speed_onetime / (1 + r)
            benefit_pvs["Speed-to-Serve (One-time)"] = pv_speed

    # --- Calculate totals ---
    pv_costs = sum(c / (1 + r) ** t for t, c in enumerate(annual_costs))
    pv_benefits = sum(b / (1 + r) ** t for t, b in enumerate(annual_benefits))
    npv = pv_benefits - pv_costs

    # BCR
    bcr = calculate_bcr(pv_benefits, pv_costs) if pv_costs > 0 else 0.0

    # Net cash flows for IRR and payback
    annual_net = [annual_benefits[t] - annual_costs[t] for t in range(n + 1)]
    irr = calculate_irr(annual_net)
    payback = _calculate_payback(annual_net)

    # LCOS
    annual_energy = [0.0] * (n + 1)
    for t in range(1, n + 1):
        degradation_factor = (1 - tech.degradation_rate_annual) ** (t - 1)
        annual_energy[t] = (
            basics.capacity_mwh * tech.cycles_per_day * 365 * tech.round_trip_efficiency * degradation_factor
        )
    lcos = calculate_lcos(annual_costs, annual_energy, r)

    # Breakeven CapEx: the CapEx/kWh where BCR = 1.0
    # BCR=1 means PV_benefits = PV_costs
    # PV_costs = capex * capacity_kwh + PV_other_costs
    # So breakeven_capex = (PV_benefits - PV_other_costs) / capacity_kwh
    pv_other_costs = sum(c / (1 + r) ** t for t, c in enumerate(annual_costs)) - (
        annual_costs[0] / (1 + r) ** 0
    )
    breakeven_capex = (pv_benefits - pv_other_costs) / capacity_kwh if capacity_kwh > 0 else 0.0

    # Benefit breakdown (percentage of total PV)
    benefit_breakdown = {}
    if pv_benefits > 0:
        for name, pv_val in benefit_pvs.items():
            benefit_breakdown[name] = (pv_val / pv_benefits) * 100

    return FinancialResults(
        pv_benefits=pv_benefits,
        pv_costs=pv_costs,
        npv=npv,
        bcr=bcr,
        irr=irr,
        payback_years=payback,
        lcos_per_mwh=lcos,
        breakeven_capex_per_kwh=breakeven_capex,
        benefit_breakdown=benefit_breakdown,
        annual_costs=annual_costs,
        annual_benefits=annual_benefits,
        annual_net=annual_net,
    )
