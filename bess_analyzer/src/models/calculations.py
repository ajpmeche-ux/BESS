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
    BuildSchedule,
    CostInputs,
    FinancialResults,
    FinancingInputs,
    Project,
    ProjectBasics,
    SpecialBenefitInputs,
    TDDeferralSchedule,
    TechnologySpecs,
    UOSInputs,
)
from src.models.rate_base import CostOfCapital, RateBaseInputs, calculate_revenue_requirement
from src.models.avoided_costs import AvoidedCosts
from src.models.wires_comparison import (
    WiresAlternative, NWAParameters, compare_wires_vs_nwa,
)
from src.models.sod_check import SODInputs, check_sod_feasibility


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


def _calculate_cohort_costs(
    cod_year: int,
    capacity_mw: float,
    duration_hours: float,
    tech: TechnologySpecs,
    costs: CostInputs,
    n: int,
    global_year_0: int,
    bulk_discount: float,
    apply_learning_curve_capex: bool = False,
) -> tuple:
    """Compute annual costs for a single cohort/tranche.

    Args:
        cod_year: Calendar year this cohort comes online.
        capacity_mw: This cohort's power capacity (MW).
        duration_hours: Storage duration (hours).
        tech: Technology specifications.
        costs: Cost input parameters.
        n: Total analysis period years.
        global_year_0: Calendar year of the earliest tranche (analysis year 0).
        bulk_discount: Bulk discount multiplier (e.g., 0.95 for 5% discount).
        apply_learning_curve_capex: If True, apply learning curve to initial CapEx.
            Used for multi-tranche projects. Single-tranche uses base capex for
            backward compatibility.

    Returns:
        Tuple of (annual_costs list[0..N], cohort_capex_total float).
    """
    offset = cod_year - global_year_0  # Year within analysis when cohort comes online
    capacity_kw = capacity_mw * 1000
    capacity_mwh = capacity_mw * duration_hours
    capacity_kwh = capacity_mwh * 1000

    cohort_costs = [0.0] * (n + 1)

    # CapEx: apply learning curve only for multi-tranche projects
    if apply_learning_curve_capex:
        capex_per_kwh = costs.get_capex_at_year(cod_year)
    else:
        capex_per_kwh = costs.capex_per_kwh
    battery_capex = capex_per_kwh * capacity_kwh * bulk_discount

    # Infrastructure costs
    interconnection_cost = costs.interconnection_per_kw * capacity_kw * bulk_discount
    land_cost = costs.land_per_kw * capacity_kw * bulk_discount
    permitting_cost = costs.permitting_per_kw * capacity_kw * bulk_discount
    infrastructure_costs = interconnection_cost + land_cost + permitting_cost

    total_capex = battery_capex + infrastructure_costs

    # ITC on battery only
    total_itc_rate = costs.itc_percent + costs.itc_adders
    itc_credit = battery_capex * total_itc_rate

    # Place CapEx at cohort offset year
    if offset <= n:
        cohort_costs[offset] = total_capex - itc_credit

    # Operating costs from cohort_offset+1 through N
    for t in range(offset + 1, n + 1):
        years_operating = t - offset  # years since this cohort's COD

        # Fixed O&M
        cohort_costs[t] += costs.fom_per_kw_year * capacity_kw * bulk_discount

        # Variable O&M + Charging (with cohort-specific degradation)
        degradation_factor = (1 - tech.degradation_rate_annual) ** (years_operating - 1)
        annual_discharge_mwh = (
            capacity_mwh * tech.cycles_per_day * 365 * tech.round_trip_efficiency * degradation_factor
        )
        cohort_costs[t] += costs.vom_per_mwh * annual_discharge_mwh
        annual_charge_mwh = annual_discharge_mwh / tech.round_trip_efficiency
        cohort_costs[t] += costs.charging_cost_per_mwh * annual_charge_mwh

        # Insurance
        cohort_costs[t] += total_capex * costs.insurance_pct_of_capex

        # Property tax on remaining book value
        remaining_value = total_capex * max(0, 1 - years_operating / n)
        cohort_costs[t] += remaining_value * costs.property_tax_pct

    # Augmentation relative to cohort COD
    aug_year = offset + tech.augmentation_year
    if 1 <= aug_year <= n:
        adjusted_aug_cost = costs.get_augmentation_cost(tech.augmentation_year) * bulk_discount
        cohort_costs[aug_year] += adjusted_aug_cost * capacity_kwh

    # Decommissioning at end of analysis
    decommissioning_cost = costs.decommissioning_per_kw * capacity_kw
    residual_value = total_capex * costs.residual_value_pct
    cohort_costs[n] += decommissioning_cost - residual_value

    return cohort_costs, total_capex


def _calculate_cohort_energy(
    cod_year: int,
    capacity_mw: float,
    duration_hours: float,
    tech: TechnologySpecs,
    n: int,
    global_year_0: int,
) -> List[float]:
    """Compute annual energy discharged for a single cohort."""
    offset = cod_year - global_year_0
    capacity_mwh = capacity_mw * duration_hours
    energy = [0.0] * (n + 1)
    for t in range(offset + 1, n + 1):
        years_operating = t - offset
        degradation_factor = (1 - tech.degradation_rate_annual) ** (years_operating - 1)
        energy[t] = capacity_mwh * tech.cycles_per_day * 365 * tech.round_trip_efficiency * degradation_factor
    return energy


def _get_effective_capacity_ratios(
    tranches: List[tuple],
    tech: TechnologySpecs,
    n: int,
    global_year_0: int,
    total_capacity_mw: float,
) -> List[float]:
    """Compute effective capacity ratio at each year for benefit scaling.

    For each analysis year, sums the effective (degraded) capacity of all
    online cohorts, divided by total nominal capacity. For single-tranche
    projects this equals the standard degradation factor.

    Returns:
        List of ratios [0..N], where ratio[0] = 0 (no benefits at year 0).
    """
    ratios = [0.0] * (n + 1)
    for t in range(1, n + 1):
        effective_mw = 0.0
        for cod_year, cap_mw in tranches:
            offset = cod_year - global_year_0
            if t > offset:
                years_operating = t - offset
                degradation = (1 - tech.degradation_rate_annual) ** (years_operating - 1)
                effective_mw += cap_mw * degradation
        ratios[t] = effective_mw / total_capacity_mw if total_capacity_mw > 0 else 0.0
    return ratios


def calculate_project_economics(project: Project) -> FinancialResults:
    """Run complete economic analysis on a BESS project.

    Supports both single-asset and multi-tranche (JIT cohort) models.
    When a build_schedule is provided, costs are computed per-cohort with
    learning-curve-adjusted CapEx and cohort-specific degradation/augmentation.
    Benefits are scaled by effective online capacity at each year.

    Args:
        project: Complete Project object with all inputs populated.

    Returns:
        FinancialResults with all calculated metrics.
    """
    basics = project.basics
    tech = project.technology
    costs = project.costs
    n = basics.analysis_period_years
    r = project.get_discount_rate()

    total_capacity_mw = basics.capacity_mw
    total_capacity_kw = total_capacity_mw * 1000
    total_capacity_kwh = basics.capacity_mwh * 1000

    bulk_discount = _get_bulk_discount_factor(project)

    # Get tranches (single or multi)
    tranches = project.get_effective_tranches()
    is_multi = project.is_multi_tranche()
    global_year_0 = min(y for y, _ in tranches)

    # --- Aggregate costs across cohorts ---
    annual_costs = [0.0] * (n + 1)
    annual_energy = [0.0] * (n + 1)
    cohort_capex_list = []

    for cod_year, cap_mw in tranches:
        cohort_costs, cohort_total_capex = _calculate_cohort_costs(
            cod_year, cap_mw, basics.duration_hours,
            tech, costs, n, global_year_0, bulk_discount,
            apply_learning_curve_capex=is_multi,
        )
        cohort_energy = _calculate_cohort_energy(
            cod_year, cap_mw, basics.duration_hours,
            tech, n, global_year_0,
        )
        for t in range(n + 1):
            annual_costs[t] += cohort_costs[t]
            annual_energy[t] += cohort_energy[t]
        cohort_capex_list.append(cohort_total_capex)

    # --- Build annual benefit stream ---
    # For multi-tranche: scale benefits by effective capacity ratio to account
    # for phased deployment. For single-tranche: use values as-is (backward compat).
    if is_multi:
        capacity_ratios = _get_effective_capacity_ratios(
            tranches, tech, n, global_year_0, total_capacity_mw,
        )
    else:
        # Single tranche: no scaling on standard benefits (preserves original behavior)
        capacity_ratios = [0.0] + [1.0] * n

    annual_benefits = [0.0] * (n + 1)
    benefit_pvs = {}

    for benefit in project.benefits:
        for t in range(1, n + 1):
            if t - 1 < len(benefit.annual_values):
                val = benefit.annual_values[t - 1] * capacity_ratios[t]
            else:
                val = 0.0
            annual_benefits[t] += val

        pv_this = sum(
            (benefit.annual_values[t - 1] if t - 1 < len(benefit.annual_values) else 0.0)
            * capacity_ratios[t] / (1 + r) ** t
            for t in range(1, n + 1)
        )
        benefit_pvs[benefit.name] = pv_this

    # --- Process special benefits (formula-based) ---
    # For single-tranche, use original degradation pattern for reliability.
    # For multi-tranche, use capacity_ratios which encodes per-cohort degradation.
    special = project.special_benefits
    if special:
        if special.reliability_enabled:
            reliability_base = special.calculate_reliability_annual(basics.capacity_mwh)
            if is_multi:
                for t in range(1, n + 1):
                    annual_benefits[t] += reliability_base * capacity_ratios[t]
                pv_reliability = sum(
                    reliability_base * capacity_ratios[t] / (1 + r) ** t
                    for t in range(1, n + 1)
                )
            else:
                for t in range(1, n + 1):
                    degradation_factor = (1 - tech.degradation_rate_annual) ** (t - 1)
                    annual_benefits[t] += reliability_base * degradation_factor
                pv_reliability = sum(
                    reliability_base * (1 - tech.degradation_rate_annual) ** (t - 1) / (1 + r) ** t
                    for t in range(1, n + 1)
                )
            benefit_pvs["Reliability (Avoided Outage)"] = pv_reliability

        if special.safety_enabled:
            safety_annual = special.calculate_safety_annual(basics.capacity_mw)
            for t in range(1, n + 1):
                annual_benefits[t] += safety_annual
            pv_safety = sum(safety_annual / (1 + r) ** t for t in range(1, n + 1))
            benefit_pvs["Safety (Avoided Incident)"] = pv_safety

        if special.speed_enabled:
            speed_onetime = special.calculate_speed_onetime(total_capacity_kw)
            annual_benefits[1] += speed_onetime
            pv_speed = speed_onetime / (1 + r)
            benefit_pvs["Speed-to-Serve (One-time)"] = pv_speed

    # --- Calculate totals ---
    pv_costs = sum(c / (1 + r) ** t for t, c in enumerate(annual_costs))
    pv_benefits = sum(b / (1 + r) ** t for t, b in enumerate(annual_benefits))
    npv = pv_benefits - pv_costs

    bcr = calculate_bcr(pv_benefits, pv_costs) if pv_costs > 0 else 0.0

    annual_net = [annual_benefits[t] - annual_costs[t] for t in range(n + 1)]
    irr = calculate_irr(annual_net)
    payback = _calculate_payback(annual_net)

    lcos = calculate_lcos(annual_costs, annual_energy, r)

    pv_other_costs = pv_costs - annual_costs[0]  # PV of non-CapEx costs
    breakeven_capex = (pv_benefits - pv_other_costs) / total_capacity_kwh if total_capacity_kwh > 0 else 0.0

    benefit_breakdown = {}
    if pv_benefits > 0:
        for name, pv_val in benefit_pvs.items():
            benefit_breakdown[name] = (pv_val / pv_benefits) * 100

    # T&D deferral PV
    td_pv = 0.0
    if project.td_deferral:
        td_pv = project.td_deferral.total_pv(r)

    # Flexibility value (multi-tranche only)
    flex_value = 0.0
    if project.is_multi_tranche():
        flex_value = calculate_flexibility_value(project)

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
        flexibility_value=flex_value,
        td_deferral_pv=td_pv,
        cohort_capex=cohort_capex_list,
        num_tranches=len(tranches),
    )


def calculate_flexibility_value(project: Project) -> float:
    """Calculate Flexibility Value = PV(cost_upfront) - PV(cost_phased).

    Compares building all capacity at the earliest COD year vs the
    actual phased build schedule. The difference represents the value
    of flexibility from deferred capital and learning curve savings.

    Args:
        project: Project with multi-tranche build schedule.

    Returns:
        Flexibility value in $. Positive when phased is cheaper.
    """
    if not project.is_multi_tranche():
        return 0.0

    # Create upfront variant: all capacity at first COD year
    tranches = project.get_effective_tranches()
    first_year = min(y for y, _ in tranches)
    total_mw = sum(mw for _, mw in tranches)

    upfront_schedule = BuildSchedule(tranches=[(first_year, total_mw)])

    # Clone project with upfront schedule
    data = project.to_dict()
    upfront_project = Project.from_dict(data)
    upfront_project.build_schedule = upfront_schedule

    # Calculate costs for both scenarios (avoid infinite recursion by
    # computing costs directly without calling calculate_project_economics)
    r = project.get_discount_rate()
    n = project.basics.analysis_period_years
    bulk_discount = _get_bulk_discount_factor(project)

    # Upfront costs (learning curve applied since this is multi-tranche context)
    upfront_costs = [0.0] * (n + 1)
    for cod_year, cap_mw in upfront_schedule.tranches:
        cohort_costs, _ = _calculate_cohort_costs(
            cod_year, cap_mw, project.basics.duration_hours,
            project.technology, project.costs, n, first_year, bulk_discount,
            apply_learning_curve_capex=True,
        )
        for t in range(n + 1):
            upfront_costs[t] += cohort_costs[t]
    pv_upfront = sum(c / (1 + r) ** t for t, c in enumerate(upfront_costs))

    # Phased costs
    phased_costs = [0.0] * (n + 1)
    for cod_year, cap_mw in tranches:
        cohort_costs, _ = _calculate_cohort_costs(
            cod_year, cap_mw, project.basics.duration_hours,
            project.technology, project.costs, n, first_year, bulk_discount,
            apply_learning_curve_capex=True,
        )
        for t in range(n + 1):
            phased_costs[t] += cohort_costs[t]
    pv_phased = sum(c / (1 + r) ** t for t, c in enumerate(phased_costs))

    return pv_upfront - pv_phased


def calculate_uos_analysis(project: Project) -> dict:
    """Run Utility-Owned Storage (UOS) revenue requirement analysis.

    Performs:
    1. Rate Base / Revenue Requirement calculation
    2. Avoided Cost (ACC) benefit stream
    3. Wires vs NWA comparison
    4. Slice-of-Day feasibility check
    5. Ratepayer impact analysis

    Args:
        project: Project with uos_inputs populated.

    Returns:
        Dictionary with:
            - rate_base_results: RateBaseResults
            - avoided_costs_annual: List of annual avoided costs
            - wires_comparison: ComparisonResult
            - sod_result: SODResult
            - ratepayer_impact: Annual net impact (avoided cost - revenue requirement)
            - cumulative_savings: Cumulative ratepayer savings
    """
    uos = project.uos_inputs
    if not uos or not uos.enabled:
        return {}

    basics = project.basics
    tech = project.technology
    costs = project.costs
    n = basics.analysis_period_years
    capacity_kw = basics.capacity_mw * 1000
    capacity_kwh = basics.capacity_mwh * 1000

    # --- Build Cost of Capital ---
    coc = CostOfCapital(
        roe=uos.roe,
        cost_of_debt=uos.cost_of_debt,
        cost_of_preferred=uos.cost_of_preferred,
        equity_ratio=uos.equity_ratio,
        debt_ratio=uos.debt_ratio,
        preferred_ratio=uos.preferred_ratio,
        ror=uos.ror,
        federal_tax_rate=uos.federal_tax_rate,
        state_tax_rate=uos.state_tax_rate,
        property_tax_rate=uos.property_tax_rate,
    )

    # --- Gross Plant calculation ---
    gross_plant = costs.capex_per_kwh * capacity_kwh
    infra = (costs.interconnection_per_kw + costs.land_per_kw +
             costs.permitting_per_kw) * capacity_kw
    total_plant = gross_plant + infra

    # Annual O&M
    annual_om = costs.fom_per_kw_year * capacity_kw

    # --- Revenue Requirement ---
    rb_inputs = RateBaseInputs(
        gross_plant=total_plant,
        book_life_years=uos.book_life_years,
        macrs_class=uos.macrs_class,
        itc_rate=costs.itc_percent + costs.itc_adders,
        itc_basis_reduction=True,
        cost_of_capital=coc,
        annual_om=annual_om,
        analysis_years=n,
        bonus_depreciation_pct=uos.bonus_depreciation_pct,
    )
    rb_results = calculate_revenue_requirement(rb_inputs)

    # --- Avoided Costs (ACC) ---
    acc = AvoidedCosts()  # Uses default SCE values
    avoided_annual = acc.get_annual_avoided_costs(
        capacity_kw=capacity_kw,
        capacity_mwh=basics.capacity_mwh,
        rte=tech.round_trip_efficiency,
        degradation_rate=tech.degradation_rate_annual,
        cycles_per_day=tech.cycles_per_day,
        n_years=n,
        include_distribution=False,
    )

    # --- Wires vs NWA Comparison ---
    wires = WiresAlternative(
        cost_per_kw=uos.wires_cost_per_kw,
        capacity_kw=capacity_kw,
        book_life_years=uos.wires_book_life,
        lead_time_years=uos.wires_lead_time,
        macrs_class=20,
    )
    nwa_params = NWAParameters(
        deferral_years=uos.nwa_deferral_years,
        incrementality_flag=uos.nwa_incrementality,
        bess_gross_plant=total_plant,
        bess_book_life_years=uos.book_life_years,
        bess_macrs_class=uos.macrs_class,
        bess_annual_om=annual_om,
        bess_itc_rate=costs.itc_percent + costs.itc_adders,
        avoided_cost_annual=sum(avoided_annual) / n if avoided_annual else 0.0,
    )
    wires_result = compare_wires_vs_nwa(wires, nwa_params, coc, n)

    # --- Slice-of-Day Feasibility ---
    sod_inputs = SODInputs(
        capacity_mw=basics.capacity_mw,
        duration_hours=basics.duration_hours,
        round_trip_efficiency=tech.round_trip_efficiency,
        degradation_rate=tech.degradation_rate_annual,
        analysis_year=1,
        min_qualifying_hours=uos.sod_min_hours,
        deration_threshold=uos.sod_deration_threshold,
    )
    sod_result = check_sod_feasibility(sod_inputs)

    # --- Ratepayer Impact ---
    rr_annual = rb_results.get_annual_revenue_requirements()
    ratepayer_impact = []
    cumulative_savings = []
    running = 0.0
    for yr in range(n):
        rr = rr_annual[yr] if yr < len(rr_annual) else 0.0
        avoided = avoided_annual[yr] if yr < len(avoided_annual) else 0.0
        net = avoided - rr  # positive = savings for ratepayers
        ratepayer_impact.append(net)
        running += net
        cumulative_savings.append(running)

    return {
        "rate_base_results": rb_results,
        "avoided_costs_annual": avoided_annual,
        "wires_comparison": wires_result,
        "sod_result": sod_result,
        "ratepayer_impact": ratepayer_impact,
        "cumulative_savings": cumulative_savings,
        "revenue_requirement_annual": rr_annual,
    }
