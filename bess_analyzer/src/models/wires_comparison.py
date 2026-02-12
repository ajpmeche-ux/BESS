"""Wires vs Non-Wires Alternative (NWA) Comparison Module.

Compares the cost of traditional distribution/transmission infrastructure
("wires") against a BESS-based Non-Wires Alternative using the
Real Economic Carrying Charge (RECC) method.

The RECC method levelizes capital costs into an equivalent annual
revenue requirement, enabling apples-to-apples comparison between
assets with different lifetimes, tax treatments, and financing structures.

References:
    - E3. 2024 Avoided Cost Calculator. RECC Methodology. CPUC.
    - CPUC D.18-02-004. Distribution Resource Plans (DRP).
    - SCE Distribution Resource Plan 2024.
"""

from dataclasses import dataclass, field
from typing import List, Optional

from src.models.rate_base import (
    CostOfCapital,
    RateBaseInputs,
    RateBaseResults,
    calculate_revenue_requirement,
)


@dataclass
class WiresAlternative:
    """Traditional wires (poles & wires) infrastructure parameters.

    Attributes:
        total_cost: Total capital cost of the wires project ($).
        cost_per_kw: Capital cost per kW of capacity addressed ($/kW).
        capacity_kw: Capacity of the wires solution (kW).
        book_life_years: Book depreciation life (typically 40 years for T&D).
        lead_time_years: Years from approval to in-service.
        annual_om: Annual O&M cost ($).
        macrs_class: MACRS property class (typically 15 or 20 for T&D).
    """

    total_cost: float = 0.0
    cost_per_kw: float = 500.0
    capacity_kw: float = 100_000.0  # 100 MW default
    book_life_years: int = 40
    lead_time_years: int = 5
    annual_om: float = 0.0
    macrs_class: int = 20  # T&D typically 20-year MACRS

    def __post_init__(self):
        if self.total_cost == 0.0 and self.cost_per_kw > 0:
            self.total_cost = self.cost_per_kw * self.capacity_kw

    def to_dict(self) -> dict:
        return {
            "total_cost": self.total_cost,
            "cost_per_kw": self.cost_per_kw,
            "capacity_kw": self.capacity_kw,
            "book_life_years": self.book_life_years,
            "lead_time_years": self.lead_time_years,
            "annual_om": self.annual_om,
            "macrs_class": self.macrs_class,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "WiresAlternative":
        data = dict(data)
        return cls(**{k: v for k, v in data.items()
                      if k in cls.__dataclass_fields__})


@dataclass
class NWAParameters:
    """Non-Wires Alternative (BESS) comparison parameters.

    Attributes:
        deferral_years: Number of years the NWA defers the wires project.
        incrementality_flag: If True, NWA benefits are only the incremental
            avoided cost above what the BESS would provide anyway.
        bess_gross_plant: Total BESS capital cost ($).
        bess_book_life_years: BESS book depreciation life.
        bess_macrs_class: BESS MACRS class (typically 7-year).
        bess_annual_om: BESS annual O&M ($).
        bess_itc_rate: BESS Investment Tax Credit rate.
        avoided_cost_annual: Annual avoided cost from BESS benefits ($).
    """

    deferral_years: int = 5
    incrementality_flag: bool = True
    bess_gross_plant: float = 0.0
    bess_book_life_years: int = 20
    bess_macrs_class: int = 7
    bess_annual_om: float = 0.0
    bess_itc_rate: float = 0.30
    avoided_cost_annual: float = 0.0

    def to_dict(self) -> dict:
        return {
            "deferral_years": self.deferral_years,
            "incrementality_flag": self.incrementality_flag,
            "bess_gross_plant": self.bess_gross_plant,
            "bess_book_life_years": self.bess_book_life_years,
            "bess_macrs_class": self.bess_macrs_class,
            "bess_annual_om": self.bess_annual_om,
            "bess_itc_rate": self.bess_itc_rate,
            "avoided_cost_annual": self.avoided_cost_annual,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "NWAParameters":
        data = dict(data)
        return cls(**{k: v for k, v in data.items()
                      if k in cls.__dataclass_fields__})


@dataclass
class ComparisonResult:
    """Results of a Wires vs NWA comparison.

    Attributes:
        wires_recc: Wires RECC (levelized annual revenue requirement, $/yr).
        nwa_recc: NWA RECC (levelized annual revenue requirement, $/yr).
        wires_total_rr: Total wires revenue requirement over analysis period ($).
        nwa_total_rr: Total NWA revenue requirement over analysis period ($).
        annual_savings: Annual savings from choosing NWA over wires ($/yr).
        total_savings: Total savings over the deferral period ($).
        cumulative_savings: Cumulative savings by year ($).
        nwa_is_economic: Whether NWA is more economic than wires.
        deferral_value: Present value of deferring the wires project ($).
        wires_annual_rr: Annual wires revenue requirement schedule.
        nwa_annual_rr: Annual NWA revenue requirement schedule.
    """

    wires_recc: float = 0.0
    nwa_recc: float = 0.0
    wires_total_rr: float = 0.0
    nwa_total_rr: float = 0.0
    annual_savings: float = 0.0
    total_savings: float = 0.0
    cumulative_savings: List[float] = field(default_factory=list)
    nwa_is_economic: bool = False
    deferral_value: float = 0.0
    wires_annual_rr: List[float] = field(default_factory=list)
    nwa_annual_rr: List[float] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "wires_recc": self.wires_recc,
            "nwa_recc": self.nwa_recc,
            "wires_total_rr": self.wires_total_rr,
            "nwa_total_rr": self.nwa_total_rr,
            "annual_savings": self.annual_savings,
            "total_savings": self.total_savings,
            "cumulative_savings": list(self.cumulative_savings),
            "nwa_is_economic": self.nwa_is_economic,
            "deferral_value": self.deferral_value,
        }


def calculate_recc(total_revenue_requirement: float,
                   analysis_years: int,
                   discount_rate: float) -> float:
    """Calculate Real Economic Carrying Charge (levelized annual RR).

    RECC = PV(Revenue Requirement) / Annuity Factor

    This converts a stream of uneven annual revenue requirements into
    an equivalent level annual payment.

    Args:
        total_revenue_requirement: Sum of all annual RRs ($).
        analysis_years: Number of years.
        discount_rate: Discount rate for levelization.

    Returns:
        Levelized annual revenue requirement ($/yr).
    """
    if analysis_years <= 0 or discount_rate <= 0:
        return total_revenue_requirement / max(1, analysis_years)

    annuity_factor = (1 - (1 + discount_rate) ** -analysis_years) / discount_rate
    if annuity_factor <= 0:
        return 0.0
    return total_revenue_requirement / annuity_factor


def calculate_deferral_value(wires_cost: float,
                             deferral_years: int,
                             discount_rate: float) -> float:
    """Calculate the present value benefit of deferring a wires project.

    Deferral Value = Wires Cost - PV(Wires Cost deferred N years)
                   = Wires Cost * (1 - 1/(1+r)^N)

    Args:
        wires_cost: Total capital cost of the wires project ($).
        deferral_years: Number of years the project is deferred.
        discount_rate: Discount rate.

    Returns:
        Present value of the deferral benefit ($).
    """
    if deferral_years <= 0 or discount_rate <= 0:
        return 0.0
    return wires_cost * (1 - 1 / (1 + discount_rate) ** deferral_years)


def compare_wires_vs_nwa(
    wires: WiresAlternative,
    nwa: NWAParameters,
    cost_of_capital: CostOfCapital,
    analysis_years: int = 20,
) -> ComparisonResult:
    """Compare traditional wires infrastructure vs BESS Non-Wires Alternative.

    Uses the RECC method to levelized both alternatives and compare
    their annual revenue requirements.

    Args:
        wires: Traditional wires infrastructure parameters.
        nwa: Non-Wires Alternative (BESS) parameters.
        cost_of_capital: CPUC-authorized cost of capital.
        analysis_years: Comparison analysis period.

    Returns:
        ComparisonResult with RECC values and savings analysis.
    """
    ror = cost_of_capital.ror

    # --- Calculate Wires Revenue Requirement ---
    wires_rb_inputs = RateBaseInputs(
        gross_plant=wires.total_cost,
        book_life_years=wires.book_life_years,
        macrs_class=wires.macrs_class,
        itc_rate=0.0,  # No ITC for traditional T&D
        itc_basis_reduction=False,
        cost_of_capital=cost_of_capital,
        annual_om=wires.annual_om,
        analysis_years=analysis_years,
    )
    wires_rr = calculate_revenue_requirement(wires_rb_inputs)
    wires_annual = wires_rr.get_annual_revenue_requirements()

    # --- Calculate NWA (BESS) Revenue Requirement ---
    nwa_rb_inputs = RateBaseInputs(
        gross_plant=nwa.bess_gross_plant,
        book_life_years=nwa.bess_book_life_years,
        macrs_class=nwa.bess_macrs_class,
        itc_rate=nwa.bess_itc_rate,
        itc_basis_reduction=True,
        cost_of_capital=cost_of_capital,
        annual_om=nwa.bess_annual_om,
        analysis_years=analysis_years,
    )
    nwa_rr = calculate_revenue_requirement(nwa_rb_inputs)
    nwa_annual = nwa_rr.get_annual_revenue_requirements()

    # If incrementality flag is set, subtract baseline avoided costs
    # (benefits the BESS provides regardless of NWA status)
    if nwa.incrementality_flag and nwa.avoided_cost_annual > 0:
        nwa_annual = [max(0.0, rr - nwa.avoided_cost_annual) for rr in nwa_annual]

    # --- RECC Calculation ---
    wires_recc = wires_rr.levelized_revenue_requirement
    nwa_recc = calculate_recc(sum(nwa_annual), analysis_years, ror)

    # --- Deferral Value ---
    deferral_val = calculate_deferral_value(wires.total_cost, nwa.deferral_years, ror)

    # --- Savings Analysis ---
    # Compare over the deferral period (or full analysis if shorter)
    compare_years = min(nwa.deferral_years, analysis_years)
    annual_savings = wires_recc - nwa_recc

    cumulative = []
    running_total = 0.0
    for yr in range(analysis_years):
        if yr < len(wires_annual) and yr < len(nwa_annual):
            yr_savings = wires_annual[yr] - nwa_annual[yr]
        else:
            yr_savings = 0.0
        running_total += yr_savings
        cumulative.append(running_total)

    total_savings = sum(
        wires_annual[i] - nwa_annual[i]
        for i in range(min(compare_years, len(wires_annual), len(nwa_annual)))
    )

    return ComparisonResult(
        wires_recc=wires_recc,
        nwa_recc=nwa_recc,
        wires_total_rr=wires_rr.total_revenue_requirement,
        nwa_total_rr=sum(nwa_annual),
        annual_savings=annual_savings,
        total_savings=total_savings,
        cumulative_savings=cumulative,
        nwa_is_economic=nwa_recc < wires_recc,
        deferral_value=deferral_val,
        wires_annual_rr=wires_annual,
        nwa_annual_rr=nwa_annual,
    )
