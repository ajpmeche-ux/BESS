"""Rate Base and Revenue Requirement Engine for Utility-Owned Storage.

Implements CPUC-style revenue requirement methodology:
- Plant in Service (Gross Plant)
- Book Depreciation (straight-line over book life)
- Tax Depreciation (MACRS 5-year or 7-year schedule)
- Accumulated Deferred Income Taxes (ADIT)
- Net Rate Base = Gross Plant - Accumulated Book Depreciation - ADIT
- Revenue Requirement = Return on Rate Base + Depreciation + Taxes + O&M

References:
    - CPUC Decision D.25-12-003 (SCE 2026-2028 Cost of Capital)
    - IRS Publication 946 (MACRS depreciation schedules)
    - SCE GRC methodology for utility-owned generation assets
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional

# MACRS depreciation percentages by property class
# IRS Publication 946, Table A-1 (200% declining balance, half-year convention)
MACRS_SCHEDULES = {
    5: [0.2000, 0.3200, 0.1920, 0.1152, 0.1152, 0.0576],
    7: [0.1429, 0.2449, 0.1749, 0.1249, 0.0893, 0.0892, 0.0893, 0.0446],
    15: [0.0500, 0.0950, 0.0855, 0.0770, 0.0693, 0.0623,
         0.0590, 0.0590, 0.0591, 0.0590, 0.0591, 0.0590,
         0.0591, 0.0590, 0.0591, 0.0295],
    20: [0.0375, 0.0722, 0.0668, 0.0618, 0.0571, 0.0528,
         0.0489, 0.0452, 0.0447, 0.0447, 0.0446, 0.0446,
         0.0446, 0.0446, 0.0446, 0.0446, 0.0446, 0.0446,
         0.0446, 0.0446, 0.0223],
}


@dataclass
class CostOfCapital:
    """CPUC-authorized cost of capital parameters.

    Default values from CPUC Decision D.25-12-003 (SCE 2026-2028).

    Attributes:
        roe: Authorized return on equity (10.03%).
        cost_of_debt: Embedded cost of long-term debt (4.71%).
        cost_of_preferred: Cost of preferred stock (5.48%).
        equity_ratio: Common equity share of rate base (52.00%).
        debt_ratio: Long-term debt share of rate base (43.47%).
        preferred_ratio: Preferred stock share (4.53%).
        ror: Authorized rate of return (weighted, 7.59%).
        federal_tax_rate: Federal corporate income tax rate.
        state_tax_rate: California corporate income tax rate.
        composite_tax_rate: Combined effective tax rate.
    """

    roe: float = 0.1003          # 10.03% per D.25-12-003
    cost_of_debt: float = 0.0471  # 4.71%
    cost_of_preferred: float = 0.0548  # 5.48%
    equity_ratio: float = 0.5200  # 52.00%
    debt_ratio: float = 0.4347   # 43.47%
    preferred_ratio: float = 0.0453  # 4.53%
    ror: float = 0.0759          # 7.59% authorized ROR
    federal_tax_rate: float = 0.21  # 21% federal
    state_tax_rate: float = 0.0884  # 8.84% California
    property_tax_rate: float = 0.01  # ~1% of assessed value

    @property
    def composite_tax_rate(self) -> float:
        """Combined federal + state tax rate (state is deductible for federal).

        Formula: T_eff = T_state + T_fed * (1 - T_state)
        """
        return self.state_tax_rate + self.federal_tax_rate * (1 - self.state_tax_rate)

    @property
    def net_to_gross_multiplier(self) -> float:
        """Net-to-gross multiplier for grossing up equity return for taxes.

        Formula: NTG = 1 / (1 - T_eff * equity_ratio / (equity_ratio))
        Simplified: NTG = 1 / (1 - T_eff) for the equity portion.

        For SCE ~1.38 based on D.25-12-003.
        """
        t = self.composite_tax_rate
        # Only the equity return is subject to income tax
        # Revenue Requirement = (Rate Base * ROR + Depreciation + O&M) / (1 - T * equity_share_of_return)
        equity_share_of_return = (self.equity_ratio * self.roe) / self.ror if self.ror > 0 else 0
        return 1.0 / (1.0 - t * equity_share_of_return)

    def calculate_ror(self) -> float:
        """Calculate weighted average authorized rate of return.

        Formula: ROR = E% * ROE + D% * COD + P% * COP
        """
        return (self.equity_ratio * self.roe +
                self.debt_ratio * self.cost_of_debt +
                self.preferred_ratio * self.cost_of_preferred)

    def to_dict(self) -> dict:
        return {
            "roe": self.roe,
            "cost_of_debt": self.cost_of_debt,
            "cost_of_preferred": self.cost_of_preferred,
            "equity_ratio": self.equity_ratio,
            "debt_ratio": self.debt_ratio,
            "preferred_ratio": self.preferred_ratio,
            "ror": self.ror,
            "federal_tax_rate": self.federal_tax_rate,
            "state_tax_rate": self.state_tax_rate,
            "property_tax_rate": self.property_tax_rate,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "CostOfCapital":
        data = dict(data)
        return cls(**{k: v for k, v in data.items()
                      if k in cls.__dataclass_fields__})


@dataclass
class RateBaseInputs:
    """Inputs for the utility revenue requirement rate base calculation.

    Attributes:
        gross_plant: Total plant in service ($) - the depreciable capital cost.
        book_life_years: Book depreciation life (straight-line).
        macrs_class: MACRS property class for tax depreciation (5, 7, 15, or 20 year).
        itc_rate: Investment Tax Credit rate (e.g., 0.30 for 30%).
        itc_basis_reduction: Whether ITC reduces depreciable basis for tax (True for ITC).
        cost_of_capital: CPUC-authorized cost of capital parameters.
        annual_om: Annual O&M expense ($) - not capitalized, flows through as expense.
        analysis_years: Number of years for the revenue requirement schedule.
        bonus_depreciation_pct: Bonus depreciation percentage (0-1.0), applied in Year 1.
    """

    gross_plant: float = 0.0
    book_life_years: int = 20
    macrs_class: int = 7  # BESS typically 7-year MACRS
    itc_rate: float = 0.30
    itc_basis_reduction: bool = True  # ITC reduces depreciable tax basis by 50% of ITC
    cost_of_capital: CostOfCapital = field(default_factory=CostOfCapital)
    annual_om: float = 0.0
    analysis_years: int = 20
    bonus_depreciation_pct: float = 0.0  # Bonus depreciation (IRA provisions)

    def to_dict(self) -> dict:
        return {
            "gross_plant": self.gross_plant,
            "book_life_years": self.book_life_years,
            "macrs_class": self.macrs_class,
            "itc_rate": self.itc_rate,
            "itc_basis_reduction": self.itc_basis_reduction,
            "cost_of_capital": self.cost_of_capital.to_dict(),
            "annual_om": self.annual_om,
            "analysis_years": self.analysis_years,
            "bonus_depreciation_pct": self.bonus_depreciation_pct,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "RateBaseInputs":
        data = dict(data)
        if "cost_of_capital" in data and isinstance(data["cost_of_capital"], dict):
            data["cost_of_capital"] = CostOfCapital.from_dict(data["cost_of_capital"])
        return cls(**{k: v for k, v in data.items()
                      if k in cls.__dataclass_fields__})


@dataclass
class AnnualRateBaseResult:
    """Single year of rate base calculation results.

    Attributes:
        year: Year number (1-based).
        gross_plant: Gross plant in service ($).
        book_depreciation: Annual book depreciation expense ($).
        accumulated_book_depreciation: Cumulative book depreciation ($).
        tax_depreciation: Annual tax depreciation (MACRS) ($).
        accumulated_tax_depreciation: Cumulative tax depreciation ($).
        deferred_tax_expense: Annual deferred tax expense ($).
        adit: Accumulated deferred income taxes ($).
        net_rate_base: Net rate base = Gross Plant - Accum Book Depr - ADIT ($).
        return_on_rate_base: Rate base * authorized ROR ($).
        book_depreciation_expense: Book depreciation flowing into revenue requirement ($).
        income_tax_expense: Gross-up for income taxes on equity return ($).
        property_tax_expense: Property tax on net plant ($).
        om_expense: O&M expense ($).
        revenue_requirement: Total annual revenue requirement ($).
    """

    year: int = 0
    gross_plant: float = 0.0
    book_depreciation: float = 0.0
    accumulated_book_depreciation: float = 0.0
    tax_depreciation: float = 0.0
    accumulated_tax_depreciation: float = 0.0
    deferred_tax_expense: float = 0.0
    adit: float = 0.0
    net_rate_base: float = 0.0
    return_on_rate_base: float = 0.0
    book_depreciation_expense: float = 0.0
    income_tax_expense: float = 0.0
    property_tax_expense: float = 0.0
    om_expense: float = 0.0
    revenue_requirement: float = 0.0


@dataclass
class RateBaseResults:
    """Complete rate base and revenue requirement schedule.

    Attributes:
        annual_results: List of AnnualRateBaseResult for each year.
        total_revenue_requirement: Sum of all annual revenue requirements ($).
        levelized_revenue_requirement: Levelized (annuity-equivalent) annual RR ($).
        itc_amount: Investment tax credit amount ($).
        total_book_depreciation: Total book depreciation over analysis period ($).
        total_tax_depreciation: Total tax depreciation over MACRS life ($).
    """

    annual_results: List[AnnualRateBaseResult] = field(default_factory=list)
    total_revenue_requirement: float = 0.0
    levelized_revenue_requirement: float = 0.0
    itc_amount: float = 0.0
    total_book_depreciation: float = 0.0
    total_tax_depreciation: float = 0.0

    def get_annual_revenue_requirements(self) -> List[float]:
        """Return list of annual revenue requirement values."""
        return [r.revenue_requirement for r in self.annual_results]

    def to_dict(self) -> dict:
        return {
            "total_revenue_requirement": self.total_revenue_requirement,
            "levelized_revenue_requirement": self.levelized_revenue_requirement,
            "itc_amount": self.itc_amount,
            "total_book_depreciation": self.total_book_depreciation,
            "total_tax_depreciation": self.total_tax_depreciation,
            "annual_count": len(self.annual_results),
        }


def calculate_book_depreciation(gross_plant: float, book_life: int,
                                 analysis_years: int) -> List[float]:
    """Calculate straight-line book depreciation schedule.

    Args:
        gross_plant: Total depreciable plant cost ($).
        book_life: Depreciation life in years.
        analysis_years: Length of analysis period.

    Returns:
        List of annual book depreciation amounts (length = analysis_years).
    """
    if book_life <= 0:
        return [0.0] * analysis_years

    annual_depr = gross_plant / book_life
    result = []
    for yr in range(1, analysis_years + 1):
        if yr <= book_life:
            result.append(annual_depr)
        else:
            result.append(0.0)
    return result


def calculate_tax_depreciation(depreciable_basis: float, macrs_class: int,
                                analysis_years: int,
                                bonus_pct: float = 0.0) -> List[float]:
    """Calculate MACRS tax depreciation schedule.

    Args:
        depreciable_basis: Tax-depreciable basis ($). May be reduced by ITC.
        macrs_class: MACRS property class (5, 7, 15, or 20).
        analysis_years: Length of analysis period.
        bonus_pct: Bonus depreciation percentage (0-1.0) applied in Year 1.

    Returns:
        List of annual tax depreciation amounts (length = analysis_years).
    """
    if macrs_class not in MACRS_SCHEDULES:
        raise ValueError(f"Unsupported MACRS class: {macrs_class}. "
                         f"Supported: {list(MACRS_SCHEDULES.keys())}")

    schedule = MACRS_SCHEDULES[macrs_class]
    result = [0.0] * analysis_years

    # Apply bonus depreciation first
    bonus_amount = depreciable_basis * bonus_pct
    remaining_basis = depreciable_basis - bonus_amount

    if analysis_years > 0 and bonus_amount > 0:
        result[0] = bonus_amount

    # Apply MACRS to remaining basis
    for i, pct in enumerate(schedule):
        if i < analysis_years:
            result[i] += remaining_basis * pct

    return result


def calculate_adit(book_depreciation: List[float],
                   tax_depreciation: List[float],
                   composite_tax_rate: float) -> List[float]:
    """Calculate Accumulated Deferred Income Taxes (ADIT).

    ADIT arises from timing differences between book and tax depreciation.
    When tax depreciation > book depreciation, ADIT increases (deferred
    tax liability), reducing rate base.

    Args:
        book_depreciation: Annual book depreciation schedule.
        tax_depreciation: Annual tax depreciation schedule.
        composite_tax_rate: Combined federal + state tax rate.

    Returns:
        List of cumulative ADIT balances for each year.
    """
    n = len(book_depreciation)
    adit = [0.0] * n
    cumulative = 0.0

    for i in range(n):
        # Deferred tax = (tax_depr - book_depr) * tax_rate
        timing_diff = tax_depreciation[i] - book_depreciation[i]
        deferred_tax = timing_diff * composite_tax_rate
        cumulative += deferred_tax
        adit[i] = max(0.0, cumulative)  # ADIT cannot be negative for rate base

    return adit


def calculate_revenue_requirement(inputs: RateBaseInputs) -> RateBaseResults:
    """Calculate complete revenue requirement schedule.

    Revenue Requirement for each year:
        RR = Return on Rate Base + Book Depreciation + Income Taxes + Property Tax + O&M

    Where:
        - Return on Rate Base = Net Rate Base * Authorized ROR
        - Net Rate Base = Gross Plant - Accumulated Book Depreciation - ADIT
        - Income Tax = (Equity Return) * Tax Rate / (1 - Tax Rate)
          (gross-up to make equity holders whole after tax)
        - Property Tax = Net Plant * Property Tax Rate

    Args:
        inputs: RateBaseInputs with all parameters.

    Returns:
        RateBaseResults with annual schedule and summary metrics.
    """
    coc = inputs.cost_of_capital
    n = inputs.analysis_years
    gross_plant = inputs.gross_plant

    # Calculate ITC
    itc_amount = gross_plant * inputs.itc_rate

    # Tax-depreciable basis (ITC reduces basis by 50% of credit amount)
    if inputs.itc_basis_reduction:
        tax_basis = gross_plant - (itc_amount * 0.5)
    else:
        tax_basis = gross_plant

    # Calculate depreciation schedules
    book_depr = calculate_book_depreciation(gross_plant, inputs.book_life_years, n)
    tax_depr = calculate_tax_depreciation(
        tax_basis, inputs.macrs_class, n, inputs.bonus_depreciation_pct
    )

    # Calculate ADIT
    composite_tax = coc.composite_tax_rate
    adit_schedule = calculate_adit(book_depr, tax_depr, composite_tax)

    # Build annual results
    annual_results = []
    accum_book_depr = 0.0
    accum_tax_depr = 0.0

    for yr in range(n):
        year_num = yr + 1
        accum_book_depr += book_depr[yr]
        accum_tax_depr += tax_depr[yr]

        # Net rate base
        net_rate_base = gross_plant - accum_book_depr - adit_schedule[yr]
        net_rate_base = max(0.0, net_rate_base)

        # Return on rate base
        return_on_rb = net_rate_base * coc.ror

        # Income tax gross-up on equity return only
        equity_return = net_rate_base * coc.equity_ratio * coc.roe
        # Tax gross-up: equity_return / (1 - tax_rate) - equity_return
        if composite_tax < 1.0:
            income_tax = equity_return * composite_tax / (1.0 - composite_tax)
        else:
            income_tax = 0.0

        # Subtract deferred tax benefit (ADIT change reduces current tax)
        timing_diff = tax_depr[yr] - book_depr[yr]
        deferred_tax_expense = timing_diff * composite_tax
        income_tax -= deferred_tax_expense

        # Property tax on net plant value
        net_plant = gross_plant - accum_book_depr
        property_tax = max(0.0, net_plant) * coc.property_tax_rate

        # O&M expense
        om = inputs.annual_om

        # Total revenue requirement
        rr = return_on_rb + book_depr[yr] + income_tax + property_tax + om

        annual_results.append(AnnualRateBaseResult(
            year=year_num,
            gross_plant=gross_plant,
            book_depreciation=book_depr[yr],
            accumulated_book_depreciation=accum_book_depr,
            tax_depreciation=tax_depr[yr],
            accumulated_tax_depreciation=accum_tax_depr,
            deferred_tax_expense=deferred_tax_expense,
            adit=adit_schedule[yr],
            net_rate_base=net_rate_base,
            return_on_rate_base=return_on_rb,
            book_depreciation_expense=book_depr[yr],
            income_tax_expense=income_tax,
            property_tax_expense=property_tax,
            om_expense=om,
            revenue_requirement=rr,
        ))

    # Summary metrics
    total_rr = sum(r.revenue_requirement for r in annual_results)

    # Levelized revenue requirement (annuity-equivalent)
    if coc.ror > 0 and n > 0:
        annuity_factor = (1 - (1 + coc.ror) ** -n) / coc.ror
        levelized_rr = total_rr / annuity_factor if annuity_factor > 0 else 0.0
        # Actually, levelized = PV(RR) / annuity_factor
        pv_rr = sum(r.revenue_requirement / (1 + coc.ror) ** r.year
                     for r in annual_results)
        levelized_rr = pv_rr / annuity_factor if annuity_factor > 0 else 0.0
    else:
        levelized_rr = total_rr / n if n > 0 else 0.0

    return RateBaseResults(
        annual_results=annual_results,
        total_revenue_requirement=total_rr,
        levelized_revenue_requirement=levelized_rr,
        itc_amount=itc_amount,
        total_book_depreciation=sum(book_depr),
        total_tax_depreciation=sum(tax_depr),
    )
