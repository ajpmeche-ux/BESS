"""Data models for BESS Analyzer projects.

Defines dataclasses for project inputs, technology specifications,
cost structures, benefit streams, and financial results. All models
support JSON serialization via to_dict()/from_dict() methods.
"""

from dataclasses import dataclass, field
from datetime import date
from typing import Dict, List, Optional, Tuple


@dataclass
class ProjectBasics:
    """Basic project identification and sizing parameters.

    Attributes:
        name: Project name (e.g., "Moss Landing BESS").
        project_id: Unique identifier.
        location: Site or market location (e.g., "CAISO NP15").
        capacity_mw: Nameplate power capacity in megawatts.
        duration_hours: Storage duration in hours.
        capacity_mwh: Energy capacity in MWh (auto-calculated).
        in_service_date: Expected commercial operation date.
        analysis_period_years: Economic analysis horizon (default 20).
        discount_rate: Nominal discount rate for NPV (default 0.07).
        ownership_type: "utility" or "merchant" - affects WACC and tax treatment.
    """

    name: str = ""
    project_id: str = ""
    location: str = ""
    capacity_mw: float = 100.0
    duration_hours: float = 4.0
    capacity_mwh: float = 0.0
    in_service_date: date = field(default_factory=lambda: date(2027, 1, 1))
    analysis_period_years: int = 20
    discount_rate: float = 0.07
    ownership_type: str = "utility"  # "utility" or "merchant"

    def __post_init__(self):
        self.capacity_mwh = self.capacity_mw * self.duration_hours
        if self.capacity_mw <= 0:
            raise ValueError(f"capacity_mw must be > 0, got {self.capacity_mw}")
        if self.duration_hours <= 0:
            raise ValueError(f"duration_hours must be > 0, got {self.duration_hours}")
        if self.analysis_period_years < 1:
            raise ValueError(f"analysis_period_years must be >= 1, got {self.analysis_period_years}")
        if not 0 < self.discount_rate < 1:
            raise ValueError(f"discount_rate must be between 0 and 1, got {self.discount_rate}")
        if self.ownership_type not in ("utility", "merchant"):
            raise ValueError(f"ownership_type must be 'utility' or 'merchant', got {self.ownership_type}")

    def to_dict(self) -> dict:
        return {
            "name": self.name,
            "project_id": self.project_id,
            "location": self.location,
            "capacity_mw": self.capacity_mw,
            "duration_hours": self.duration_hours,
            "capacity_mwh": self.capacity_mwh,
            "in_service_date": self.in_service_date.isoformat(),
            "analysis_period_years": self.analysis_period_years,
            "discount_rate": self.discount_rate,
            "ownership_type": self.ownership_type,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "ProjectBasics":
        data = dict(data)
        if "in_service_date" in data and isinstance(data["in_service_date"], str):
            data["in_service_date"] = date.fromisoformat(data["in_service_date"])
        data.pop("capacity_mwh", None)
        data.setdefault("ownership_type", "utility")
        return cls(**data)


@dataclass
class TechnologySpecs:
    """Battery technology specifications.

    Attributes:
        chemistry: Battery chemistry type (LFP, NMC, Other).
        round_trip_efficiency: AC-AC round-trip efficiency (0-1).
        degradation_rate_annual: Annual capacity degradation fraction.
        cycle_life: Number of full-depth cycles before end-of-life.
        warranty_years: Manufacturer warranty period.
        augmentation_year: Year in which battery augmentation occurs.
        cycles_per_day: Average number of full charge-discharge cycles per day.
    """

    chemistry: str = "LFP"
    round_trip_efficiency: float = 0.85
    degradation_rate_annual: float = 0.025
    cycle_life: int = 6000
    warranty_years: int = 10
    augmentation_year: int = 12
    cycles_per_day: float = 1.0

    def __post_init__(self):
        if not 0.5 <= self.round_trip_efficiency <= 1.0:
            raise ValueError(
                f"round_trip_efficiency must be 0.5-1.0, got {self.round_trip_efficiency}"
            )
        if not 0 <= self.degradation_rate_annual <= 0.10:
            raise ValueError(
                f"degradation_rate_annual must be 0-0.10, got {self.degradation_rate_annual}"
            )
        if not 0.1 <= self.cycles_per_day <= 3.0:
            raise ValueError(
                f"cycles_per_day must be 0.1-3.0, got {self.cycles_per_day}"
            )

    def to_dict(self) -> dict:
        return {
            "chemistry": self.chemistry,
            "round_trip_efficiency": self.round_trip_efficiency,
            "degradation_rate_annual": self.degradation_rate_annual,
            "cycle_life": self.cycle_life,
            "warranty_years": self.warranty_years,
            "augmentation_year": self.augmentation_year,
            "cycles_per_day": self.cycles_per_day,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "TechnologySpecs":
        data = dict(data)
        data.setdefault("cycles_per_day", 1.0)
        return cls(**data)


@dataclass
class CostInputs:
    """Project cost parameters.

    All costs are in real (constant) dollars. Includes learning curve
    parameters to model cost declines over time, plus common infrastructure
    costs that apply to all utility projects.

    Attributes:
        capex_per_kwh: Installed capital cost per kWh of energy capacity.
        fom_per_kw_year: Fixed operations & maintenance cost per kW-year.
        vom_per_mwh: Variable operations & maintenance cost per MWh discharged.
        augmentation_per_kwh: Battery replacement/augmentation cost per kWh.
        decommissioning_per_kw: End-of-life decommissioning cost per kW.
        learning_rate: Annual cost decline rate (0.10 = 10% annual decline).
        cost_base_year: Reference year for base costs (learning applied from this year).

        # Tax credits (BESS-specific)
        itc_percent: Investment Tax Credit percentage (0.30 = 30%). IRA base is 30%.
        itc_adders: Additional ITC adders (energy community, domestic content, etc.).

        # One-time costs (Common to all utility projects)
        interconnection_per_kw: Interconnection/network upgrade costs ($/kW).
        land_per_kw: Land acquisition or lease capitalized cost ($/kW).
        permitting_per_kw: Permitting and environmental review costs ($/kW).

        # Annual costs (Common to all utility projects)
        insurance_pct_of_capex: Annual insurance as % of CapEx (0.01 = 1%).
        property_tax_pct: Annual property tax as % of asset value (0.01 = 1%).
        charging_cost_per_mwh: Cost of electricity for charging ($/MWh).

        # End of life
        residual_value_pct: Residual value at end of analysis as % of CapEx (0.10 = 10%).
    """

    capex_per_kwh: float = 160.0
    fom_per_kw_year: float = 25.0
    vom_per_mwh: float = 0.0
    augmentation_per_kwh: float = 55.0
    decommissioning_per_kw: float = 10.0
    learning_rate: float = 0.10
    cost_base_year: int = 2024

    # Tax credits (BESS-specific under IRA)
    itc_percent: float = 0.30  # 30% base ITC
    itc_adders: float = 0.0  # Additional adders (energy community +10%, domestic content +10%)

    # One-time costs (Common to all utility infrastructure projects)
    interconnection_per_kw: float = 100.0  # $/kW - network upgrades, studies
    land_per_kw: float = 10.0  # $/kW - site acquisition
    permitting_per_kw: float = 15.0  # $/kW - permits, environmental review

    # Annual costs (Common to all utility infrastructure projects)
    insurance_pct_of_capex: float = 0.005  # 0.5% of CapEx annually
    property_tax_pct: float = 0.01  # 1% of asset value annually
    charging_cost_per_mwh: float = 30.0  # $/MWh for grid charging

    # End of life value
    residual_value_pct: float = 0.10  # 10% residual value at end of analysis

    # Bulk discount for fleet purchases
    bulk_discount_rate: float = 0.0  # e.g., 0.10 = 10% discount on all costs
    bulk_discount_threshold_mwh: float = 0.0  # Minimum capacity to trigger discount (0 = disabled)

    def __post_init__(self):
        if self.capex_per_kwh < 0:
            raise ValueError(f"capex_per_kwh must be >= 0, got {self.capex_per_kwh}")
        if not 0 <= self.learning_rate <= 0.30:
            raise ValueError(f"learning_rate must be 0-0.30, got {self.learning_rate}")
        if not 0 <= self.itc_percent <= 0.50:
            raise ValueError(f"itc_percent must be 0-0.50, got {self.itc_percent}")
        if not 0 <= self.itc_adders <= 0.20:
            raise ValueError(f"itc_adders must be 0-0.20, got {self.itc_adders}")
        if self.charging_cost_per_mwh < 0:
            raise ValueError(f"charging_cost_per_mwh must be >= 0, got {self.charging_cost_per_mwh}")
        if not 0 <= self.residual_value_pct <= 0.50:
            raise ValueError(f"residual_value_pct must be 0-0.50, got {self.residual_value_pct}")
        if not 0 <= self.bulk_discount_rate <= 0.30:
            raise ValueError(f"bulk_discount_rate must be 0-0.30, got {self.bulk_discount_rate}")
        if self.bulk_discount_threshold_mwh < 0:
            raise ValueError(f"bulk_discount_threshold_mwh must be >= 0, got {self.bulk_discount_threshold_mwh}")

    def get_augmentation_cost(self, years_from_base: int) -> float:
        """Calculate augmentation cost adjusted for learning curve.

        Args:
            years_from_base: Number of years from cost_base_year.

        Returns:
            Adjusted augmentation cost per kWh accounting for cost decline.
        """
        decline_factor = (1 - self.learning_rate) ** max(0, years_from_base)
        return self.augmentation_per_kwh * decline_factor

    def get_capex_at_year(self, year: int) -> float:
        """Calculate CapEx at a future year adjusted for learning curve.

        Useful for fleet expansion or replacement cost analysis.

        Args:
            year: Calendar year (e.g., 2030).

        Returns:
            Projected CapEx per kWh at the specified year.
        """
        years_from_base = year - self.cost_base_year
        decline_factor = (1 - self.learning_rate) ** max(0, years_from_base)
        return self.capex_per_kwh * decline_factor

    def to_dict(self) -> dict:
        return {
            "capex_per_kwh": self.capex_per_kwh,
            "fom_per_kw_year": self.fom_per_kw_year,
            "vom_per_mwh": self.vom_per_mwh,
            "augmentation_per_kwh": self.augmentation_per_kwh,
            "decommissioning_per_kw": self.decommissioning_per_kw,
            "learning_rate": self.learning_rate,
            "cost_base_year": self.cost_base_year,
            "itc_percent": self.itc_percent,
            "itc_adders": self.itc_adders,
            "interconnection_per_kw": self.interconnection_per_kw,
            "land_per_kw": self.land_per_kw,
            "permitting_per_kw": self.permitting_per_kw,
            "insurance_pct_of_capex": self.insurance_pct_of_capex,
            "property_tax_pct": self.property_tax_pct,
            "charging_cost_per_mwh": self.charging_cost_per_mwh,
            "residual_value_pct": self.residual_value_pct,
            "bulk_discount_rate": self.bulk_discount_rate,
            "bulk_discount_threshold_mwh": self.bulk_discount_threshold_mwh,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "CostInputs":
        # Handle legacy data without newer fields
        data = dict(data)
        data.setdefault("learning_rate", 0.10)
        data.setdefault("cost_base_year", 2024)
        data.setdefault("itc_percent", 0.30)
        data.setdefault("itc_adders", 0.0)
        data.setdefault("interconnection_per_kw", 100.0)
        data.setdefault("land_per_kw", 10.0)
        data.setdefault("permitting_per_kw", 15.0)
        data.setdefault("insurance_pct_of_capex", 0.005)
        data.setdefault("property_tax_pct", 0.01)
        data.setdefault("charging_cost_per_mwh", 30.0)
        data.setdefault("residual_value_pct", 0.10)
        data.setdefault("bulk_discount_rate", 0.0)
        data.setdefault("bulk_discount_threshold_mwh", 0.0)
        return cls(**data)


@dataclass
class FinancingInputs:
    """Project financing structure parameters.

    Enables WACC calculation based on debt/equity mix for more accurate
    discount rate derivation. Optional - if not provided, uses project
    discount_rate directly.

    Attributes:
        debt_percent: Percentage of project financed with debt (0.60 = 60%).
        interest_rate: Annual interest rate on debt (0.05 = 5%).
        loan_term_years: Loan amortization period in years.
        cost_of_equity: Required return on equity (0.10 = 10%).
        tax_rate: Corporate tax rate for interest deduction (0.21 = 21%).
    """

    debt_percent: float = 0.60  # 60% debt / 40% equity typical for utility
    interest_rate: float = 0.05  # 5% interest rate
    loan_term_years: int = 15  # 15-year loan term
    cost_of_equity: float = 0.10  # 10% required return on equity
    tax_rate: float = 0.21  # 21% federal corporate tax rate

    def __post_init__(self):
        if not 0 <= self.debt_percent <= 1.0:
            raise ValueError(f"debt_percent must be 0-1.0, got {self.debt_percent}")
        if not 0 <= self.interest_rate <= 0.20:
            raise ValueError(f"interest_rate must be 0-0.20, got {self.interest_rate}")
        if not 1 <= self.loan_term_years <= 30:
            raise ValueError(f"loan_term_years must be 1-30, got {self.loan_term_years}")
        if not 0 <= self.cost_of_equity <= 0.30:
            raise ValueError(f"cost_of_equity must be 0-0.30, got {self.cost_of_equity}")
        if not 0 <= self.tax_rate <= 0.50:
            raise ValueError(f"tax_rate must be 0-0.50, got {self.tax_rate}")

    def calculate_wacc(self) -> float:
        """Calculate weighted average cost of capital.

        Formula:
            WACC = (E/V) * Re + (D/V) * Rd * (1 - Tc)

        Where:
            E/V = equity weight
            Re = cost of equity
            D/V = debt weight
            Rd = cost of debt (interest rate)
            Tc = corporate tax rate

        Returns:
            WACC as a decimal (e.g., 0.07 for 7%).
        """
        equity_weight = 1 - self.debt_percent
        after_tax_debt_cost = self.interest_rate * (1 - self.tax_rate)
        return (equity_weight * self.cost_of_equity) + (self.debt_percent * after_tax_debt_cost)

    def to_dict(self) -> dict:
        return {
            "debt_percent": self.debt_percent,
            "interest_rate": self.interest_rate,
            "loan_term_years": self.loan_term_years,
            "cost_of_equity": self.cost_of_equity,
            "tax_rate": self.tax_rate,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "FinancingInputs":
        data = dict(data)
        data.setdefault("debt_percent", 0.60)
        data.setdefault("interest_rate", 0.05)
        data.setdefault("loan_term_years", 15)
        data.setdefault("cost_of_equity", 0.10)
        data.setdefault("tax_rate", 0.21)
        return cls(**data)


@dataclass
class BenefitStream:
    """A single revenue or benefit stream over the project life.

    Attributes:
        name: Benefit category name (e.g., "Resource Adequacy").
        annual_values: List of annual benefit values in $/year for each year (1..N).
        description: Brief description of the benefit.
        data_source: Source organization or publication.
        citation: Full citation string for reports.
    """

    name: str = ""
    annual_values: List[float] = field(default_factory=list)
    description: str = ""
    data_source: str = ""
    citation: str = ""

    def to_dict(self) -> dict:
        return {
            "name": self.name,
            "annual_values": list(self.annual_values),
            "description": self.description,
            "data_source": self.data_source,
            "citation": self.citation,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "BenefitStream":
        return cls(**data)


@dataclass
class SpecialBenefitInputs:
    """Custom inputs for formula-based benefit calculations.

    These benefits have different calculation methods than standard
    $/kW-year benefits. Each requires specific parameters.

    Reliability Benefits (Avoided Outage Cost):
        Based on customer interruption costs and expected outage hours.
        Formula: outage_hours × customer_cost_per_kwh × capacity_mwh × 1000 × backup_pct

    Safety Benefits (Avoided Incident Cost):
        Based on probability of incidents and associated costs.
        Formula: incident_probability × incident_cost × risk_reduction_factor

    Speed-to-Serve Benefits (Faster Deployment Value):
        One-time benefit for faster deployment compared to alternatives.
        Formula: months_saved × value_per_kw_month × capacity_kw
    """

    # Reliability Benefits (Avoided Outage Cost)
    reliability_enabled: bool = False
    outage_hours_per_year: float = 4.0  # Expected outage hours mitigated per year
    customer_cost_per_kwh: float = 10.0  # Value of lost load ($/kWh)
    backup_capacity_pct: float = 0.50  # Portion of capacity providing backup (0-1)

    # Safety Benefits (Avoided Incident Cost)
    safety_enabled: bool = False
    incident_probability: float = 0.001  # Annual probability of safety incident
    incident_cost: float = 1_000_000.0  # Cost per incident ($)
    risk_reduction_factor: float = 0.50  # Risk reduction from BESS (0-1)

    # Speed-to-Serve Benefits (One-time, not annual)
    speed_enabled: bool = False
    months_saved: int = 24  # Months faster than alternative (e.g., peaker plant)
    value_per_kw_month: float = 5.0  # $/kW-month value of early capacity

    def __post_init__(self):
        if not 0 <= self.backup_capacity_pct <= 1.0:
            raise ValueError(f"backup_capacity_pct must be 0-1.0, got {self.backup_capacity_pct}")
        if not 0 <= self.risk_reduction_factor <= 1.0:
            raise ValueError(f"risk_reduction_factor must be 0-1.0, got {self.risk_reduction_factor}")
        if self.months_saved < 0:
            raise ValueError(f"months_saved must be >= 0, got {self.months_saved}")
        if self.outage_hours_per_year < 0:
            raise ValueError(f"outage_hours_per_year must be >= 0, got {self.outage_hours_per_year}")
        if self.customer_cost_per_kwh < 0:
            raise ValueError(f"customer_cost_per_kwh must be >= 0, got {self.customer_cost_per_kwh}")

    def calculate_reliability_annual(self, capacity_mwh: float) -> float:
        """Calculate annual reliability benefit value.

        Args:
            capacity_mwh: Project energy capacity in MWh.

        Returns:
            Annual reliability benefit in $.
        """
        if not self.reliability_enabled:
            return 0.0
        # Annual Value = outage_hours × customer_cost × capacity_kWh × backup_pct
        return (self.outage_hours_per_year * self.customer_cost_per_kwh *
                capacity_mwh * 1000 * self.backup_capacity_pct)

    def calculate_safety_annual(self, capacity_mw: float) -> float:
        """Calculate annual safety benefit value.

        Args:
            capacity_mw: Project power capacity in MW (used as scaling factor).

        Returns:
            Annual safety benefit in $.
        """
        if not self.safety_enabled:
            return 0.0
        # Annual Value = incident_prob × incident_cost × risk_reduction
        # Scaled by capacity as a proxy for risk exposure
        capacity_factor = capacity_mw / 100.0  # Normalize to 100 MW baseline
        return (self.incident_probability * self.incident_cost *
                self.risk_reduction_factor * capacity_factor)

    def calculate_speed_onetime(self, capacity_kw: float) -> float:
        """Calculate one-time speed-to-serve benefit (applied in Year 1).

        Args:
            capacity_kw: Project power capacity in kW.

        Returns:
            One-time speed benefit in $.
        """
        if not self.speed_enabled:
            return 0.0
        # One-time = months_saved × value_per_kw_month × capacity_kw
        return self.months_saved * self.value_per_kw_month * capacity_kw

    def to_dict(self) -> dict:
        return {
            "reliability_enabled": self.reliability_enabled,
            "outage_hours_per_year": self.outage_hours_per_year,
            "customer_cost_per_kwh": self.customer_cost_per_kwh,
            "backup_capacity_pct": self.backup_capacity_pct,
            "safety_enabled": self.safety_enabled,
            "incident_probability": self.incident_probability,
            "incident_cost": self.incident_cost,
            "risk_reduction_factor": self.risk_reduction_factor,
            "speed_enabled": self.speed_enabled,
            "months_saved": self.months_saved,
            "value_per_kw_month": self.value_per_kw_month,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "SpecialBenefitInputs":
        # Only include fields that exist in the dataclass
        valid_fields = {f.name for f in cls.__dataclass_fields__.values()}
        filtered = {k: v for k, v in data.items() if k in valid_fields}
        return cls(**filtered)


@dataclass
class BuildSchedule:
    """Phased build schedule for JIT multi-tranche deployment.

    Each tranche represents a cohort of capacity with its own COD year
    and MW allocation. CapEx for each tranche is determined by the
    learning curve at its COD year.

    Attributes:
        tranches: List of (cod_year, capacity_mw) tuples.
    """

    tranches: List[Tuple[int, float]] = field(default_factory=list)

    def __post_init__(self):
        for i, (year, mw) in enumerate(self.tranches):
            if mw <= 0:
                raise ValueError(f"Tranche {i}: capacity_mw must be > 0, got {mw}")
            if year < 2020 or year > 2060:
                raise ValueError(f"Tranche {i}: cod_year must be 2020-2060, got {year}")

    @property
    def total_capacity_mw(self) -> float:
        return sum(mw for _, mw in self.tranches)

    @property
    def first_cod_year(self) -> int:
        return min(y for y, _ in self.tranches) if self.tranches else 0

    @property
    def last_cod_year(self) -> int:
        return max(y for y, _ in self.tranches) if self.tranches else 0

    def to_dict(self) -> dict:
        return {
            "tranches": [{"cod_year": y, "capacity_mw": mw} for y, mw in self.tranches],
        }

    @classmethod
    def from_dict(cls, data: dict) -> "BuildSchedule":
        tranches = [(t["cod_year"], t["capacity_mw"]) for t in data.get("tranches", [])]
        return cls(tranches=tranches)


@dataclass
class TDDeferralInputs:
    """T&D Capital Deferral inputs for phased deployment benefit.

    Formula: PV = K * [1 - ((1+g)/(1+r))^n]

    Attributes:
        deferred_capital_cost: K - total deferred T&D capital cost ($).
        load_growth_rate: g - annual load growth rate (decimal).
        discount_rate: r - discount rate for T&D deferral PV (decimal).
        deferral_years: n - number of years the capital spend is deferred.
    """

    deferred_capital_cost: float = 0.0
    load_growth_rate: float = 0.01
    discount_rate: float = 0.07
    deferral_years: int = 5

    def __post_init__(self):
        if self.deferred_capital_cost < 0:
            raise ValueError(f"deferred_capital_cost must be >= 0, got {self.deferred_capital_cost}")
        if not 0 <= self.load_growth_rate <= 0.20:
            raise ValueError(f"load_growth_rate must be 0-0.20, got {self.load_growth_rate}")
        if not 0 < self.discount_rate < 1:
            raise ValueError(f"discount_rate must be between 0 and 1, got {self.discount_rate}")
        if self.deferral_years < 0:
            raise ValueError(f"deferral_years must be >= 0, got {self.deferral_years}")

    def calculate_deferral_pv(self) -> float:
        """Calculate T&D deferral present value.

        PV = K * [1 - ((1+g)/(1+r))^n]
        """
        if self.deferred_capital_cost <= 0 or self.deferral_years <= 0:
            return 0.0
        ratio = (1 + self.load_growth_rate) / (1 + self.discount_rate)
        return self.deferred_capital_cost * (1 - ratio ** self.deferral_years)

    def to_dict(self) -> dict:
        return {
            "deferred_capital_cost": self.deferred_capital_cost,
            "load_growth_rate": self.load_growth_rate,
            "discount_rate": self.discount_rate,
            "deferral_years": self.deferral_years,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "TDDeferralInputs":
        valid_fields = {f.name for f in cls.__dataclass_fields__.values()}
        filtered = {k: v for k, v in data.items() if k in valid_fields}
        return cls(**filtered)


@dataclass
class FinancialResults:
    """Calculated financial metrics from project economics.

    Attributes:
        pv_benefits: Present value of all benefits ($).
        pv_costs: Present value of all costs ($).
        npv: Net present value ($).
        bcr: Benefit-cost ratio.
        irr: Internal rate of return (decimal), or None.
        payback_years: Simple payback period in years, or None.
        lcos_per_mwh: Levelized cost of storage ($/MWh).
        breakeven_capex_per_kwh: Maximum CapEx for BCR >= 1.0 ($/kWh).
        benefit_breakdown: Category name -> percentage of total PV benefits.
        annual_costs: List of annual cost values for each year (0..N).
        annual_benefits: List of annual benefit values for each year (0..N).
        annual_net: List of annual net cash flow for each year (0..N).
    """

    pv_benefits: float = 0.0
    pv_costs: float = 0.0
    npv: float = 0.0
    bcr: float = 0.0
    irr: Optional[float] = None
    payback_years: Optional[float] = None
    lcos_per_mwh: float = 0.0
    breakeven_capex_per_kwh: float = 0.0
    benefit_breakdown: Dict[str, float] = field(default_factory=dict)
    annual_costs: List[float] = field(default_factory=list)
    annual_benefits: List[float] = field(default_factory=list)
    annual_net: List[float] = field(default_factory=list)
    flexibility_value: float = 0.0
    td_deferral_pv: float = 0.0
    cohort_capex: List[float] = field(default_factory=list)
    num_tranches: int = 1

    def to_dict(self) -> dict:
        return {
            "pv_benefits": self.pv_benefits,
            "pv_costs": self.pv_costs,
            "npv": self.npv,
            "bcr": self.bcr,
            "irr": self.irr,
            "payback_years": self.payback_years,
            "lcos_per_mwh": self.lcos_per_mwh,
            "breakeven_capex_per_kwh": self.breakeven_capex_per_kwh,
            "benefit_breakdown": dict(self.benefit_breakdown),
            "annual_costs": list(self.annual_costs),
            "annual_benefits": list(self.annual_benefits),
            "annual_net": list(self.annual_net),
            "flexibility_value": self.flexibility_value,
            "td_deferral_pv": self.td_deferral_pv,
            "cohort_capex": list(self.cohort_capex),
            "num_tranches": self.num_tranches,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "FinancialResults":
        data = dict(data)
        data.setdefault("flexibility_value", 0.0)
        data.setdefault("td_deferral_pv", 0.0)
        data.setdefault("cohort_capex", [])
        data.setdefault("num_tranches", 1)
        return cls(**data)


@dataclass
class UOSInputs:
    """Utility-Owned Storage (UOS) revenue requirement inputs.

    Contains SCE-specific regulatory parameters for rate base analysis,
    avoided cost calculation, wires comparison, and SOD feasibility.

    Attributes:
        enabled: Whether UOS analysis is active.
        roe: Authorized return on equity (per D.25-12-003).
        cost_of_debt: Embedded cost of long-term debt.
        cost_of_preferred: Cost of preferred stock.
        equity_ratio: Common equity share of capital structure.
        debt_ratio: Long-term debt share.
        preferred_ratio: Preferred stock share.
        ror: Authorized rate of return.
        federal_tax_rate: Federal corporate income tax rate.
        state_tax_rate: California state income tax rate.
        property_tax_rate: Property tax rate on assessed value.
        book_life_years: Book depreciation life for rate base.
        macrs_class: MACRS property class (5, 7, 15, or 20).
        bonus_depreciation_pct: Bonus depreciation percentage.
        wires_cost_per_kw: Traditional wires alternative cost ($/kW).
        wires_book_life: Wires asset book life (years).
        wires_lead_time: Wires project lead time (years).
        nwa_deferral_years: NWA deferral period (years).
        nwa_incrementality: Whether to apply incrementality adjustment.
        sod_min_hours: Minimum qualifying hours for SOD RA.
        sod_deration_threshold: Minimum capacity factor for SOD qualification.
        use_acc_trajectory: Use ACC declining generation capacity trajectory.
    """

    enabled: bool = False

    # Cost of Capital (D.25-12-003 defaults)
    roe: float = 0.1003
    cost_of_debt: float = 0.0471
    cost_of_preferred: float = 0.0548
    equity_ratio: float = 0.5200
    debt_ratio: float = 0.4347
    preferred_ratio: float = 0.0453
    ror: float = 0.0759
    federal_tax_rate: float = 0.21
    state_tax_rate: float = 0.0884
    property_tax_rate: float = 0.01

    # Rate Base parameters
    book_life_years: int = 20
    macrs_class: int = 7
    bonus_depreciation_pct: float = 0.0

    # Wires vs NWA comparison
    wires_cost_per_kw: float = 500.0
    wires_book_life: int = 40
    wires_lead_time: int = 5
    nwa_deferral_years: int = 5
    nwa_incrementality: bool = True

    # Slice-of-Day
    sod_min_hours: int = 4
    sod_deration_threshold: float = 0.50

    # ACC integration
    use_acc_trajectory: bool = True

    def to_dict(self) -> dict:
        return {
            "enabled": self.enabled,
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
            "book_life_years": self.book_life_years,
            "macrs_class": self.macrs_class,
            "bonus_depreciation_pct": self.bonus_depreciation_pct,
            "wires_cost_per_kw": self.wires_cost_per_kw,
            "wires_book_life": self.wires_book_life,
            "wires_lead_time": self.wires_lead_time,
            "nwa_deferral_years": self.nwa_deferral_years,
            "nwa_incrementality": self.nwa_incrementality,
            "sod_min_hours": self.sod_min_hours,
            "sod_deration_threshold": self.sod_deration_threshold,
            "use_acc_trajectory": self.use_acc_trajectory,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "UOSInputs":
        valid_fields = {f.name for f in cls.__dataclass_fields__.values()}
        filtered = {k: v for k, v in data.items() if k in valid_fields}
        return cls(**filtered)


@dataclass
class Project:
    """Complete BESS project combining all inputs and results.

    Attributes:
        basics: Project identification and sizing.
        technology: Battery technology specifications.
        costs: Cost input parameters.
        financing: Project financing structure (optional).
        benefits: List of benefit/revenue streams.
        special_benefits: Formula-based benefits (reliability, safety, speed) (optional).
        uos_inputs: Utility-Owned Storage regulatory inputs (optional).
        results: Calculated financial results (populated after analysis).
        assumption_library: Name of loaded assumption library, if any.
        library_version: Version of loaded assumption library, if any.
    """

    basics: ProjectBasics = field(default_factory=ProjectBasics)
    technology: TechnologySpecs = field(default_factory=TechnologySpecs)
    costs: CostInputs = field(default_factory=CostInputs)
    financing: Optional[FinancingInputs] = None
    benefits: List[BenefitStream] = field(default_factory=list)
    special_benefits: Optional[SpecialBenefitInputs] = None
    uos_inputs: Optional[UOSInputs] = None
    build_schedule: Optional[BuildSchedule] = None
    td_deferral: Optional[TDDeferralInputs] = None
    results: Optional[FinancialResults] = None
    assumption_library: str = ""
    library_version: str = ""

    def get_discount_rate(self) -> float:
        """Return effective discount rate (WACC if financing provided, else basics.discount_rate)."""
        if self.financing:
            return self.financing.calculate_wacc()
        return self.basics.discount_rate

    def is_multi_tranche(self) -> bool:
        """Return True if project uses a phased build schedule with multiple tranches."""
        return (self.build_schedule is not None
                and len(self.build_schedule.tranches) > 1)

    def get_effective_tranches(self) -> List[Tuple[int, float]]:
        """Return tranches list; single tranche from basics if no build_schedule."""
        if self.build_schedule and self.build_schedule.tranches:
            return list(self.build_schedule.tranches)
        return [(self.basics.in_service_date.year, self.basics.capacity_mw)]

    def to_dict(self) -> dict:
        return {
            "basics": self.basics.to_dict(),
            "technology": self.technology.to_dict(),
            "costs": self.costs.to_dict(),
            "financing": self.financing.to_dict() if self.financing else None,
            "benefits": [b.to_dict() for b in self.benefits],
            "special_benefits": self.special_benefits.to_dict() if self.special_benefits else None,
            "uos_inputs": self.uos_inputs.to_dict() if self.uos_inputs else None,
            "build_schedule": self.build_schedule.to_dict() if self.build_schedule else None,
            "td_deferral": self.td_deferral.to_dict() if self.td_deferral else None,
            "results": self.results.to_dict() if self.results else None,
            "assumption_library": self.assumption_library,
            "library_version": self.library_version,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "Project":
        financing_data = data.get("financing")
        special_benefits_data = data.get("special_benefits")
        uos_data = data.get("uos_inputs")
        build_schedule_data = data.get("build_schedule")
        td_deferral_data = data.get("td_deferral")
        return cls(
            basics=ProjectBasics.from_dict(data["basics"]),
            technology=TechnologySpecs.from_dict(data["technology"]),
            costs=CostInputs.from_dict(data["costs"]),
            financing=FinancingInputs.from_dict(financing_data) if financing_data else None,
            benefits=[BenefitStream.from_dict(b) for b in data.get("benefits", [])],
            special_benefits=SpecialBenefitInputs.from_dict(special_benefits_data) if special_benefits_data else None,
            uos_inputs=UOSInputs.from_dict(uos_data) if uos_data else None,
            build_schedule=BuildSchedule.from_dict(build_schedule_data) if build_schedule_data else None,
            td_deferral=TDDeferralInputs.from_dict(td_deferral_data) if td_deferral_data else None,
            results=FinancialResults.from_dict(data["results"]) if data.get("results") else None,
            assumption_library=data.get("assumption_library", ""),
            library_version=data.get("library_version", ""),
        )
