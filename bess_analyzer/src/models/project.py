"""Data models for BESS Analyzer projects.

Defines dataclasses for project inputs, technology specifications,
cost structures, benefit streams, and financial results. All models
support JSON serialization via to_dict()/from_dict() methods.
"""

from dataclasses import dataclass, field
from datetime import date
from typing import Dict, List, Optional


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
    """

    chemistry: str = "LFP"
    round_trip_efficiency: float = 0.85
    degradation_rate_annual: float = 0.025
    cycle_life: int = 6000
    warranty_years: int = 10
    augmentation_year: int = 12

    def __post_init__(self):
        if not 0.5 <= self.round_trip_efficiency <= 1.0:
            raise ValueError(
                f"round_trip_efficiency must be 0.5-1.0, got {self.round_trip_efficiency}"
            )
        if not 0 <= self.degradation_rate_annual <= 0.10:
            raise ValueError(
                f"degradation_rate_annual must be 0-0.10, got {self.degradation_rate_annual}"
            )

    def to_dict(self) -> dict:
        return {
            "chemistry": self.chemistry,
            "round_trip_efficiency": self.round_trip_efficiency,
            "degradation_rate_annual": self.degradation_rate_annual,
            "cycle_life": self.cycle_life,
            "warranty_years": self.warranty_years,
            "augmentation_year": self.augmentation_year,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "TechnologySpecs":
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

    def __post_init__(self):
        if self.capex_per_kwh < 0:
            raise ValueError(f"capex_per_kwh must be >= 0, got {self.capex_per_kwh}")
        if not 0 <= self.learning_rate <= 0.30:
            raise ValueError(f"learning_rate must be 0-0.30, got {self.learning_rate}")
        if not 0 <= self.itc_percent <= 0.50:
            raise ValueError(f"itc_percent must be 0-0.50, got {self.itc_percent}")
        if not 0 <= self.itc_adders <= 0.20:
            raise ValueError(f"itc_adders must be 0-0.20, got {self.itc_adders}")

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
        }

    @classmethod
    def from_dict(cls, data: dict) -> "FinancialResults":
        return cls(**data)


@dataclass
class Project:
    """Complete BESS project combining all inputs and results.

    Attributes:
        basics: Project identification and sizing.
        technology: Battery technology specifications.
        costs: Cost input parameters.
        benefits: List of benefit/revenue streams.
        results: Calculated financial results (populated after analysis).
        assumption_library: Name of loaded assumption library, if any.
        library_version: Version of loaded assumption library, if any.
    """

    basics: ProjectBasics = field(default_factory=ProjectBasics)
    technology: TechnologySpecs = field(default_factory=TechnologySpecs)
    costs: CostInputs = field(default_factory=CostInputs)
    benefits: List[BenefitStream] = field(default_factory=list)
    results: Optional[FinancialResults] = None
    assumption_library: str = ""
    library_version: str = ""

    def to_dict(self) -> dict:
        return {
            "basics": self.basics.to_dict(),
            "technology": self.technology.to_dict(),
            "costs": self.costs.to_dict(),
            "benefits": [b.to_dict() for b in self.benefits],
            "results": self.results.to_dict() if self.results else None,
            "assumption_library": self.assumption_library,
            "library_version": self.library_version,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "Project":
        return cls(
            basics=ProjectBasics.from_dict(data["basics"]),
            technology=TechnologySpecs.from_dict(data["technology"]),
            costs=CostInputs.from_dict(data["costs"]),
            benefits=[BenefitStream.from_dict(b) for b in data.get("benefits", [])],
            results=FinancialResults.from_dict(data["results"]) if data.get("results") else None,
            assumption_library=data.get("assumption_library", ""),
            library_version=data.get("library_version", ""),
        )
