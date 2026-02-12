"""Slice-of-Day (SOD) Feasibility Check for Resource Adequacy.

Validates whether a BESS project can meet the CPUC Slice-of-Day RA
framework requirements by checking if the battery's energy capacity
and duration can satisfy the 24-hour load profile constraints.

Under the SOD framework (D.23-06-029), RA resources must demonstrate
they can provide capacity during specific hourly slices. A 4-hour
battery can only count for hours where cumulative discharge does
not exceed its energy capacity.

References:
    - CPUC D.23-06-029. Slice-of-Day RA Framework. June 2023.
    - CPUC D.24-02-045. RA Reform Implementation. February 2024.
"""

from dataclasses import dataclass, field
from typing import List, Optional, Tuple


# Default SCE summer peak day load shape (capacity factors by hour 0-23)
# 1.0 = peak demand hour, 0.0 = no RA obligation
DEFAULT_SCE_LOAD_SHAPE = [
    0.00, 0.00, 0.00, 0.00, 0.00, 0.00,  # HE 1-6 (midnight-6am)
    0.00, 0.00, 0.10, 0.20, 0.40, 0.60,  # HE 7-12 (morning ramp)
    0.80, 0.90, 1.00, 1.00, 1.00, 0.95,  # HE 13-18 (afternoon peak)
    0.85, 0.70, 0.50, 0.30, 0.10, 0.00,  # HE 19-24 (evening ramp-down)
]


@dataclass
class SODInputs:
    """Inputs for Slice-of-Day feasibility check.

    Attributes:
        capacity_mw: Battery nameplate power capacity (MW).
        duration_hours: Battery energy duration (hours).
        round_trip_efficiency: AC-AC round-trip efficiency.
        degradation_rate: Annual capacity degradation rate.
        analysis_year: Year within project life to evaluate (for degradation).
        hourly_capacity_factors: 24-hour load shape (capacity factors 0-1).
        min_qualifying_hours: Minimum hours the battery must serve to qualify.
        deration_threshold: Minimum capacity factor to count as a qualifying hour.
    """

    capacity_mw: float = 100.0
    duration_hours: float = 4.0
    round_trip_efficiency: float = 0.85
    degradation_rate: float = 0.025
    analysis_year: int = 1
    hourly_capacity_factors: List[float] = field(
        default_factory=lambda: list(DEFAULT_SCE_LOAD_SHAPE)
    )
    min_qualifying_hours: int = 4
    deration_threshold: float = 0.50

    @property
    def energy_capacity_mwh(self) -> float:
        """Effective energy capacity after degradation."""
        deg = (1 - self.degradation_rate) ** max(0, self.analysis_year - 1)
        return self.capacity_mw * self.duration_hours * deg

    def to_dict(self) -> dict:
        return {
            "capacity_mw": self.capacity_mw,
            "duration_hours": self.duration_hours,
            "round_trip_efficiency": self.round_trip_efficiency,
            "degradation_rate": self.degradation_rate,
            "analysis_year": self.analysis_year,
            "hourly_capacity_factors": list(self.hourly_capacity_factors),
            "min_qualifying_hours": self.min_qualifying_hours,
            "deration_threshold": self.deration_threshold,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "SODInputs":
        data = dict(data)
        return cls(**{k: v for k, v in data.items()
                      if k in cls.__dataclass_fields__})


@dataclass
class SODResult:
    """Results of Slice-of-Day feasibility check.

    Attributes:
        feasible: Whether the BESS meets SOD requirements.
        qualifying_hours: Number of hours the BESS can serve.
        required_hours: Minimum hours required.
        max_continuous_hours: Longest continuous block the battery can serve.
        effective_capacity_mw: Derated capacity for RA counting.
        deration_factor: Ratio of effective to nameplate capacity.
        hourly_dispatch: Hour-by-hour dispatch profile (MW by hour 0-23).
        hourly_soc: Hour-by-hour state of charge (MWh by hour 0-23).
        energy_shortfall_mwh: Energy shortfall if infeasible (0 if feasible).
        notes: Explanatory notes about the result.
    """

    feasible: bool = False
    qualifying_hours: int = 0
    required_hours: int = 4
    max_continuous_hours: int = 0
    effective_capacity_mw: float = 0.0
    deration_factor: float = 0.0
    hourly_dispatch: List[float] = field(default_factory=lambda: [0.0] * 24)
    hourly_soc: List[float] = field(default_factory=lambda: [0.0] * 24)
    energy_shortfall_mwh: float = 0.0
    notes: str = ""

    def to_dict(self) -> dict:
        return {
            "feasible": self.feasible,
            "qualifying_hours": self.qualifying_hours,
            "required_hours": self.required_hours,
            "max_continuous_hours": self.max_continuous_hours,
            "effective_capacity_mw": self.effective_capacity_mw,
            "deration_factor": self.deration_factor,
            "hourly_dispatch": list(self.hourly_dispatch),
            "hourly_soc": list(self.hourly_soc),
            "energy_shortfall_mwh": self.energy_shortfall_mwh,
            "notes": self.notes,
        }


def check_sod_feasibility(inputs: SODInputs) -> SODResult:
    """Check if a BESS project meets Slice-of-Day RA requirements.

    Algorithm:
        1. Identify hours with capacity factor >= deration_threshold.
        2. Sort qualifying hours by capacity factor (descending) to prioritize
           peak hours.
        3. Simulate dispatch: discharge at full power during qualifying hours
           until energy capacity is exhausted.
        4. Count qualifying hours served and compute deration factor.

    Args:
        inputs: SODInputs with battery specs and load shape.

    Returns:
        SODResult with feasibility determination and dispatch profile.
    """
    if len(inputs.hourly_capacity_factors) != 24:
        return SODResult(
            feasible=False,
            notes="Load shape must have exactly 24 hourly values.",
        )

    capacity_mw = inputs.capacity_mw
    energy_mwh = inputs.energy_capacity_mwh
    cf = inputs.hourly_capacity_factors

    # Identify hours requiring RA capacity (above deration threshold)
    demand_hours = []
    for hour, factor in enumerate(cf):
        if factor >= inputs.deration_threshold:
            demand_hours.append((hour, factor))

    # Sort by capacity factor descending (serve peak hours first)
    demand_hours.sort(key=lambda x: x[1], reverse=True)

    # Simulate dispatch
    hourly_dispatch = [0.0] * 24
    hourly_soc = [0.0] * 24
    remaining_energy = energy_mwh
    qualifying_hours = 0

    # Track which hours are served
    served_hours = set()

    for hour, factor in demand_hours:
        # Discharge at full capacity for this hour
        discharge_mw = min(capacity_mw, capacity_mw * factor)
        energy_needed = discharge_mw * 1.0  # 1 hour

        if remaining_energy >= energy_needed:
            hourly_dispatch[hour] = discharge_mw
            remaining_energy -= energy_needed
            qualifying_hours += 1
            served_hours.add(hour)

    # Calculate SOC profile (starting from full charge)
    soc = energy_mwh
    for hour in range(24):
        if hour in served_hours:
            soc -= hourly_dispatch[hour]
        hourly_soc[hour] = max(0.0, soc)

    # Find longest continuous block of served hours
    max_continuous = 0
    current_run = 0
    for hour in range(24):
        if hour in served_hours:
            current_run += 1
            max_continuous = max(max_continuous, current_run)
        else:
            current_run = 0

    # Calculate energy shortfall
    total_demand_energy = sum(
        capacity_mw * factor for hour, factor in demand_hours
    )
    shortfall = max(0.0, total_demand_energy - energy_mwh)

    # Deration factor
    total_demand_hours = len(demand_hours)
    deration = qualifying_hours / total_demand_hours if total_demand_hours > 0 else 0.0
    effective_mw = capacity_mw * deration

    # Feasibility check
    feasible = qualifying_hours >= inputs.min_qualifying_hours

    # Build notes
    notes_parts = []
    if feasible:
        notes_parts.append(
            f"PASS: {qualifying_hours} qualifying hours >= "
            f"{inputs.min_qualifying_hours} required."
        )
    else:
        notes_parts.append(
            f"FAIL: {qualifying_hours} qualifying hours < "
            f"{inputs.min_qualifying_hours} required."
        )

    notes_parts.append(
        f"Battery: {capacity_mw:.0f} MW x {inputs.duration_hours:.0f}h = "
        f"{energy_mwh:.0f} MWh effective capacity "
        f"(Year {inputs.analysis_year}, {inputs.degradation_rate*100:.1f}%/yr degradation)."
    )

    if shortfall > 0:
        notes_parts.append(
            f"Energy shortfall: {shortfall:.1f} MWh. "
            f"Consider longer duration or larger capacity."
        )

    return SODResult(
        feasible=feasible,
        qualifying_hours=qualifying_hours,
        required_hours=inputs.min_qualifying_hours,
        max_continuous_hours=max_continuous,
        effective_capacity_mw=effective_mw,
        deration_factor=deration,
        hourly_dispatch=hourly_dispatch,
        hourly_soc=hourly_soc,
        energy_shortfall_mwh=shortfall,
        notes=" ".join(notes_parts),
    )


def check_sod_over_lifetime(inputs: SODInputs,
                             analysis_years: int = 20) -> List[SODResult]:
    """Run SOD feasibility check for each year of the project life.

    As the battery degrades, it may lose the ability to meet SOD
    requirements in later years.

    Args:
        inputs: Base SODInputs (analysis_year will be overridden).
        analysis_years: Number of years to evaluate.

    Returns:
        List of SODResult for each year (1..N).
    """
    results = []
    for yr in range(1, analysis_years + 1):
        yr_inputs = SODInputs(
            capacity_mw=inputs.capacity_mw,
            duration_hours=inputs.duration_hours,
            round_trip_efficiency=inputs.round_trip_efficiency,
            degradation_rate=inputs.degradation_rate,
            analysis_year=yr,
            hourly_capacity_factors=list(inputs.hourly_capacity_factors),
            min_qualifying_hours=inputs.min_qualifying_hours,
            deration_threshold=inputs.deration_threshold,
        )
        results.append(check_sod_feasibility(yr_inputs))
    return results
