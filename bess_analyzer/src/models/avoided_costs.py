"""2024 Avoided Cost Calculator (ACC) integration for SCE service territory.

Provides avoided cost trajectories from the E3 2024 ACC:
- Generation Capacity ($/kW-yr): declining trajectory as storage saturates
- Distribution Capacity ($/kW-yr): for NWA-qualifying projects
- Energy Value ($/MWh): time-of-use weighted
- GHG emissions value ($/ton CO2e)
- Ancillary services and transmission deferral

References:
    - E3. 2024 Avoided Cost Calculator. Prepared for CPUC. Version 2a. October 2024.
    - CPUC D.24-06-050. Resource Adequacy Program. 2024.
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional


@dataclass
class GenerationCapacityCost:
    """Avoided generation capacity cost trajectory.

    The ACC provides a declining trajectory for generation capacity value
    as storage market penetration increases. Values start at ~$89.48/kW-yr
    and decline to ~$39/kW-yr over 20 years.

    Attributes:
        values_per_kw_year: Annual avoided cost values ($/kW-yr) for each year.
        start_year: Calendar year for the first value.
    """

    values_per_kw_year: List[float] = field(default_factory=lambda: [
        89.48, 82.00, 75.00, 68.50, 62.50, 57.00, 52.00, 48.00, 44.50, 41.50,
        39.50, 39.00, 39.00, 39.00, 39.00, 39.00, 39.00, 39.00, 39.00, 39.00
    ])
    start_year: int = 2026

    def get_value(self, year_index: int) -> float:
        """Get avoided generation capacity cost for a given year index (0-based).

        If year_index exceeds the trajectory, returns the last value.
        """
        if year_index < 0:
            return 0.0
        if year_index < len(self.values_per_kw_year):
            return self.values_per_kw_year[year_index]
        return self.values_per_kw_year[-1] if self.values_per_kw_year else 0.0

    def get_trajectory(self, n_years: int) -> List[float]:
        """Get full trajectory extended to n_years."""
        result = []
        for i in range(n_years):
            result.append(self.get_value(i))
        return result

    def to_dict(self) -> dict:
        return {
            "values_per_kw_year": list(self.values_per_kw_year),
            "start_year": self.start_year,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "GenerationCapacityCost":
        return cls(
            values_per_kw_year=data.get("values_per_kw_year", []),
            start_year=data.get("start_year", 2026),
        )


@dataclass
class DistributionCapacityCost:
    """Avoided distribution capacity cost.

    Applies only to NWA-qualifying projects where storage defers
    distribution infrastructure investment.

    Attributes:
        value_per_kw_year: Base avoided cost ($/kW-yr).
        escalation_rate: Annual escalation rate.
        applies_to_nwa_only: Whether this cost only applies to NWA projects.
    """

    value_per_kw_year: float = 77.30
    escalation_rate: float = 0.02
    applies_to_nwa_only: bool = True

    def get_trajectory(self, n_years: int) -> List[float]:
        """Get escalating distribution capacity cost trajectory."""
        return [self.value_per_kw_year * (1 + self.escalation_rate) ** i
                for i in range(n_years)]

    def to_dict(self) -> dict:
        return {
            "value_per_kw_year": self.value_per_kw_year,
            "escalation_rate": self.escalation_rate,
            "applies_to_nwa_only": self.applies_to_nwa_only,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "DistributionCapacityCost":
        return cls(**{k: v for k, v in data.items()
                      if k in cls.__dataclass_fields__})


@dataclass
class EnergyValue:
    """Avoided energy cost by time-of-use period.

    Attributes:
        on_peak_per_mwh: On-peak avoided energy cost ($/MWh).
        off_peak_per_mwh: Off-peak avoided energy cost ($/MWh).
        super_off_peak_per_mwh: Super off-peak cost ($/MWh).
        weighted_average_per_mwh: Weighted average across TOU periods ($/MWh).
        escalation_rate: Annual escalation rate.
    """

    on_peak_per_mwh: float = 85.0
    off_peak_per_mwh: float = 35.0
    super_off_peak_per_mwh: float = 15.0
    weighted_average_per_mwh: float = 55.0
    escalation_rate: float = 0.025

    def get_arbitrage_spread(self) -> float:
        """Calculate arbitrage spread (charge off-peak, discharge on-peak)."""
        return self.on_peak_per_mwh - self.off_peak_per_mwh

    def get_trajectory(self, n_years: int) -> List[float]:
        """Get escalating weighted average energy value trajectory."""
        return [self.weighted_average_per_mwh * (1 + self.escalation_rate) ** i
                for i in range(n_years)]

    def to_dict(self) -> dict:
        return {
            "on_peak_per_mwh": self.on_peak_per_mwh,
            "off_peak_per_mwh": self.off_peak_per_mwh,
            "super_off_peak_per_mwh": self.super_off_peak_per_mwh,
            "weighted_average_per_mwh": self.weighted_average_per_mwh,
            "escalation_rate": self.escalation_rate,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "EnergyValue":
        return cls(**{k: v for k, v in data.items()
                      if k in cls.__dataclass_fields__})


@dataclass
class AvoidedCosts:
    """Complete avoided cost inputs from the 2024 ACC.

    Aggregates all ACC cost categories into a single structure.

    Attributes:
        generation_capacity: Generation capacity avoided costs.
        distribution_capacity: Distribution capacity avoided costs (NWA only).
        energy_value: Energy value by TOU period.
        ghg_value_per_ton: GHG avoided cost ($/ton CO2e).
        ghg_escalation_rate: Annual GHG price escalation.
        ghg_emission_factor: Grid emission factor (tons CO2e/MWh).
        ancillary_per_kw_year: Ancillary services value ($/kW-yr).
        ancillary_escalation: Ancillary services escalation rate.
        transmission_per_kw_year: Transmission deferral value ($/kW-yr).
        transmission_escalation: Transmission deferral escalation rate.
    """

    generation_capacity: GenerationCapacityCost = field(
        default_factory=GenerationCapacityCost
    )
    distribution_capacity: DistributionCapacityCost = field(
        default_factory=DistributionCapacityCost
    )
    energy_value: EnergyValue = field(default_factory=EnergyValue)

    ghg_value_per_ton: float = 52.0
    ghg_escalation_rate: float = 0.03
    ghg_emission_factor: float = 0.35  # tons CO2e per MWh

    ancillary_per_kw_year: float = 10.0
    ancillary_escalation: float = 0.01

    transmission_per_kw_year: float = 25.0
    transmission_escalation: float = 0.015

    def calculate_total_avoided_cost(
        self,
        capacity_kw: float,
        annual_discharge_mwh: float,
        year_index: int,
        include_distribution: bool = False,
    ) -> float:
        """Calculate total avoided cost for a single year.

        Args:
            capacity_kw: Project capacity in kW.
            annual_discharge_mwh: Annual energy discharged in MWh.
            year_index: Year index (0-based) into the analysis period.
            include_distribution: Whether to include distribution capacity
                (only for NWA-qualifying projects).

        Returns:
            Total avoided cost for the year ($).
        """
        total = 0.0

        # Generation capacity
        gen_cap = self.generation_capacity.get_value(year_index)
        total += gen_cap * capacity_kw

        # Distribution capacity (NWA only)
        if include_distribution:
            dist_cap = self.distribution_capacity.value_per_kw_year * (
                1 + self.distribution_capacity.escalation_rate
            ) ** year_index
            total += dist_cap * capacity_kw

        # Energy value
        energy_val = self.energy_value.weighted_average_per_mwh * (
            1 + self.energy_value.escalation_rate
        ) ** year_index
        total += energy_val * annual_discharge_mwh

        # GHG value
        ghg_per_mwh = (self.ghg_value_per_ton * self.ghg_emission_factor *
                       (1 + self.ghg_escalation_rate) ** year_index)
        total += ghg_per_mwh * annual_discharge_mwh

        # Ancillary services
        anc = self.ancillary_per_kw_year * (1 + self.ancillary_escalation) ** year_index
        total += anc * capacity_kw

        # Transmission deferral
        trans = self.transmission_per_kw_year * (1 + self.transmission_escalation) ** year_index
        total += trans * capacity_kw

        return total

    def get_annual_avoided_costs(
        self,
        capacity_kw: float,
        capacity_mwh: float,
        rte: float,
        degradation_rate: float,
        cycles_per_day: float,
        n_years: int,
        include_distribution: bool = False,
    ) -> List[float]:
        """Calculate avoided cost trajectory over the analysis period.

        Args:
            capacity_kw: Project capacity in kW.
            capacity_mwh: Project energy capacity in MWh.
            rte: Round-trip efficiency.
            degradation_rate: Annual capacity degradation rate.
            cycles_per_day: Daily charge-discharge cycles.
            n_years: Analysis period length.
            include_distribution: Include distribution capacity value.

        Returns:
            List of annual avoided costs ($) for each year (1..N).
        """
        result = []
        for yr in range(n_years):
            deg_factor = (1 - degradation_rate) ** yr
            annual_discharge = capacity_mwh * cycles_per_day * 365 * rte * deg_factor
            avoided = self.calculate_total_avoided_cost(
                capacity_kw, annual_discharge, yr, include_distribution
            )
            result.append(avoided)
        return result

    def to_dict(self) -> dict:
        return {
            "generation_capacity": self.generation_capacity.to_dict(),
            "distribution_capacity": self.distribution_capacity.to_dict(),
            "energy_value": self.energy_value.to_dict(),
            "ghg_value_per_ton": self.ghg_value_per_ton,
            "ghg_escalation_rate": self.ghg_escalation_rate,
            "ghg_emission_factor": self.ghg_emission_factor,
            "ancillary_per_kw_year": self.ancillary_per_kw_year,
            "ancillary_escalation": self.ancillary_escalation,
            "transmission_per_kw_year": self.transmission_per_kw_year,
            "transmission_escalation": self.transmission_escalation,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "AvoidedCosts":
        d = dict(data)
        gen = d.pop("generation_capacity", None)
        dist = d.pop("distribution_capacity", None)
        energy = d.pop("energy_value", None)

        result = cls(**{k: v for k, v in d.items()
                        if k in cls.__dataclass_fields__})
        if gen:
            result.generation_capacity = GenerationCapacityCost.from_dict(gen)
        if dist:
            result.distribution_capacity = DistributionCapacityCost.from_dict(dist)
        if energy:
            result.energy_value = EnergyValue.from_dict(energy)
        return result
