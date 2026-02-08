"""Assumption library loader for BESS Analyzer.

Loads industry-standard cost and benefit assumptions from JSON files
and applies them to Project objects. Supports NREL ATB, Lazard LCOS,
and CPUC California assumption sets.
"""

import json
import sys
from pathlib import Path
from typing import Dict, List

from src.models.project import BenefitStream, CostInputs, FinancingInputs, Project, TechnologySpecs


def _get_resource_path(relative_path: str) -> Path:
    """Get absolute path to resource, works for dev and for PyInstaller."""
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        # Running as PyInstaller bundle
        base_path = Path(sys._MEIPASS)
    else:
        # Running in development
        base_path = Path(__file__).resolve().parent.parent.parent
    return base_path / relative_path


# Default library directory - handles both dev and bundled modes
_DEFAULT_LIBRARY_DIR = _get_resource_path("resources/libraries")


class AssumptionLibrary:
    """Manages loading and applying assumption libraries.

    Scans a directory for JSON library files and provides methods to
    apply them to Project objects.

    Args:
        library_dir: Path to directory containing library JSON files.
            Defaults to resources/libraries/.
    """

    def __init__(self, library_dir: str = ""):
        self.library_dir = Path(library_dir) if library_dir else _DEFAULT_LIBRARY_DIR
        self._libraries: Dict[str, dict] = {}
        self._load_all()

    def _load_all(self) -> None:
        """Load all JSON files from the library directory."""
        if not self.library_dir.exists():
            return
        for path in sorted(self.library_dir.glob("*.json")):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                key = data.get("name", path.stem)
                self._libraries[key] = data
            except (json.JSONDecodeError, KeyError):
                continue

    def get_library_names(self) -> List[str]:
        """Return sorted list of available library names."""
        return sorted(self._libraries.keys())

    def get_library_metadata(self, name: str) -> Dict[str, str]:
        """Return metadata for a library (source, version, date, notes).

        Args:
            name: Library name as returned by get_library_names().

        Returns:
            Dict with keys: source, version, date_published, url, notes.
        """
        lib = self._libraries.get(name, {})
        return {
            "source": lib.get("source", ""),
            "version": lib.get("version", ""),
            "date_published": lib.get("date_published", ""),
            "url": lib.get("url", ""),
            "notes": lib.get("notes", ""),
        }

    def apply_library_to_project(self, project: Project, library_name: str) -> None:
        """Apply a library's assumptions to a project.

        Updates the project's costs, technology specs, and benefit streams
        in-place. Benefit per-kW values are scaled to the project's capacity
        and escalated over the analysis period.

        Args:
            project: Project to update.
            library_name: Name of the library to apply.

        Raises:
            KeyError: If library_name is not found.
        """
        if library_name not in self._libraries:
            raise KeyError(f"Library '{library_name}' not found. Available: {self.get_library_names()}")

        lib = self._libraries[library_name]
        n = project.basics.analysis_period_years
        capacity_kw = project.basics.capacity_mw * 1000

        # Apply costs (including learning curve projections and infrastructure costs)
        cost_data = lib.get("costs", {})
        cost_projections = lib.get("cost_projections", {})
        project.costs = CostInputs(
            capex_per_kwh=cost_data.get("capex_per_kwh", project.costs.capex_per_kwh),
            fom_per_kw_year=cost_data.get("fom_per_kw_year", project.costs.fom_per_kw_year),
            vom_per_mwh=cost_data.get("vom_per_mwh", project.costs.vom_per_mwh),
            augmentation_per_kwh=cost_data.get("augmentation_per_kwh", project.costs.augmentation_per_kwh),
            decommissioning_per_kw=cost_data.get("decommissioning_per_kw", project.costs.decommissioning_per_kw),
            learning_rate=cost_projections.get("learning_rate", 0.10),
            cost_base_year=cost_projections.get("base_year", 2024),
            # Tax credits (BESS-specific)
            itc_percent=cost_data.get("itc_percent", 0.30),
            itc_adders=cost_data.get("itc_adders", 0.0),
            # Infrastructure costs (common to all utility projects)
            interconnection_per_kw=cost_data.get("interconnection_per_kw", 100.0),
            land_per_kw=cost_data.get("land_per_kw", 10.0),
            permitting_per_kw=cost_data.get("permitting_per_kw", 15.0),
            insurance_pct_of_capex=cost_data.get("insurance_pct_of_capex", 0.005),
            property_tax_pct=cost_data.get("property_tax_pct", 0.01),
            # Charging cost and residual value
            charging_cost_per_mwh=cost_data.get("charging_cost_per_mwh", 30.0),
            residual_value_pct=cost_data.get("residual_value_pct", 0.10),
        )

        # Apply technology specs
        tech_data = lib.get("technology", {})
        project.technology = TechnologySpecs(
            chemistry=tech_data.get("chemistry", project.technology.chemistry),
            round_trip_efficiency=tech_data.get("round_trip_efficiency", project.technology.round_trip_efficiency),
            degradation_rate_annual=tech_data.get("degradation_rate_annual", project.technology.degradation_rate_annual),
            cycle_life=tech_data.get("cycle_life", project.technology.cycle_life),
            warranty_years=tech_data.get("warranty_years", project.technology.warranty_years),
            augmentation_year=tech_data.get("augmentation_year", project.technology.augmentation_year),
            cycles_per_day=tech_data.get("cycles_per_day", 1.0),
        )

        # Apply financing structure if provided
        financing_data = lib.get("financing", {})
        if financing_data:
            project.financing = FinancingInputs(
                debt_percent=financing_data.get("debt_percent", 0.60),
                interest_rate=financing_data.get("interest_rate", 0.05),
                loan_term_years=financing_data.get("loan_term_years", 15),
                cost_of_equity=financing_data.get("cost_of_equity", 0.10),
                tax_rate=financing_data.get("tax_rate", 0.21),
            )

        # Build benefit streams
        project.benefits = []
        for b_data in lib.get("benefits", []):
            value_per_kw = b_data.get("annual_value_per_kw", 0)
            escalation = b_data.get("escalation", 0)

            # Build annual values: scale by capacity, escalate each year
            annual_values = []
            for t in range(n):
                year_value = value_per_kw * capacity_kw * (1 + escalation) ** t
                annual_values.append(year_value)

            stream = BenefitStream(
                name=b_data.get("name", ""),
                annual_values=annual_values,
                description=b_data.get("description", ""),
                data_source=b_data.get("data_source", ""),
                citation=b_data.get("citation", ""),
            )
            project.benefits.append(stream)

        project.assumption_library = library_name
        project.library_version = lib.get("version", "")
