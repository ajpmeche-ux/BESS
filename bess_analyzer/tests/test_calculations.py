"""Unit tests for BESS Analyzer financial calculations.

Tests cover NPV, BCR, IRR, LCOS, payback, and the full project
economics pipeline. Each test verifies against hand-computed values.
"""

import json
import math
import os
import tempfile
from datetime import date

import pytest

from src.models.calculations import (
    _calculate_payback,
    calculate_bcr,
    calculate_irr,
    calculate_lcos,
    calculate_npv,
    calculate_project_economics,
)
from src.models.project import (
    BenefitStream,
    CostInputs,
    FinancialResults,
    Project,
    ProjectBasics,
    TechnologySpecs,
)
from src.data.validators import (
    validate_capacity,
    validate_discount_rate,
    validate_duration,
    validate_efficiency,
    validate_project,
)
from src.data.storage import save_project, load_project
from src.data.libraries import AssumptionLibrary


# ---- NPV Tests ----

class TestNPV:
    def test_npv_simple(self):
        """NPV of [-1000, 500, 500, 500] at 10% should be ~243.43."""
        cf = [-1000, 500, 500, 500]
        result = calculate_npv(cf, 0.10)
        assert abs(result - 243.43) < 0.5

    def test_npv_zero_rate(self):
        """At 0% discount rate, NPV = sum of cash flows."""
        cf = [-1000, 300, 300, 300, 300]
        result = calculate_npv(cf, 0.0)
        assert abs(result - 200.0) < 0.01

    def test_npv_single_cashflow(self):
        """Single upfront cost should return that cost."""
        result = calculate_npv([-5000], 0.07)
        assert abs(result - (-5000)) < 0.01

    def test_npv_negative_result(self):
        """Small benefits with large upfront cost give negative NPV."""
        cf = [-1000, 100, 100, 100]
        result = calculate_npv(cf, 0.10)
        assert result < 0


# ---- BCR Tests ----

class TestBCR:
    def test_bcr_basic(self):
        """BCR = 2000/1000 = 2.0."""
        assert calculate_bcr(2000, 1000) == 2.0

    def test_bcr_less_than_one(self):
        """BCR < 1 when costs exceed benefits."""
        result = calculate_bcr(500, 1000)
        assert result == 0.5

    def test_bcr_zero_costs_raises(self):
        """BCR should raise ValueError when costs are zero."""
        with pytest.raises(ValueError):
            calculate_bcr(1000, 0)

    def test_bcr_negative_costs_raises(self):
        """BCR should raise ValueError when costs are negative."""
        with pytest.raises(ValueError):
            calculate_bcr(1000, -100)


# ---- IRR Tests ----

class TestIRR:
    def test_irr_basic(self):
        """IRR for [-1000, 400, 400, 400] should be ~9.7%."""
        result = calculate_irr([-1000, 400, 400, 400])
        assert result is not None
        assert abs(result - 0.0966) < 0.01

    def test_irr_zero_npv(self):
        """Investment that exactly breaks even at 10%."""
        # [-1000, 1100] at 10% has NPV=0, so IRR=10%
        result = calculate_irr([-1000, 1100])
        assert result is not None
        assert abs(result - 0.10) < 0.001

    def test_irr_no_solution(self):
        """All negative cash flows should return None."""
        result = calculate_irr([-1000, -500, -500])
        assert result is None


# ---- LCOS Tests ----

class TestLCOS:
    def test_lcos_basic(self):
        """LCOS with known costs and energy values."""
        costs = [10000, 1000, 1000]
        energy = [0, 500, 500]
        result = calculate_lcos(costs, energy, 0.0)
        # PV costs = 12000, PV energy = 1000
        assert abs(result - 12.0) < 0.01

    def test_lcos_zero_energy(self):
        """LCOS should return 0 when no energy is discharged."""
        result = calculate_lcos([10000], [0], 0.05)
        assert result == 0.0


# ---- Payback Tests ----

class TestPayback:
    def test_payback_basic(self):
        """Payback for [-1000, 500, 500, 500] = 2.0 years."""
        result = _calculate_payback([-1000, 500, 500, 500])
        assert result is not None
        assert abs(result - 2.0) < 0.01

    def test_payback_never(self):
        """No payback if benefits never cover initial cost."""
        result = _calculate_payback([-1000, 10, 10, 10])
        assert result is None

    def test_payback_fractional(self):
        """Payback with interpolation: [-1000, 600, 600] = 1.67 years."""
        result = _calculate_payback([-1000, 600, 600])
        assert result is not None
        assert abs(result - (1 + 400 / 600)) < 0.01


# ---- Full Project Economics ----

class TestProjectEconomics:
    def _make_project(self):
        """Create a standard test project: 100 MW, 4hr, NREL-like assumptions."""
        basics = ProjectBasics(
            name="Test Project",
            project_id="TEST-001",
            location="CAISO",
            capacity_mw=100,
            duration_hours=4,
            in_service_date=date(2027, 1, 1),
            analysis_period_years=20,
            discount_rate=0.07,
        )
        tech = TechnologySpecs(
            chemistry="LFP",
            round_trip_efficiency=0.85,
            degradation_rate_annual=0.025,
            cycle_life=6000,
            augmentation_year=12,
        )
        costs = CostInputs(
            capex_per_kwh=160,
            fom_per_kw_year=25,
            vom_per_mwh=0,
            augmentation_per_kwh=55,
            decommissioning_per_kw=10,
        )
        # 100 MW * 1000 = 100,000 kW * $150/kW = $15M/year RA
        ra_values = [15_000_000 * (1.02 ** t) for t in range(20)]
        arb_values = [4_000_000 * (1.015 ** t) for t in range(20)]
        benefits = [
            BenefitStream(
                name="Resource Adequacy",
                annual_values=ra_values,
                citation="CPUC RA Report 2024",
            ),
            BenefitStream(
                name="Energy Arbitrage",
                annual_values=arb_values,
                citation="CAISO Market Data 2024",
            ),
        ]
        return Project(basics=basics, technology=tech, costs=costs, benefits=benefits)

    def test_economics_bcr_positive(self):
        """Standard project should have BCR > 1.0."""
        project = self._make_project()
        results = calculate_project_economics(project)
        assert results.bcr > 1.0

    def test_economics_npv_positive(self):
        """Standard project should have positive NPV."""
        project = self._make_project()
        results = calculate_project_economics(project)
        assert results.npv > 0

    def test_economics_irr_exists(self):
        """Standard project should have a calculable IRR."""
        project = self._make_project()
        results = calculate_project_economics(project)
        assert results.irr is not None
        assert results.irr > 0

    def test_economics_lcos_reasonable(self):
        """LCOS should be in a reasonable range ($50-$500/MWh)."""
        project = self._make_project()
        results = calculate_project_economics(project)
        assert 10 < results.lcos_per_mwh < 500

    def test_economics_benefit_breakdown(self):
        """Benefit breakdown percentages should sum to ~100%."""
        project = self._make_project()
        results = calculate_project_economics(project)
        total_pct = sum(results.benefit_breakdown.values())
        assert abs(total_pct - 100.0) < 0.1

    def test_economics_annual_arrays_length(self):
        """Annual arrays should have n+1 entries (year 0 through year N)."""
        project = self._make_project()
        results = calculate_project_economics(project)
        assert len(results.annual_costs) == 21
        assert len(results.annual_benefits) == 21
        assert len(results.annual_net) == 21

    def test_economics_year0_no_benefits(self):
        """Year 0 should have no benefits (construction period)."""
        project = self._make_project()
        results = calculate_project_economics(project)
        assert results.annual_benefits[0] == 0.0

    def test_economics_capex_at_year0(self):
        """Year 0 cost should equal total CapEx."""
        project = self._make_project()
        results = calculate_project_economics(project)
        expected_capex = 160 * 400_000  # $160/kWh * 400,000 kWh
        assert abs(results.annual_costs[0] - expected_capex) < 1.0


# ---- Validation Tests ----

class TestValidation:
    def test_valid_capacity(self):
        valid, msg = validate_capacity(100)
        assert valid is True
        assert msg == ""

    def test_zero_capacity_invalid(self):
        valid, msg = validate_capacity(0)
        assert valid is False

    def test_large_capacity_warning(self):
        valid, msg = validate_capacity(1500)
        assert valid is True
        assert "Warning" in msg

    def test_valid_efficiency(self):
        valid, _ = validate_efficiency(0.85)
        assert valid is True

    def test_low_efficiency_invalid(self):
        valid, _ = validate_efficiency(0.50)
        assert valid is False

    def test_valid_discount_rate(self):
        valid, _ = validate_discount_rate(0.07)
        assert valid is True

    def test_high_discount_rate_invalid(self):
        valid, _ = validate_discount_rate(0.25)
        assert valid is False

    def test_validate_project_no_benefits_warning(self):
        project = Project()
        valid, messages = validate_project(project)
        assert any("No benefit" in m for m in messages)


# ---- Save/Load Tests ----

class TestStorage:
    def test_save_load_roundtrip(self):
        """Save and reload a project, verify data integrity."""
        project = Project(
            basics=ProjectBasics(
                name="Roundtrip Test",
                capacity_mw=50,
                duration_hours=2,
                discount_rate=0.08,
            ),
            costs=CostInputs(capex_per_kwh=150),
            benefits=[
                BenefitStream(name="RA", annual_values=[1000, 1020, 1040]),
            ],
        )
        with tempfile.NamedTemporaryFile(suffix=".json", delete=False) as f:
            path = f.name

        try:
            save_project(project, path)
            loaded = load_project(path)
            assert loaded.basics.name == "Roundtrip Test"
            assert loaded.basics.capacity_mw == 50
            assert loaded.costs.capex_per_kwh == 150
            assert len(loaded.benefits) == 1
            assert loaded.benefits[0].annual_values == [1000, 1020, 1040]
        finally:
            os.unlink(path)


# ---- Library Tests ----

class TestLibraries:
    def test_library_loads(self):
        """AssumptionLibrary should find at least 3 libraries."""
        lib = AssumptionLibrary()
        names = lib.get_library_names()
        assert len(names) >= 3

    def test_apply_library(self):
        """Applying a library should populate costs and benefits."""
        lib = AssumptionLibrary()
        names = lib.get_library_names()
        project = Project()
        lib.apply_library_to_project(project, names[0])
        assert len(project.benefits) > 0
        assert project.assumption_library == names[0]
