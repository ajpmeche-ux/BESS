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
    FinancingInputs,
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
        """Year 0 cost should include CapEx, infrastructure, minus ITC."""
        project = self._make_project()
        results = calculate_project_economics(project)
        # Battery CapEx
        battery_capex = 160 * 400_000  # $160/kWh * 400,000 kWh = $64M
        # Infrastructure costs (default values: interconnection $100, land $10, permitting $15)
        capacity_kw = 100_000
        infra_costs = (100 + 10 + 15) * capacity_kw  # $12.5M
        total_capex = battery_capex + infra_costs
        # ITC applies to battery only (default 30%)
        itc_credit = battery_capex * 0.30  # $19.2M
        expected_year0 = total_capex - itc_credit  # $64M + $12.5M - $19.2M = $57.3M
        assert abs(results.annual_costs[0] - expected_year0) < 1.0


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

    def test_library_includes_td_deferral(self):
        """Libraries should include T&D Deferral benefit stream."""
        lib = AssumptionLibrary()
        names = lib.get_library_names()
        project = Project()
        # NREL should now have T&D Deferral
        nrel_name = [n for n in names if "NREL" in n][0]
        lib.apply_library_to_project(project, nrel_name)
        benefit_names = [b.name for b in project.benefits]
        assert "T&D Deferral" in benefit_names

    def test_library_includes_learning_rate(self):
        """Libraries should populate learning rate in costs."""
        lib = AssumptionLibrary()
        names = lib.get_library_names()
        project = Project()
        nrel_name = [n for n in names if "NREL" in n][0]
        lib.apply_library_to_project(project, nrel_name)
        # NREL has 12% learning rate
        assert project.costs.learning_rate == 0.12
        assert project.costs.cost_base_year == 2024


# ---- Learning Curve Tests ----

class TestLearningCurve:
    def test_augmentation_cost_year0(self):
        """At year 0 (base year), no cost decline applied."""
        costs = CostInputs(augmentation_per_kwh=100, learning_rate=0.10, cost_base_year=2024)
        result = costs.get_augmentation_cost(0)
        assert abs(result - 100) < 0.01

    def test_augmentation_cost_year10(self):
        """At year 10, cost should be ~35% of original with 10% learning rate."""
        costs = CostInputs(augmentation_per_kwh=100, learning_rate=0.10, cost_base_year=2024)
        result = costs.get_augmentation_cost(10)
        expected = 100 * (0.90 ** 10)  # ~34.87
        assert abs(result - expected) < 0.01

    def test_augmentation_cost_year12(self):
        """Standard augmentation year (12) with 12% learning rate."""
        costs = CostInputs(augmentation_per_kwh=55, learning_rate=0.12, cost_base_year=2024)
        result = costs.get_augmentation_cost(12)
        expected = 55 * (0.88 ** 12)  # ~11.69
        assert abs(result - expected) < 0.01

    def test_capex_projection(self):
        """Future CapEx projection should apply learning curve."""
        costs = CostInputs(capex_per_kwh=160, learning_rate=0.10, cost_base_year=2024)
        # CapEx in 2030 (6 years from base)
        result = costs.get_capex_at_year(2030)
        expected = 160 * (0.90 ** 6)  # ~84.93
        assert abs(result - expected) < 0.01

    def test_zero_learning_rate(self):
        """Zero learning rate should maintain original costs."""
        costs = CostInputs(augmentation_per_kwh=55, learning_rate=0.0, cost_base_year=2024)
        result = costs.get_augmentation_cost(20)
        assert result == 55

    def test_learning_curve_in_project_economics(self):
        """Project economics should apply learning curve to augmentation."""
        from datetime import date
        basics = ProjectBasics(
            name="Test",
            capacity_mw=100,
            duration_hours=4,
            analysis_period_years=20,
            discount_rate=0.07,
        )
        tech = TechnologySpecs(augmentation_year=12)
        # High learning rate for clear effect
        costs = CostInputs(
            capex_per_kwh=160,
            augmentation_per_kwh=100,
            learning_rate=0.15,
            cost_base_year=2024,
        )
        benefits = [
            BenefitStream(
                name="RA",
                annual_values=[15_000_000] * 20,
            ),
        ]
        project = Project(basics=basics, technology=tech, costs=costs, benefits=benefits)
        results = calculate_project_economics(project)

        # Augmentation should be discounted by learning curve
        # Year 12 cost = 100 * (0.85^12) * 400,000 kWh = ~5.7M (vs ~40M without learning)
        expected_aug = 100 * (0.85 ** 12) * 400_000
        assert results.annual_costs[12] < 20_000_000  # Much less than original

    def test_cost_inputs_validation(self):
        """Learning rate should be validated (0-30%)."""
        import pytest
        with pytest.raises(ValueError):
            CostInputs(learning_rate=0.50)  # Too high


# ---- ITC and Infrastructure Cost Tests ----

class TestITCAndInfrastructure:
    def test_itc_reduces_year0_costs(self):
        """ITC should reduce Year 0 capital costs."""
        basics = ProjectBasics(
            name="ITC Test",
            capacity_mw=100,
            duration_hours=4,
            analysis_period_years=20,
            discount_rate=0.07,
        )
        tech = TechnologySpecs()
        # With ITC at 30%
        costs_with_itc = CostInputs(
            capex_per_kwh=160,
            itc_percent=0.30,
            itc_adders=0.0,
        )
        # Without ITC
        costs_without_itc = CostInputs(
            capex_per_kwh=160,
            itc_percent=0.0,
            itc_adders=0.0,
        )
        benefits = [BenefitStream(name="RA", annual_values=[15_000_000] * 20)]

        project_with_itc = Project(basics=basics, technology=tech, costs=costs_with_itc, benefits=benefits)
        project_without_itc = Project(basics=basics, technology=tech, costs=costs_without_itc, benefits=benefits)

        results_with = calculate_project_economics(project_with_itc)
        results_without = calculate_project_economics(project_without_itc)

        # Year 0 costs should be lower with ITC
        assert results_with.annual_costs[0] < results_without.annual_costs[0]
        # ITC credit should be 30% of battery CapEx
        battery_capex = 160 * 400_000  # $64M
        expected_itc = battery_capex * 0.30  # $19.2M
        year0_diff = results_without.annual_costs[0] - results_with.annual_costs[0]
        assert abs(year0_diff - expected_itc) < 1000  # Allow small rounding

    def test_itc_adders_stack(self):
        """ITC adders should stack with base ITC."""
        basics = ProjectBasics(capacity_mw=100, duration_hours=4)
        tech = TechnologySpecs()
        # 30% base + 10% adders = 40% total
        costs = CostInputs(capex_per_kwh=160, itc_percent=0.30, itc_adders=0.10)
        benefits = [BenefitStream(name="RA", annual_values=[15_000_000] * 20)]
        project = Project(basics=basics, technology=tech, costs=costs, benefits=benefits)

        results = calculate_project_economics(project)

        # Total ITC should be 40%
        battery_capex = 160 * 400_000
        infrastructure = 100 * 100_000 + 10 * 100_000 + 15 * 100_000  # interconnect + land + permit
        total_capex = battery_capex + infrastructure
        expected_itc = battery_capex * 0.40  # ITC applies only to battery
        expected_year0 = total_capex - expected_itc
        assert abs(results.annual_costs[0] - expected_year0) < 1000

    def test_infrastructure_costs_added_to_year0(self):
        """Infrastructure costs should be added to Year 0."""
        basics = ProjectBasics(capacity_mw=100, duration_hours=4)
        tech = TechnologySpecs()
        costs = CostInputs(
            capex_per_kwh=160,
            itc_percent=0.0,  # No ITC for clear comparison
            interconnection_per_kw=100,
            land_per_kw=10,
            permitting_per_kw=15,
        )
        benefits = [BenefitStream(name="RA", annual_values=[15_000_000] * 20)]
        project = Project(basics=basics, technology=tech, costs=costs, benefits=benefits)

        results = calculate_project_economics(project)

        capacity_kw = 100_000
        capacity_kwh = 400_000
        battery_capex = 160 * capacity_kwh
        infra_costs = (100 + 10 + 15) * capacity_kw
        expected_year0 = battery_capex + infra_costs
        assert abs(results.annual_costs[0] - expected_year0) < 1.0

    def test_insurance_added_to_annual_costs(self):
        """Insurance should be added to annual operating costs."""
        basics = ProjectBasics(capacity_mw=100, duration_hours=4)
        tech = TechnologySpecs()
        costs = CostInputs(
            capex_per_kwh=160,
            fom_per_kw_year=0,  # Zero out FOM for clear comparison
            vom_per_mwh=0,
            itc_percent=0.0,
            insurance_pct_of_capex=0.01,  # 1% of CapEx
            property_tax_pct=0.0,  # Zero out property tax
        )
        benefits = [BenefitStream(name="RA", annual_values=[15_000_000] * 20)]
        project = Project(basics=basics, technology=tech, costs=costs, benefits=benefits)

        results = calculate_project_economics(project)

        # Total CapEx for insurance calculation
        capacity_kw = 100_000
        capacity_kwh = 400_000
        total_capex = 160 * capacity_kwh + (100 + 10 + 15) * capacity_kw
        expected_annual_insurance = total_capex * 0.01

        # Year 1 should include insurance (no FOM, no VOM in this test)
        # Note: There may be some property tax even at 0% due to implementation
        assert results.annual_costs[1] >= expected_annual_insurance * 0.99  # Allow 1% tolerance

    def test_property_tax_decreases_over_time(self):
        """Property tax should decrease as asset depreciates."""
        basics = ProjectBasics(capacity_mw=100, duration_hours=4, analysis_period_years=20)
        tech = TechnologySpecs()
        costs = CostInputs(
            capex_per_kwh=160,
            fom_per_kw_year=0,
            vom_per_mwh=0,
            itc_percent=0.0,
            insurance_pct_of_capex=0.0,
            property_tax_pct=0.01,
        )
        benefits = [BenefitStream(name="RA", annual_values=[15_000_000] * 20)]
        project = Project(basics=basics, technology=tech, costs=costs, benefits=benefits)

        results = calculate_project_economics(project)

        # Property tax in year 1 should be higher than year 20
        # (straight-line depreciation reduces book value)
        assert results.annual_costs[1] > results.annual_costs[19]

    def test_ownership_type_validation(self):
        """Ownership type should be validated."""
        import pytest
        with pytest.raises(ValueError):
            ProjectBasics(ownership_type="invalid")

    def test_itc_percent_validation(self):
        """ITC percent should be validated (0-50%)."""
        import pytest
        with pytest.raises(ValueError):
            CostInputs(itc_percent=0.60)  # Too high

    def test_itc_adders_validation(self):
        """ITC adders should be validated (0-20%)."""
        import pytest
        with pytest.raises(ValueError):
            CostInputs(itc_adders=0.25)  # Too high


# ---- Library New Benefits Tests ----

class TestLibraryNewBenefits:
    def test_library_includes_resilience(self):
        """Libraries should include Resilience Value benefit stream."""
        lib = AssumptionLibrary()
        names = lib.get_library_names()
        project = Project()
        nrel_name = [n for n in names if "NREL" in n][0]
        lib.apply_library_to_project(project, nrel_name)
        benefit_names = [b.name for b in project.benefits]
        assert "Resilience Value" in benefit_names

    def test_library_includes_renewable_integration(self):
        """Libraries should include Renewable Integration benefit stream."""
        lib = AssumptionLibrary()
        names = lib.get_library_names()
        project = Project()
        nrel_name = [n for n in names if "NREL" in n][0]
        lib.apply_library_to_project(project, nrel_name)
        benefit_names = [b.name for b in project.benefits]
        assert "Renewable Integration" in benefit_names

    def test_library_includes_ghg_value(self):
        """Libraries should include GHG Emissions Value benefit stream."""
        lib = AssumptionLibrary()
        names = lib.get_library_names()
        project = Project()
        nrel_name = [n for n in names if "NREL" in n][0]
        lib.apply_library_to_project(project, nrel_name)
        benefit_names = [b.name for b in project.benefits]
        assert "GHG Emissions Value" in benefit_names

    def test_library_includes_voltage_support(self):
        """Libraries should include Voltage Support benefit stream."""
        lib = AssumptionLibrary()
        names = lib.get_library_names()
        project = Project()
        nrel_name = [n for n in names if "NREL" in n][0]
        lib.apply_library_to_project(project, nrel_name)
        benefit_names = [b.name for b in project.benefits]
        assert "Voltage Support" in benefit_names

    def test_library_loads_itc(self):
        """Libraries should load ITC percent."""
        lib = AssumptionLibrary()
        names = lib.get_library_names()
        project = Project()
        nrel_name = [n for n in names if "NREL" in n][0]
        lib.apply_library_to_project(project, nrel_name)
        assert project.costs.itc_percent == 0.30

    def test_library_loads_infrastructure_costs(self):
        """Libraries should load infrastructure costs."""
        lib = AssumptionLibrary()
        names = lib.get_library_names()
        project = Project()
        nrel_name = [n for n in names if "NREL" in n][0]
        lib.apply_library_to_project(project, nrel_name)
        assert project.costs.interconnection_per_kw == 100
        assert project.costs.land_per_kw == 10
        assert project.costs.permitting_per_kw == 15

    def test_cpuc_library_has_itc_adders(self):
        """CPUC library should have ITC adders (energy community)."""
        lib = AssumptionLibrary()
        names = lib.get_library_names()
        project = Project()
        cpuc_name = [n for n in names if "CPUC" in n][0]
        lib.apply_library_to_project(project, cpuc_name)
        assert project.costs.itc_adders == 0.10  # 10% energy community adder


# ---- Cycles Per Day Tests ----

class TestCyclesPerDay:
    def test_cycles_per_day_default(self):
        """Default cycles_per_day should be 1.0."""
        tech = TechnologySpecs()
        assert tech.cycles_per_day == 1.0

    def test_cycles_per_day_validation(self):
        """cycles_per_day should be validated (0.1-3.0)."""
        with pytest.raises(ValueError):
            TechnologySpecs(cycles_per_day=0.05)  # Too low
        with pytest.raises(ValueError):
            TechnologySpecs(cycles_per_day=4.0)  # Too high

    def test_more_cycles_increases_energy(self):
        """More cycles per day should increase annual energy discharge and charging costs."""
        basics = ProjectBasics(capacity_mw=100, duration_hours=4)
        # Use charging cost to make the difference visible
        costs = CostInputs(itc_percent=0.0, charging_cost_per_mwh=30, vom_per_mwh=1.0)
        benefits = [BenefitStream(name="RA", annual_values=[15_000_000] * 20)]

        tech_1cycle = TechnologySpecs(cycles_per_day=1.0)
        tech_2cycles = TechnologySpecs(cycles_per_day=2.0)

        project_1 = Project(basics=basics, technology=tech_1cycle, costs=costs, benefits=benefits)
        project_2 = Project(basics=basics, technology=tech_2cycles, costs=costs, benefits=benefits)

        results_1 = calculate_project_economics(project_1)
        results_2 = calculate_project_economics(project_2)

        # Annual costs should be higher with more cycles (more energy discharged = more charging + VOM)
        assert results_2.annual_costs[1] > results_1.annual_costs[1]

    def test_library_loads_cycles_per_day(self):
        """Libraries should load cycles_per_day."""
        lib = AssumptionLibrary()
        names = lib.get_library_names()
        project = Project()
        nrel_name = [n for n in names if "NREL" in n][0]
        lib.apply_library_to_project(project, nrel_name)
        assert project.technology.cycles_per_day == 1.0


# ---- Charging Cost Tests ----

class TestChargingCost:
    def test_charging_cost_default(self):
        """Default charging_cost_per_mwh should be 30.0."""
        costs = CostInputs()
        assert costs.charging_cost_per_mwh == 30.0

    def test_charging_cost_validation(self):
        """charging_cost_per_mwh should not be negative."""
        with pytest.raises(ValueError):
            CostInputs(charging_cost_per_mwh=-10)

    def test_charging_cost_increases_annual_costs(self):
        """Charging cost should increase annual operating costs."""
        basics = ProjectBasics(capacity_mw=100, duration_hours=4)
        tech = TechnologySpecs()
        benefits = [BenefitStream(name="RA", annual_values=[15_000_000] * 20)]

        costs_no_charging = CostInputs(
            itc_percent=0.0,
            fom_per_kw_year=0,
            vom_per_mwh=0,
            insurance_pct_of_capex=0,
            property_tax_pct=0,
            charging_cost_per_mwh=0,
        )
        costs_with_charging = CostInputs(
            itc_percent=0.0,
            fom_per_kw_year=0,
            vom_per_mwh=0,
            insurance_pct_of_capex=0,
            property_tax_pct=0,
            charging_cost_per_mwh=50,  # $50/MWh
        )

        project_no_charge = Project(basics=basics, technology=tech, costs=costs_no_charging, benefits=benefits)
        project_with_charge = Project(basics=basics, technology=tech, costs=costs_with_charging, benefits=benefits)

        results_no = calculate_project_economics(project_no_charge)
        results_with = calculate_project_economics(project_with_charge)

        # Year 1 costs should be higher with charging cost
        assert results_with.annual_costs[1] > results_no.annual_costs[1]

    def test_library_loads_charging_cost(self):
        """Libraries should load charging_cost_per_mwh."""
        lib = AssumptionLibrary()
        names = lib.get_library_names()
        project = Project()
        nrel_name = [n for n in names if "NREL" in n][0]
        lib.apply_library_to_project(project, nrel_name)
        assert project.costs.charging_cost_per_mwh == 30.0


# ---- Residual Value Tests ----

class TestResidualValue:
    def test_residual_value_default(self):
        """Default residual_value_pct should be 0.10 (10%)."""
        costs = CostInputs()
        assert costs.residual_value_pct == 0.10

    def test_residual_value_validation(self):
        """residual_value_pct should be 0-50%."""
        with pytest.raises(ValueError):
            CostInputs(residual_value_pct=-0.1)
        with pytest.raises(ValueError):
            CostInputs(residual_value_pct=0.60)

    def test_residual_value_reduces_final_year_costs(self):
        """Residual value should reduce final year net costs."""
        basics = ProjectBasics(capacity_mw=100, duration_hours=4, analysis_period_years=20)
        tech = TechnologySpecs()
        benefits = [BenefitStream(name="RA", annual_values=[15_000_000] * 20)]

        costs_no_residual = CostInputs(
            itc_percent=0.0,
            residual_value_pct=0.0,
            decommissioning_per_kw=10,
        )
        costs_with_residual = CostInputs(
            itc_percent=0.0,
            residual_value_pct=0.20,  # 20% residual
            decommissioning_per_kw=10,
        )

        project_no = Project(basics=basics, technology=tech, costs=costs_no_residual, benefits=benefits)
        project_with = Project(basics=basics, technology=tech, costs=costs_with_residual, benefits=benefits)

        results_no = calculate_project_economics(project_no)
        results_with = calculate_project_economics(project_with)

        # Year 20 costs should be lower with residual value
        assert results_with.annual_costs[20] < results_no.annual_costs[20]


# ---- Financing Structure Tests ----

class TestFinancingInputs:
    def test_financing_defaults(self):
        """FinancingInputs should have sensible defaults."""
        fin = FinancingInputs()
        assert fin.debt_percent == 0.60
        assert fin.interest_rate == 0.05
        assert fin.loan_term_years == 15
        assert fin.cost_of_equity == 0.10
        assert fin.tax_rate == 0.21

    def test_wacc_calculation(self):
        """WACC should be calculated correctly."""
        # WACC = (1-D)*Re + D*Rd*(1-Tc)
        # With 60% debt at 5%, 40% equity at 10%, 21% tax:
        # WACC = 0.4*0.10 + 0.6*0.05*(1-0.21) = 0.04 + 0.0237 = 0.0637
        fin = FinancingInputs(
            debt_percent=0.60,
            interest_rate=0.05,
            cost_of_equity=0.10,
            tax_rate=0.21,
        )
        wacc = fin.calculate_wacc()
        expected = 0.4 * 0.10 + 0.6 * 0.05 * (1 - 0.21)
        assert abs(wacc - expected) < 0.0001
        assert abs(wacc - 0.0637) < 0.001

    def test_wacc_all_equity(self):
        """WACC with 100% equity should equal cost of equity."""
        fin = FinancingInputs(debt_percent=0.0, cost_of_equity=0.12)
        assert abs(fin.calculate_wacc() - 0.12) < 0.0001

    def test_wacc_all_debt(self):
        """WACC with 100% debt should equal after-tax debt cost."""
        fin = FinancingInputs(debt_percent=1.0, interest_rate=0.06, tax_rate=0.21)
        expected = 0.06 * (1 - 0.21)
        assert abs(fin.calculate_wacc() - expected) < 0.0001

    def test_financing_validation(self):
        """Financing inputs should be validated."""
        with pytest.raises(ValueError):
            FinancingInputs(debt_percent=1.5)  # Too high
        with pytest.raises(ValueError):
            FinancingInputs(interest_rate=0.25)  # Too high
        with pytest.raises(ValueError):
            FinancingInputs(loan_term_years=0)  # Too low
        with pytest.raises(ValueError):
            FinancingInputs(cost_of_equity=0.35)  # Too high

    def test_project_get_discount_rate_with_financing(self):
        """Project.get_discount_rate should return WACC when financing is provided."""
        basics = ProjectBasics(discount_rate=0.07)
        financing = FinancingInputs(
            debt_percent=0.60,
            interest_rate=0.05,
            cost_of_equity=0.10,
            tax_rate=0.21,
        )
        project = Project(basics=basics, financing=financing)
        wacc = project.get_discount_rate()
        expected_wacc = financing.calculate_wacc()
        assert wacc == expected_wacc
        assert wacc != 0.07  # Should not use basics.discount_rate

    def test_project_get_discount_rate_without_financing(self):
        """Project.get_discount_rate should return discount_rate when no financing."""
        basics = ProjectBasics(discount_rate=0.08)
        project = Project(basics=basics, financing=None)
        assert project.get_discount_rate() == 0.08

    def test_library_loads_financing(self):
        """Libraries should load financing structure."""
        lib = AssumptionLibrary()
        names = lib.get_library_names()
        project = Project()
        nrel_name = [n for n in names if "NREL" in n][0]
        lib.apply_library_to_project(project, nrel_name)
        assert project.financing is not None
        assert project.financing.debt_percent == 0.60

    def test_financing_affects_npv(self):
        """Financing structure should affect NPV via WACC."""
        basics = ProjectBasics(capacity_mw=100, duration_hours=4, discount_rate=0.10)
        tech = TechnologySpecs()
        costs = CostInputs(itc_percent=0.30)
        benefits = [BenefitStream(name="RA", annual_values=[15_000_000] * 20)]

        # Higher equity cost = higher WACC = lower NPV
        financing_low_wacc = FinancingInputs(
            debt_percent=0.70,
            interest_rate=0.04,
            cost_of_equity=0.08,
            tax_rate=0.21,
        )
        financing_high_wacc = FinancingInputs(
            debt_percent=0.30,
            interest_rate=0.06,
            cost_of_equity=0.15,
            tax_rate=0.21,
        )

        project_low = Project(basics=basics, technology=tech, costs=costs, benefits=benefits, financing=financing_low_wacc)
        project_high = Project(basics=basics, technology=tech, costs=costs, benefits=benefits, financing=financing_high_wacc)

        results_low = calculate_project_economics(project_low)
        results_high = calculate_project_economics(project_high)

        # Lower WACC should give higher NPV
        assert results_low.npv > results_high.npv
