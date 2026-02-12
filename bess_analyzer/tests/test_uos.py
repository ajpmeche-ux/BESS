"""Tests for Utility-Owned Storage (UOS) modules.

Tests cover:
- Rate Base / Revenue Requirement calculations
- Avoided Cost Calculator (ACC) integration
- Wires vs NWA comparison (RECC method)
- Slice-of-Day (SOD) feasibility check
- End-to-end UOS analysis via calculate_uos_analysis
"""

import math
import pytest

from src.models.rate_base import (
    MACRS_SCHEDULES,
    CostOfCapital,
    RateBaseInputs,
    calculate_book_depreciation,
    calculate_tax_depreciation,
    calculate_adit,
    calculate_revenue_requirement,
)
from src.models.avoided_costs import (
    AvoidedCosts,
    GenerationCapacityCost,
    DistributionCapacityCost,
    EnergyValue,
)
from src.models.wires_comparison import (
    WiresAlternative,
    NWAParameters,
    calculate_recc,
    calculate_deferral_value,
    compare_wires_vs_nwa,
)
from src.models.sod_check import (
    SODInputs,
    check_sod_feasibility,
    check_sod_over_lifetime,
    DEFAULT_SCE_LOAD_SHAPE,
)
from src.models.project import (
    Project,
    ProjectBasics,
    TechnologySpecs,
    CostInputs,
    UOSInputs,
    BenefitStream,
)
from src.models.calculations import calculate_uos_analysis


# =============================================================================
# Rate Base Tests
# =============================================================================

class TestCostOfCapital:
    def test_sce_defaults(self):
        """SCE cost of capital defaults should match D.25-12-003."""
        coc = CostOfCapital()
        assert coc.roe == 0.1003
        assert coc.cost_of_debt == 0.0471
        assert coc.equity_ratio == 0.5200
        assert coc.ror == 0.0759

    def test_composite_tax_rate(self):
        """Composite tax rate should be ~27.98%."""
        coc = CostOfCapital()
        # T_eff = 0.0884 + 0.21 * (1 - 0.0884) = 0.0884 + 0.19136 = 0.27976
        assert abs(coc.composite_tax_rate - 0.2798) < 0.001

    def test_calculate_ror(self):
        """Calculated ROR should match weighted average."""
        coc = CostOfCapital()
        expected = (0.52 * 0.1003 + 0.4347 * 0.0471 + 0.0453 * 0.0548)
        assert abs(coc.calculate_ror() - expected) < 0.001

    def test_net_to_gross_multiplier(self):
        """NTG multiplier should be > 1.0 (taxes gross up equity return)."""
        coc = CostOfCapital()
        ntg = coc.net_to_gross_multiplier
        assert 1.1 < ntg < 1.5

    def test_serialization(self):
        """Cost of capital should roundtrip through dict."""
        coc = CostOfCapital(roe=0.12, cost_of_debt=0.05)
        d = coc.to_dict()
        coc2 = CostOfCapital.from_dict(d)
        assert coc2.roe == 0.12
        assert coc2.cost_of_debt == 0.05


class TestBookDepreciation:
    def test_straight_line(self):
        """$1M over 20 years = $50,000/year."""
        depr = calculate_book_depreciation(1_000_000, 20, 20)
        assert len(depr) == 20
        assert all(abs(d - 50_000) < 0.01 for d in depr)

    def test_shorter_analysis(self):
        """Analysis shorter than book life should still work."""
        depr = calculate_book_depreciation(1_000_000, 20, 10)
        assert len(depr) == 10
        assert all(abs(d - 50_000) < 0.01 for d in depr)

    def test_longer_analysis(self):
        """Years beyond book life should have zero depreciation."""
        depr = calculate_book_depreciation(1_000_000, 10, 20)
        assert len(depr) == 20
        # 1M / 10 years = 100K/year for first 10, then 0
        assert abs(depr[0] - 100_000) < 0.01
        assert all(abs(d - 100_000) < 0.01 for d in depr[:10])
        assert all(d == 0.0 for d in depr[10:])

    def test_total_equals_gross_plant(self):
        """Total depreciation should equal gross plant when analysis >= book life."""
        depr = calculate_book_depreciation(500_000, 20, 20)
        assert abs(sum(depr) - 500_000) < 0.01


class TestTaxDepreciation:
    def test_macrs_7_year(self):
        """7-year MACRS should sum to 100% of basis."""
        depr = calculate_tax_depreciation(1_000_000, 7, 20)
        assert abs(sum(depr) - 1_000_000) < 0.01

    def test_macrs_5_year(self):
        """5-year MACRS should sum to 100% of basis."""
        depr = calculate_tax_depreciation(1_000_000, 5, 20)
        assert abs(sum(depr) - 1_000_000) < 0.01

    def test_macrs_percentages(self):
        """MACRS percentages should match IRS schedule."""
        schedule = MACRS_SCHEDULES[7]
        assert abs(schedule[0] - 0.1429) < 0.0001
        assert abs(sum(schedule) - 1.0) < 0.001

    def test_invalid_macrs_class(self):
        """Invalid MACRS class should raise ValueError."""
        with pytest.raises(ValueError):
            calculate_tax_depreciation(1_000_000, 10, 20)

    def test_bonus_depreciation(self):
        """Bonus depreciation should front-load deductions."""
        depr_no_bonus = calculate_tax_depreciation(1_000_000, 7, 20, bonus_pct=0.0)
        depr_bonus = calculate_tax_depreciation(1_000_000, 7, 20, bonus_pct=0.5)
        # Year 1 with bonus should be much larger
        assert depr_bonus[0] > depr_no_bonus[0]
        # Totals should still equal basis
        assert abs(sum(depr_bonus) - 1_000_000) < 0.01


class TestADIT:
    def test_adit_positive_when_tax_exceeds_book(self):
        """ADIT should be positive when tax depreciation > book depreciation."""
        book = [50_000] * 20
        tax = [200_000, 200_000, 200_000, 200_000, 100_000, 50_000, 50_000, 50_000] + [0] * 12
        adit = calculate_adit(book, tax, 0.28)
        # Early years: tax > book, so ADIT increases
        assert adit[0] > 0
        assert adit[1] > adit[0]

    def test_adit_sums_correctly(self):
        """ADIT should be cumulative timing differences * tax rate."""
        book = [100_000] * 5
        tax = [200_000, 150_000, 100_000, 50_000, 0]
        adit = calculate_adit(book, tax, 0.30)
        # Year 1: (200K - 100K) * 0.30 = 30K
        assert abs(adit[0] - 30_000) < 0.01


class TestRevenueRequirement:
    def test_basic_revenue_requirement(self):
        """Revenue requirement should include return, depreciation, taxes, O&M."""
        inputs = RateBaseInputs(
            gross_plant=10_000_000,
            book_life_years=20,
            macrs_class=7,
            itc_rate=0.0,
            cost_of_capital=CostOfCapital(),
            annual_om=500_000,
            analysis_years=20,
        )
        results = calculate_revenue_requirement(inputs)
        assert len(results.annual_results) == 20
        assert results.total_revenue_requirement > 0
        assert results.levelized_revenue_requirement > 0

        # Each year should have positive RR
        for r in results.annual_results:
            assert r.revenue_requirement > 0
            assert r.return_on_rate_base >= 0
            assert r.book_depreciation >= 0

    def test_itc_reduces_tax_basis(self):
        """30% ITC should reduce tax-depreciable basis by 15% (half of credit)."""
        inputs_no_itc = RateBaseInputs(
            gross_plant=10_000_000, itc_rate=0.0,
            cost_of_capital=CostOfCapital(), analysis_years=20,
        )
        inputs_with_itc = RateBaseInputs(
            gross_plant=10_000_000, itc_rate=0.30,
            cost_of_capital=CostOfCapital(), analysis_years=20,
        )
        results_no = calculate_revenue_requirement(inputs_no_itc)
        results_with = calculate_revenue_requirement(inputs_with_itc)

        assert results_with.itc_amount == 3_000_000
        # Tax depreciation with ITC should be lower (reduced basis)
        assert results_with.total_tax_depreciation < results_no.total_tax_depreciation

    def test_rr_declines_over_time(self):
        """Revenue requirement should generally decline as rate base shrinks."""
        inputs = RateBaseInputs(
            gross_plant=10_000_000, cost_of_capital=CostOfCapital(),
            analysis_years=20, annual_om=0,
        )
        results = calculate_revenue_requirement(inputs)
        rr = results.get_annual_revenue_requirements()
        # Year 1 RR should be higher than Year 20 RR
        assert rr[0] > rr[-1]

    def test_serialization(self):
        """RateBaseInputs should roundtrip through dict."""
        inputs = RateBaseInputs(gross_plant=5_000_000, book_life_years=15)
        d = inputs.to_dict()
        inputs2 = RateBaseInputs.from_dict(d)
        assert inputs2.gross_plant == 5_000_000
        assert inputs2.book_life_years == 15


# =============================================================================
# Avoided Cost Tests
# =============================================================================

class TestGenerationCapacity:
    def test_default_trajectory(self):
        """Default trajectory should start at ~$89.48 and decline."""
        gc = GenerationCapacityCost()
        assert gc.get_value(0) == 89.48
        assert gc.get_value(19) == 39.00
        # Should decline
        assert gc.get_value(5) < gc.get_value(0)

    def test_trajectory_extension(self):
        """Values beyond trajectory should return last value."""
        gc = GenerationCapacityCost()
        assert gc.get_value(25) == 39.00

    def test_get_trajectory(self):
        """get_trajectory should return correct length."""
        gc = GenerationCapacityCost()
        traj = gc.get_trajectory(25)
        assert len(traj) == 25
        assert traj[0] == 89.48


class TestAvoidedCosts:
    def test_total_avoided_cost(self):
        """Total avoided cost should include all components."""
        acc = AvoidedCosts()
        total = acc.calculate_total_avoided_cost(
            capacity_kw=100_000,
            annual_discharge_mwh=100_000,
            year_index=0,
        )
        assert total > 0
        # Should include gen cap ($89.48 * 100K) + energy + GHG + AS + Trans
        gen_cap_component = 89.48 * 100_000
        assert total > gen_cap_component  # Other components add to it

    def test_distribution_only_with_flag(self):
        """Distribution capacity should only be included when flagged."""
        acc = AvoidedCosts()
        without_dist = acc.calculate_total_avoided_cost(100_000, 100_000, 0, False)
        with_dist = acc.calculate_total_avoided_cost(100_000, 100_000, 0, True)
        assert with_dist > without_dist

    def test_annual_trajectory(self):
        """Annual avoided costs should form a trajectory over N years."""
        acc = AvoidedCosts()
        annual = acc.get_annual_avoided_costs(
            capacity_kw=100_000, capacity_mwh=400,
            rte=0.85, degradation_rate=0.025,
            cycles_per_day=1.0, n_years=20,
        )
        assert len(annual) == 20
        assert all(a > 0 for a in annual)

    def test_serialization(self):
        """AvoidedCosts should roundtrip through dict."""
        acc = AvoidedCosts(ghg_value_per_ton=60.0)
        d = acc.to_dict()
        acc2 = AvoidedCosts.from_dict(d)
        assert acc2.ghg_value_per_ton == 60.0


# =============================================================================
# Wires vs NWA Tests
# =============================================================================

class TestWiresComparison:
    def test_recc_calculation(self):
        """RECC should convert total RR into levelized annual amount."""
        recc = calculate_recc(10_000_000, 20, 0.0759)
        assert recc > 0
        # Should be roughly total_rr / annuity_factor
        annuity = (1 - 1.0759 ** -20) / 0.0759
        expected = 10_000_000 / annuity
        assert abs(recc - expected) < 0.01

    def test_deferral_value(self):
        """Deferral value should be positive and bounded."""
        dv = calculate_deferral_value(50_000_000, 5, 0.0759)
        assert dv > 0
        assert dv < 50_000_000  # Can't exceed wires cost

    def test_deferral_value_zero_years(self):
        """Zero deferral years should give zero value."""
        dv = calculate_deferral_value(50_000_000, 0, 0.0759)
        assert dv == 0.0

    def test_nwa_comparison(self):
        """NWA comparison should produce valid results."""
        coc = CostOfCapital()
        wires = WiresAlternative(cost_per_kw=500, capacity_kw=100_000)
        nwa = NWAParameters(
            bess_gross_plant=64_000_000,  # $160/kWh * 400MWh
            bess_book_life_years=20,
            bess_macrs_class=7,
            bess_annual_om=2_500_000,
            bess_itc_rate=0.30,
        )
        result = compare_wires_vs_nwa(wires, nwa, coc, 20)

        assert result.wires_recc > 0
        assert result.nwa_recc > 0
        assert len(result.cumulative_savings) == 20
        assert len(result.wires_annual_rr) == 20
        assert len(result.nwa_annual_rr) == 20

    def test_wires_serialization(self):
        """WiresAlternative should roundtrip through dict."""
        w = WiresAlternative(cost_per_kw=600, capacity_kw=50_000)
        d = w.to_dict()
        w2 = WiresAlternative.from_dict(d)
        assert w2.cost_per_kw == 600
        assert w2.total_cost == 600 * 50_000


# =============================================================================
# Slice-of-Day Tests
# =============================================================================

class TestSODFeasibility:
    def test_4hour_battery_passes(self):
        """100MW x 4h battery should pass SOD with default load shape."""
        inputs = SODInputs(
            capacity_mw=100, duration_hours=4,
            min_qualifying_hours=4,
        )
        result = check_sod_feasibility(inputs)
        assert result.feasible
        assert result.qualifying_hours >= 4

    def test_1hour_battery_fails(self):
        """100MW x 1h battery should fail SOD (not enough energy)."""
        inputs = SODInputs(
            capacity_mw=100, duration_hours=1,
            min_qualifying_hours=4,
        )
        result = check_sod_feasibility(inputs)
        assert not result.feasible
        assert result.qualifying_hours < 4

    def test_degradation_reduces_capacity(self):
        """Battery in Year 10 should have less effective capacity."""
        inputs_yr1 = SODInputs(
            capacity_mw=100, duration_hours=4,
            degradation_rate=0.025, analysis_year=1,
        )
        inputs_yr10 = SODInputs(
            capacity_mw=100, duration_hours=4,
            degradation_rate=0.025, analysis_year=10,
        )
        assert inputs_yr10.energy_capacity_mwh < inputs_yr1.energy_capacity_mwh

    def test_hourly_dispatch_profile(self):
        """Dispatch profile should have values only during served hours."""
        inputs = SODInputs(capacity_mw=100, duration_hours=4)
        result = check_sod_feasibility(inputs)
        assert len(result.hourly_dispatch) == 24
        assert len(result.hourly_soc) == 24
        # Some hours should have non-zero dispatch
        assert sum(1 for d in result.hourly_dispatch if d > 0) > 0

    def test_lifetime_check(self):
        """Lifetime SOD check should return results for each year."""
        inputs = SODInputs(capacity_mw=100, duration_hours=4)
        results = check_sod_over_lifetime(inputs, 20)
        assert len(results) == 20
        # Year 1 should pass
        assert results[0].feasible

    def test_invalid_load_shape(self):
        """Non-24-hour load shape should return infeasible."""
        inputs = SODInputs(
            capacity_mw=100, duration_hours=4,
            hourly_capacity_factors=[0.5] * 12,  # Wrong length
        )
        result = check_sod_feasibility(inputs)
        assert not result.feasible

    def test_default_load_shape(self):
        """Default SCE load shape should have 24 values."""
        assert len(DEFAULT_SCE_LOAD_SHAPE) == 24
        assert max(DEFAULT_SCE_LOAD_SHAPE) == 1.0
        assert min(DEFAULT_SCE_LOAD_SHAPE) == 0.0

    def test_serialization(self):
        """SODInputs should roundtrip through dict."""
        inputs = SODInputs(capacity_mw=200, duration_hours=6)
        d = inputs.to_dict()
        inputs2 = SODInputs.from_dict(d)
        assert inputs2.capacity_mw == 200
        assert inputs2.duration_hours == 6


# =============================================================================
# UOS Project Integration Tests
# =============================================================================

class TestUOSInputs:
    def test_defaults(self):
        """UOSInputs should have correct SCE defaults."""
        uos = UOSInputs()
        assert uos.roe == 0.1003
        assert uos.ror == 0.0759
        assert uos.macrs_class == 7
        assert uos.wires_cost_per_kw == 500.0

    def test_serialization(self):
        """UOSInputs should roundtrip through dict."""
        uos = UOSInputs(enabled=True, roe=0.11)
        d = uos.to_dict()
        uos2 = UOSInputs.from_dict(d)
        assert uos2.enabled is True
        assert uos2.roe == 0.11

    def test_project_with_uos(self):
        """Project with UOS inputs should serialize correctly."""
        project = Project(
            basics=ProjectBasics(name="UOS Test", capacity_mw=100),
            uos_inputs=UOSInputs(enabled=True),
        )
        d = project.to_dict()
        assert d["uos_inputs"] is not None
        assert d["uos_inputs"]["enabled"] is True

        loaded = Project.from_dict(d)
        assert loaded.uos_inputs is not None
        assert loaded.uos_inputs.enabled is True


class TestCalculateUOSAnalysis:
    def test_disabled_returns_empty(self):
        """UOS analysis should return empty dict when disabled."""
        project = Project(
            basics=ProjectBasics(name="Test", capacity_mw=100),
        )
        result = calculate_uos_analysis(project)
        assert result == {}

    def test_enabled_returns_results(self):
        """UOS analysis should return all result components when enabled."""
        project = Project(
            basics=ProjectBasics(name="SCE UOS", capacity_mw=100, duration_hours=4),
            technology=TechnologySpecs(),
            costs=CostInputs(capex_per_kwh=160),
            uos_inputs=UOSInputs(enabled=True),
        )
        result = calculate_uos_analysis(project)

        assert "rate_base_results" in result
        assert "avoided_costs_annual" in result
        assert "wires_comparison" in result
        assert "sod_result" in result
        assert "ratepayer_impact" in result
        assert "cumulative_savings" in result
        assert "revenue_requirement_annual" in result

        # Revenue requirement should be a 20-year schedule
        rr = result["revenue_requirement_annual"]
        assert len(rr) == 20
        assert all(r > 0 for r in rr)

        # Avoided costs should be a 20-year schedule
        ac = result["avoided_costs_annual"]
        assert len(ac) == 20
        assert all(a > 0 for a in ac)

        # SOD should pass for 4h battery
        sod = result["sod_result"]
        assert sod.feasible

    def test_ratepayer_impact_consistency(self):
        """Ratepayer impact should equal avoided costs minus revenue requirement."""
        project = Project(
            basics=ProjectBasics(name="Impact Test", capacity_mw=100),
            costs=CostInputs(capex_per_kwh=160),
            uos_inputs=UOSInputs(enabled=True),
        )
        result = calculate_uos_analysis(project)

        rr = result["revenue_requirement_annual"]
        ac = result["avoided_costs_annual"]
        impact = result["ratepayer_impact"]

        for i in range(len(impact)):
            expected = ac[i] - rr[i]
            assert abs(impact[i] - expected) < 0.01

    def test_cumulative_savings_accumulate(self):
        """Cumulative savings should be running sum of ratepayer impact."""
        project = Project(
            basics=ProjectBasics(name="Cumulative Test", capacity_mw=100),
            costs=CostInputs(capex_per_kwh=160),
            uos_inputs=UOSInputs(enabled=True),
        )
        result = calculate_uos_analysis(project)

        impact = result["ratepayer_impact"]
        cumulative = result["cumulative_savings"]

        running = 0.0
        for i in range(len(cumulative)):
            running += impact[i]
            assert abs(cumulative[i] - running) < 0.01
