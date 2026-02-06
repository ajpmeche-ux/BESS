# BESS Economic Analyzer - User Guide

## For Electric Utility Engineering and Finance Professionals

**Version 2.0 | February 2026**

---

## Table of Contents

1. [Introduction](#1-introduction)
2. [Quick Start](#2-quick-start)
3. [Input Parameters](#3-input-parameters)
4. [Assumption Libraries](#4-assumption-libraries)
5. [Cost Structure](#5-cost-structure)
6. [Benefit Streams](#6-benefit-streams)
7. [Financial Metrics](#7-financial-metrics)
8. [Methodology](#8-methodology)
9. [Interpreting Results](#9-interpreting-results)
10. [References](#10-references)

---

## 1. Introduction

### 1.1 Purpose

The BESS Economic Analyzer is a benefit-cost analysis tool designed to evaluate utility-scale battery energy storage system (BESS) investments. It implements the California Public Utilities Commission (CPUC) Standard Practice Manual methodology and incorporates industry-standard assumptions from NREL, Lazard, and E3.

### 1.2 Key Features

- **Comprehensive Cost Modeling**: Includes battery CapEx, infrastructure costs (interconnection, land, permitting), operating costs, and tax credits
- **Multiple Benefit Streams**: Resource adequacy, energy arbitrage, ancillary services, T&D deferral, resilience, renewable integration, GHG emissions value, and voltage support
- **Technology Learning Curve**: Projects future cost declines based on industry experience curves
- **Utility vs. Merchant Ownership**: Supports different discount rates and tax treatment by ownership structure
- **Categorized Benefits**: Distinguishes between benefits common to all infrastructure projects and those specific to BESS

### 1.3 Intended Use

This tool is designed for:
- Integrated Resource Planning (IRP) screening analysis
- Distribution System Planning (DSP) non-wires alternatives evaluation
- Investment committee presentations
- Rate case filings and regulatory proceedings
- Competitive procurement bid evaluation

---

## 2. Quick Start

### 2.1 Using the Excel Workbook

1. Open `BESS_Analyzer.xlsx` in Microsoft Excel
2. Navigate to the **Inputs** sheet
3. Select an assumption library (NREL, Lazard, or CPUC)
4. Enter project-specific parameters:
   - Capacity (MW)
   - Duration (hours)
   - Location
   - Analysis period
5. Review populated costs and benefits (adjust if needed)
6. View results on the **Results** sheet
7. Generate reports using VBA macros (see VBA_Instructions sheet)

### 2.2 Using the Python Application

```bash
cd bess_analyzer
python -m pytest tests/  # Verify installation
python main.py           # Launch GUI application
```

### 2.3 Programmatic Analysis

```python
from src.models.project import Project, ProjectBasics, CostInputs
from src.models.calculations import calculate_project_economics
from src.data.libraries import AssumptionLibrary

# Create project
basics = ProjectBasics(
    name="Example 100 MW / 400 MWh BESS",
    capacity_mw=100,
    duration_hours=4,
    discount_rate=0.07,
    ownership_type="utility"
)

# Load industry assumptions
lib = AssumptionLibrary()
project = Project(basics=basics)
lib.apply_library_to_project(project, "NREL ATB 2024 - Moderate Scenario")

# Calculate economics
results = calculate_project_economics(project)
print(f"BCR: {results.bcr:.2f}")
print(f"NPV: ${results.npv:,.0f}")
```

---

## 3. Input Parameters

### 3.1 Project Basics

| Parameter | Description | Typical Range | Units |
|-----------|-------------|---------------|-------|
| **Capacity (MW)** | Nameplate power rating | 10 - 500 | MW |
| **Duration (hours)** | Energy-to-power ratio | 2 - 8 | hours |
| **Energy Capacity (MWh)** | Auto-calculated: MW × hours | - | MWh |
| **Analysis Period** | Economic evaluation horizon | 15 - 25 | years |
| **Discount Rate** | Nominal WACC for NPV | 6% - 10% | % |
| **Ownership Type** | Utility or Merchant | - | - |

### 3.2 Ownership Type Impact

| Parameter | Utility-Owned | Merchant |
|-----------|---------------|----------|
| Typical WACC | 6.0% - 7.5% | 8.0% - 12.0% |
| Tax Treatment | Rate-based recovery | Market revenues |
| Risk Profile | Lower (regulated) | Higher (market exposure) |
| ITC Utilization | May require tax equity | Direct utilization |

**Guidance**: For utility-owned projects subject to rate regulation, use the utility's authorized return on equity (ROE) and capital structure to derive WACC. For merchant projects, use project finance hurdle rates reflecting market risk.

---

## 4. Assumption Libraries

### 4.1 NREL ATB 2024 - Moderate Scenario

**Source**: National Renewable Energy Laboratory Annual Technology Baseline 2024

The NREL ATB provides technology cost and performance projections used in capacity expansion models across the industry. The "Moderate" scenario represents the median of 16 industry data sources.

| Parameter | Value | Notes |
|-----------|-------|-------|
| CapEx | $160/kWh | 4-hour LFP system, 2024 dollars |
| Fixed O&M | $25/kW-year | Includes site maintenance, monitoring |
| Round-trip Efficiency | 85% | AC-AC, including inverter losses |
| Degradation | 2.5%/year | Capacity fade |
| Cycle Life | 6,000 cycles | Full depth of discharge equivalent |
| Learning Rate | 12%/year | Annual cost decline |

**Best Used For**: Long-term planning studies, IRP analysis, regulatory filings requiring defensible third-party sources.

### 4.2 Lazard LCOS v10.0 - 2025

**Source**: Lazard's Levelized Cost of Storage Analysis, Version 10.0

Lazard provides financial advisor perspective on storage economics, widely cited in investment banking and project finance.

| Parameter | Value | Notes |
|-----------|-------|-------|
| CapEx | $145/kWh | Reflects 2025 market pricing |
| Fixed O&M | $22/kW-year | Conservative estimate |
| Round-trip Efficiency | 86% | Latest LFP performance |
| Degradation | 2.0%/year | Improved cell chemistry |
| Learning Rate | 10%/year | Conservative projection |

**Best Used For**: Investment committee presentations, competitive bid evaluation, merchant project analysis.

### 4.3 CPUC California 2024

**Source**: California Public Utilities Commission, E3 Avoided Cost Calculator

California-specific assumptions reflecting the state's high renewable penetration, capacity market premiums, and regulatory requirements.

| Parameter | Value | Notes |
|-----------|-------|-------|
| CapEx | $155/kWh | California market premium |
| RA Value | $180/kW-year | Reflects CA capacity shortfall |
| ITC Adders | +10% | Energy community bonus |
| Property Tax | 1.05% | California average |

**Best Used For**: California utility planning, CPUC filings, projects in CAISO footprint.

---

## 5. Cost Structure

### 5.1 Capital Costs (Year 0)

#### Battery System CapEx
The installed cost of the battery energy storage system, including:
- Battery modules and racks
- Power conversion system (inverters)
- Battery management system (BMS)
- Thermal management (HVAC)
- Enclosures and containers
- Balance of plant
- Engineering, procurement, construction (EPC)

**Current Market Range**: $130 - $200/kWh (2024-2025)

#### Infrastructure Costs (Common to All Projects)

| Cost Component | Default | Range | Description |
|----------------|---------|-------|-------------|
| **Interconnection** | $100/kW | $50 - $300/kW | Network upgrades, studies, metering |
| **Land** | $10/kW | $5 - $50/kW | Site acquisition or capitalized lease |
| **Permitting** | $15/kW | $10 - $50/kW | Environmental review, permits |

*Note: These costs apply to any utility infrastructure project (BESS, gas peaker, T&D upgrade) and should be included in like-for-like comparisons.*

### 5.2 Tax Credits (BESS-Specific)

#### Investment Tax Credit (ITC) under Inflation Reduction Act

| ITC Component | Rate | Requirements |
|---------------|------|--------------|
| **Base ITC** | 30% | Standalone storage eligible (as of 2023) |
| **Energy Community Adder** | +10% | Located in energy community |
| **Domestic Content Adder** | +10% | US-manufactured components |
| **Maximum ITC** | 50% | All adders combined |

**Important**: ITC applies only to the battery system cost, not infrastructure costs. For utility-owned projects, ITC may require tax equity partnerships if the utility lacks sufficient tax appetite.

### 5.3 Operating Costs (Years 1-N)

| Cost Component | Typical Value | Escalation | Description |
|----------------|---------------|------------|-------------|
| **Fixed O&M** | $20-30/kW-year | 2%/year | Maintenance, monitoring, site costs |
| **Variable O&M** | $0-2/MWh | 2%/year | Throughput-dependent (minimal for LFP) |
| **Insurance** | 0.5% of CapEx | Fixed | Property and liability coverage |
| **Property Tax** | 1.0% of book value | Declining | Decreases with depreciation |

### 5.4 Major Maintenance (Augmentation)

Battery capacity degrades over time due to:
- Calendar aging (time-based)
- Cycle aging (throughput-based)
- Temperature stress

**Augmentation** replaces degraded modules to restore nameplate capacity, typically in Year 10-12. The cost benefits from the technology learning curve:

```
Augmentation Cost (Year t) = Base Cost × (1 - Learning Rate)^t
```

**Example**: $55/kWh base cost with 12% learning rate
- Year 12 cost: $55 × (0.88)^12 = $11.69/kWh

### 5.5 End-of-Life Costs

| Component | Typical Value | Notes |
|-----------|---------------|-------|
| **Decommissioning** | $10/kW | Site restoration, equipment removal |
| **Battery Recycling** | Net zero to positive | Lithium and cobalt recovery value may offset costs |

---

## 6. Benefit Streams

### 6.1 Benefit Categories

Benefits are categorized to support like-for-like comparisons with traditional alternatives:

- **Common**: Benefits achievable by any resource (BESS, gas peaker, T&D upgrade)
- **BESS-Specific**: Benefits unique to battery storage technology

### 6.2 Resource Adequacy (Common)

**Definition**: Capacity value for meeting system peak demand and reserve margin requirements.

| Market | Typical Value | Trend |
|--------|---------------|-------|
| CAISO (California) | $150-200/kW-year | Increasing |
| PJM | $50-100/kW-year | Volatile |
| ERCOT | $0 (energy-only) | N/A |
| SPP/MISO | $40-80/kW-year | Stable |

**Methodology**: Value equals the avoided cost of the marginal capacity resource (typically a combustion turbine) or the clearing price in organized capacity markets.

**Reference**: CPUC Resource Adequacy Program, D.24-06-050

### 6.3 Energy Arbitrage (Common)

**Definition**: Revenue from charging during low-price periods and discharging during high-price periods.

**Calculation**:
```
Annual Arbitrage = Capacity (MWh) × Cycles/year × Price Spread × RTE
```

**Typical Values**: $30-60/kW-year depending on:
- Market price volatility
- Renewable penetration (duck curve)
- Natural gas prices
- Congestion patterns

**Reference**: Historical LMP data from CAISO OASIS, PJM Data Miner, ERCOT MIS

### 6.4 Ancillary Services (Common)

**Definition**: Revenue from providing grid services including:
- Frequency regulation (fastest response)
- Spinning reserves (10-minute response)
- Non-spinning reserves (30-minute response)
- Flexible ramping product (CAISO)

**Typical Values**: $10-20/kW-year

**Note**: BESS has competitive advantage in regulation due to fast response, but market depth is limited.

**Reference**: FERC Order 841, CAISO ESDER initiatives

### 6.5 T&D Deferral (Common)

**Definition**: Value of deferring or avoiding transmission and distribution infrastructure investments through peak load reduction.

**Methodology** (E3 Avoided Cost Calculator):
```
Deferral Value = Avoided T&D Cost × Deferral Period × BESS Contribution Factor
```

Where:
- **Avoided T&D Cost**: Marginal cost of distribution capacity (~$100-300/kW)
- **Deferral Period**: Years of investment deferral (typically 3-10 years)
- **BESS Contribution Factor**: Effective load carrying capability (70-90%)

**Typical Values**: $15-40/kW-year in constrained areas; $0 in unconstrained areas

**Important**: This value is highly location-specific. Verify with distribution planning that the proposed BESS location has deferral opportunities.

**Reference**: E3 CPUC Avoided Cost Calculator 2024

### 6.6 Resilience Value (Common)

**Definition**: Avoided outage costs and improved reliability for critical loads.

**Methodology** (LBNL Interruption Cost Estimate Calculator):
```
Resilience Value = VOLL × Avoided Outage Hours × Load Served
```

Where:
- **VOLL** (Value of Lost Load): $5,000 - $50,000/MWh depending on customer class
- **Avoided Outage Hours**: Based on historical SAIDI/SAIFI improvements
- **Load Served**: Critical load supported during outages

**Typical Values**: $40-80/kW-year for utility-owned systems supporting critical infrastructure

**Reference**: LBNL ICE Calculator, IEEE 1366 reliability indices

### 6.7 Renewable Integration (BESS-Specific)

**Definition**: Value of avoiding renewable curtailment and improving variable generation utilization.

**Methodology**:
```
Integration Value = Curtailed MWh Avoided × (Wholesale Price + REC Value)
```

**Typical Values**: $15-40/kW-year in high-renewable regions

**Note**: This benefit is specific to storage and should not be included when comparing to gas peakers or T&D upgrades.

**Reference**: NREL Grid Integration Studies, CAISO Oversupply Reports

### 6.8 GHG Emissions Value (BESS-Specific)

**Definition**: Value of avoided CO2 emissions by displacing marginal fossil generation.

**Methodology**:
```
GHG Value = Displaced Emissions (tons) × Carbon Price ($/ton)
```

**Carbon Price Assumptions**:
| Source | 2025 Price | 2030 Price |
|--------|------------|------------|
| CARB Cap-and-Trade | $35/ton | $50/ton |
| EPA Social Cost of Carbon | $51/ton | $62/ton |
| IEA Net Zero Pathway | $75/ton | $130/ton |

**Typical Values**: $10-30/kW-year

**Reference**: EPA Social Cost of Greenhouse Gases Technical Support Document

### 6.9 Voltage Support (Common)

**Definition**: Distribution-level voltage regulation and power quality improvement services.

**Services Provided**:
- Volt-VAR optimization
- Power factor correction
- Harmonic filtering (with advanced inverters)

**Typical Values**: $5-15/kW-year

**Reference**: EPRI Distribution System Studies, IEEE 1547-2018

---

## 7. Financial Metrics

### 7.1 Net Present Value (NPV)

**Formula**:
```
NPV = Σ (CFt / (1+r)^t) for t = 0 to N
```

**Interpretation**:
- NPV > 0: Project creates value
- NPV < 0: Project destroys value
- NPV = 0: Project breaks even

**Use Case**: Primary decision metric for capital budgeting

### 7.2 Benefit-Cost Ratio (BCR)

**Formula**:
```
BCR = PV(Benefits) / PV(Costs)
```

**Interpretation** (CPUC Standard Practice Manual):
| BCR | Recommendation |
|-----|----------------|
| ≥ 1.5 | **Approve** - Strong economic case |
| 1.0 - 1.5 | **Further Study** - Marginal economics |
| < 1.0 | **Reject** - Costs exceed benefits |

**Use Case**: Regulatory proceedings, IRP screening, demand-side management evaluation

### 7.3 Internal Rate of Return (IRR)

**Formula**: The discount rate r where NPV = 0

**Interpretation**:
- IRR > WACC: Project earns above cost of capital
- IRR < WACC: Project earns below cost of capital

**Use Case**: Investment committee presentations, comparison across projects with different scales

### 7.4 Levelized Cost of Storage (LCOS)

**Formula** (Lazard Methodology):
```
LCOS = PV(Lifetime Costs) / PV(Lifetime Energy Discharged)
```

**Components**:
- Capital cost (levelized)
- Fixed O&M (levelized)
- Charging cost (electricity)
- Efficiency losses

**Typical Range**: $100-200/MWh for 4-hour systems

**Use Case**: Technology comparison, procurement bid evaluation

### 7.5 Payback Period

**Formula**: Year t where cumulative cash flow ≥ 0

**Use Case**: Quick screening metric; often used for corporate investment decisions

**Limitation**: Does not account for time value of money or cash flows beyond payback

---

## 8. Methodology

### 8.1 Cash Flow Model Structure

```
Year 0:  - Battery CapEx
         - Infrastructure (interconnection, land, permitting)
         + ITC credit (applies to battery only)
         = Net Capital Cost

Years 1-N: - Fixed O&M
           - Variable O&M
           - Insurance
           - Property Tax (declining)
           + Benefits (escalating)
           = Annual Net Cash Flow

Year 12:   - Augmentation cost (learning curve adjusted)

Year N:    - Decommissioning
```

### 8.2 Degradation Modeling

Battery capacity degrades according to:
```
Capacity(t) = Capacity(0) × (1 - degradation_rate)^t
```

This affects:
- Energy-based benefits (arbitrage, renewable integration)
- Capacity-based benefits if derated (RA in some markets)

### 8.3 Escalation and Inflation

| Category | Default Escalation | Rationale |
|----------|-------------------|-----------|
| Benefits | 1.5% - 3.0% | Electricity price inflation |
| Fixed O&M | 2.0% | General inflation |
| Insurance | 0% (fixed) | Negotiated contracts |
| Property Tax | N/A (declining) | Based on depreciation |

### 8.4 Technology Learning Curve

Battery costs have declined approximately 15% per year over the past decade. Projections:

| Source | Learning Rate | Application |
|--------|---------------|-------------|
| NREL ATB | 12%/year | Planning studies |
| BloombergNEF | 10-15%/year | Market analysis |
| Lazard | 10%/year | Conservative |

The model applies learning curve to:
- Augmentation costs (future battery replacement)
- CapEx projections for fleet expansion analysis

---

## 9. Interpreting Results

### 9.1 Results Dashboard

The Results sheet displays:

| Metric | How to Interpret |
|--------|------------------|
| **BCR** | Primary go/no-go metric for utility projects |
| **NPV** | Dollar value created; compare across projects |
| **IRR** | Must exceed WACC to create value |
| **Payback** | Rule of thumb: <7 years is strong |
| **LCOS** | Compare to avoided cost of alternatives |
| **Breakeven CapEx** | Maximum CapEx where BCR = 1.0 |

### 9.2 Sensitivity Analysis

Recommended sensitivities to test:

| Parameter | Low Case | Base Case | High Case |
|-----------|----------|-----------|-----------|
| CapEx | -20% | Base | +20% |
| Discount Rate | Base - 1% | Base | Base + 1% |
| RA Value | -25% | Base | +25% |
| Learning Rate | 5% | Base | 15% |

### 9.3 Common Issues and Solutions

**BCR < 1.0 (Costs Exceed Benefits)**:
- Verify ITC is being captured
- Check interconnection costs aren't excessive
- Confirm T&D deferral opportunity exists at location
- Consider longer duration (higher RA value per MW)

**IRR < WACC**:
- Benefits may be conservatively estimated
- Consider additional revenue streams (frequency regulation)
- Evaluate merchant operation in adjacent markets

**Payback > 10 Years**:
- Expected for utility-scale BESS
- Focus on NPV and BCR for decision-making
- Long payback is acceptable for rate-based assets

---

## 10. References

### Primary Data Sources

1. **NREL Annual Technology Baseline 2024**
   - URL: https://atb.nrel.gov/electricity/2024/utility-scale_battery_storage
   - Use: Cost projections, technology parameters

2. **Lazard's Levelized Cost of Storage Analysis, Version 10.0**
   - URL: https://www.lazard.com/research-insights/levelized-cost-of-storage/
   - Use: LCOS methodology, merchant economics

3. **CPUC Standard Practice Manual**
   - URL: https://www.cpuc.ca.gov/industries-and-topics/electrical-energy/demand-side-management
   - Use: BCR methodology, avoided cost framework

4. **E3 Avoided Cost Calculator 2024**
   - URL: https://www.cpuc.ca.gov/industries-and-topics/electrical-energy/demand-side-management/acc
   - Use: T&D deferral values, GHG costs

5. **LBNL Interruption Cost Estimate (ICE) Calculator**
   - URL: https://icecalculator.com/
   - Use: Resilience value, VOLL estimates

### Regulatory References

6. **CPUC Resource Adequacy Program**
   - Decision D.24-06-050
   - Use: RA capacity values, ELCC methodology

7. **FERC Order 841**
   - Electric Storage Participation in Markets
   - Use: Market participation rules

8. **IRS Guidance on ITC for Energy Storage**
   - Notice 2023-XX (IRA Implementation)
   - Use: Tax credit eligibility, adder requirements

### Academic and Industry References

9. **Brealey, Myers, Allen. Principles of Corporate Finance (13th ed.)**
   - McGraw-Hill, 2020
   - Use: NPV, IRR, WACC methodology

10. **NREL Storage Futures Study**
    - Technical Report NREL/TP-6A20-77449
    - Use: Grid integration value, deployment projections

11. **EPRI Grid Energy Storage**
    - Technical Report 3002020045
    - Use: Distribution services value

12. **BloombergNEF Battery Price Survey 2025**
    - Use: Technology cost trends, learning curves

---

## Appendix A: Glossary

| Term | Definition |
|------|------------|
| **BCR** | Benefit-Cost Ratio |
| **BESS** | Battery Energy Storage System |
| **CapEx** | Capital Expenditure |
| **ELCC** | Effective Load Carrying Capability |
| **FOM** | Fixed Operations & Maintenance |
| **ITC** | Investment Tax Credit |
| **LCOS** | Levelized Cost of Storage |
| **LFP** | Lithium Iron Phosphate (battery chemistry) |
| **NMC** | Nickel Manganese Cobalt (battery chemistry) |
| **NPV** | Net Present Value |
| **O&M** | Operations & Maintenance |
| **RA** | Resource Adequacy |
| **RTE** | Round-Trip Efficiency |
| **T&D** | Transmission & Distribution |
| **VOLL** | Value of Lost Load |
| **VOM** | Variable Operations & Maintenance |
| **WACC** | Weighted Average Cost of Capital |

---

## Appendix B: Sample Project Analysis

### Project: 100 MW / 400 MWh LFP BESS

**Assumptions** (NREL ATB 2024 - Moderate):

| Parameter | Value |
|-----------|-------|
| Capacity | 100 MW / 400 MWh |
| Duration | 4 hours |
| CapEx | $160/kWh ($64M battery) |
| Infrastructure | $12.5M |
| ITC | 30% ($19.2M credit) |
| Net Capital Cost | $57.3M |
| Discount Rate | 7% |
| Analysis Period | 20 years |

**Annual Benefits** (Year 1):

| Benefit | Value | Category |
|---------|-------|----------|
| Resource Adequacy | $15.0M | Common |
| Energy Arbitrage | $4.0M | Common |
| Ancillary Services | $1.5M | Common |
| T&D Deferral | $2.5M | Common |
| Resilience | $5.0M | Common |
| Renewable Integration | $2.5M | BESS-specific |
| GHG Value | $1.5M | BESS-specific |
| Voltage Support | $0.8M | Common |
| **Total Year 1** | **$32.8M** | |

**Results**:

| Metric | Value | Assessment |
|--------|-------|------------|
| BCR | 2.1 | Strong - Approve |
| NPV | $89M | Substantial value creation |
| IRR | 18% | Well above 7% WACC |
| Payback | 4.2 years | Excellent |
| LCOS | $142/MWh | Competitive |

---

*Document prepared for internal use. Verify all assumptions with current market data before use in regulatory filings or investment decisions.*
