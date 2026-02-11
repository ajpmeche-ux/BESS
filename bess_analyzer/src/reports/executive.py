"""Executive summary PDF report generation using ReportLab.

Generates a comprehensive 8-10 page professional PDF containing:
- Executive summary with key findings
- Project overview and configuration
- Detailed financial metrics with interpretations
- Complete cost analysis with breakdowns
- Benefit analysis with individual streams
- Cash flow projections and charts
- Sensitivity analysis results
- Full methodology documentation
- Complete assumption library with citations
"""

import tempfile
from datetime import datetime
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (
    Image,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
    KeepTogether,
)

from src.models.project import FinancialResults, Project
from src.reports.charts import create_benefit_pie_chart, create_cashflow_chart
from src.utils.formatters import format_currency, format_percent, format_years


def _get_recommendation(bcr: float) -> tuple:
    """Return recommendation text and color based on BCR threshold."""
    if bcr >= 1.5:
        return "APPROVE - Strong economic case", colors.green, (
            "The project demonstrates robust economics with benefits significantly exceeding costs. "
            "The BCR of {:.2f} exceeds the 1.5 threshold typically required for regulatory approval. "
            "The project is expected to create substantial value for ratepayers."
        )
    elif bcr >= 1.0:
        return "FURTHER STUDY - Marginal economics", colors.orange, (
            "The project shows marginal economics with benefits slightly exceeding costs. "
            "While the BCR exceeds 1.0, additional analysis may be warranted to confirm assumptions. "
            "Consider sensitivity analysis on key inputs before proceeding."
        )
    else:
        return "REJECT - Costs exceed benefits", colors.red, (
            "The project does not meet economic viability thresholds. "
            "Costs exceed benefits based on current assumptions. "
            "Consider alternative configurations or await more favorable market conditions."
        )


def _create_styles():
    """Create custom paragraph styles for the report."""
    styles = getSampleStyleSheet()

    custom_styles = {
        'title': ParagraphStyle(
            "CustomTitle", parent=styles["Title"],
            fontSize=24, spaceAfter=6, textColor=colors.HexColor("#1565c0")
        ),
        'subtitle': ParagraphStyle(
            "Subtitle", parent=styles["Heading2"],
            fontSize=14, spaceAfter=12, textColor=colors.HexColor("#424242")
        ),
        'heading1': ParagraphStyle(
            "CustomHeading1", parent=styles["Heading1"],
            fontSize=16, spaceAfter=10, spaceBefore=16,
            textColor=colors.HexColor("#1565c0"), borderPadding=4
        ),
        'heading2': ParagraphStyle(
            "CustomHeading2", parent=styles["Heading2"],
            fontSize=13, spaceAfter=8, spaceBefore=12,
            textColor=colors.HexColor("#1976d2")
        ),
        'heading3': ParagraphStyle(
            "CustomHeading3", parent=styles["Heading3"],
            fontSize=11, spaceAfter=6, spaceBefore=8,
            textColor=colors.HexColor("#424242")
        ),
        'body': ParagraphStyle(
            "CustomBody", parent=styles["Normal"],
            fontSize=10, spaceAfter=8, leading=14
        ),
        'body_indent': ParagraphStyle(
            "BodyIndent", parent=styles["Normal"],
            fontSize=10, spaceAfter=6, leading=14, leftIndent=20
        ),
        'small': ParagraphStyle(
            "Small", parent=styles["Normal"],
            fontSize=8, textColor=colors.grey, spaceAfter=4
        ),
        'citation': ParagraphStyle(
            "Citation", parent=styles["Normal"],
            fontSize=8, textColor=colors.HexColor("#616161"),
            leftIndent=20, spaceAfter=3
        ),
        'formula': ParagraphStyle(
            "Formula", parent=styles["Normal"],
            fontSize=10, fontName="Courier",
            leftIndent=30, spaceAfter=6, textColor=colors.HexColor("#1565c0")
        ),
        'note': ParagraphStyle(
            "Note", parent=styles["Normal"],
            fontSize=9, textColor=colors.HexColor("#616161"),
            leftIndent=10, rightIndent=10, spaceAfter=8,
            backColor=colors.HexColor("#f5f5f5"), borderPadding=8
        ),
        'table_note': ParagraphStyle(
            "TableNote", parent=styles["Normal"],
            fontSize=8, textColor=colors.grey, spaceAfter=12
        ),
    }

    return custom_styles


def _create_table_style(header_color="#1565c0"):
    """Create consistent table styling."""
    return TableStyle([
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(header_color)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#e0e0e0")),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#fafafa")]),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ])


def generate_executive_summary(
    project: Project, results: FinancialResults, output_path: str
) -> None:
    """Generate a comprehensive executive summary PDF report.

    Creates an 8-10 page PDF with:
    - Page 1: Executive Summary with key findings
    - Page 2: Project Configuration & Technology
    - Page 3: Cost Analysis (CapEx, O&M, Infrastructure)
    - Page 4: Benefit Analysis (all streams detailed)
    - Page 5: Financial Metrics & Interpretation
    - Page 6: Cash Flow Analysis with charts
    - Page 7: Sensitivity Analysis
    - Page 8-9: Methodology & Formulas
    - Page 10: Complete Assumptions & Citations

    Args:
        project: Project with all inputs populated.
        results: Calculated FinancialResults.
        output_path: File path for the output PDF.
    """
    doc = SimpleDocTemplate(
        output_path,
        pagesize=letter,
        topMargin=0.6 * inch,
        bottomMargin=0.6 * inch,
        leftMargin=0.7 * inch,
        rightMargin=0.7 * inch,
    )

    styles = _create_styles()
    elements = []

    basics = project.basics
    costs = project.costs
    tech = project.technology
    financing = project.financing

    # =========================================================================
    # PAGE 1: EXECUTIVE SUMMARY
    # =========================================================================
    elements.append(Paragraph("Battery Energy Storage System", styles['title']))
    elements.append(Paragraph("Economic Analysis Report", styles['subtitle']))
    elements.append(Spacer(1, 6))

    # Report metadata
    meta_data = [
        ["Report Date", datetime.now().strftime("%B %d, %Y")],
        ["Project", basics.name or "Unnamed Project"],
        ["Assumption Library", project.assumption_library or "Custom"],
    ]
    meta_table = Table(meta_data, colWidths=[1.5 * inch, 5.5 * inch])
    meta_table.setStyle(TableStyle([
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("TEXTCOLOR", (0, 0), (0, -1), colors.grey),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    elements.append(meta_table)
    elements.append(Spacer(1, 16))

    # Key Finding Box
    rec_text, rec_color, rec_detail = _get_recommendation(results.bcr)

    finding_data = [[Paragraph(f"<b>RECOMMENDATION: {rec_text}</b>", styles['body'])]]
    finding_table = Table(finding_data, colWidths=[7 * inch])
    finding_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#e3f2fd")),
        ("BOX", (0, 0), (-1, -1), 2, rec_color),
        ("TOPPADDING", (0, 0), (-1, -1), 12),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 12),
        ("LEFTPADDING", (0, 0), (-1, -1), 12),
    ]))
    elements.append(finding_table)
    elements.append(Spacer(1, 8))
    elements.append(Paragraph(rec_detail.format(results.bcr), styles['body']))
    elements.append(Spacer(1, 16))

    # Key Metrics Summary
    elements.append(Paragraph("Key Financial Metrics", styles['heading2']))

    capex_total = results.annual_costs[0] if results.annual_costs else 0
    metrics_data = [
        ["Metric", "Value", "Interpretation"],
        ["Benefit-Cost Ratio (BCR)", f"{results.bcr:.2f}",
         "Pass (≥1.0)" if results.bcr >= 1.0 else "Fail (<1.0)"],
        ["Net Present Value (NPV)", format_currency(results.npv, decimals=1),
         "Value created" if results.npv > 0 else "Value destroyed"],
        ["Internal Rate of Return (IRR)",
         format_percent(results.irr) if results.irr else "N/A",
         f"{'Exceeds' if results.irr and results.irr > basics.discount_rate else 'Below'} hurdle rate"
         if results.irr else "Cannot calculate"],
        ["Simple Payback", format_years(results.payback_years),
         "Within analysis period" if results.payback_years and results.payback_years <= basics.analysis_period_years else "Beyond horizon"],
        ["LCOS ($/MWh)", f"${results.lcos_per_mwh:,.1f}",
         "Competitive" if results.lcos_per_mwh < 200 else "Above market"],
        ["Breakeven CapEx", f"${results.breakeven_capex_per_kwh:,.0f}/kWh",
         f"{'Above' if results.breakeven_capex_per_kwh > costs.capex_per_kwh else 'Below'} current cost"],
    ]

    metrics_table = Table(metrics_data, colWidths=[2.2 * inch, 1.8 * inch, 3 * inch])
    metrics_table.setStyle(_create_table_style())
    elements.append(metrics_table)
    elements.append(Spacer(1, 12))

    # Project Summary
    elements.append(Paragraph("Project Summary", styles['heading2']))
    summary_data = [
        ["Parameter", "Value"],
        ["System Capacity", f"{basics.capacity_mw:,.1f} MW / {basics.capacity_mwh:,.1f} MWh"],
        ["Duration", f"{basics.duration_hours:.1f} hours"],
        ["Total Capital Investment", format_currency(capex_total, decimals=1)],
        ["Analysis Period", f"{basics.analysis_period_years} years"],
        ["Discount Rate (WACC)", format_percent(basics.discount_rate)],
        ["PV of Total Costs", format_currency(results.pv_costs, decimals=1)],
        ["PV of Total Benefits", format_currency(results.pv_benefits, decimals=1)],
    ]
    summary_table = Table(summary_data, colWidths=[3.5 * inch, 3.5 * inch])
    summary_table.setStyle(_create_table_style("#424242"))
    elements.append(summary_table)

    # =========================================================================
    # PAGE 2: PROJECT CONFIGURATION & TECHNOLOGY
    # =========================================================================
    elements.append(PageBreak())
    elements.append(Paragraph("Project Configuration", styles['heading1']))

    # Technology Specifications
    elements.append(Paragraph("Technology Specifications", styles['heading2']))
    tech_text = (
        f"The proposed system utilizes <b>{tech.chemistry}</b> battery technology, which offers "
        f"favorable cycle life characteristics for grid-scale applications. The system is designed "
        f"for {basics.duration_hours}-hour duration with a nameplate capacity of {basics.capacity_mw:,.0f} MW, "
        f"providing {basics.capacity_mwh:,.0f} MWh of energy storage capacity."
    )
    elements.append(Paragraph(tech_text, styles['body']))

    tech_data = [
        ["Parameter", "Value", "Notes"],
        ["Battery Chemistry", tech.chemistry, "LFP preferred for long-duration, high-cycle applications"],
        ["Round-Trip Efficiency", format_percent(tech.round_trip_efficiency), "AC-to-AC efficiency including inverter losses"],
        ["Annual Degradation", format_percent(tech.degradation_rate_annual), "Capacity fade per year of operation"],
        ["Cycle Life", f"{tech.cycle_life:,}", "Full depth-of-discharge cycles to 80% capacity"],
        ["Augmentation Year", f"Year {tech.augmentation_year}", "Planned battery module replacement"],
        ["Cycles per Day", f"{tech.cycles_per_day:.1f}", "Average daily discharge cycles assumed"],
    ]
    tech_table = Table(tech_data, colWidths=[2 * inch, 1.5 * inch, 3.5 * inch])
    tech_table.setStyle(_create_table_style("#1976d2"))
    elements.append(tech_table)
    elements.append(Spacer(1, 12))

    # Financing Structure
    elements.append(Paragraph("Financing Structure", styles['heading2']))
    if financing:
        wacc = (1 - financing.debt_percent) * financing.cost_of_equity + \
               financing.debt_percent * financing.interest_rate * (1 - financing.tax_rate)

        fin_text = (
            f"The project assumes a capital structure of {format_percent(financing.debt_percent)} debt "
            f"and {format_percent(1 - financing.debt_percent)} equity. Based on the financing assumptions, "
            f"the weighted average cost of capital (WACC) is calculated at {format_percent(wacc)}."
        )
        elements.append(Paragraph(fin_text, styles['body']))

        fin_data = [
            ["Parameter", "Value", "Source/Basis"],
            ["Debt Percentage", format_percent(financing.debt_percent), "Typical utility project financing"],
            ["Interest Rate", format_percent(financing.interest_rate), "Current market rates"],
            ["Loan Term", f"{financing.loan_term_years} years", "Standard project finance term"],
            ["Cost of Equity", format_percent(financing.cost_of_equity), "Required return on equity"],
            ["Tax Rate", format_percent(financing.tax_rate), "Federal corporate tax rate"],
            ["Calculated WACC", format_percent(wacc), "Used as discount rate"],
        ]
        fin_table = Table(fin_data, colWidths=[2 * inch, 1.5 * inch, 3.5 * inch])
        fin_table.setStyle(_create_table_style("#1976d2"))
        elements.append(fin_table)
    elements.append(Spacer(1, 12))

    # =========================================================================
    # PAGE 3: COST ANALYSIS
    # =========================================================================
    elements.append(PageBreak())
    elements.append(Paragraph("Cost Analysis", styles['heading1']))

    elements.append(Paragraph("Capital Expenditures", styles['heading2']))

    # Calculate cost components
    battery_capex = costs.capex_per_kwh * basics.capacity_mwh * 1000
    infra_cost = (costs.interconnection_per_kw + costs.land_per_kw + costs.permitting_per_kw) * basics.capacity_mw * 1000
    total_capex = battery_capex + infra_cost
    itc_credit = battery_capex * (costs.itc_percent + costs.itc_adders)
    net_capex = total_capex - itc_credit

    capex_text = (
        f"Total capital expenditure is estimated at {format_currency(total_capex, decimals=1)}, "
        f"consisting of battery system costs ({format_currency(battery_capex, decimals=1)}) "
        f"and infrastructure costs ({format_currency(infra_cost, decimals=1)}). "
        f"After applying the Investment Tax Credit of {format_currency(itc_credit, decimals=1)}, "
        f"the net Year 0 investment is {format_currency(net_capex, decimals=1)}."
    )
    elements.append(Paragraph(capex_text, styles['body']))

    capex_data = [
        ["Cost Component", "$/Unit", "Quantity", "Total"],
        ["Battery System ($/kWh)", f"${costs.capex_per_kwh:,.0f}", f"{basics.capacity_mwh * 1000:,.0f} kWh", format_currency(battery_capex, decimals=1)],
        ["Interconnection ($/kW)", f"${costs.interconnection_per_kw:,.0f}", f"{basics.capacity_mw * 1000:,.0f} kW", format_currency(costs.interconnection_per_kw * basics.capacity_mw * 1000, decimals=1)],
        ["Land Acquisition ($/kW)", f"${costs.land_per_kw:,.0f}", f"{basics.capacity_mw * 1000:,.0f} kW", format_currency(costs.land_per_kw * basics.capacity_mw * 1000, decimals=1)],
        ["Permitting ($/kW)", f"${costs.permitting_per_kw:,.0f}", f"{basics.capacity_mw * 1000:,.0f} kW", format_currency(costs.permitting_per_kw * basics.capacity_mw * 1000, decimals=1)],
        ["Gross Capital Cost", "", "", format_currency(total_capex, decimals=1)],
        [f"Less: ITC ({format_percent(costs.itc_percent + costs.itc_adders)})", "", "", f"({format_currency(itc_credit, decimals=1)})"],
        ["Net Capital Investment", "", "", format_currency(net_capex, decimals=1)],
    ]
    capex_table = Table(capex_data, colWidths=[2.5 * inch, 1.3 * inch, 1.5 * inch, 1.7 * inch])
    capex_table.setStyle(_create_table_style())
    # Bold the totals
    capex_table.setStyle(TableStyle([
        ("FONTNAME", (0, -3), (-1, -1), "Helvetica-Bold"),
        ("LINEABOVE", (0, -3), (-1, -3), 1, colors.black),
    ]))
    elements.append(capex_table)
    elements.append(Spacer(1, 8))

    # Add bulk discount note if applicable
    if (costs.bulk_discount_rate > 0 and
        costs.bulk_discount_threshold_mwh > 0 and
        basics.capacity_mwh >= costs.bulk_discount_threshold_mwh):
        discount_note = (
            f"<b>Bulk Discount Applied:</b> A {format_percent(costs.bulk_discount_rate)} fleet purchase discount "
            f"has been applied to all costs (threshold: {costs.bulk_discount_threshold_mwh:,.0f} MWh). "
            f"This discount reflects economies of scale when purchasing multiple units."
        )
        elements.append(Paragraph(discount_note, styles['note']))
    elements.append(Spacer(1, 12))

    # Operating Costs
    elements.append(Paragraph("Annual Operating Costs", styles['heading2']))

    annual_fom = costs.fom_per_kw_year * basics.capacity_mw * 1000
    annual_charging = costs.charging_cost_per_mwh * basics.capacity_mwh * tech.cycles_per_day * 365 * tech.round_trip_efficiency
    annual_insurance = battery_capex * costs.insurance_pct_of_capex
    annual_property_tax = total_capex * costs.property_tax_pct

    opex_data = [
        ["Operating Cost", "Annual Amount", "Basis"],
        ["Fixed O&M", format_currency(annual_fom, decimals=1), f"${costs.fom_per_kw_year:,.0f}/kW-year"],
        ["Charging Costs", format_currency(annual_charging, decimals=1), f"${costs.charging_cost_per_mwh:,.0f}/MWh at {tech.cycles_per_day} cycles/day"],
        ["Insurance", format_currency(annual_insurance, decimals=1), f"{format_percent(costs.insurance_pct_of_capex)} of battery CapEx"],
        ["Property Tax", format_currency(annual_property_tax, decimals=1), f"{format_percent(costs.property_tax_pct)} of total CapEx (declining)"],
        ["Total Annual OpEx", format_currency(annual_fom + annual_charging + annual_insurance + annual_property_tax, decimals=1), ""],
    ]
    opex_table = Table(opex_data, colWidths=[2.5 * inch, 2 * inch, 2.5 * inch])
    opex_table.setStyle(_create_table_style("#1976d2"))
    elements.append(opex_table)
    elements.append(Spacer(1, 8))

    # Special costs
    elements.append(Paragraph("Non-Recurring Costs", styles['heading3']))
    special_data = [
        ["Cost Item", "Year", "Amount", "Notes"],
        ["Battery Augmentation", f"Year {tech.augmentation_year}",
         format_currency(costs.augmentation_per_kwh * basics.capacity_mwh * 1000, decimals=1),
         "Battery module replacement"],
        ["Decommissioning", f"Year {basics.analysis_period_years}",
         format_currency(costs.decommissioning_per_kw * basics.capacity_mw * 1000, decimals=1),
         "End-of-life removal (net of residual value)"],
    ]
    special_table = Table(special_data, colWidths=[2 * inch, 1 * inch, 2 * inch, 2 * inch])
    special_table.setStyle(_create_table_style("#757575"))
    elements.append(special_table)

    # =========================================================================
    # PAGE 4: BENEFIT ANALYSIS
    # =========================================================================
    elements.append(PageBreak())
    elements.append(Paragraph("Benefit Analysis", styles['heading1']))

    elements.append(Paragraph(
        "The following benefit streams have been identified and quantified for this project. "
        "All values are shown as Year 1 amounts and escalated annually as indicated.",
        styles['body']
    ))
    elements.append(Spacer(1, 8))

    # Benefit stream details
    ben_data = [["Benefit Category", "Year 1 Value", "Escalation", "PV Total", "% of Total"]]
    for benefit in project.benefits:
        year1_val = benefit.annual_values[0] if benefit.annual_values else 0
        # Calculate implied escalation rate from annual values
        if len(benefit.annual_values) >= 2 and benefit.annual_values[0] > 0:
            escalation = (benefit.annual_values[1] / benefit.annual_values[0]) - 1
        else:
            escalation = 0.0
        pv_pct = results.benefit_breakdown.get(benefit.name, 0)
        pv_val = results.pv_benefits * pv_pct / 100
        ben_data.append([
            benefit.name,
            format_currency(year1_val, decimals=1),
            format_percent(escalation),
            format_currency(pv_val, decimals=1),
            f"{pv_pct:.1f}%"
        ])

    # Total row
    total_year1 = sum(b.annual_values[0] if b.annual_values else 0 for b in project.benefits)
    ben_data.append(["Total Benefits", format_currency(total_year1, decimals=1), "", format_currency(results.pv_benefits, decimals=1), "100.0%"])

    ben_table = Table(ben_data, colWidths=[2.2 * inch, 1.4 * inch, 1 * inch, 1.4 * inch, 1 * inch])
    ben_table.setStyle(_create_table_style("#2e7d32"))
    ben_table.setStyle(TableStyle([
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("LINEABOVE", (0, -1), (-1, -1), 1, colors.black),
    ]))
    elements.append(ben_table)
    elements.append(Spacer(1, 12))

    # Benefit descriptions
    elements.append(Paragraph("Benefit Stream Descriptions", styles['heading2']))

    benefit_descriptions = {
        "Resource Adequacy": "Capacity payments for providing reliable power during peak demand periods. Based on CPUC Resource Adequacy program requirements.",
        "Energy Arbitrage": "Revenue from buying electricity during low-price periods and selling during high-price periods. Based on historical CAISO price spreads.",
        "Ancillary Services": "Payments for providing frequency regulation, spinning reserves, and other grid services. Based on CAISO AS market prices.",
        "T&D Deferral": "Avoided costs from deferring transmission and distribution upgrades. Based on E3 Avoided Cost Calculator.",
        "Resilience Value": "Value of backup power during outages. Based on LBNL ICE Calculator customer interruption costs.",
        "Renewable Integration": "Value of enabling increased renewable penetration by providing firming and shifting services.",
        "GHG Emissions Value": "Monetized value of avoided greenhouse gas emissions. Based on EPA Social Cost of Carbon.",
        "Voltage Support": "Value of providing reactive power and voltage regulation services to the distribution system.",
        # Special benefits (formula-based)
        "Reliability (Avoided Outage)": "Avoided customer outage costs based on expected outage hours and value of lost load. Calculated as: outage_hours × customer_cost_per_kWh × capacity × backup_percentage.",
        "Safety (Avoided Incident)": "Avoided safety incident costs from improved grid stability. Calculated as: incident_probability × incident_cost × risk_reduction_factor.",
        "Speed-to-Serve (One-time)": "One-time value of faster deployment compared to alternative generation sources. BESS can be deployed in 12-18 months vs. 36+ months for gas peakers.",
    }

    for benefit in project.benefits:
        desc = benefit_descriptions.get(benefit.name, benefit.description or "No description available.")
        elements.append(Paragraph(f"<b>{benefit.name}:</b> {desc}", styles['body_indent']))

    # Add special benefits descriptions if enabled
    special = project.special_benefits
    if special:
        if special.reliability_enabled:
            desc = benefit_descriptions.get("Reliability (Avoided Outage)", "")
            elements.append(Paragraph(f"<b>Reliability (Avoided Outage):</b> {desc}", styles['body_indent']))
        if special.safety_enabled:
            desc = benefit_descriptions.get("Safety (Avoided Incident)", "")
            elements.append(Paragraph(f"<b>Safety (Avoided Incident):</b> {desc}", styles['body_indent']))
        if special.speed_enabled:
            desc = benefit_descriptions.get("Speed-to-Serve (One-time)", "")
            elements.append(Paragraph(f"<b>Speed-to-Serve (One-time):</b> {desc}", styles['body_indent']))

    # =========================================================================
    # PAGE 5: FINANCIAL METRICS DETAIL
    # =========================================================================
    elements.append(PageBreak())
    elements.append(Paragraph("Financial Metrics Analysis", styles['heading1']))

    # BCR Analysis
    elements.append(Paragraph("Benefit-Cost Ratio (BCR)", styles['heading2']))
    bcr_text = (
        f"The Benefit-Cost Ratio is calculated as the ratio of present value of benefits "
        f"to present value of costs. A BCR greater than 1.0 indicates that benefits exceed costs.<br/><br/>"
        f"<b>Calculation:</b><br/>"
        f"BCR = PV(Benefits) / PV(Costs)<br/>"
        f"BCR = {format_currency(results.pv_benefits, decimals=1)} / {format_currency(results.pv_costs, decimals=1)}<br/>"
        f"<b>BCR = {results.bcr:.2f}</b>"
    )
    elements.append(Paragraph(bcr_text, styles['body']))
    elements.append(Spacer(1, 8))

    # NPV Analysis
    elements.append(Paragraph("Net Present Value (NPV)", styles['heading2']))
    npv_text = (
        f"Net Present Value represents the difference between the present value of benefits "
        f"and present value of costs, discounted at the project's cost of capital ({format_percent(basics.discount_rate)}).<br/><br/>"
        f"<b>Calculation:</b><br/>"
        f"NPV = PV(Benefits) - PV(Costs)<br/>"
        f"NPV = {format_currency(results.pv_benefits, decimals=1)} - {format_currency(results.pv_costs, decimals=1)}<br/>"
        f"<b>NPV = {format_currency(results.npv, decimals=1)}</b><br/><br/>"
        f"A positive NPV indicates the project creates value above the required rate of return."
    )
    elements.append(Paragraph(npv_text, styles['body']))
    elements.append(Spacer(1, 8))

    # IRR Analysis
    elements.append(Paragraph("Internal Rate of Return (IRR)", styles['heading2']))
    if results.irr is not None:
        irr_text = (
            f"The Internal Rate of Return is the discount rate at which NPV equals zero. "
            f"It represents the project's expected return on investment.<br/><br/>"
            f"<b>Result: IRR = {format_percent(results.irr)}</b><br/><br/>"
            f"The IRR of {format_percent(results.irr)} {'exceeds' if results.irr > basics.discount_rate else 'is below'} "
            f"the required return of {format_percent(basics.discount_rate)}, indicating the project "
            f"{'meets' if results.irr > basics.discount_rate else 'does not meet'} the minimum return threshold."
        )
    else:
        irr_text = "IRR could not be calculated. This typically occurs when cash flows do not change sign or when the project does not achieve positive cumulative returns."
    elements.append(Paragraph(irr_text, styles['body']))
    elements.append(Spacer(1, 8))

    # LCOS Analysis
    elements.append(Paragraph("Levelized Cost of Storage (LCOS)", styles['heading2']))
    lcos_text = (
        f"LCOS represents the all-in cost per MWh of energy discharged over the project lifetime, "
        f"following the Lazard methodology.<br/><br/>"
        f"<b>Result: LCOS = ${results.lcos_per_mwh:,.1f}/MWh</b><br/><br/>"
        f"This metric enables comparison across different storage technologies and project configurations. "
        f"For context, current utility-scale battery LCOS typically ranges from $120-180/MWh."
    )
    elements.append(Paragraph(lcos_text, styles['body']))

    # =========================================================================
    # PAGE 6: CASH FLOW ANALYSIS
    # =========================================================================
    elements.append(PageBreak())
    elements.append(Paragraph("Cash Flow Analysis", styles['heading1']))

    with tempfile.TemporaryDirectory() as tmpdir:
        # Cash Flow Table (selected years)
        elements.append(Paragraph("Annual Cash Flows", styles['heading2']))

        n_years = len(results.annual_costs)
        years_to_show = [0, 1, 5, 10, 15, min(20, n_years - 1)]
        years_to_show = sorted(set(y for y in years_to_show if y < n_years))

        cf_data = [["Year"] + [f"Year {y}" for y in years_to_show]]
        cf_data.append(["Costs"] + [format_currency(results.annual_costs[y], decimals=0) for y in years_to_show])
        cf_data.append(["Benefits"] + [format_currency(results.annual_benefits[y], decimals=0) for y in years_to_show])
        cf_data.append(["Net CF"] + [format_currency(results.annual_net[y], decimals=0) for y in years_to_show])

        cf_table = Table(cf_data, colWidths=[1.2 * inch] + [1 * inch] * len(years_to_show))
        cf_table.setStyle(_create_table_style())
        elements.append(cf_table)
        elements.append(Spacer(1, 12))

        # Cash Flow Chart
        cf_path = str(Path(tmpdir) / "cashflow.png")
        if results.annual_costs and results.annual_benefits:
            create_cashflow_chart(results.annual_costs, results.annual_benefits, cf_path)
            elements.append(Image(cf_path, width=6.5 * inch, height=3.5 * inch))
        elements.append(Spacer(1, 12))

        # Benefit Pie Chart
        elements.append(Paragraph("Benefit Distribution", styles['heading2']))
        pie_path = str(Path(tmpdir) / "benefit_pie.png")
        if results.benefit_breakdown:
            create_benefit_pie_chart(results.benefit_breakdown, pie_path)
            elements.append(Image(pie_path, width=5 * inch, height=4 * inch))

        # =========================================================================
        # PAGE 7: SENSITIVITY ANALYSIS
        # =========================================================================
        elements.append(PageBreak())
        elements.append(Paragraph("Sensitivity Analysis", styles['heading1']))

        elements.append(Paragraph(
            "The following tables show how key metrics change under different assumptions. "
            "This analysis helps identify the key value drivers and risk factors for the project.",
            styles['body']
        ))
        elements.append(Spacer(1, 12))

        # BCR Sensitivity Table
        elements.append(Paragraph("BCR Sensitivity: CapEx vs. Benefit Levels", styles['heading2']))

        capex_levels = [100, 120, 140, 160, 180, 200, 220]
        benefit_mults = [0.8, 0.9, 1.0, 1.1, 1.2]

        sens_header = ["CapEx $/kWh"] + [f"{int(m*100)}% Benefits" for m in benefit_mults]
        sens_data = [sens_header]

        for capex in capex_levels:
            row = [f"${capex}"]
            capex_ratio = capex / costs.capex_per_kwh if costs.capex_per_kwh > 0 else 1.0
            for ben_mult in benefit_mults:
                adj_pv_benefits = results.pv_benefits * ben_mult
                adj_pv_costs = results.pv_costs * (1 + (capex_ratio - 1) * 0.7)
                bcr = adj_pv_benefits / adj_pv_costs if adj_pv_costs > 0 else 0
                row.append(f"{bcr:.2f}")
            sens_data.append(row)

        sens_table = Table(sens_data, colWidths=[1.2 * inch] + [1.16 * inch] * len(benefit_mults))
        sens_table.setStyle(_create_table_style("#424242"))

        # Color code BCR cells
        for row_idx in range(1, len(sens_data)):
            for col_idx in range(1, len(sens_data[row_idx])):
                bcr_val = float(sens_data[row_idx][col_idx])
                if bcr_val >= 1.5:
                    color = colors.HexColor("#c8e6c9")
                elif bcr_val >= 1.0:
                    color = colors.HexColor("#fff9c4")
                else:
                    color = colors.HexColor("#ffcdd2")
                sens_table.setStyle(TableStyle([("BACKGROUND", (col_idx, row_idx), (col_idx, row_idx), color)]))

        elements.append(sens_table)
        elements.append(Paragraph("Green: BCR ≥ 1.5 (Strong) | Yellow: BCR 1.0-1.5 (Marginal) | Red: BCR < 1.0 (Fail)", styles['table_note']))

        # =========================================================================
        # PAGE 8: METHODOLOGY
        # =========================================================================
        elements.append(PageBreak())
        elements.append(Paragraph("Methodology & Formulas", styles['heading1']))

        elements.append(Paragraph("Discounted Cash Flow Analysis", styles['heading2']))
        elements.append(Paragraph(
            "This analysis employs standard discounted cash flow (DCF) methodology as described in "
            "Brealey, Myers & Allen (2020) and consistent with CPUC Standard Practice Manual guidelines. "
            "All future costs and benefits are discounted to present value using the project's weighted "
            "average cost of capital (WACC).",
            styles['body']
        ))
        elements.append(Spacer(1, 8))

        # Formulas
        elements.append(Paragraph("Key Formulas", styles['heading2']))

        formulas = [
            ("Present Value", "PV = FV / (1 + r)^t", "where r = discount rate, t = time period"),
            ("Net Present Value", "NPV = Σ(CFt / (1 + r)^t) for t = 0 to N", "Sum of discounted cash flows"),
            ("Benefit-Cost Ratio", "BCR = PV(Benefits) / PV(Costs)", "Ratio of present values"),
            ("Internal Rate of Return", "Solve for r where: Σ(CFt / (1 + r)^t) = 0", "Discount rate at NPV = 0"),
            ("LCOS", "LCOS = PV(Costs) / PV(Energy)", "Levelized cost per MWh discharged"),
            ("WACC", "WACC = E/(D+E) × Re + D/(D+E) × Rd × (1-Tc)", "Weighted average cost of capital"),
        ]

        for name, formula, desc in formulas:
            elements.append(Paragraph(f"<b>{name}:</b>", styles['body']))
            elements.append(Paragraph(formula, styles['formula']))
            elements.append(Paragraph(desc, styles['body_indent']))
            elements.append(Spacer(1, 4))

        # Degradation
        elements.append(Paragraph("Battery Degradation", styles['heading2']))
        elements.append(Paragraph(
            f"Battery capacity is assumed to degrade at {format_percent(tech.degradation_rate_annual)} per year. "
            f"This affects both the energy output available for arbitrage and ancillary services, and the capacity "
            f"available for resource adequacy. The analysis assumes battery augmentation (module replacement) "
            f"in Year {tech.augmentation_year} to restore capacity.",
            styles['body']
        ))

        # =========================================================================
        # PAGE 9: ASSUMPTIONS & CITATIONS
        # =========================================================================
        elements.append(PageBreak())
        elements.append(Paragraph("Complete Assumption Set", styles['heading1']))

        # All assumptions table
        assumptions_data = [
            ["Category", "Parameter", "Value", "Source"],
            ["Project", "Capacity", f"{basics.capacity_mw} MW / {basics.capacity_mwh} MWh", "Project specification"],
            ["Project", "Duration", f"{basics.duration_hours} hours", "Project specification"],
            ["Project", "Analysis Period", f"{basics.analysis_period_years} years", "Standard utility practice"],
            ["Technology", "Chemistry", tech.chemistry, project.assumption_library or "Industry standard"],
            ["Technology", "Round-Trip Efficiency", format_percent(tech.round_trip_efficiency), "Manufacturer specifications"],
            ["Technology", "Annual Degradation", format_percent(tech.degradation_rate_annual), "NREL ATB 2024"],
            ["Technology", "Cycle Life", f"{tech.cycle_life:,} cycles", "Manufacturer warranty"],
            ["Cost", "CapEx", f"${costs.capex_per_kwh}/kWh", project.assumption_library or "Market data"],
            ["Cost", "Fixed O&M", f"${costs.fom_per_kw_year}/kW-year", project.assumption_library or "NREL ATB"],
            ["Cost", "Interconnection", f"${costs.interconnection_per_kw}/kW", "Utility estimate"],
            ["Tax Credit", "ITC Rate", format_percent(costs.itc_percent), "IRA 2022 (26 USC §48)"],
            ["Tax Credit", "ITC Adders", format_percent(costs.itc_adders), "Energy Community/Domestic Content"],
        ]

        if financing:
            assumptions_data.extend([
                ["Finance", "Debt %", format_percent(financing.debt_percent), "Project financing structure"],
                ["Finance", "Interest Rate", format_percent(financing.interest_rate), "Current market rates"],
                ["Finance", "Cost of Equity", format_percent(financing.cost_of_equity), "CAPM estimate"],
            ])

        assumptions_table = Table(assumptions_data, colWidths=[1.2 * inch, 1.8 * inch, 1.5 * inch, 2.5 * inch])
        assumptions_table.setStyle(_create_table_style("#616161"))
        elements.append(assumptions_table)
        elements.append(Spacer(1, 16))

        # Citations
        elements.append(Paragraph("References & Citations", styles['heading2']))

        citations = [
            "NREL. Annual Technology Baseline 2024. National Renewable Energy Laboratory. https://atb.nrel.gov/",
            "Lazard. Lazard's Levelized Cost of Storage Analysis, Version 10.0. March 2025.",
            "California Public Utilities Commission. Standard Practice Manual: Economic Analysis of Demand-Side Programs and Projects. October 2001.",
            "E3 (Energy + Environmental Economics). CPUC Avoided Cost Calculator 2024.",
            "Lawrence Berkeley National Laboratory. Interruption Cost Estimate (ICE) Calculator. https://icecalculator.com/",
            "U.S. Environmental Protection Agency. Social Cost of Greenhouse Gases. 2024.",
            "Brealey, R., Myers, S., & Allen, F. Principles of Corporate Finance (13th ed.). McGraw-Hill, 2020.",
            "NREL. Storage Futures Study. NREL/TP-6A20-77449. 2021.",
            "Internal Revenue Code, 26 USC §48 (Investment Tax Credit as amended by IRA 2022).",
        ]

        # Add benefit-specific citations
        for benefit in project.benefits:
            if benefit.citation and benefit.citation not in citations:
                citations.append(benefit.citation)

        for i, cite in enumerate(sorted(set(citations)), 1):
            elements.append(Paragraph(f"[{i}] {cite}", styles['citation']))

        elements.append(Spacer(1, 20))

        # Footer
        elements.append(Paragraph(
            f"<i>Report generated by BESS Analyzer v1.0 on {datetime.now().strftime('%Y-%m-%d %H:%M')}. "
            "All calculations are reproducible from the documented inputs and methodology above. "
            "This analysis is provided for informational purposes and should be validated before use in investment decisions.</i>",
            styles['small']
        ))

        # Build PDF (must happen while tmpdir exists for chart images)
        doc.build(elements)
