"""Executive summary PDF report generation using ReportLab.

Generates a 3-4 page professional PDF containing project overview,
key financial metrics, benefit breakdown, cost analysis, and
methodology documentation with full citations.
"""

import tempfile
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
)

from src.models.project import FinancialResults, Project
from src.reports.charts import create_benefit_pie_chart, create_cashflow_chart
from src.utils.formatters import format_currency, format_percent, format_years


def _get_recommendation(bcr: float) -> tuple:
    """Return recommendation text and color based on BCR threshold."""
    if bcr >= 1.5:
        return "APPROVE - Strong economic case", colors.green
    elif bcr >= 1.0:
        return "FURTHER STUDY - Marginal economics", colors.orange
    else:
        return "REJECT - Costs exceed benefits", colors.red


def generate_executive_summary(
    project: Project, results: FinancialResults, output_path: str
) -> None:
    """Generate an executive summary PDF report.

    Creates a 3-4 page PDF with:
    - Page 1: Project overview and key metrics
    - Page 2: Financial analysis with cost/benefit breakdown
    - Page 3: Methodology, assumptions, and citations

    Args:
        project: Project with all inputs populated.
        results: Calculated FinancialResults.
        output_path: File path for the output PDF.
    """
    doc = SimpleDocTemplate(
        output_path,
        pagesize=letter,
        topMargin=0.75 * inch,
        bottomMargin=0.75 * inch,
        leftMargin=0.75 * inch,
        rightMargin=0.75 * inch,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "CustomTitle", parent=styles["Title"], fontSize=18, spaceAfter=6
    )
    heading_style = ParagraphStyle(
        "CustomHeading", parent=styles["Heading2"], fontSize=14,
        spaceAfter=8, spaceBefore=12, textColor=colors.HexColor("#1565c0"),
    )
    body_style = styles["Normal"]
    small_style = ParagraphStyle(
        "Small", parent=body_style, fontSize=8, textColor=colors.grey,
    )

    elements = []

    # --- PAGE 1: Overview & Key Findings ---
    elements.append(Paragraph("BESS Economic Analysis", title_style))
    elements.append(Paragraph("Executive Summary", styles["Heading3"]))
    elements.append(Spacer(1, 12))

    # Project info
    basics = project.basics
    capex_total = results.annual_costs[0] if results.annual_costs else 0
    info_data = [
        ["Project Name", basics.name or "Unnamed"],
        ["Location", basics.location or "Not specified"],
        ["Capacity", f"{basics.capacity_mw:,.1f} MW / {basics.capacity_mwh:,.1f} MWh"],
        ["Duration", f"{basics.duration_hours:.1f} hours"],
        ["Total Investment", format_currency(capex_total, decimals=1)],
        ["Analysis Period", f"{basics.analysis_period_years} years"],
        ["Discount Rate", format_percent(basics.discount_rate)],
        ["Assumption Library", project.assumption_library or "Custom"],
    ]
    info_table = Table(info_data, colWidths=[2.5 * inch, 4.5 * inch])
    info_table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("LINEBELOW", (0, -1), (-1, -1), 1, colors.grey),
    ]))
    elements.append(info_table)
    elements.append(Spacer(1, 16))

    # Key metrics
    elements.append(Paragraph("Key Financial Metrics", heading_style))

    rec_text, rec_color = _get_recommendation(results.bcr)
    metrics_data = [
        ["Metric", "Value", "Assessment"],
        ["Benefit-Cost Ratio (BCR)", f"{results.bcr:.2f}",
         "Pass" if results.bcr >= 1.0 else "Fail"],
        ["Net Present Value (NPV)", format_currency(results.npv, decimals=1),
         "Positive" if results.npv > 0 else "Negative"],
        ["Internal Rate of Return",
         format_percent(results.irr) if results.irr is not None else "N/A",
         f"{'>' if results.irr and results.irr > basics.discount_rate else '<'} discount rate"
         if results.irr is not None else ""],
        ["Payback Period", format_years(results.payback_years),
         f"{'Within' if results.payback_years and results.payback_years <= basics.analysis_period_years else 'Beyond'} analysis period"
         if results.payback_years is not None else "Not achieved"],
        ["LCOS", f"${results.lcos_per_mwh:,.1f}/MWh", ""],
        ["Breakeven CapEx", f"${results.breakeven_capex_per_kwh:,.0f}/kWh", ""],
    ]
    metrics_table = Table(metrics_data, colWidths=[2.5 * inch, 2 * inch, 2.5 * inch])
    metrics_table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1565c0")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("ALIGN", (1, 1), (1, -1), "RIGHT"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f5f5f5")]),
    ]))
    elements.append(metrics_table)
    elements.append(Spacer(1, 12))

    # Recommendation
    elements.append(Paragraph(f"<b>Recommendation:</b> {rec_text}", body_style))

    # --- PAGE 2: Financial Analysis ---
    elements.append(PageBreak())
    elements.append(Paragraph("Financial Analysis", heading_style))

    # Cost breakdown
    elements.append(Paragraph("Cost Breakdown", styles["Heading3"]))
    cost_data = [["Cost Category", "Value"]]
    cost_data.append(["Capital Expenditure (Year 0)", format_currency(results.annual_costs[0], decimals=1)])

    # Calculate PV of O&M
    r = basics.discount_rate
    pv_om = sum(
        results.annual_costs[t] / (1 + r) ** t
        for t in range(1, len(results.annual_costs))
    )
    cost_data.append(["PV of O&M & Other Costs", format_currency(pv_om, decimals=1)])
    cost_data.append(["Total PV of Costs", format_currency(results.pv_costs, decimals=1)])

    cost_table = Table(cost_data, colWidths=[4 * inch, 3 * inch])
    cost_table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e0e0e0")),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
    ]))
    elements.append(cost_table)
    elements.append(Spacer(1, 12))

    # Benefit breakdown
    elements.append(Paragraph("Benefit Breakdown", styles["Heading3"]))
    ben_data = [["Benefit Category", "PV ($)", "% of Total"]]
    for name, pct in results.benefit_breakdown.items():
        pv_val = results.pv_benefits * pct / 100
        ben_data.append([name, format_currency(pv_val, decimals=1), f"{pct:.1f}%"])
    ben_data.append(["Total PV of Benefits", format_currency(results.pv_benefits, decimals=1), "100.0%"])

    ben_table = Table(ben_data, colWidths=[3 * inch, 2 * inch, 2 * inch])
    ben_table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e0e0e0")),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
    ]))
    elements.append(ben_table)
    elements.append(Spacer(1, 12))

    # Pie chart
    with tempfile.TemporaryDirectory() as tmpdir:
        pie_path = str(Path(tmpdir) / "benefit_pie.png")
        cf_path = str(Path(tmpdir) / "cashflow.png")

        if results.benefit_breakdown:
            create_benefit_pie_chart(results.benefit_breakdown, pie_path)
            elements.append(Image(pie_path, width=4.5 * inch, height=3.75 * inch))

        # Cash flow chart
        elements.append(PageBreak())
        elements.append(Paragraph("Cash Flow Analysis", heading_style))

        if results.annual_costs and results.annual_benefits:
            create_cashflow_chart(results.annual_costs, results.annual_benefits, cf_path)
            elements.append(Image(cf_path, width=6.5 * inch, height=3.5 * inch))
        elements.append(Spacer(1, 12))

        # --- PAGE 3: Methodology & Citations ---
        elements.append(Paragraph("Methodology & Assumptions", heading_style))

        tech = project.technology
        method_text = (
            f"This analysis uses a standard discounted cash flow (DCF) methodology "
            f"to evaluate the economic viability of a {basics.capacity_mw:,.0f} MW / "
            f"{basics.capacity_mwh:,.0f} MWh battery energy storage system. "
            f"All costs and benefits are discounted at a nominal rate of "
            f"{format_percent(basics.discount_rate)} over a {basics.analysis_period_years}-year "
            f"analysis period.<br/><br/>"
            f"<b>Technology Assumptions:</b> {tech.chemistry} battery chemistry with "
            f"{format_percent(tech.round_trip_efficiency)} round-trip efficiency and "
            f"{format_percent(tech.degradation_rate_annual)} annual degradation. "
            f"Battery augmentation occurs in Year {tech.augmentation_year}.<br/><br/>"
            f"<b>Key Formulas:</b><br/>"
            f"&bull; NPV = Sum of CF_t / (1+r)^t for t = 0 to N<br/>"
            f"&bull; BCR = PV(Benefits) / PV(Costs) [CPUC Standard Practice Manual]<br/>"
            f"&bull; IRR = Discount rate where NPV = 0<br/>"
            f"&bull; LCOS = PV(Costs) / PV(Energy Discharged) [Lazard methodology]<br/>"
        )
        elements.append(Paragraph(method_text, body_style))
        elements.append(Spacer(1, 12))

        # Citations
        elements.append(Paragraph("Data Sources & Citations", heading_style))
        citations = set()
        for benefit in project.benefits:
            if benefit.citation:
                citations.add(benefit.citation)

        # Standard references
        citations.add(
            "NREL. Annual Technology Baseline 2024. National Renewable Energy Laboratory. "
            "https://atb.nrel.gov/"
        )
        citations.add(
            "Lazard. Lazard's Levelized Cost of Storage Analysis, Version 10.0. 2025."
        )
        citations.add(
            "California Public Utilities Commission. California Standard Practice Manual: "
            "Economic Analysis of Demand-Side Programs and Projects. 2001."
        )
        citations.add(
            "Brealey, R., Myers, S., & Allen, F. Principles of Corporate Finance (13th ed.). "
            "McGraw-Hill, 2020."
        )

        for i, cite in enumerate(sorted(citations), 1):
            elements.append(Paragraph(f"[{i}] {cite}", small_style))

        elements.append(Spacer(1, 20))
        elements.append(Paragraph(
            "<i>Report generated by BESS Analyzer v1.0. All calculations are reproducible "
            "from the documented inputs and methodology above.</i>",
            small_style,
        ))

        # Build PDF (must happen while tmpdir exists for chart images)
        doc.build(elements)
