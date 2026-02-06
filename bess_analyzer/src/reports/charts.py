"""Chart generation for BESS Analyzer reports.

Creates publication-quality matplotlib charts for benefit breakdown,
cash flows, and sensitivity analysis. Charts are saved as PNG files
for embedding in PDF reports.
"""

from typing import Dict, List

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker


def create_benefit_pie_chart(benefit_breakdown: Dict[str, float], output_path: str) -> None:
    """Create a pie chart showing benefit breakdown by category.

    Args:
        benefit_breakdown: Dict mapping category name to percentage of total PV.
        output_path: File path to save the PNG chart.
    """
    if not benefit_breakdown:
        return

    fig, ax = plt.subplots(figsize=(6, 5), dpi=150)
    labels = list(benefit_breakdown.keys())
    sizes = list(benefit_breakdown.values())
    colors = ["#1565c0", "#2e7d32", "#ef6c00", "#6a1b9a", "#c62828", "#00838f", "#4e342e"]

    wedges, texts, autotexts = ax.pie(
        sizes,
        labels=labels,
        autopct="%1.1f%%",
        colors=colors[:len(labels)],
        textprops={"fontsize": 10},
        startangle=90,
    )
    for autotext in autotexts:
        autotext.set_fontweight("bold")

    ax.set_title("Benefit Breakdown by Category", fontsize=13, fontweight="bold", pad=15)
    fig.tight_layout()
    fig.savefig(output_path, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def create_cashflow_chart(
    annual_costs: List[float],
    annual_benefits: List[float],
    output_path: str,
) -> None:
    """Create a bar chart showing annual costs and benefits over time.

    Args:
        annual_costs: Annual cost values for years 0..N.
        annual_benefits: Annual benefit values for years 0..N.
        output_path: File path to save the PNG chart.
    """
    n = len(annual_costs)
    years = list(range(n))

    fig, ax = plt.subplots(figsize=(8, 4.5), dpi=150)
    bar_width = 0.35

    ax.bar(
        [y - bar_width / 2 for y in years],
        [-c / 1e6 for c in annual_costs],
        bar_width,
        label="Costs",
        color="#c62828",
        alpha=0.8,
    )
    ax.bar(
        [y + bar_width / 2 for y in years],
        [b / 1e6 for b in annual_benefits],
        bar_width,
        label="Benefits",
        color="#2e7d32",
        alpha=0.8,
    )

    ax.set_xlabel("Year", fontsize=11)
    ax.set_ylabel("$ Millions", fontsize=11)
    ax.set_title("Annual Costs and Benefits", fontsize=13, fontweight="bold")
    ax.legend(fontsize=10)
    ax.axhline(y=0, color="black", linewidth=0.5)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"${x:,.0f}M"))
    ax.grid(axis="y", alpha=0.3)

    fig.tight_layout()
    fig.savefig(output_path, bbox_inches="tight", facecolor="white")
    plt.close(fig)
