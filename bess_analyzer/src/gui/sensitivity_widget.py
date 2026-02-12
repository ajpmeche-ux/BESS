"""Sensitivity analysis widget for BESS Analyzer.

Displays NPV and BCR sensitivity tables, single-variable impacts,
and a tornado chart showing which parameters move BCR the most.
"""

import copy
import io

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QBrush, QColor, QPixmap
from PyQt6.QtWidgets import (
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QScrollArea,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from src.models.calculations import calculate_project_economics
from src.models.project import FinancialResults, Project
from src.utils.formatters import format_currency


# Color palette for cell backgrounds (readable on both light and dark themes)
COLOR_GREEN = QColor("#c6efce")
COLOR_YELLOW = QColor("#ffeb9c")
COLOR_RED = QColor("#ffc7ce")
COLOR_DARK_TEXT = QColor("#000000")
COLOR_BASE_ROW = QColor("#d6e4f0")  # Highlight for the base-case row


def _clone_project(project: Project) -> Project:
    """Deep-clone a Project for sensitivity runs."""
    d = project.to_dict()
    return Project.from_dict(d)


def _vary_project(project: Project, param_key: str, multiplier: float) -> Project:
    """Clone a project and multiply one parameter by multiplier.

    param_key format: 'section.field' e.g. 'costs.capex_per_kwh'
    """
    p = _clone_project(project)
    section, field = param_key.split(".", 1)
    obj = getattr(p, section)
    base_val = getattr(obj, field)
    setattr(obj, field, base_val * multiplier)
    return p


# Parameters to test in the tornado analysis
# (display_name, param_key, base_display_fn)
TORNADO_PARAMS = [
    ("CapEx ($/kWh)", "costs.capex_per_kwh",
     lambda p: f"${p.costs.capex_per_kwh:,.0f}"),
    ("Fixed O&M ($/kW-yr)", "costs.fom_per_kw_year",
     lambda p: f"${p.costs.fom_per_kw_year:,.1f}"),
    ("Charging Cost ($/MWh)", "costs.charging_cost_per_mwh",
     lambda p: f"${p.costs.charging_cost_per_mwh:,.0f}"),
    ("Interconnection ($/kW)", "costs.interconnection_per_kw",
     lambda p: f"${p.costs.interconnection_per_kw:,.0f}"),
    ("ITC Rate", "costs.itc_percent",
     lambda p: f"{p.costs.itc_percent*100:.0f}%"),
    ("Discount Rate", "basics.discount_rate",
     lambda p: f"{p.basics.discount_rate*100:.1f}%"),
    ("Round-Trip Efficiency", "technology.round_trip_efficiency",
     lambda p: f"{p.technology.round_trip_efficiency*100:.0f}%"),
    ("Degradation Rate", "technology.degradation_rate_annual",
     lambda p: f"{p.technology.degradation_rate_annual*100:.1f}%"),
    ("Cycles per Day", "technology.cycles_per_day",
     lambda p: f"{p.technology.cycles_per_day:.1f}"),
]


class SensitivityWidget(QWidget):
    """Sensitivity analysis display with NPV/BCR tables and tornado chart."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._init_ui()

    def _init_ui(self):
        outer = QVBoxLayout(self)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        container = QWidget()
        layout = QVBoxLayout(container)

        # Placeholder
        self.placeholder = QLabel("Run 'Calculate Economics' to see sensitivity analysis.")
        self.placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.placeholder.setStyleSheet("font-size: 14px; color: #999; padding: 40px;")
        layout.addWidget(self.placeholder)

        # Content container
        self.content = QWidget()
        self.content.setVisible(False)
        content_layout = QVBoxLayout(self.content)

        # Instructions
        instructions = QLabel(
            "These tables show how NPV and BCR change with different CapEx levels "
            "and benefit scaling factors. Green = BCR \u2265 1.5, Yellow = 1.0\u20131.5, "
            "Red = < 1.0. The highlighted row is your current CapEx input."
        )
        instructions.setWordWrap(True)
        instructions.setStyleSheet("color: #666; padding: 8px; background: #f5f5f5; border-radius: 4px;")
        content_layout.addWidget(instructions)

        # Tornado Chart (BCR Drivers)
        tornado_group = QGroupBox("Tornado Analysis: What Moves the BCR Most? (\u00b120%)")
        tornado_layout = QVBoxLayout(tornado_group)
        self.tornado_chart_label = QLabel()
        self.tornado_chart_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.tornado_chart_label.setMinimumSize(700, 400)
        tornado_layout.addWidget(self.tornado_chart_label)
        content_layout.addWidget(tornado_group)

        # Tornado detail table
        tornado_detail_group = QGroupBox("Single-Variable Sensitivity Detail (\u00b120%)")
        tornado_detail_layout = QVBoxLayout(tornado_detail_group)
        self.tornado_table = QTableWidget()
        self.tornado_table.setMinimumHeight(280)
        tornado_detail_layout.addWidget(self.tornado_table)
        content_layout.addWidget(tornado_detail_group)

        # NPV Table
        npv_group = QGroupBox("NPV Sensitivity ($)")
        npv_layout = QVBoxLayout(npv_group)
        self.npv_table = QTableWidget()
        self.npv_table.setMinimumHeight(250)
        npv_layout.addWidget(self.npv_table)
        content_layout.addWidget(npv_group)

        # BCR Table
        bcr_group = QGroupBox("BCR Sensitivity (Benefit-Cost Ratio)")
        bcr_layout = QVBoxLayout(bcr_group)
        self.bcr_table = QTableWidget()
        self.bcr_table.setMinimumHeight(250)
        bcr_layout.addWidget(self.bcr_table)
        content_layout.addWidget(bcr_group)

        layout.addWidget(self.content)
        scroll.setWidget(container)
        outer.addWidget(scroll)

    def display_sensitivity(self, project: Project, results: FinancialResults):
        """Calculate and display sensitivity analysis tables."""
        self.placeholder.setVisible(False)
        self.content.setVisible(True)

        base_capex = project.costs.capex_per_kwh

        # Build CapEx levels centered on the user's input value
        capex_step = max(10, round(base_capex * 0.1 / 10) * 10)
        capex_levels = [
            round(base_capex - 3 * capex_step),
            round(base_capex - 2 * capex_step),
            round(base_capex - 1 * capex_step),
            round(base_capex),
            round(base_capex + 1 * capex_step),
            round(base_capex + 2 * capex_step),
            round(base_capex + 3 * capex_step),
        ]
        capex_levels = [max(10, c) for c in capex_levels]
        base_row_idx = 3

        benefit_multipliers = [0.7, 0.8, 0.9, 1.0, 1.1, 1.2, 1.3]

        # Calculate sensitivity matrices
        npv_matrix = []
        bcr_matrix = []

        for capex in capex_levels:
            npv_row = []
            bcr_row = []
            capex_ratio = capex / base_capex if base_capex > 0 else 1.0

            for ben_mult in benefit_multipliers:
                adjusted_pv_benefits = results.pv_benefits * ben_mult
                adjusted_pv_costs = results.pv_costs * (1 + (capex_ratio - 1) * 0.7)

                npv = adjusted_pv_benefits - adjusted_pv_costs
                bcr = adjusted_pv_benefits / adjusted_pv_costs if adjusted_pv_costs > 0 else 0

                npv_row.append(npv)
                bcr_row.append(bcr)

            npv_matrix.append(npv_row)
            bcr_matrix.append(bcr_row)

        # Fill NPV table
        self._fill_sensitivity_table(
            self.npv_table, capex_levels, benefit_multipliers,
            npv_matrix, is_bcr=False, base_row_idx=base_row_idx,
        )

        # Fill BCR table
        self._fill_sensitivity_table(
            self.bcr_table, capex_levels, benefit_multipliers,
            bcr_matrix, is_bcr=True, base_row_idx=base_row_idx,
        )

        # Tornado analysis
        self._run_tornado_analysis(project, results)

    def _fill_sensitivity_table(
        self, table, capex_levels, benefit_mults, matrix, is_bcr, base_row_idx=3,
    ):
        """Fill a sensitivity table with colored values."""
        table.clear()
        table.setRowCount(len(capex_levels))
        table.setColumnCount(len(benefit_mults) + 1)

        headers = ["CapEx"] + [f"{int(m*100)}% Benefits" for m in benefit_mults]
        table.setHorizontalHeaderLabels(headers)

        text_brush = QBrush(COLOR_DARK_TEXT)

        for row_idx, capex in enumerate(capex_levels):
            is_base_row = (row_idx == base_row_idx)

            label = f"${capex}/kWh"
            if is_base_row:
                label += " \u25c0"
            capex_item = QTableWidgetItem(label)
            capex_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            capex_item.setForeground(text_brush)
            if is_base_row:
                capex_item.setBackground(QBrush(COLOR_BASE_ROW))
            table.setItem(row_idx, 0, capex_item)

            for col_idx, value in enumerate(matrix[row_idx]):
                if is_bcr:
                    text = f"{value:.2f}"
                    if value >= 1.5:
                        bg_color = COLOR_GREEN
                    elif value >= 1.0:
                        bg_color = COLOR_YELLOW
                    else:
                        bg_color = COLOR_RED
                else:
                    text = format_currency(value, decimals=1)
                    bg_color = COLOR_GREEN if value >= 0 else COLOR_RED

                item = QTableWidgetItem(text)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                item.setBackground(QBrush(bg_color))
                item.setForeground(text_brush)
                table.setItem(row_idx, col_idx + 1, item)

        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        table.verticalHeader().setVisible(True)

    def _run_tornado_analysis(self, project: Project, results: FinancialResults):
        """Run tornado sensitivity: vary each parameter +/-20%, recalculate BCR."""
        base_bcr = results.bcr
        base_npv = results.npv
        tornado_data = []  # (name, base_display, bcr_low, bcr_high, npv_low, npv_high)

        for display_name, param_key, base_fn in TORNADO_PARAMS:
            try:
                # Check if base value is zero - skip if so
                section, field = param_key.split(".", 1)
                base_val = getattr(getattr(project, section), field)
                if base_val == 0:
                    continue

                # -20%
                p_low = _vary_project(project, param_key, 0.8)
                r_low = calculate_project_economics(p_low)
                # +20%
                p_high = _vary_project(project, param_key, 1.2)
                r_high = calculate_project_economics(p_high)

                tornado_data.append((
                    display_name,
                    base_fn(project),
                    r_low.bcr,
                    r_high.bcr,
                    r_low.npv,
                    r_high.npv,
                ))
            except Exception:
                continue

        # Also vary each benefit stream individually
        for i, benefit in enumerate(project.benefits):
            if not benefit.annual_values or all(v == 0 for v in benefit.annual_values):
                continue
            try:
                # -20% on this benefit
                p_low = _clone_project(project)
                p_low.benefits[i].annual_values = [v * 0.8 for v in benefit.annual_values]
                r_low = calculate_project_economics(p_low)

                # +20%
                p_high = _clone_project(project)
                p_high.benefits[i].annual_values = [v * 1.2 for v in benefit.annual_values]
                r_high = calculate_project_economics(p_high)

                base_pv = sum(
                    v / (1 + project.get_discount_rate()) ** (t + 1)
                    for t, v in enumerate(benefit.annual_values)
                )
                tornado_data.append((
                    f"Benefit: {benefit.name}",
                    format_currency(base_pv, decimals=1),
                    r_low.bcr,
                    r_high.bcr,
                    r_low.npv,
                    r_high.npv,
                ))
            except Exception:
                continue

        # Sort by BCR swing (largest first)
        tornado_data.sort(key=lambda x: abs(x[3] - x[2]), reverse=True)

        # Draw tornado chart
        self._draw_tornado_chart(tornado_data, base_bcr)

        # Fill detail table
        self._fill_tornado_table(tornado_data, base_bcr, base_npv)

    def _draw_tornado_chart(self, tornado_data, base_bcr):
        """Draw horizontal tornado bar chart for BCR sensitivity."""
        if not tornado_data:
            self.tornado_chart_label.setText("No parameters to vary.")
            return

        fig, ax = plt.subplots(figsize=(8, max(3.5, 0.45 * len(tornado_data) + 1)), dpi=100)

        names = [d[0] for d in tornado_data]
        bcr_lows = [d[2] for d in tornado_data]
        bcr_highs = [d[3] for d in tornado_data]

        y_pos = list(range(len(names)))

        # For each parameter, draw two bars from base_bcr
        for i, (name, _, bcr_low, bcr_high, _, _) in enumerate(tornado_data):
            # Left bar (whichever direction goes below base)
            left_val = min(bcr_low, bcr_high)
            right_val = max(bcr_low, bcr_high)

            # Bar going left of base (red = worse)
            if left_val < base_bcr:
                ax.barh(i, left_val - base_bcr, left=base_bcr,
                        color="#c62828", alpha=0.7, height=0.6)
            # Bar going right of base (green = better)
            if right_val > base_bcr:
                ax.barh(i, right_val - base_bcr, left=base_bcr,
                        color="#2e7d32", alpha=0.7, height=0.6)

            # Add value labels at the ends
            ax.text(left_val - 0.01, i, f"{left_val:.2f}",
                    va="center", ha="right", fontsize=8, color="#333")
            ax.text(right_val + 0.01, i, f"{right_val:.2f}",
                    va="center", ha="left", fontsize=8, color="#333")

        # Base line
        ax.axvline(x=base_bcr, color="#333", linewidth=1.5, linestyle="--", label=f"Base BCR: {base_bcr:.2f}")

        ax.set_yticks(y_pos)
        ax.set_yticklabels(names, fontsize=9)
        ax.set_xlabel("Benefit-Cost Ratio (BCR)", fontsize=10)
        ax.set_title("Tornado Chart: BCR Sensitivity to \u00b120% Parameter Changes",
                      fontsize=11, fontweight="bold")
        ax.legend(loc="lower right", fontsize=8)
        ax.invert_yaxis()
        fig.tight_layout()

        buf = io.BytesIO()
        fig.savefig(buf, format="png", bbox_inches="tight")
        plt.close(fig)
        buf.seek(0)
        pixmap = QPixmap()
        pixmap.loadFromData(buf.read())
        self.tornado_chart_label.setPixmap(pixmap)

    def _fill_tornado_table(self, tornado_data, base_bcr, base_npv):
        """Fill tornado detail table with exact values."""
        self.tornado_table.clear()
        n = len(tornado_data)
        self.tornado_table.setRowCount(n)
        self.tornado_table.setColumnCount(7)
        self.tornado_table.setHorizontalHeaderLabels([
            "Parameter", "Base Value",
            "BCR (-20%)", "BCR (+20%)", "BCR Swing",
            "NPV (-20%)", "NPV (+20%)",
        ])

        text_brush = QBrush(COLOR_DARK_TEXT)

        for row, (name, base_display, bcr_low, bcr_high, npv_low, npv_high) in enumerate(tornado_data):
            swing = abs(bcr_high - bcr_low)
            cells = [
                name,
                base_display,
                f"{bcr_low:.3f}",
                f"{bcr_high:.3f}",
                f"{swing:.3f}",
                format_currency(npv_low, decimals=1),
                format_currency(npv_high, decimals=1),
            ]
            for col, text in enumerate(cells):
                item = QTableWidgetItem(text)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                item.setForeground(text_brush)

                # Color BCR cells
                if col == 2:  # BCR -20%
                    if bcr_low >= 1.5:
                        item.setBackground(QBrush(COLOR_GREEN))
                    elif bcr_low >= 1.0:
                        item.setBackground(QBrush(COLOR_YELLOW))
                    else:
                        item.setBackground(QBrush(COLOR_RED))
                elif col == 3:  # BCR +20%
                    if bcr_high >= 1.5:
                        item.setBackground(QBrush(COLOR_GREEN))
                    elif bcr_high >= 1.0:
                        item.setBackground(QBrush(COLOR_YELLOW))
                    else:
                        item.setBackground(QBrush(COLOR_RED))
                elif col == 4:  # Swing
                    item.setBackground(QBrush(COLOR_BASE_ROW))
                elif col >= 5:  # NPV cells
                    val = npv_low if col == 5 else npv_high
                    item.setBackground(QBrush(COLOR_GREEN if val >= 0 else COLOR_RED))

                self.tornado_table.setItem(row, col, item)

        self.tornado_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
