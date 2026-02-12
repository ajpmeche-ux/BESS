"""Sensitivity analysis widget for BESS Analyzer.

Displays NPV and BCR sensitivity tables showing how metrics change
with different CapEx and benefit multiplier assumptions.
"""

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QBrush, QColor
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

from src.models.project import FinancialResults, Project
from src.utils.formatters import format_currency


# Color palette for cell backgrounds (readable on both light and dark themes)
COLOR_GREEN = QColor("#c6efce")
COLOR_YELLOW = QColor("#ffeb9c")
COLOR_RED = QColor("#ffc7ce")
COLOR_DARK_TEXT = QColor("#000000")
COLOR_BASE_ROW = QColor("#d6e4f0")  # Highlight for the base-case row


class SensitivityWidget(QWidget):
    """Sensitivity analysis display with NPV and BCR tables."""

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

        # Single Variable Impacts
        single_group = QGroupBox("Single Variable Impacts (\u00b120%)")
        single_layout = QVBoxLayout(single_group)
        self.single_table = QTableWidget()
        self.single_table.setMinimumHeight(150)
        single_layout.addWidget(self.single_table)
        content_layout.addWidget(single_group)

        layout.addWidget(self.content)
        scroll.setWidget(container)
        outer.addWidget(scroll)

    def display_sensitivity(self, project: Project, results: FinancialResults):
        """Calculate and display sensitivity analysis tables."""
        self.placeholder.setVisible(False)
        self.content.setVisible(True)

        base_capex = project.costs.capex_per_kwh

        # Build CapEx levels centered on the user's input value
        # Use ~10% steps, rounded to nearest 10
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
        capex_levels = [max(10, c) for c in capex_levels]  # no negatives
        base_row_idx = 3  # center row is the user's input

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
            self.npv_table,
            capex_levels,
            benefit_multipliers,
            npv_matrix,
            is_bcr=False,
            base_row_idx=base_row_idx,
        )

        # Fill BCR table
        self._fill_sensitivity_table(
            self.bcr_table,
            capex_levels,
            benefit_multipliers,
            bcr_matrix,
            is_bcr=True,
            base_row_idx=base_row_idx,
        )

        # Fill single variable table
        self._fill_single_variable_table(project, results)

    def _fill_sensitivity_table(
        self,
        table: QTableWidget,
        capex_levels: list,
        benefit_mults: list,
        matrix: list,
        is_bcr: bool,
        base_row_idx: int = 3,
    ):
        """Fill a sensitivity table with colored values."""
        table.clear()
        table.setRowCount(len(capex_levels))
        table.setColumnCount(len(benefit_mults) + 1)

        # Headers
        headers = ["CapEx"] + [f"{int(m*100)}% Benefits" for m in benefit_mults]
        table.setHorizontalHeaderLabels(headers)

        text_brush = QBrush(COLOR_DARK_TEXT)

        for row_idx, capex in enumerate(capex_levels):
            is_base_row = (row_idx == base_row_idx)

            # Row header (CapEx level)
            label = f"${capex}/kWh"
            if is_base_row:
                label += " \u25c0"  # arrow marker for base case
            capex_item = QTableWidgetItem(label)
            capex_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            capex_item.setForeground(text_brush)
            if is_base_row:
                capex_item.setBackground(QBrush(COLOR_BASE_ROW))
            table.setItem(row_idx, 0, capex_item)

            # Values
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
                    if value >= 0:
                        bg_color = COLOR_GREEN
                    else:
                        bg_color = COLOR_RED

                item = QTableWidgetItem(text)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                item.setBackground(QBrush(bg_color))
                item.setForeground(text_brush)
                table.setItem(row_idx, col_idx + 1, item)

        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        table.verticalHeader().setVisible(True)

    def _fill_single_variable_table(self, project: Project, results: FinancialResults):
        """Fill single variable sensitivity table."""
        self.single_table.clear()
        self.single_table.setRowCount(4)
        self.single_table.setColumnCount(4)
        self.single_table.setHorizontalHeaderLabels(
            ["Parameter", "Base Value", "-20% NPV", "+20% NPV"]
        )

        base_npv = results.npv
        text_brush = QBrush(COLOR_DARK_TEXT)

        rows = [
            (
                "CapEx ($/kWh)",
                f"${project.costs.capex_per_kwh:.0f}",
                base_npv + results.annual_costs[0] * 0.2,
                base_npv - results.annual_costs[0] * 0.2
            ),
            (
                "Total Benefits",
                format_currency(results.pv_benefits, decimals=1),
                base_npv - results.pv_benefits * 0.2,
                base_npv + results.pv_benefits * 0.2
            ),
            (
                "Discount Rate",
                f"{project.basics.discount_rate * 100:.1f}%",
                base_npv * 1.15,
                base_npv * 0.85
            ),
            (
                "Cycles per Day",
                f"{project.technology.cycles_per_day:.1f}",
                base_npv * 0.85,
                base_npv * 1.15
            ),
        ]

        for row_idx, (param, base, low, high) in enumerate(rows):
            items = [
                QTableWidgetItem(param),
                QTableWidgetItem(base),
                QTableWidgetItem(format_currency(low, decimals=1)),
                QTableWidgetItem(format_currency(high, decimals=1)),
            ]
            for col_idx, item in enumerate(items):
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                item.setForeground(text_brush)
                # Color the NPV cells
                if col_idx >= 2:
                    val = low if col_idx == 2 else high
                    item.setBackground(QBrush(COLOR_GREEN if val >= 0 else COLOR_RED))
                self.single_table.setItem(row_idx, col_idx, item)

        self.single_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
