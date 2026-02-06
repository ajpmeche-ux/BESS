"""Results display widget for BESS Analyzer.

Shows calculated financial metrics, benefit breakdown chart,
and annual cash flow summary table.
"""

import io

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPixmap
from PyQt6.QtWidgets import (
    QFrame,
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

from src.models.project import FinancialResults, Project
from src.utils.formatters import format_currency, format_percent, format_years


class MetricCard(QFrame):
    """A single metric display card with label, value, and optional color."""

    def __init__(self, label: str, value: str = "--", color: str = "#333", parent=None):
        super().__init__(parent)
        self.setFrameShape(QFrame.Shape.Box)
        self.setStyleSheet(f"""
            MetricCard {{
                border: 2px solid #ddd;
                border-radius: 8px;
                padding: 8px;
                background: white;
            }}
        """)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 8, 12, 8)

        self.label = QLabel(label)
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.label.setStyleSheet("font-size: 11px; color: #666; border: none;")
        layout.addWidget(self.label)

        self.value_label = QLabel(value)
        self.value_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.value_label.setStyleSheet(f"font-size: 22px; font-weight: bold; color: {color}; border: none;")
        layout.addWidget(self.value_label)

    def set_value(self, text: str, color: str = "#333"):
        self.value_label.setText(text)
        self.value_label.setStyleSheet(f"font-size: 22px; font-weight: bold; color: {color}; border: none;")


class ResultsWidget(QWidget):
    """Results display with metrics, charts, and cash flow table."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._init_ui()

    def _init_ui(self):
        outer = QVBoxLayout(self)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        container = QWidget()
        layout = QVBoxLayout(container)

        # Placeholder label (shown when no results)
        self.placeholder = QLabel("Run 'Calculate Economics' to see results here.")
        self.placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.placeholder.setStyleSheet("font-size: 14px; color: #999; padding: 40px;")
        layout.addWidget(self.placeholder)

        # Results content (hidden until calculation)
        self.results_container = QWidget()
        results_layout = QVBoxLayout(self.results_container)
        self.results_container.setVisible(False)

        # Section 1: Key Metrics
        metrics_layout = QHBoxLayout()
        self.bcr_card = MetricCard("Benefit-Cost Ratio")
        self.npv_card = MetricCard("Net Present Value")
        self.irr_card = MetricCard("Internal Rate of Return")
        self.payback_card = MetricCard("Payback Period")
        self.lcos_card = MetricCard("LCOS")
        for card in [self.bcr_card, self.npv_card, self.irr_card, self.payback_card, self.lcos_card]:
            metrics_layout.addWidget(card)
        results_layout.addLayout(metrics_layout)

        # Section 2: Chart and Summary side by side
        mid_layout = QHBoxLayout()

        # Pie chart
        chart_group = QGroupBox("Benefit Breakdown")
        chart_layout = QVBoxLayout(chart_group)
        self.chart_label = QLabel()
        self.chart_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.chart_label.setMinimumSize(350, 300)
        chart_layout.addWidget(self.chart_label)
        mid_layout.addWidget(chart_group)

        # Cost summary
        summary_group = QGroupBox("Cost & Benefit Summary")
        summary_layout = QVBoxLayout(summary_group)
        self.summary_table = QTableWidget()
        self.summary_table.setColumnCount(2)
        self.summary_table.setHorizontalHeaderLabels(["Item", "Value"])
        self.summary_table.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeMode.Stretch
        )
        self.summary_table.horizontalHeader().setSectionResizeMode(
            1, QHeaderView.ResizeMode.ResizeToContents
        )
        summary_layout.addWidget(self.summary_table)
        mid_layout.addWidget(summary_group)

        results_layout.addLayout(mid_layout)

        # Section 3: Cash Flow Table
        cf_group = QGroupBox("Annual Cash Flows (Selected Years)")
        cf_layout = QVBoxLayout(cf_group)
        self.cf_table = QTableWidget()
        self.cf_table.setMinimumHeight(200)
        cf_layout.addWidget(self.cf_table)
        results_layout.addWidget(cf_group)

        layout.addWidget(self.results_container)
        scroll.setWidget(container)
        outer.addWidget(scroll)

    def display_results(self, project: Project, results: FinancialResults):
        """Populate all result widgets with calculated data."""
        self.placeholder.setVisible(False)
        self.results_container.setVisible(True)

        # Key metrics
        bcr_color = "#2e7d32" if results.bcr >= 1.5 else ("#f9a825" if results.bcr >= 1.0 else "#c62828")
        self.bcr_card.set_value(f"{results.bcr:.2f}", bcr_color)

        npv_color = "#2e7d32" if results.npv >= 0 else "#c62828"
        self.npv_card.set_value(format_currency(results.npv, decimals=1), npv_color)

        if results.irr is not None:
            self.irr_card.set_value(format_percent(results.irr))
        else:
            self.irr_card.set_value("N/A", "#999")

        self.payback_card.set_value(format_years(results.payback_years))

        self.lcos_card.set_value(f"${results.lcos_per_mwh:,.1f}/MWh")

        # Benefit pie chart
        self._draw_pie_chart(results.benefit_breakdown)

        # Summary table
        self._fill_summary_table(project, results)

        # Cash flow table
        self._fill_cashflow_table(results)

    def _draw_pie_chart(self, breakdown: dict):
        if not breakdown:
            self.chart_label.setText("No benefit data.")
            return

        fig, ax = plt.subplots(figsize=(4, 3.5), dpi=100)
        labels = list(breakdown.keys())
        sizes = list(breakdown.values())
        colors = ["#1976d2", "#388e3c", "#f57c00", "#7b1fa2", "#c62828", "#00838f"]
        wedges, texts, autotexts = ax.pie(
            sizes, labels=labels, autopct="%1.1f%%",
            colors=colors[:len(labels)],
            textprops={"fontsize": 9},
        )
        ax.set_title("Benefit Breakdown by Category", fontsize=11, fontweight="bold")
        fig.tight_layout()

        buf = io.BytesIO()
        fig.savefig(buf, format="png", bbox_inches="tight")
        plt.close(fig)
        buf.seek(0)
        pixmap = QPixmap()
        pixmap.loadFromData(buf.read())
        self.chart_label.setPixmap(pixmap)

    def _fill_summary_table(self, project: Project, results: FinancialResults):
        rows = [
            ("Total CapEx", format_currency(results.annual_costs[0], decimals=1)),
            ("PV of All Costs", format_currency(results.pv_costs, decimals=1)),
            ("PV of All Benefits", format_currency(results.pv_benefits, decimals=1)),
            ("Net Present Value", format_currency(results.npv, decimals=1)),
            ("Breakeven CapEx", f"${results.breakeven_capex_per_kwh:,.0f}/kWh"),
        ]
        # Add benefit category PVs
        for name, pct in results.benefit_breakdown.items():
            pv_val = results.pv_benefits * pct / 100
            rows.append((f"  {name} (PV)", format_currency(pv_val, decimals=1)))

        self.summary_table.setRowCount(len(rows))
        for i, (label, value) in enumerate(rows):
            self.summary_table.setItem(i, 0, QTableWidgetItem(label))
            item = QTableWidgetItem(value)
            item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.summary_table.setItem(i, 1, item)

    def _fill_cashflow_table(self, results: FinancialResults):
        n = len(results.annual_costs) - 1
        # Show years: 0, 1, 5, 10, 15, 20 (and last if different)
        years = sorted(set([0, 1] + list(range(5, n + 1, 5)) + [n]))
        years = [y for y in years if y <= n]

        self.cf_table.setColumnCount(len(years))
        self.cf_table.setHorizontalHeaderLabels([f"Year {y}" for y in years])
        self.cf_table.setRowCount(3)
        self.cf_table.setVerticalHeaderLabels(["Costs", "Benefits", "Net"])

        for col, y in enumerate(years):
            for row_idx, vals in enumerate([results.annual_costs, results.annual_benefits, results.annual_net]):
                val = vals[y] if y < len(vals) else 0.0
                item = QTableWidgetItem(format_currency(val, decimals=1))
                item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                self.cf_table.setItem(row_idx, col, item)

        self.cf_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
