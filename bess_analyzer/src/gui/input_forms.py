"""Input forms for BESS Analyzer project parameters.

Provides a scrollable widget with grouped input sections for project
basics, technology specs, costs, and benefit streams.
"""

from datetime import date

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QComboBox,
    QDateEdit,
    QDoubleSpinBox,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QPushButton,
    QScrollArea,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from src.data.libraries import AssumptionLibrary
from src.models.project import BenefitStream, CostInputs, Project, ProjectBasics, TechnologySpecs


class InputFormWidget(QWidget):
    """Scrollable input form for all BESS project parameters."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._library = AssumptionLibrary()
        self._init_ui()

    def _init_ui(self):
        outer = QVBoxLayout(self)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        container = QWidget()
        layout = QVBoxLayout(container)

        # Section 1: Project Basics
        layout.addWidget(self._create_basics_section())

        # Section 2: Library Selector
        layout.addWidget(self._create_library_section())

        # Section 3: Technology Specs
        layout.addWidget(self._create_technology_section())

        # Section 4: Costs
        layout.addWidget(self._create_costs_section())

        # Section 5: Benefits
        layout.addWidget(self._create_benefits_section())

        layout.addStretch()
        scroll.setWidget(container)
        outer.addWidget(scroll)

    # --- Section Builders ---

    def _create_basics_section(self) -> QGroupBox:
        group = QGroupBox("Project Basics")
        layout = QVBoxLayout(group)

        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("e.g., Moss Landing BESS")
        layout.addLayout(self._row("Project Name:", self.name_edit))

        self.project_id_edit = QLineEdit()
        self.project_id_edit.setPlaceholderText("e.g., BESS-001")
        layout.addLayout(self._row("Project ID:", self.project_id_edit))

        self.location_edit = QLineEdit()
        self.location_edit.setPlaceholderText("e.g., CAISO NP15")
        layout.addLayout(self._row("Location:", self.location_edit))

        self.capacity_spin = QDoubleSpinBox()
        self.capacity_spin.setRange(0.1, 2000.0)
        self.capacity_spin.setValue(100.0)
        self.capacity_spin.setSuffix(" MW")
        self.capacity_spin.setDecimals(1)
        self.capacity_spin.valueChanged.connect(self._update_energy)
        layout.addLayout(self._row("Capacity:", self.capacity_spin))

        self.duration_spin = QDoubleSpinBox()
        self.duration_spin.setRange(0.5, 24.0)
        self.duration_spin.setValue(4.0)
        self.duration_spin.setSuffix(" hours")
        self.duration_spin.setDecimals(1)
        self.duration_spin.valueChanged.connect(self._update_energy)
        layout.addLayout(self._row("Duration:", self.duration_spin))

        self.energy_label = QLabel("400.0 MWh")
        self.energy_label.setStyleSheet("font-weight: bold;")
        layout.addLayout(self._row("Energy Capacity:", self.energy_label))

        self.date_edit = QDateEdit()
        self.date_edit.setDate(date(2027, 1, 1))
        self.date_edit.setCalendarPopup(True)
        layout.addLayout(self._row("In-Service Date:", self.date_edit))

        self.period_spin = QSpinBox()
        self.period_spin.setRange(10, 30)
        self.period_spin.setValue(20)
        self.period_spin.setSuffix(" years")
        layout.addLayout(self._row("Analysis Period:", self.period_spin))

        self.discount_spin = QDoubleSpinBox()
        self.discount_spin.setRange(1.0, 20.0)
        self.discount_spin.setValue(7.0)
        self.discount_spin.setSuffix(" %")
        self.discount_spin.setDecimals(1)
        layout.addLayout(self._row("Discount Rate:", self.discount_spin))

        return group

    def _create_library_section(self) -> QGroupBox:
        group = QGroupBox("Assumption Library")
        layout = QVBoxLayout(group)

        row = QHBoxLayout()
        self.library_combo = QComboBox()
        self.library_combo.addItem("-- Select Library --")
        for name in self._library.get_library_names():
            self.library_combo.addItem(name)
        self.library_combo.currentTextChanged.connect(self._on_library_selected)
        row.addWidget(self.library_combo, stretch=1)

        self.load_library_btn = QPushButton("Load Library")
        self.load_library_btn.clicked.connect(self.load_library)
        row.addWidget(self.load_library_btn)
        layout.addLayout(row)

        self.library_info_label = QLabel("")
        self.library_info_label.setWordWrap(True)
        self.library_info_label.setStyleSheet("color: #666; font-size: 11px;")
        layout.addWidget(self.library_info_label)

        return group

    def _create_technology_section(self) -> QGroupBox:
        group = QGroupBox("Technology Specifications")
        layout = QVBoxLayout(group)

        self.chemistry_combo = QComboBox()
        self.chemistry_combo.addItems(["LFP", "NMC", "Other"])
        layout.addLayout(self._row("Chemistry:", self.chemistry_combo))

        self.rte_spin = QDoubleSpinBox()
        self.rte_spin.setRange(70.0, 95.0)
        self.rte_spin.setValue(85.0)
        self.rte_spin.setSuffix(" %")
        self.rte_spin.setDecimals(1)
        layout.addLayout(self._row("Round-Trip Efficiency:", self.rte_spin))

        self.degradation_spin = QDoubleSpinBox()
        self.degradation_spin.setRange(0.5, 5.0)
        self.degradation_spin.setValue(2.5)
        self.degradation_spin.setSuffix(" %/yr")
        self.degradation_spin.setDecimals(2)
        layout.addLayout(self._row("Annual Degradation:", self.degradation_spin))

        self.cycle_life_spin = QSpinBox()
        self.cycle_life_spin.setRange(1000, 10000)
        self.cycle_life_spin.setValue(6000)
        self.cycle_life_spin.setSuffix(" cycles")
        layout.addLayout(self._row("Cycle Life:", self.cycle_life_spin))

        self.augmentation_year_spin = QSpinBox()
        self.augmentation_year_spin.setRange(1, 25)
        self.augmentation_year_spin.setValue(12)
        self.augmentation_year_spin.setSuffix(" year")
        layout.addLayout(self._row("Augmentation Year:", self.augmentation_year_spin))

        return group

    def _create_costs_section(self) -> QGroupBox:
        group = QGroupBox("Cost Inputs")
        layout = QVBoxLayout(group)

        self.capex_spin = QDoubleSpinBox()
        self.capex_spin.setRange(0, 500.0)
        self.capex_spin.setValue(160.0)
        self.capex_spin.setPrefix("$")
        self.capex_spin.setSuffix(" /kWh")
        self.capex_spin.setDecimals(0)
        layout.addLayout(self._row("CapEx:", self.capex_spin))

        self.fom_spin = QDoubleSpinBox()
        self.fom_spin.setRange(0, 100.0)
        self.fom_spin.setValue(25.0)
        self.fom_spin.setPrefix("$")
        self.fom_spin.setSuffix(" /kW-yr")
        self.fom_spin.setDecimals(1)
        layout.addLayout(self._row("Fixed O&M:", self.fom_spin))

        self.vom_spin = QDoubleSpinBox()
        self.vom_spin.setRange(0, 50.0)
        self.vom_spin.setValue(0.0)
        self.vom_spin.setPrefix("$")
        self.vom_spin.setSuffix(" /MWh")
        self.vom_spin.setDecimals(2)
        layout.addLayout(self._row("Variable O&M:", self.vom_spin))

        self.aug_cost_spin = QDoubleSpinBox()
        self.aug_cost_spin.setRange(0, 300.0)
        self.aug_cost_spin.setValue(55.0)
        self.aug_cost_spin.setPrefix("$")
        self.aug_cost_spin.setSuffix(" /kWh")
        self.aug_cost_spin.setDecimals(0)
        layout.addLayout(self._row("Augmentation Cost:", self.aug_cost_spin))

        self.decom_spin = QDoubleSpinBox()
        self.decom_spin.setRange(0, 50.0)
        self.decom_spin.setValue(10.0)
        self.decom_spin.setPrefix("$")
        self.decom_spin.setSuffix(" /kW")
        self.decom_spin.setDecimals(1)
        layout.addLayout(self._row("Decommissioning:", self.decom_spin))

        return group

    def _create_benefits_section(self) -> QGroupBox:
        group = QGroupBox("Benefit Streams")
        layout = QVBoxLayout(group)

        self.benefits_table = QTableWidget(0, 4)
        self.benefits_table.setHorizontalHeaderLabels([
            "Name", "Year 1 Value ($/yr)", "Escalation (%)", "Citation"
        ])
        header = self.benefits_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        self.benefits_table.setMinimumHeight(150)
        layout.addWidget(self.benefits_table)

        btn_row = QHBoxLayout()
        add_btn = QPushButton("Add Benefit")
        add_btn.clicked.connect(self._add_benefit_row)
        btn_row.addWidget(add_btn)
        remove_btn = QPushButton("Remove Selected")
        remove_btn.clicked.connect(self._remove_benefit_row)
        btn_row.addWidget(remove_btn)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        return group

    # --- Helper Methods ---

    @staticmethod
    def _row(label_text: str, widget) -> QHBoxLayout:
        row = QHBoxLayout()
        lbl = QLabel(label_text)
        lbl.setFixedWidth(160)
        row.addWidget(lbl)
        row.addWidget(widget)
        return row

    def _update_energy(self):
        mwh = self.capacity_spin.value() * self.duration_spin.value()
        self.energy_label.setText(f"{mwh:,.1f} MWh")

    def _on_library_selected(self, name: str):
        if name == "-- Select Library --":
            self.library_info_label.setText("")
            return
        meta = self._library.get_library_metadata(name)
        self.library_info_label.setText(
            f"Source: {meta['source']} | Version: {meta['version']} | "
            f"Published: {meta['date_published']}\n{meta['notes']}"
        )

    def _add_benefit_row(self):
        row = self.benefits_table.rowCount()
        self.benefits_table.insertRow(row)
        self.benefits_table.setItem(row, 0, QTableWidgetItem(""))
        self.benefits_table.setItem(row, 1, QTableWidgetItem("0"))
        self.benefits_table.setItem(row, 2, QTableWidgetItem("2.0"))
        self.benefits_table.setItem(row, 3, QTableWidgetItem(""))

    def _remove_benefit_row(self):
        row = self.benefits_table.currentRow()
        if row >= 0:
            self.benefits_table.removeRow(row)

    # --- Public API ---

    def load_library(self):
        """Apply the selected assumption library to the form fields."""
        name = self.library_combo.currentText()
        if name == "-- Select Library --":
            return

        # Create a temporary project to apply library
        project = self.get_project()
        self._library.apply_library_to_project(project, name)
        self.load_project(project)

    def load_project(self, project: Project):
        """Populate all form fields from a Project object."""
        b = project.basics
        self.name_edit.setText(b.name)
        self.project_id_edit.setText(b.project_id)
        self.location_edit.setText(b.location)
        self.capacity_spin.setValue(b.capacity_mw)
        self.duration_spin.setValue(b.duration_hours)
        self.date_edit.setDate(b.in_service_date)
        self.period_spin.setValue(b.analysis_period_years)
        self.discount_spin.setValue(b.discount_rate * 100)

        t = project.technology
        idx = self.chemistry_combo.findText(t.chemistry)
        if idx >= 0:
            self.chemistry_combo.setCurrentIndex(idx)
        self.rte_spin.setValue(t.round_trip_efficiency * 100)
        self.degradation_spin.setValue(t.degradation_rate_annual * 100)
        self.cycle_life_spin.setValue(t.cycle_life)
        self.augmentation_year_spin.setValue(t.augmentation_year)

        c = project.costs
        self.capex_spin.setValue(c.capex_per_kwh)
        self.fom_spin.setValue(c.fom_per_kw_year)
        self.vom_spin.setValue(c.vom_per_mwh)
        self.aug_cost_spin.setValue(c.augmentation_per_kwh)
        self.decom_spin.setValue(c.decommissioning_per_kw)

        # Benefits table
        self.benefits_table.setRowCount(0)
        for benefit in project.benefits:
            row = self.benefits_table.rowCount()
            self.benefits_table.insertRow(row)
            self.benefits_table.setItem(row, 0, QTableWidgetItem(benefit.name))
            year1_val = benefit.annual_values[0] if benefit.annual_values else 0
            self.benefits_table.setItem(row, 1, QTableWidgetItem(f"{year1_val:,.0f}"))

            # Infer escalation from first two years
            if len(benefit.annual_values) >= 2 and benefit.annual_values[0] > 0:
                esc = (benefit.annual_values[1] / benefit.annual_values[0] - 1) * 100
            else:
                esc = 0.0
            self.benefits_table.setItem(row, 2, QTableWidgetItem(f"{esc:.1f}"))
            self.benefits_table.setItem(row, 3, QTableWidgetItem(benefit.citation))

        if project.assumption_library:
            idx = self.library_combo.findText(project.assumption_library)
            if idx >= 0:
                self.library_combo.setCurrentIndex(idx)

    def get_project(self) -> Project:
        """Read all form values and construct a Project object."""
        basics = ProjectBasics(
            name=self.name_edit.text(),
            project_id=self.project_id_edit.text(),
            location=self.location_edit.text(),
            capacity_mw=self.capacity_spin.value(),
            duration_hours=self.duration_spin.value(),
            in_service_date=self.date_edit.date().toPyDate(),
            analysis_period_years=self.period_spin.value(),
            discount_rate=self.discount_spin.value() / 100,
        )

        technology = TechnologySpecs(
            chemistry=self.chemistry_combo.currentText(),
            round_trip_efficiency=self.rte_spin.value() / 100,
            degradation_rate_annual=self.degradation_spin.value() / 100,
            cycle_life=self.cycle_life_spin.value(),
            augmentation_year=self.augmentation_year_spin.value(),
        )

        costs = CostInputs(
            capex_per_kwh=self.capex_spin.value(),
            fom_per_kw_year=self.fom_spin.value(),
            vom_per_mwh=self.vom_spin.value(),
            augmentation_per_kwh=self.aug_cost_spin.value(),
            decommissioning_per_kw=self.decom_spin.value(),
        )

        # Build benefit streams from table
        benefits = []
        n = basics.analysis_period_years
        for row in range(self.benefits_table.rowCount()):
            name = self.benefits_table.item(row, 0).text() if self.benefits_table.item(row, 0) else ""
            try:
                year1_str = self.benefits_table.item(row, 1).text().replace(",", "")
                year1_val = float(year1_str)
            except (ValueError, AttributeError):
                year1_val = 0.0
            try:
                esc_str = self.benefits_table.item(row, 2).text().replace(",", "")
                esc = float(esc_str) / 100
            except (ValueError, AttributeError):
                esc = 0.0
            citation = self.benefits_table.item(row, 3).text() if self.benefits_table.item(row, 3) else ""

            annual_values = [year1_val * (1 + esc) ** t for t in range(n)]
            benefits.append(BenefitStream(
                name=name,
                annual_values=annual_values,
                citation=citation,
            ))

        lib_name = self.library_combo.currentText()
        if lib_name == "-- Select Library --":
            lib_name = ""

        return Project(
            basics=basics,
            technology=technology,
            costs=costs,
            benefits=benefits,
            assumption_library=lib_name,
        )
