"""Main application window for BESS Analyzer.

Provides the top-level window with menu bar, tabbed interface
(Inputs and Results), and action buttons for calculation and
report generation.
"""

from PyQt6.QtCore import QSize, Qt
from PyQt6.QtWidgets import (
    QFileDialog,
    QHBoxLayout,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QStatusBar,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from src.data.storage import load_project, save_project
from src.data.validators import validate_project
from src.gui.input_forms import InputFormWidget
from src.gui.results_display import ResultsWidget
from src.gui.sensitivity_widget import SensitivityWidget
from src.models.calculations import calculate_project_economics, calculate_uos_analysis
from src.models.project import Project
from src.reports.executive import generate_executive_summary


class MainWindow(QMainWindow):
    """Main application window with tabbed inputs and results."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("BESS Analyzer - Battery Energy Storage Economics")
        self.setMinimumSize(QSize(1000, 700))
        self._current_project = None
        self._current_results = None
        self._current_file = None
        self._init_ui()
        self._init_menu()
        self.statusBar().showMessage("Ready. Load an assumption library or enter project data.")

    def _init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        # Tab widget
        self.tabs = QTabWidget()
        self.input_form = InputFormWidget()
        self.results_widget = ResultsWidget()
        self.sensitivity_widget = SensitivityWidget()
        self.tabs.addTab(self.input_form, "Project Inputs")
        self.tabs.addTab(self.results_widget, "Results")
        self.tabs.addTab(self.sensitivity_widget, "Sensitivity Analysis")
        layout.addWidget(self.tabs)

        # Action buttons
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        self.calc_btn = QPushButton("Calculate Economics")
        self.calc_btn.setMinimumHeight(36)
        self.calc_btn.setStyleSheet("font-weight: bold; padding: 6px 20px;")
        self.calc_btn.clicked.connect(self._run_analysis)
        btn_layout.addWidget(self.calc_btn)

        self.report_btn = QPushButton("Generate Report")
        self.report_btn.setMinimumHeight(36)
        self.report_btn.setStyleSheet("padding: 6px 20px;")
        self.report_btn.setEnabled(False)
        self.report_btn.clicked.connect(self._generate_report)
        btn_layout.addWidget(self.report_btn)

        self.excel_btn = QPushButton("Export to Excel")
        self.excel_btn.setMinimumHeight(36)
        self.excel_btn.setStyleSheet("padding: 6px 20px;")
        self.excel_btn.setEnabled(False)
        self.excel_btn.clicked.connect(self._export_to_excel)
        btn_layout.addWidget(self.excel_btn)

        btn_layout.addStretch()
        layout.addLayout(btn_layout)

    def _init_menu(self):
        menubar = self.menuBar()

        # File menu
        file_menu = menubar.addMenu("File")
        file_menu.addAction("New Project", self._new_project)
        file_menu.addAction("Open...", self._open_project)
        file_menu.addAction("Save", self._save_project)
        file_menu.addAction("Save As...", self._save_project_as)
        file_menu.addSeparator()
        file_menu.addAction("Exit", self.close)

        # Analysis menu
        analysis_menu = menubar.addMenu("Analysis")
        analysis_menu.addAction("Calculate Economics", self._run_analysis)

        # Reports menu
        reports_menu = menubar.addMenu("Reports")
        reports_menu.addAction("Executive Summary PDF", self._generate_report)
        reports_menu.addAction("Export to Excel", self._export_to_excel)

        # Help menu
        help_menu = menubar.addMenu("Help")
        help_menu.addAction("About", self._show_about)

    # --- Actions ---

    def _new_project(self):
        self._current_file = None
        self._current_project = None
        self._current_results = None
        self.input_form.load_project(Project())
        self.report_btn.setEnabled(False)
        self.excel_btn.setEnabled(False)
        self.tabs.setCurrentIndex(0)
        self.statusBar().showMessage("New project created.")

    def _open_project(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Open Project", "", "JSON Files (*.json);;All Files (*)"
        )
        if not path:
            return
        try:
            project = load_project(path)
            self.input_form.load_project(project)
            self._current_file = path
            self._current_project = project
            self.statusBar().showMessage(f"Loaded: {path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load project:\n{e}")

    def _save_project(self):
        if self._current_file:
            self._do_save(self._current_file)
        else:
            self._save_project_as()

    def _save_project_as(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Project", "", "JSON Files (*.json);;All Files (*)"
        )
        if path:
            self._do_save(path)

    def _do_save(self, path: str):
        try:
            project = self.input_form.get_project()
            project.results = self._current_results
            save_project(project, path)
            self._current_file = path
            self.statusBar().showMessage(f"Saved: {path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save project:\n{e}")

    def _run_analysis(self):
        try:
            project = self.input_form.get_project()
        except Exception as e:
            QMessageBox.critical(self, "Input Error", f"Invalid inputs:\n{e}")
            return

        # Validate
        is_valid, messages = validate_project(project)
        if not is_valid:
            errors = "\n".join(m for m in messages if not m.startswith("Warning"))
            QMessageBox.critical(self, "Validation Error", f"Fix these errors:\n\n{errors}")
            return

        warnings = [m for m in messages if m.startswith("Warning")]
        if warnings:
            reply = QMessageBox.question(
                self, "Warnings",
                "The following warnings were found:\n\n" + "\n".join(warnings) +
                "\n\nContinue with calculation?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if reply != QMessageBox.StandardButton.Yes:
                return

        # Calculate
        try:
            results = calculate_project_economics(project)
            self._current_project = project
            self._current_results = results
            project.results = results

            # Run UOS analysis if enabled
            uos_results = None
            if project.uos_inputs and project.uos_inputs.enabled:
                uos_results = calculate_uos_analysis(project)

            self.results_widget.display_results(project, results)
            if uos_results:
                self.results_widget.display_uos_results(uos_results)
            else:
                self.results_widget.display_uos_results(None)
            self.sensitivity_widget.display_sensitivity(project, results)
            self.tabs.setCurrentIndex(1)
            self.report_btn.setEnabled(True)
            self.excel_btn.setEnabled(True)

            status_msg = f"Analysis complete. BCR: {results.bcr:.2f} | NPV: ${results.npv / 1e6:,.1f}M"
            if uos_results:
                rb = uos_results.get("rate_base_results")
                sod = uos_results.get("sod_result")
                if rb:
                    status_msg += f" | Levelized RR: ${rb.levelized_revenue_requirement / 1e6:,.1f}M/yr"
                if sod:
                    status_msg += f" | SOD: {'PASS' if sod.feasible else 'FAIL'}"
            self.statusBar().showMessage(status_msg)
        except Exception as e:
            QMessageBox.critical(self, "Calculation Error", f"Error during calculation:\n{e}")

    def _generate_report(self):
        if not self._current_project or not self._current_results:
            QMessageBox.warning(self, "No Results", "Run 'Calculate Economics' first.")
            return

        path, _ = QFileDialog.getSaveFileName(
            self, "Save Executive Summary", "", "PDF Files (*.pdf);;All Files (*)"
        )
        if not path:
            return

        try:
            generate_executive_summary(
                self._current_project, self._current_results, path
            )
            self.statusBar().showMessage(f"Report saved: {path}")
            QMessageBox.information(self, "Report Generated", f"Executive summary saved to:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Report Error", f"Failed to generate report:\n{e}")

    def _export_to_excel(self):
        if not self._current_project or not self._current_results:
            QMessageBox.warning(self, "No Results", "Run 'Calculate Economics' first.")
            return

        path, _ = QFileDialog.getSaveFileName(
            self, "Export to Excel", "BESS_Analysis.xlsx",
            "Excel Files (*.xlsx);;All Files (*)"
        )
        if not path:
            return

        try:
            from excel_generator import create_workbook
            create_workbook(path)
            self.statusBar().showMessage(f"Excel exported: {path}")
            QMessageBox.information(
                self, "Export Complete",
                f"Excel workbook saved to:\n{path}\n\n"
                "Note: Open the VBA_Code sheet for macro instructions."
            )
        except Exception as e:
            QMessageBox.critical(self, "Export Error", f"Failed to export Excel:\n{e}")

    def _show_about(self):
        QMessageBox.about(
            self,
            "About BESS Analyzer",
            "<b>BESS Analyzer v1.0</b><br><br>"
            "Battery Energy Storage Economic Analysis Application<br><br>"
            "Calculates NPV, BCR, IRR, and LCOS for utility-scale<br>"
            "battery storage projects using industry-standard assumptions.<br><br>"
            "Supports NREL ATB, Lazard LCOS, and CPUC assumption libraries.<br><br>"
            "<b>Features:</b><br>"
            "- Financing structure with WACC calculation<br>"
            "- Infrastructure costs (interconnection, land, permitting)<br>"
            "- Investment Tax Credit (ITC) with adders<br>"
            "- Sensitivity analysis tables<br>"
            "- PDF reports and Excel export<br><br>"
            "<i>Built for utility planners and regulatory analysts.</i>",
        )
