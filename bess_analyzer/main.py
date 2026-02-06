"""BESS Analyzer - Battery Energy Storage Economic Analysis Application.

A professional desktop application for calculating the economic viability
of battery energy storage projects. Supports NPV, BCR, IRR, and LCOS
calculations with industry-standard assumption libraries.
"""

import sys
from PyQt6.QtWidgets import QApplication
from src.gui.main_window import MainWindow


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("BESS Analyzer")
    app.setOrganizationName("RESS")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
