# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec file for BESS Analyzer.

Builds a standalone Windows executable with all dependencies bundled.
Run with: pyinstaller --noconfirm bess_analyzer.spec
"""

import os
import sys

block_cipher = None

# Get the project root directory
spec_dir = os.path.dirname(os.path.abspath(SPEC))

a = Analysis(
    ['main.py'],
    pathex=[spec_dir],
    binaries=[],
    datas=[
        # Include assumption library JSON files
        (os.path.join('resources', 'libraries', '*.json'),
         os.path.join('resources', 'libraries')),
        # Include the Excel template
        ('BESS_Analyzer.xlsx', '.'),
    ],
    hiddenimports=[
        'PyQt6',
        'PyQt6.QtCore',
        'PyQt6.QtGui',
        'PyQt6.QtWidgets',
        'numpy',
        'numpy_financial',
        'pandas',
        'matplotlib',
        'matplotlib.backends.backend_agg',
        'reportlab',
        'reportlab.lib',
        'reportlab.platypus',
        'xlsxwriter',
        'src',
        'src.models',
        'src.models.project',
        'src.models.calculations',
        'src.models.rate_base',
        'src.models.avoided_costs',
        'src.models.wires_comparison',
        'src.models.sod_check',
        'src.gui',
        'src.gui.main_window',
        'src.gui.input_forms',
        'src.gui.results_display',
        'src.gui.sensitivity_widget',
        'src.data',
        'src.data.libraries',
        'src.data.storage',
        'src.data.validators',
        'src.reports',
        'src.reports.executive',
        'src.reports.charts',
        'src.utils',
        'src.utils.formatters',
        'excel_generator',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',
        'unittest',
        'test',
        'pytest',
        'IPython',
        'jupyter',
        'notebook',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='BESS_Analyzer',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # No console window — GUI only
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # Add 'icon.ico' here if you have one
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='BESS_Analyzer',
)
