# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec file for BESS Analyzer.

This spec file configures PyInstaller to create a standalone executable
for Windows (and other platforms). It bundles all dependencies including
PyQt6, matplotlib, numpy, and the assumption library JSON files.

Build command:
    pyinstaller bess_analyzer.spec

For Windows build on Windows:
    pyinstaller bess_analyzer.spec --clean

For single-file executable:
    pyinstaller bess_analyzer.spec --onefile
"""

import sys
from pathlib import Path

block_cipher = None

# Get the project root directory
project_root = Path(SPECPATH)

# Data files to include (assumption libraries, etc.)
datas = [
    (str(project_root / 'resources' / 'libraries'), 'resources/libraries'),
]

# Hidden imports that PyInstaller might miss
hiddenimports = [
    'numpy',
    'numpy_financial',
    'pandas',
    'matplotlib',
    'matplotlib.backends.backend_agg',
    'matplotlib.figure',
    'reportlab',
    'reportlab.lib',
    'reportlab.lib.colors',
    'reportlab.lib.pagesizes',
    'reportlab.lib.styles',
    'reportlab.lib.units',
    'reportlab.platypus',
    'reportlab.graphics',
    'openpyxl',
    'xlsxwriter',
    'PyQt6',
    'PyQt6.QtCore',
    'PyQt6.QtGui',
    'PyQt6.QtWidgets',
]

a = Analysis(
    ['main.py'],
    pathex=[str(project_root)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',
        'test',
        'tests',
        'pytest',
        'pytest_cov',
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
    console=False,  # Set to True for debugging
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # Add 'resources/icon.ico' if you have an icon file
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
