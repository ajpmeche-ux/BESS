#!/usr/bin/env python3
"""Helper script to create vbaProject.bin for macro-enabled Excel files.

Usage:
    python create_vba_template.py <path_to_xlsm>

To create vbaProject.bin:
1. Open Microsoft Excel
2. Create a new workbook
3. Press Alt+F11 to open VBA Editor
4. Insert > Module
5. Paste the VBA code from get_vba_code() in excel_generator.py
6. Save as 'template.xlsm' (Macro-Enabled Workbook)
7. Run: python create_vba_template.py template.xlsm
"""

import sys
import zipfile
from pathlib import Path


def extract_vba_bin(xlsm_path):
    """Extract vbaProject.bin from an .xlsm file."""
    resources_dir = Path('./resources')
    resources_dir.mkdir(exist_ok=True)

    with zipfile.ZipFile(xlsm_path, 'r') as zf:
        try:
            vba_data = zf.read('xl/vbaProject.bin')
            output_path = resources_dir / 'vbaProject.bin'
            with open(output_path, 'wb') as f:
                f.write(vba_data)
            print(f"Successfully extracted: {output_path}")
            return True
        except KeyError:
            print("Error: No vbaProject.bin found in the .xlsm file")
            print("Make sure the workbook contains VBA macros.")
            return False


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
    else:
        extract_vba_bin(sys.argv[1])
