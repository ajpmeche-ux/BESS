#!/usr/bin/env python3
"""Build script for creating BESS Analyzer executable.

This script handles the PyInstaller build process for creating
standalone executables on Windows, macOS, or Linux.

Usage:
    python build_exe.py          # Build for current platform
    python build_exe.py --onefile   # Create single-file executable
    python build_exe.py --clean     # Clean build (remove previous build artifacts)
"""

import argparse
import os
import platform
import shutil
import subprocess
import sys
from pathlib import Path


def check_pyinstaller():
    """Check if PyInstaller is installed."""
    try:
        import PyInstaller
        print(f"PyInstaller version: {PyInstaller.__version__}")
        return True
    except ImportError:
        print("ERROR: PyInstaller is not installed.")
        print("Install it with: pip install pyinstaller")
        return False


def clean_build_dirs():
    """Remove previous build artifacts."""
    dirs_to_remove = ['build', 'dist', '__pycache__']
    for dir_name in dirs_to_remove:
        dir_path = Path(dir_name)
        if dir_path.exists():
            print(f"Removing {dir_path}...")
            shutil.rmtree(dir_path)

    # Remove .spec file cache
    for spec_cache in Path('.').glob('*.spec.bak'):
        spec_cache.unlink()


def build_executable(onefile=False, clean=False, debug=False):
    """Build the executable using PyInstaller."""
    if not check_pyinstaller():
        return False

    if clean:
        clean_build_dirs()

    # Base PyInstaller command
    cmd = ['pyinstaller']

    # Use spec file
    spec_file = Path('bess_analyzer.spec')

    if spec_file.exists() and not onefile:
        cmd.append(str(spec_file))
    else:
        # Build without spec file (simple mode)
        cmd.extend([
            '--name', 'BESS_Analyzer',
            '--windowed',  # No console window
            '--add-data', f'resources/libraries{os.pathsep}resources/libraries',
        ])

        if onefile:
            cmd.append('--onefile')

        # Hidden imports
        hidden_imports = [
            'numpy', 'numpy_financial', 'pandas', 'matplotlib',
            'matplotlib.backends.backend_agg', 'reportlab',
            'reportlab.platypus', 'openpyxl', 'xlsxwriter',
        ]
        for imp in hidden_imports:
            cmd.extend(['--hidden-import', imp])

        # Exclude unnecessary modules
        cmd.extend(['--exclude-module', 'tkinter'])
        cmd.extend(['--exclude-module', 'pytest'])

        cmd.append('main.py')

    if clean:
        cmd.append('--clean')

    if debug:
        cmd.append('--debug=all')

    print(f"Running: {' '.join(cmd)}")
    print("-" * 60)

    result = subprocess.run(cmd)

    if result.returncode == 0:
        print("-" * 60)
        print("Build successful!")

        # Show output location
        if onefile:
            if platform.system() == 'Windows':
                exe_path = Path('dist') / 'BESS_Analyzer.exe'
            else:
                exe_path = Path('dist') / 'BESS_Analyzer'
        else:
            exe_path = Path('dist') / 'BESS_Analyzer'

        print(f"Executable location: {exe_path.absolute()}")

        if platform.system() == 'Windows':
            print("\nTo run: dist\\BESS_Analyzer\\BESS_Analyzer.exe")
        else:
            print(f"\nNote: This builds for {platform.system()}.")
            print("To build a Windows .exe, run this script on a Windows machine.")

        return True
    else:
        print("-" * 60)
        print("Build failed!")
        return False


def main():
    parser = argparse.ArgumentParser(
        description='Build BESS Analyzer executable'
    )
    parser.add_argument(
        '--onefile', '-o',
        action='store_true',
        help='Create a single-file executable (slower startup but easier to distribute)'
    )
    parser.add_argument(
        '--clean', '-c',
        action='store_true',
        help='Clean build directories before building'
    )
    parser.add_argument(
        '--debug', '-d',
        action='store_true',
        help='Enable debug mode for troubleshooting'
    )

    args = parser.parse_args()

    print("=" * 60)
    print("BESS Analyzer - Build Executable")
    print("=" * 60)
    print(f"Platform: {platform.system()} {platform.machine()}")
    print(f"Python: {sys.version}")
    print("=" * 60)

    # Change to script directory
    os.chdir(Path(__file__).parent)

    success = build_executable(
        onefile=args.onefile,
        clean=args.clean,
        debug=args.debug
    )

    sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()
