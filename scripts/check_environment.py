"""
Quick environment check to verify packages are accessible and Tesseract is properly configured
"""

import sys
import os
from pathlib import Path

print("="*60)
print("ENVIRONMENT CHECK")
print("="*60)

print(f"\n1. Python Information:")
print(f"   Version: {sys.version}")
print(f"   Executable: {sys.executable}")
print(f"   Path: {sys.path[0]}")

print(f"\n2. Checking installed packages:")

packages = [
    ('pandas', None),
    ('openpyxl', None),
    ('pdfplumber', None),
    ('fitz', 'PyMuPDF'),
    ('pytesseract', None),
    ('pdf2image', None),
    ('PIL', 'Pillow')
]

all_installed = True
for package, display_name in packages:
    try:
        mod = __import__(package)
        version = getattr(mod, '__version__', 'unknown')
        name = display_name or package
        print(f"   ✓ {name:15s} version: {version}")
    except ImportError:
        name = display_name or package
        print(f"   ✗ {name:15s} NOT INSTALLED")
        all_installed = False

print(f"\n3. Conda/Environment Info:")
conda_env = os.environ.get('CONDA_DEFAULT_ENV', 'not using conda')
conda_prefix = os.environ.get('CONDA_PREFIX', 'N/A')
print(f"   Conda environment: {conda_env}")
print(f"   Conda prefix: {conda_prefix}")

print(f"\n4. Tesseract Configuration:")
tessdata_prefix = os.environ.get('TESSDATA_PREFIX', 'NOT SET')
print(f"   TESSDATA_PREFIX: {tessdata_prefix}")

# Check if tessdata directory exists
if conda_prefix and conda_prefix != 'N/A':
    expected_tessdata = Path(conda_prefix) / "Library" / "share" / "tessdata"
    if expected_tessdata.exists():
        print(f"   ✓ Tessdata directory exists: {expected_tessdata}")
        # Check for eng.traineddata
        eng_file = expected_tessdata / "eng.traineddata"
        if eng_file.exists():
            size_mb = eng_file.stat().st_size / (1024*1024)
            print(f"   ✓ English language data found ({size_mb:.1f} MB)")
        else:
            print(f"   ✗ English language data (eng.traineddata) NOT FOUND")
            print(f"   → Run: python setup_tessdata.py")
    else:
        print(f"   ✗ Tessdata directory missing: {expected_tessdata}")
        print(f"   → Run: python setup_tessdata.py")

# Try running tesseract
print(f"\n5. Testing Tesseract:")
try:
    import subprocess
    result = subprocess.run(['tesseract', '--version'], capture_output=True, text=True, timeout=5)
    if result.returncode == 0:
        version_line = result.stdout.split('\n')[0]
        print(f"   ✓ Tesseract installed: {version_line}")
    else:
        print(f"   ✗ Tesseract command failed")
except FileNotFoundError:
    print(f"   ✗ Tesseract NOT FOUND in PATH")
except Exception as e:
    print(f"   ✗ Error checking tesseract: {e}")

print("\n" + "="*60)
if all_installed:
    print("✓✓✓ ALL PACKAGES INSTALLED - Ready to go!")
else:
    print("✗✗✗ SOME PACKAGES MISSING")
    print("\nTo install in current environment:")
    print("  conda install -c conda-forge pandas openpyxl pdfplumber pymupdf pytesseract pdf2image pillow poppler tesseract")
    print("\nTo setup Tesseract language data:")
    print("  python setup_tessdata.py")
print("="*60)
