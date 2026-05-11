"""
Automated setup script for Tesseract language data

This script downloads the English language data file for Tesseract OCR
and sets up the TESSDATA_PREFIX environment variable.
"""

import os
import sys
from pathlib import Path
import urllib.request


def setup_tessdata():
    """Download and setup Tesseract language data"""

    print("="*60)
    print("TESSERACT LANGUAGE DATA SETUP")
    print("="*60)

    # Determine tessdata directory
    conda_prefix = os.environ.get('CONDA_PREFIX')

    if conda_prefix:
        tessdata_dir = Path(conda_prefix) / "Library" / "share" / "tessdata"
        print(f"\n✓ Found Conda environment: {conda_prefix}")
    else:
        print("\n✗ CONDA_PREFIX not set - are you in a conda environment?")
        print("  Try: conda activate your_env_name")

        # Fallback to common locations
        possible_paths = [
            Path("C:/Program Files/Tesseract-OCR/tessdata"),
            Path("/usr/share/tesseract-ocr/5/tessdata"),
            Path("/usr/local/share/tessdata"),
        ]

        for path in possible_paths:
            if path.exists():
                tessdata_dir = path
                print(f"\n✓ Found Tesseract at: {tessdata_dir}")
                break
        else:
            print("\n✗ Could not find Tesseract installation directory")
            print("  Please install Tesseract first:")
            print("    conda install -c conda-forge tesseract")
            sys.exit(1)

    # Create directory if it doesn't exist
    tessdata_dir.mkdir(parents=True, exist_ok=True)
    print(f"✓ Tessdata directory: {tessdata_dir}")

    # Download English language data
    eng_file = tessdata_dir / "eng.traineddata"

    if eng_file.exists():
        size_mb = eng_file.stat().st_size / (1024*1024)
        print(f"\n✓ English language data already exists ({size_mb:.1f} MB)")
        print(f"  Location: {eng_file}")
    else:
        print(f"\n→ Downloading English language data...")
        url = "https://github.com/tesseract-ocr/tessdata/raw/main/eng.traineddata"

        try:
            urllib.request.urlretrieve(url, str(eng_file))
            size_mb = eng_file.stat().st_size / (1024*1024)
            print(f"✓ Downloaded successfully ({size_mb:.1f} MB)")
            print(f"  Location: {eng_file}")
        except Exception as e:
            print(f"✗ Download failed: {e}")
            print(f"\n  Try manual download:")
            print(f"    URL: {url}")
            print(f"    Save to: {eng_file}")
            sys.exit(1)

    # Set environment variable
    print(f"\n→ Setting TESSDATA_PREFIX environment variable...")
    os.environ['TESSDATA_PREFIX'] = str(tessdata_dir)
    print(f"✓ Set for current session: {tessdata_dir}")

    # Provide instructions for permanent setup
    print(f"\n→ To make this permanent:")
    if os.name == 'nt':  # Windows
        print(f"\n  Run in PowerShell:")
        print(f'  [Environment]::SetEnvironmentVariable("TESSDATA_PREFIX", "{tessdata_dir}", "User")')
        print(f'  $env:TESSDATA_PREFIX = "{tessdata_dir}"')
    else:  # Linux/Mac
        print(f"\n  Add to ~/.bashrc or ~/.zshrc:")
        print(f'  export TESSDATA_PREFIX="{tessdata_dir}"')

    # Test Tesseract
    print(f"\n→ Testing Tesseract...")
    try:
        import subprocess
        result = subprocess.run(
            ['tesseract', '--list-langs'],
            capture_output=True,
            text=True,
            timeout=5
        )

        if result.returncode == 0:
            langs = [line.strip() for line in result.stdout.split('\n') if line.strip() and line.strip() != 'List of available languages in "/usr/share/tesseract-ocr/5/tessdata/":' and not line.startswith('List of')]
            print(f"✓ Tesseract is working!")
            print(f"  Available languages: {', '.join(langs)}")
        else:
            print(f"✗ Tesseract test failed: {result.stderr}")
    except FileNotFoundError:
        print(f"✗ Tesseract not found in PATH")
        print(f"  Install with: conda install -c conda-forge tesseract")
    except Exception as e:
        print(f"✗ Error testing Tesseract: {e}")

    print("\n" + "="*60)
    print("✓✓✓ SETUP COMPLETE!")
    print("="*60)
    print("\nYou can now run:")
    print("  python test_single_poster.py")
    print("  python poster_pipeline.py ...")
    print("="*60)


if __name__ == '__main__':
    setup_tessdata()
