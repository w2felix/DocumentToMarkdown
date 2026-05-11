"""
Simple diagnostic script to check what's happening with text extraction
"""

print("="*60)
print("DIAGNOSTIC: Checking poster 160.pdf text extraction")
print("="*60)

# Check imports
print("\n1. Checking required packages...")
try:
    import pdfplumber
    print(f"   ✓ pdfplumber version: {pdfplumber.__version__}")
except ImportError as e:
    print(f"   ✗ pdfplumber not installed: {e}")

try:
    import fitz
    print(f"   ✓ PyMuPDF (fitz) version: {fitz.__version__}")
except ImportError as e:
    print(f"   ✗ PyMuPDF not installed: {e}")

try:
    import pandas as pd
    print(f"   ✓ pandas version: {pd.__version__}")
except ImportError as e:
    print(f"   ✗ pandas not installed: {e}")

# Try to extract text
print("\n2. Attempting text extraction from 160.pdf...")

# Get project root (parent of scripts folder)
try:
    import pdfplumber
    from pathlib import Path

    project_root = Path(__file__).parent.parent
    PDF_PATH = project_root / "test_poster" / "160.pdf"

    pdf_path = Path(PDF_PATH)
    if not pdf_path.exists():
        print(f"   ✗ PDF not found at: {pdf_path}")
    else:
        print(f"   ✓ PDF found: {pdf_path.name} ({pdf_path.stat().st_size / (1024*1024):.1f} MB)")

        with pdfplumber.open(pdf_path) as pdf:
            num_pages = len(pdf.pages)
            print(f"   ✓ PDF has {num_pages} page(s)")

            text = ""
            for i, page in enumerate(pdf.pages):
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n\n"
                    print(f"   ✓ Page {i+1}: {len(page_text)} characters")
                else:
                    print(f"   ✗ Page {i+1}: No text extracted")

            print(f"\n3. Text extraction summary:")
            print(f"   Total characters: {len(text)}")
            print(f"   Total words: {len(text.split())}")

            # Quality metrics
            alpha_chars = sum(c.isalpha() for c in text)
            digit_chars = sum(c.isdigit() for c in text)
            space_chars = sum(c.isspace() for c in text)
            other_chars = len(text) - alpha_chars - digit_chars - space_chars

            print(f"   Alpha characters: {alpha_chars} ({alpha_chars/len(text)*100:.1f}%)")
            print(f"   Digit characters: {digit_chars} ({digit_chars/len(text)*100:.1f}%)")
            print(f"   Space characters: {space_chars} ({space_chars/len(text)*100:.1f}%)")
            print(f"   Other characters: {other_chars} ({other_chars/len(text)*100:.1f}%)")

            # Quality check
            print(f"\n4. Quality assessment:")
            alpha_ratio = alpha_chars / len(text) if len(text) > 0 else 0
            word_count = len(text.split())

            checks = {
                "Has content (>100 chars)": len(text) > 100,
                "Sufficient words (>20)": word_count > 20,
                "Good alpha ratio (>30%)": alpha_ratio > 0.3
            }

            all_pass = True
            for check, passed in checks.items():
                status = "✓" if passed else "✗"
                print(f"   {status} {check}")
                if not passed:
                    all_pass = False

            if all_pass:
                print(f"\n   ✓✓✓ TEXT QUALITY: GOOD - Should work!")
            else:
                print(f"\n   ✗✗✗ TEXT QUALITY: POOR - May need OCR")

            # Show sample
            print(f"\n5. Text sample (first 400 characters):")
            print("   " + "-"*56)
            sample = text[:400].replace('\n', '\n   ')
            print(f"   {sample}")
            print("   " + "-"*56)

except Exception as e:
    print(f"\n   ✗ Error: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "="*60)
print("DIAGNOSTIC COMPLETE")
print("="*60)
