"""
Debug script to see what text is actually being extracted from 160.pdf
"""

import pdfplumber
from pathlib import Path

# Get project root (parent of scripts folder)
project_root = Path(__file__).parent.parent
PDF_PATH = project_root / "test_poster" / "160.pdf"

def debug_text_extraction():
    print("="*60)
    print("Debugging Text Extraction for 160.pdf")
    print("="*60)

    pdf_path = Path(PDF_PATH)

    if not pdf_path.exists():
        print(f"ERROR: File not found: {pdf_path}")
        return

    print(f"\nFile found: {pdf_path.name}")
    print(f"File size: {pdf_path.stat().st_size / (1024*1024):.2f} MB")

    try:
        with pdfplumber.open(pdf_path) as pdf:
            print(f"\nNumber of pages: {len(pdf.pages)}")

            # Extract text from all pages
            full_text = ""
            for i, page in enumerate(pdf.pages):
                page_text = page.extract_text()
                if page_text:
                    full_text += page_text + "\n\n"
                    print(f"\nPage {i+1}: {len(page_text)} characters")
                else:
                    print(f"\nPage {i+1}: NO TEXT EXTRACTED")

            print(f"\n{'='*60}")
            print(f"TOTAL TEXT LENGTH: {len(full_text)} characters")
            print(f"{'='*60}")

            # Quality checks
            words = full_text.split()
            print(f"\nWord count: {len(words)}")

            alpha_chars = sum(c.isalpha() for c in full_text)
            total_chars = len(full_text)
            alpha_ratio = alpha_chars / total_chars if total_chars > 0 else 0

            print(f"Alpha characters: {alpha_chars}/{total_chars} ({alpha_ratio:.2%})")
            print(f"Is printable: {full_text.isprintable()}")

            # Show first 500 characters
            print(f"\n{'='*60}")
            print("FIRST 500 CHARACTERS:")
            print(f"{'='*60}")
            print(full_text[:500])
            print("...")

            # Show last 300 characters
            print(f"\n{'='*60}")
            print("LAST 300 CHARACTERS:")
            print(f"{'='*60}")
            print("...")
            print(full_text[-300:])

            # Check quality
            print(f"\n{'='*60}")
            print("QUALITY ASSESSMENT:")
            print(f"{'='*60}")

            checks = {
                "Length > 100 chars": len(full_text.strip()) >= 100,
                "Word count > 50": len(words) >= 50,
                "Alpha ratio > 50%": alpha_ratio >= 0.5,
            }

            for check, passed in checks.items():
                status = "✓ PASS" if passed else "✗ FAIL"
                print(f"{status}: {check}")

            all_passed = all(checks.values())
            print(f"\n{'='*60}")
            if all_passed:
                print("✓ TEXT QUALITY: GOOD")
                print("Native extraction should work!")
            else:
                print("✗ TEXT QUALITY: POOR")
                print("Would fallback to OCR")
            print(f"{'='*60}")

    except Exception as e:
        print(f"\nERROR: {e}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    debug_text_extraction()
