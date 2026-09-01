"""Benchmark pdf-inspector frontend against the legacy pdfplumber path.

Runs ``PaperPipeline.characterize_pdf`` twice on each PDF (once with
``DOC2MD_USE_PDF_INSPECTOR=1``, once without) and prints a comparison of
extraction time, page count, total text length, and per-page OCR usage.
No AI calls are made; this benchmarks only the text-extraction stage.

Usage:
    python scripts/benchmark_frontend.py PDF [PDF ...]
    python scripts/benchmark_frontend.py --dir /path/to/pdfs
"""
from __future__ import annotations

import argparse
import os
import sys
import time
from pathlib import Path

# make repo root importable when invoked as `python scripts/benchmark_frontend.py`
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from paper_pipeline import PaperPipeline  # noqa: E402


def _run(pdf_path: Path, use_inspector: bool) -> dict:
    prev = os.environ.get("DOC2MD_USE_PDF_INSPECTOR")
    os.environ["DOC2MD_USE_PDF_INSPECTOR"] = "1" if use_inspector else "0"
    try:
        pipeline = PaperPipeline.from_file(pdf_path, max_vision_pages=0, budget=0)
        t0 = time.perf_counter()
        page_data, method = pipeline.characterize_pdf(pdf_path)
        elapsed = time.perf_counter() - t0
        total_chars = sum(p["char_count"] for p in page_data)
        low = sum(1 for p in page_data
                  if p["char_count"] < pipeline.MIN_TEXT_FOR_TEXT_PAGE)
        return {
            "method": method,
            "elapsed_s": elapsed,
            "pages": len(page_data),
            "total_chars": total_chars,
            "low_text_pages": low,
        }
    finally:
        if prev is None:
            os.environ.pop("DOC2MD_USE_PDF_INSPECTOR", None)
        else:
            os.environ["DOC2MD_USE_PDF_INSPECTOR"] = prev


def _fmt_row(label: str, r: dict) -> str:
    return (f"  {label:14s} method={r['method']:22s} "
            f"time={r['elapsed_s']:6.2f}s  pages={r['pages']:3d}  "
            f"chars={r['total_chars']:7d}  low={r['low_text_pages']}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pdfs", nargs="*", help="PDF files to benchmark")
    parser.add_argument("--dir", help="Directory to scan for PDFs (non-recursive)")
    args = parser.parse_args()

    paths: list[Path] = [Path(p) for p in args.pdfs]
    if args.dir:
        paths.extend(sorted(Path(args.dir).glob("*.pdf")))

    paths = [p for p in paths if p.exists() and p.suffix.lower() == ".pdf"]
    if not paths:
        print("No PDFs provided.", file=sys.stderr)
        return 2

    speedups: list[float] = []
    for p in paths:
        print(f"\n{p.name}")
        legacy = _run(p, use_inspector=False)
        inspector = _run(p, use_inspector=True)
        print(_fmt_row("pdfplumber", legacy))
        print(_fmt_row("pdf_inspector", inspector))
        if inspector["elapsed_s"] > 0:
            speedup = legacy["elapsed_s"] / inspector["elapsed_s"]
            speedups.append(speedup)
            char_delta = inspector["total_chars"] - legacy["total_chars"]
            print(f"  speedup={speedup:.2f}x  char_delta={char_delta:+d}")

    if speedups:
        avg = sum(speedups) / len(speedups)
        print(f"\navg speedup over {len(speedups)} files: {avg:.2f}x")
    return 0


if __name__ == "__main__":
    sys.exit(main())
