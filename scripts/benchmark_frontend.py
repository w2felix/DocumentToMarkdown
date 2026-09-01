"""Benchmark pdf-inspector frontend against the legacy pdfplumber path.

Runs ``PaperPipeline.characterize_pdf`` twice on each PDF (once with
``DOC2MD_USE_PDF_INSPECTOR=1``, once without) and prints a comparison of
extraction time, page count, total text length, and per-page OCR usage.
No AI calls are made; this benchmarks only the text-extraction stage.

Each run executes in its own subprocess so a hanging legacy extraction
does not stall the sweep. Runs are wall-clock-bounded by ``--timeout``.

Usage:
    python scripts/benchmark_frontend.py PDF [PDF ...]
    python scripts/benchmark_frontend.py --dir /path/to/pdfs
    python scripts/benchmark_frontend.py --list-file bench_pdfs.txt --timeout 120
"""
from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))


def _worker(pdf_path: str, use_inspector: str) -> None:
    """Subprocess entrypoint. Emits a single JSON line on stdout."""
    os.environ["DOC2MD_USE_PDF_INSPECTOR"] = use_inspector
    from paper_pipeline import PaperPipeline  # noqa: E402

    pipeline = PaperPipeline.from_file(Path(pdf_path), max_vision_pages=0, budget=0)
    t0 = time.perf_counter()
    page_data, method = pipeline.characterize_pdf(Path(pdf_path))
    elapsed = time.perf_counter() - t0
    total_chars = sum(p["char_count"] for p in page_data)
    low = sum(1 for p in page_data
              if p["char_count"] < pipeline.MIN_TEXT_FOR_TEXT_PAGE)
    print(json.dumps({
        "method": method,
        "elapsed_s": elapsed,
        "pages": len(page_data),
        "total_chars": total_chars,
        "low_text_pages": low,
    }))


def _run_isolated(pdf_path: Path, use_inspector: bool, timeout: float) -> dict:
    cmd = [sys.executable, __file__, "--worker",
           "--pdf", str(pdf_path),
           "--use-inspector", "1" if use_inspector else "0"]
    try:
        proc = subprocess.run(
            cmd, capture_output=True, text=True, timeout=timeout,
            cwd=str(REPO_ROOT),
        )
    except subprocess.TimeoutExpired:
        return {"error": f"timeout>{timeout:.0f}s"}

    if proc.returncode != 0:
        return {"error": (proc.stderr or "").strip()[:200] or "nonzero exit"}
    for line in reversed(proc.stdout.strip().splitlines()):
        line = line.strip()
        if line.startswith("{"):
            try:
                return json.loads(line)
            except json.JSONDecodeError:
                continue
    return {"error": "no result on stdout"}


def _fmt(r: dict) -> str:
    if "error" in r:
        return f"ERROR {r['error']}"
    return (f"method={r['method']:22s} time={r['elapsed_s']:6.2f}s  "
            f"pages={r['pages']:3d}  chars={r['total_chars']:7d}  "
            f"low={r['low_text_pages']}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pdfs", nargs="*")
    parser.add_argument("--dir")
    parser.add_argument("--list-file",
                        help="Text file with one PDF path per line")
    parser.add_argument("--timeout", type=float, default=120.0,
                        help="Per-run wall-clock timeout in seconds (default 120)")
    # worker-mode plumbing
    parser.add_argument("--worker", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--pdf", help=argparse.SUPPRESS)
    parser.add_argument("--use-inspector", help=argparse.SUPPRESS)
    args = parser.parse_args()

    if args.worker:
        _worker(args.pdf, args.use_inspector)
        return 0

    paths: list[Path] = [Path(p) for p in args.pdfs]
    if args.dir:
        paths.extend(sorted(Path(args.dir).glob("*.pdf")))
    if args.list_file:
        for line in Path(args.list_file).read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#"):
                paths.append(Path(line))
    paths = [p for p in paths if p.exists() and p.suffix.lower() == ".pdf"]
    if not paths:
        print("No PDFs provided.", file=sys.stderr)
        return 2

    speedups: list[float] = []
    char_deltas: list[int] = []
    inspector_timeouts = legacy_timeouts = 0
    rows: list[tuple[str, dict, dict]] = []

    for p in paths:
        print(f"\n{p.name}")
        legacy = _run_isolated(p, use_inspector=False, timeout=args.timeout)
        inspector = _run_isolated(p, use_inspector=True, timeout=args.timeout)
        print(f"  pdfplumber      {_fmt(legacy)}")
        print(f"  pdf_inspector   {_fmt(inspector)}")
        rows.append((p.name, legacy, inspector))

        legacy_ok = "error" not in legacy
        inspector_ok = "error" not in inspector
        if "error" in legacy and "timeout" in legacy["error"]:
            legacy_timeouts += 1
        if "error" in inspector and "timeout" in inspector["error"]:
            inspector_timeouts += 1

        if legacy_ok and inspector_ok and inspector["elapsed_s"] > 0:
            speedup = legacy["elapsed_s"] / inspector["elapsed_s"]
            char_delta = inspector["total_chars"] - legacy["total_chars"]
            speedups.append(speedup)
            char_deltas.append(char_delta)
            print(f"  speedup={speedup:.2f}x  char_delta={char_delta:+d}")

    print("\n" + "=" * 70)
    print(f"files={len(paths)}  "
          f"legacy_timeouts={legacy_timeouts}  "
          f"inspector_timeouts={inspector_timeouts}")
    if speedups:
        avg_sp = sum(speedups) / len(speedups)
        median_sp = sorted(speedups)[len(speedups) // 2]
        avg_cd = sum(char_deltas) / len(char_deltas)
        print(f"avg speedup={avg_sp:.2f}x  median={median_sp:.2f}x  "
              f"avg char delta={avg_cd:+.0f}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
