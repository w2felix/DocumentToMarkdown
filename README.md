<p align="center">
  <img src="logo.svg" alt="Doc2MD" width="150">
</p>

# DocumentToMarkdown

Convert scientific documents into structured, searchable markdown using AI-powered vision analysis.

Built for pharmaceutical research — handles conference posters, patent filings, presentation slides, and corporate presentations (PPTX + PDF) out of the box.

## What It Does

| Pipeline | Input | Output |
|----------|-------|--------|
| **Poster** | Conference poster PDFs | Structured markdown with figures, sections, and metadata |
| **Patent** | Patent filing PDFs (WO/EP/US) | Claims, chemical structures, SMILES, executive summaries |
| **Talk** | Slide-based presentation PDFs | Slide-by-slide extraction with narrative summaries |
| **Presentation** | PPTX + PDF presentations | Native text extraction, per-slide image descriptions, action items, chemical structures |

Each pipeline produces a self-contained `.md` file with YAML frontmatter, making the output easy to search, filter, and integrate into knowledge bases.

## Quick Start

> **First time?** See the [full setup guide](setup.md) for step-by-step installation from scratch (Miniconda, packages, credentials, troubleshooting).

### 1. Setup Environment

```bash
conda create -n ds_env python=3.11 -y
conda activate ds_env

conda install -c conda-forge pdfplumber pymupdf pillow -y
conda install pandas openpyxl -y
conda install -c conda-forge tesseract pytesseract -y
pip install anthropic
```

### 2. Set Credentials

```powershell
[Environment]::SetEnvironmentVariable("ANTHROPIC_AUTH_TOKEN", "your-token", "User")
[Environment]::SetEnvironmentVariable("ANTHROPIC_BASE_URL", "your-proxy-url", "User")
```

Restart your terminal after setting these.

### 3. Run a Pipeline

**Posters:**
```bash
# with metadata
python poster_pipeline.py --input "path/to/poster_pdfs" --metadata "abstracts.xlsx"
# without metadata
python poster_pipeline.py --input "path/to/poster_pdfs"
# standardized filenames
python poster_pipeline.py --input "path/to/poster_pdfs" --naming standardized --conference AACR --year 2026
```

**Patents:**
```bash
python patent_pipeline.py --input "path/to/patent_pdfs"
# or a single file:
python patent_pipeline.py --single "path/to/WO2024123456.pdf"
# recursive scan with detailed naming:
python patent_pipeline.py --input "path/to/patent_pdfs" --recursive --naming detailed
```

**Talks:**
```bash
python talk_pipeline.py --input "path/to/talk_pdfs" --metadata "abstracts.xlsx"
# or a single file:
python talk_pipeline.py --single "path/to/talk.pdf"
```

**Presentations:**
```bash
python presentation_pipeline.py --input "path/to/presentations"
# text-only (no API calls):
python presentation_pipeline.py --input "path/to/presentations" --no-vision
# single file with dated naming:
python presentation_pipeline.py --single "path/to/file.pptx" --naming dated
```

## Common Flags

All pipelines share a consistent CLI interface:

| Flag | Description |
|------|-------------|
| `--input` | Input folder containing documents (aliases: `--sharepoint`, `--talks`) |
| `--output` | Output directory for markdown files |
| `--single` | Process a single file instead of a folder |
| `--recursive` | Recursively search subfolders for files |
| `--no-skip` | Reprocess files that already exist (default: skip existing) |
| `--naming` | Output filename scheme (options vary per pipeline) |
| `--verbose` | Enable debug logging |

Each pipeline also has specialized flags — see the pipeline-specific documentation for full options:
- [Poster Pipeline flags](README_Posters.md#command-line-interface) — `--metadata`, `--force-ocr`, `--conference`, `--col-*`
- [Patent Pipeline flags](README_Patents.md#command-line-interface) — `--no-vision`, `--claims-only`, `--budget`, `--ocr-engine`
- [Talk Pipeline flags](README_Talks.md#command-line-interface) — `--metadata`
- [Presentation Pipeline flags](README_Presentations.md#command-line-interface) — `--no-vision`

## Output Structure

Each pipeline generates markdown files with:

- **YAML frontmatter** — structured metadata (dates, authors, scores, classifications)
- **Executive summary** — AI-generated overview of the document
- **Full content** — sections, claims, or slides extracted from the PDF
- **Figures & visuals** — descriptions of charts, structures, and diagrams
- **Quality score** — automated 0–10 quality assessment

Example output location:
```
output/
├── poster_1234.md
├── poster_1235.md
└── quality_log.txt

output_patents/
├── patent_WO2024123456A1.md
└── quality_log.txt

output_talks/
├── talk_04_ED03_Bunne_toward_virtual_patients.md
└── processing_log.txt

output_presentations/
├── presentation_ru_onc_operations_update_darmstadt.md
└── presentation_caris_discovery_non_con_apr.md
```

## Pipeline Comparison

### At a Glance

| | Poster | Patent | Talk | Presentation |
|---|--------|--------|------|--------------|
| **Input** | Single-page poster PDF | Multi-page patent PDF (50–300+ pages) | Multi-slide presentation PDF (screenshots) | PPTX or PDF slide decks |
| **Output** | Sections (Methods, Results, Conclusions) | Patent sections + claims + chemical data | Slide-by-slide content + narrative summary | Slide content + action items + metrics |
| **Text extraction** | Native PDF + OCR + Vision AI | Native PDF + Tesseract OCR (scanned) + selective Vision AI for garbled pages | OCR + Vision AI only (no extractable text) | Native PPTX (python-pptx) or PyMuPDF + pdfplumber tables + Vision AI fallback |
| **Metadata source** | Excel spreadsheet (optional) | Extracted from the PDF itself | Excel spreadsheet (optional) | Extracted from the file itself |
| **Quality gate** | Yes — skips FAIR/POOR | No — all patents saved | No — all talks saved | No — all presentations saved |

### Performance & Cost

| | Poster | Patent | Talk | Presentation |
|---|--------|--------|------|--------------|
| **Processing time** | 1.5–4 min | 2–3 min (text+vision), 3–5 min (scanned+Tesseract), ~15s text-only | 1.5–4 min | <1s text-only, 30–90s with vision, 2–4 min image-heavy |
| **API calls per doc** | 4 + N figures | 10–48 (text-native), 8–35 (scanned+Tesseract) | 4–9 (scales with slide count) | 1–3 (smart gating) + 5–15 (image enrichment), 0 text-only |
| **Token usage per doc** | ~30K–60K | ~40K–120K | ~18K–35K | ~8K–20K (text-only), ~50K–120K (image-heavy) |
| **Vision AI pages** | All pages (mandatory) | ~10% of pages (selective) | All slides (mandatory) | Smart gating: global analysis + per-slide image enrichment (requires PowerPoint for PPTX) |
| **Text-only mode** | No | Yes (`--no-vision`) | No | Yes (`--no-vision`) |
| **Claims-only mode** | No | Yes (`--claims-only`, ~5s) | No | No |
| **Concurrency** | Up to 5 figure workers | Up to 3 batch workers | Up to 5 batch workers | Sequential |

### Unique Capabilities

| Capability | Poster | Patent | Talk | Presentation |
|------------|:------:|:------:|:----:|:------------:|
| Two-stage figure analysis | x | | | |
| Chemical structure extraction (SMILES) | | x | | x |
| Claims dependency tree | | x | | |
| Semantic classification (target, mechanism, modality) | | x | | |
| Hybrid OCR (Tesseract + Vision AI) | | x | | |
| Text quality scoring & auto-repair | | x | | |
| OCR pre-pass as RAG context | x | | x | |
| Abstract matching from metadata | x | | x | |
| Native PPTX text extraction | | | | x |
| Per-slide image enrichment (auto-describes figures/screenshots) | | | | x |
| Smart Vision AI gating (skip when not needed) | | | | x |
| Language detection (EN/DE) + English output | | | | x |
| Classification detection (3-signal) | | | | x |
| Action items & metrics extraction | | | | x |
| Conditional summary (content-type aware) | | | | x |
| Standardized naming schemes | x | x | | x |
| Executive summary | x | x | x | x |
| Quality scoring | x | x | x | x |

### When to Use Which

- **Poster** — single-page conference posters with figures, methods, and results sections
- **Patent** — multi-page patent filings (WIPO, EPO, USPTO) with claims, chemical structures, and experimental data
- **Talk** — slide-based presentations captured as PDF screenshots (no extractable text)
- **Presentation** — PPTX or PDF slide decks with native text (corporate meetings, scientific presentations, agendas); supports English and German content

## How It Works

1. **Text Extraction** — native text via PyMuPDF/pdfplumber/python-pptx, with OCR fallback (Tesseract) for posters/talks/scanned patents
2. **Page Rendering** — pages rendered to images using PyMuPDF (PDFs) or PowerPoint COM (PPTX) when Vision AI is needed
3. **Vision AI** — Claude analyzes page images for figures, chemical structures, and garbled text (smart gating skips this when not needed)
4. **Structuring** — extracted content is parsed into logical sections
5. **Summarization** — AI generates executive summaries and key findings (always in English)
6. **Quality Scoring** — automated scoring flags documents that may need manual review

## Documentation

| Guide | Description |
|-------|-------------|
| [Setup Guide](setup.md) | Full installation from scratch — Miniconda, packages, credentials, troubleshooting |
| [Poster Pipeline](README_Posters.md) | Deep dive — architecture, processing stages, figure analysis, quality scoring |
| [Patent Pipeline](README_Patents.md) | Deep dive — claims parsing, SMILES extraction, text quality repair, semantic classification |
| [Talk Pipeline](README_Talks.md) | Deep dive — slide extraction, OCR pre-pass, summary generation, batch processing |
| [Presentation Pipeline](README_Presentations.md) | Deep dive — PPTX/PDF extraction, classification detection, action items, naming schemes |

## Requirements

- Python 3.11+ (via Conda)
- Tesseract OCR (poster/talk/patent pipelines — required for poster/talk, optional for patent scanned PDFs)
- python-pptx, PyMuPDF, pdfplumber (presentation pipeline)
- Microsoft PowerPoint (optional — enables Vision AI for PPTX slide rendering)
- Anthropic API access (Claude Sonnet 4.6)
- Windows (uses Windows Registry for credential loading; adaptable to other OS)
