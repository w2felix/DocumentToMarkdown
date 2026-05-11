# DocumentToMarkdown

Convert scientific PDFs into structured, searchable markdown using AI-powered vision analysis.

Built for pharmaceutical research — handles conference posters, patent filings, and presentation slides out of the box.

## What It Does

| Pipeline | Input | Output |
|----------|-------|--------|
| **Poster** | Conference poster PDFs | Structured markdown with figures, sections, and metadata |
| **Patent** | Patent filing PDFs (WO/EP/US) | Claims, chemical structures, SMILES, executive summaries |
| **Talk** | Slide-based presentation PDFs | Slide-by-slide extraction with narrative summaries |

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
python poster_pipeline.py --sharepoint "path/to/poster_pdfs" --metadata "abstracts.xlsx"
```

**Patents:**
```bash
python patent_pipeline.py --input "path/to/patent_pdfs"
# or a single file:
python patent_pipeline.py --single "path/to/WO2024123456.pdf"
```

**Talks:**
```bash
python talk_pipeline.py --talks "path/to/talk_pdfs" --metadata "abstracts.xlsx"
# or a single file:
python talk_pipeline.py --single "path/to/talk.pdf"
```

## Pipeline Options

### Poster Pipeline

| Flag | Description |
|------|-------------|
| `--sharepoint` | Folder containing poster PDFs (required) |
| `--metadata` | Excel file with abstract metadata (required) |
| `--output` | Output directory (default: `output`) |
| `--single` | Process a single PDF file |
| `--recursive` | Search subfolders for PDFs |
| `--no-skip` | Reprocess already-converted files |
| `--force-ocr` | Force OCR even when native text works |
| `--no-detailed-analysis` | Skip two-stage figure analysis (faster) |

### Patent Pipeline

| Flag | Description |
|------|-------------|
| `--input` | Folder containing patent PDFs |
| `--single` | Process a single patent PDF |
| `--output` | Output directory (default: `output_patents`) |
| `--no-vision` | Text-only extraction (skip AI analysis) |
| `--claims-only` | Extract only the claims section |
| `--max-figure-pages` | Limit figure pages analyzed |
| `--skip-existing` | Skip already-processed patents |

### Talk Pipeline

| Flag | Description |
|------|-------------|
| `--talks` | Folder containing talk PDFs |
| `--single` | Process a single talk PDF |
| `--metadata` | Excel file with abstract metadata |
| `--output` | Output directory (default: `output_talks`) |
| `--no-skip` | Reprocess already-converted files |

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
└── figures/
    ├── poster_1234_fig1.jpg
    └── ...

output_patents/
├── patent_WO2024123456A1.md
└── quality_log.txt

output_talks/
├── talk_04_ED03_Bunne_toward_virtual_patients.md
└── processing_log.txt
```

## How It Works

1. **PDF Rendering** — pages are rendered to images using PyMuPDF
2. **Text Extraction** — native text extraction via pdfplumber, with OCR fallback (Tesseract)
3. **Vision AI** — Claude analyzes page images for figures, chemical structures, and garbled text
4. **Structuring** — extracted content is parsed into logical sections
5. **Summarization** — AI generates executive summaries and key findings
6. **Quality Scoring** — automated scoring flags documents that may need manual review

## Documentation

| Guide | Description |
|-------|-------------|
| [Setup Guide](setup.md) | Full installation from scratch — Miniconda, packages, credentials, troubleshooting |
| [Poster Pipeline](README_Posters.md) | Deep dive — architecture, processing stages, figure analysis, quality scoring |
| [Patent Pipeline](README_Patents.md) | Deep dive — claims parsing, SMILES extraction, text quality repair, semantic classification |
| [Talk Pipeline](README_Talks.md) | Deep dive — slide extraction, OCR pre-pass, summary generation, batch processing |

## Requirements

- Python 3.11+ (via Conda)
- Tesseract OCR
- Anthropic API access (Claude Sonnet 4.6)
- Windows (uses Windows Registry for credential loading; adaptable to other OS)
