# Scientific Talk Processing Pipeline

Automated extraction and analysis of scientific conference talk slides (PDF screenshots) using AI-powered vision processing. Converts multi-slide presentation PDFs into structured, searchable markdown with comprehensive slide content extraction, executive summaries, and metadata integration.

**Powered by Claude Vision AI (Sonnet 4.6)** with OCR pre-pass for high-accuracy text extraction from slide screenshots.

---

## Table of Contents

- [Quick Start](#quick-start)
- [Installation & Setup](#installation--setup)
- [Usage](#usage)
  - [Command-Line Interface](#command-line-interface)
  - [Python API](#python-api)
  - [Processing a Single Talk](#processing-a-single-talk)
  - [Processing All Talks](#processing-all-talks)
- [Pipeline Architecture](#pipeline-architecture)
  - [Processing Stages](#processing-stages)
  - [Key Features](#key-features)
- [Output Format](#output-format)
- [Excel Metadata Requirements](#excel-metadata-requirements)
- [Processing Log & Quality Assessment](#processing-log--quality-assessment)
- [Performance & Scalability](#performance--scalability)
- [Troubleshooting](#troubleshooting)
- [Known Limitations](#known-limitations)

---

## Quick Start

```bash
# 1. Activate environment
conda activate ds_env

# 2. Process all talks (uses default paths)
python talk_pipeline.py

# 3. Process a single talk
python talk_pipeline.py --single "path/to/06_ED03_M.Moor_Towards reliable medical AI.pdf"

# 4. Custom paths
python talk_pipeline.py --talks "C:\path\to\talks" --metadata "test_poster\AACR2026_Abstracts.xlsx" --output "output_talks"

# 5. Reprocess all (ignore previously generated files)
python talk_pipeline.py --no-skip
```

The pipeline will process all PDFs and generate markdown files in `output_talks/`.

---

## Installation & Setup

See [setup.md](setup.md) for complete installation instructions.

**Prerequisites:**
- Miniconda (Python 3.11)
- Required packages: `pymupdf`, `pandas`, `openpyxl`, `anthropic`, `pytesseract`, `pillow`
- API credentials: `ANTHROPIC_AUTH_TOKEN` and `ANTHROPIC_BASE_URL`

**Additional for OCR pre-pass:**
- Tesseract OCR installed and accessible
- `TESSDATA_PREFIX` configured (auto-detected in conda environments)

---

## Usage

### Command-Line Interface

```bash
python talk_pipeline.py [OPTIONS]
```

| Option | Default | Description |
|--------|---------|-------------|
| `--input` / `--talks` | *(required)* | Folder containing talk PDFs |
| `--metadata` | `test_poster\AACR2026_Abstracts.xlsx` | Excel file with abstract metadata |
| `--output` | `output_talks` | Output directory for markdown files |
| `--single` | *(none)* | Process a single PDF file only |
| `--recursive` | `False` | Recursively search subfolders for PDF files |
| `--no-skip` | `False` | Reprocess files that already exist |
| `--naming` | `default` | Filename scheme: `default`, `detailed`, or `dated` |
| `--verbose` | `False` | Enable debug logging |

### Python API

```python
from talk_pipeline import TalkPipeline

pipeline = TalkPipeline(
    talks_folder="path/to/talks",
    metadata_excel="path/to/AACR2026_Abstracts.xlsx",
    output_dir="output_talks",
    recursive=False,
    naming="default"
)
```

### Processing a Single Talk

```python
from pathlib import Path

pipeline.process_single_talk(
    Path("path/to/04_ED03_C.Bunne_Toward virtual patients.pdf"),
    skip_existing=False
)
```

### Processing All Talks

```python
# Process all, skip already-processed files
pipeline.process_all_talks(skip_existing=True)

# Reprocess everything
pipeline.process_all_talks(skip_existing=False)
```

---

## Pipeline Architecture

The pipeline handles multi-slide presentation PDFs where each page is a screenshot image (no extractable text). It uses a batch Vision AI approach with OCR support.

### Processing Stages

#### 1. Filename Parsing & Metadata Matching

- Parses structured filename: `NN_SESSION_Speaker_Title.pdf`
- Extracts talk number, session code, speaker name, title
- Matches to abstract in Excel by speaker last name + title keyword similarity
- 44 of 77 AACR 2026 talks have matching abstracts (minisymposia oral presentations)
- 33 talks are invited/plenary speakers without submitted abstracts — processed from slides only

#### 2. PDF-to-Image Conversion

- Renders each slide page at 150 DPI using PyMuPDF
- Produces landscape images (~1754x1241 px per slide)
- Handles talks with 13-40+ slides

#### 3. OCR Pre-Pass (RAG Context)

- Runs Tesseract OCR on all slide images before Vision AI
- Extracts text as reference context (typically 200-400 chars per slide)
- Passed alongside images to Vision AI for verification and correction
- OCR text may contain errors — Vision AI uses it as a guide, not ground truth

#### 4. Batch Vision AI Slide Extraction (Core)

- Pre-encodes all slide images to base64 JPEG (max 1568px dimension)
- Processes slides in batches of 5 per API call
- Batches run concurrently (up to 5 simultaneous API calls)
- Each batch includes: slide images + OCR reference text per slide
- Output per slide: text content, visual elements, key point summary
- Format: `## Slide N: [Title/Topic]` with structured subsections

#### 5. Executive Summary & Key Takeaways (Single API Call)

- Sends full extracted slide content + abstract (if available) to Vision AI
- Smart content truncation: if content exceeds 40K chars, truncates at `## Slide N` boundaries (preserves complete slides)
- Generates both outputs in one API call to eliminate redundant token consumption
- **Executive summary**: 3-5 paragraph scientific prose covering research question, methods, key results, implications
- **Key takeaways**: 4-7 bullet points focusing on novel results, clinical implications, methodological advances

#### 6. Quality Assessment

- Scores 5 talk-specific dimensions on 0-10 scale (see [Quality Assessment](#processing-log--quality-assessment))
- Overall score: weighted combination, clamped to 0-10
- Assessment label: Excellent / Good / Fair / Poor

#### 7. Markdown Generation

- Structured format with YAML frontmatter (quality scores included)
- Slide-by-slide content with visual element descriptions
- Executive summary + abstract (if available) + key takeaways

---

## Key Features

### Vision AI with OCR Support
- **OCR pre-pass**: Tesseract extracts reference text from slide screenshots
- **Batch processing**: 5 slides per API call for efficiency
- **RAG-enhanced**: OCR text guides Vision AI for accurate extraction
- **Error correction**: Vision AI corrects OCR mistakes against visual content

### Intelligent Metadata Matching
- **Speaker-based matching**: Finds abstracts by speaker last name + title keywords
- **Graceful degradation**: Talks without abstracts proceed with slide-only extraction
- **Rich metadata**: Session title, start time, abstract URL, authors, affiliations

### Production Ready
- **Resume capability**: `skip_existing=True` allows interrupted runs to continue
- **Processing log**: Tab-separated log with quality scores for all processed talks
- **Quality frontmatter**: Scores embedded in each markdown file
- **Merck Foundry integration**: Uses internal AI proxy
- **Cached API client**: Single Anthropic client instance reused across all calls

---

## Output Format

### Generated Files

```
output_talks/
├── talk_06_ED03_Moor_towards_reliable_medical_ai.md
├── talk_52_MSMCB0701_Boija_oral_smallmolecule_condensate_modulator_cmod.md
├── processing_log.txt
└── ...
```

### Filename Convention

```
talk_{NUMBER}_{SESSION}_{Speaker}_{title_slug}.md
```

### Markdown Structure

```markdown
---
talk_number: 06
session_code: ED03
session_title: "AI and Foundation Models in Medicine"
speaker: "M. Moor"
title: "Towards reliable medical AI"
num_slides: 13
abstract_number: 12345
abstract_url: "https://www.abstractsonline.com/..."
start_time: "4/17/2026 10:00:00 AM"
processing_date: 2026-04-29
vision_model: claude-sonnet-4-6
quality_overall: 8.4/10
quality_assessment: Excellent
---

# Towards reliable medical AI

**Speaker**: M. Moor
**Session**: ED03 — AI and Foundation Models in Medicine

**Authors**: Michael Moor, MD, PhD (if abstract matched)
**Affiliations**: ETH Zurich (if abstract matched)

---

## Executive Summary

AI-generated 3-5 paragraph scientific summary capturing the full
narrative arc of the presentation...

---

## Abstract

Published abstract text from AACR (if matched from Excel metadata)...

---

## Slide Content

## Slide 1: Title Slide
**Text content**: ETH Zurich, DBSSE, Towards reliable medical AI...
**Visual elements**: Background image, university logos, presenter info
**Key point**: Introduction of the talk topic and presenter

---

## Slide 2: Healthcare AI Vision
**Text content**: "...it could change healthcare as we know it"
- Integrate multi-modal data
- Earlier & more comprehensive diagnosis
**Visual elements**: Flow diagram showing patient data pipeline...
**Key point**: Medical AI can transform healthcare through multimodal integration

[... additional slides ...]

---

## Key Takeaways

- Key finding or message #1
- Key finding or message #2
- Key finding or message #3
- Key finding or message #4
- Key finding or message #5
```

---

## Excel Metadata Requirements

The Excel file should be `AACR2026_Abstracts.xlsx` with a sheet named `Sheet1` containing:

| Column | Description |
|--------|-------------|
| `Abstract Title` | Title of the abstract (HTML-formatted) |
| `Abstract Text` | Full abstract text (HTML-formatted) |
| `Abstract Number` | Unique abstract identifier |
| `Abstract Authors` | Author list with affiliations (HTML superscripts) |
| `Abstract Companies` | Institutional affiliations |
| `Activity` | "Invited Speaker", "Abstract Submission", "Late Breaking..." |
| `Session Title` | Session name |
| `Start` | Presentation start time |
| `AbstractUrl` | Link to online abstract |

**Matching logic**: The pipeline searches `Abstract Authors` for the speaker's last name, then ranks matches by title keyword overlap (minimum score of 2 required).

---

## Processing Log & Quality Assessment

### Processing Log

After processing, check `output_talks/processing_log.txt`:

```
TIMESTAMP           TALK_NUM  SESSION     SPEAKER  SLIDES  OCR_CHARS  CONTENT_CHARS  SUMMARY_CHARS  ABSTRACT_MATCH  QUALITY  ASSESSMENT  STATUS
2026-04-29 15:23:09 06        ED03        Moor     13      3460       17355          4122           No              8.4      Excellent   SAVED
2026-04-29 15:30:15 52        MSMCB0701   Boija    18      4820       33689          3382           Yes             9.1      Excellent   SAVED
2026-04-29 15:35:00 01        ED54        Shain    16      0          0              0              No                                   FAILED_EXTRACTION
```

### Quality Dimensions

| Dimension | Weight | What it Measures | Score 10 = |
|-----------|--------|------------------|------------|
| Content Coverage | 0.30 | Chars extracted per slide | >= 800 chars/slide |
| OCR Support | 0.15 | OCR text available as ratio of final content | ~67% ratio |
| Summary Quality | 0.25 | Executive summary length + takeaway count | 3000+ chars + 5 takeaways |
| Slide Coverage | 0.20 | % of slides with `## Slide N` headers in output | 100% coverage |
| Abstract Enrichment | 0.10 | Whether abstract metadata was matched | Yes = 10, No = 5 |

### Quality Thresholds

- **Excellent** (8-10): Comprehensive extraction, rich summaries, full coverage
- **Good** (5.5-8): High quality, most slides captured, good summary
- **Fair** (4-5.5): Partial extraction, some slides missed
- **Poor** (<4): Major extraction failures

### Status Values

- `SAVED`: Successfully processed and markdown generated
- `FAILED_CONVERSION`: Could not convert PDF to images
- `FAILED_EXTRACTION`: Vision AI returned no content

---

## Performance & Scalability

### Processing Time

- **Single talk** (13 slides): ~1.5 minutes (with concurrent batches)
- **Single talk** (30+ slides): ~3-4 minutes
- **Full batch** (77 talks): ~3-4 hours estimated
- **OCR overhead**: ~5 seconds per talk (negligible)
- **Concurrency speedup**: ~2.5-3x faster than sequential batch processing

### API Usage

Each talk makes approximately:
- 3-8 API calls for batch slide extraction (5 slides per call, max 8192 output tokens each)
- 1 API call for executive summary + key takeaways combined (max 3072 output tokens)

**Total per talk**: ~4-9 API calls, ~18K-35K tokens (varies by slide count)

**Full batch estimate**: ~400-600 API calls for all 77 talks

### Optimization Features

- **Cached API client**: Single Anthropic client instance reused across all calls
- **Pre-encoded images**: All slides encoded to base64 before threading (PIL thread-safety)
- **Concurrent batches**: Up to 5 simultaneous API calls (~2.5x speedup over sequential)
- **Batch processing**: 5 slides per API call (5x fewer calls than per-slide)
- **Combined summary call**: Executive summary + key takeaways in one API call (saves 1 call per talk)
- **Smart content truncation**: Truncates at `## Slide N` boundaries, not arbitrary char positions
- **Image resizing**: Max 1568px dimension reduces API payload size
- **JPEG quality 85**: Balances image size and readability
- **Cached existing files**: Output directory scanned once at startup
- **Skip-existing mode**: Fast reruns after interruption (`--no-skip` to override)
- **TESSDATA_PREFIX auto-detection**: Configured once at initialization from conda environment

---

## Troubleshooting

### API Credentials Missing

```powershell
# Set credentials (Windows PowerShell - Administrator mode)
[Environment]::SetEnvironmentVariable("ANTHROPIC_AUTH_TOKEN", "your-token", "User")
[Environment]::SetEnvironmentVariable("ANTHROPIC_BASE_URL", "https://palantir.mcloud.merckgroup.com/language-model-service/api/proxy/anthropic", "User")

# Restart terminal to load new variables
```

### OCR Not Working

```bash
# Check Tesseract is installed
tesseract --version

# If missing in conda:
conda install -c conda-forge tesseract -y

# Verify TESSDATA_PREFIX
python -c "import pytesseract; print(pytesseract.get_tesseract_version())"
```

The pipeline gracefully handles OCR failures — it proceeds with Vision AI only (quality may be slightly lower).

### Excel Permission Errors

If `PermissionError` occurs when reading the Excel file:
- Close the file in Excel
- Or use pandas (which handles some lock scenarios): the pipeline uses `pd.read_excel()` which is more resilient

### OneDrive Sync Delays

The talks folder is on OneDrive. If files appear missing or processing is slow:
- Ensure OneDrive has synced the folder locally
- Check file availability (cloud-only files need to download first)
- Large PDFs (50-125 MB) may take time to sync

### Low Quality Scores

If talks score FAIR or POOR:
- Check if the PDF is corrupted or has unusual formatting
- Very dark slides or low-contrast text may reduce OCR quality
- Animated slides captured as screenshots may have overlapping content
- Review the `processing_log.txt` for patterns

---

## Known Limitations

1. **Image-only PDFs**: Relies entirely on Vision AI for content extraction (no native text available)
2. **Batch granularity**: 5 slides per batch may miss inter-slide context; narrative flow between batches is inferred
3. **Abstract matching**: Speaker name matching may fail for common surnames or hyphenated names
4. **Animated content**: Slides captured mid-animation may have overlapping or partial content
5. **Language**: Optimized for English scientific presentations
6. **Large presentations**: Talks with 40+ slides consume significant API tokens
7. **OneDrive latency**: File access from OneDrive synced folders can be slow for large PDFs
8. **No quality gate**: All talks generate markdown regardless of quality score (no FAIR/POOR filtering)

---

## License

Internal tool for Merck Group - Conference talk processing

**Confidential & Proprietary**
