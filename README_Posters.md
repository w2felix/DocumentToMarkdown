# Scientific Poster Processing Pipeline

Automated extraction and analysis of scientific conference posters (PDF) using AI-powered vision processing. Converts single-page poster PDFs into structured, searchable markdown with comprehensive figure analysis, quality assessment, and metadata integration.

**Powered by Claude Vision AI (Sonnet 4.6)** with mandatory RAG enhancement for high-accuracy text extraction from complex multi-column poster layouts.

---

## Table of Contents

- [Quick Start](#quick-start)
- [Installation & Setup](#installation--setup)
- [Usage](#usage)
  - [Command-Line Interface](#command-line-interface)
  - [Python API](#python-api)
  - [Processing a Single Poster](#processing-a-single-poster)
  - [Processing All Posters](#processing-all-posters)
- [Pipeline Architecture](#pipeline-architecture)
  - [Processing Stages](#processing-stages)
  - [Key Features](#key-features)
- [Output Format](#output-format)
- [Excel Metadata Requirements](#excel-metadata-requirements)
- [Processing Log & Quality Assessment](#processing-log--quality-assessment)
- [Performance & Scalability](#performance--scalability)
- [Troubleshooting](#troubleshooting)
- [Comparison with Talk Pipeline](#comparison-with-talk-pipeline)
- [Known Limitations](#known-limitations)

---

## Quick Start

```bash
# 1. Activate environment
conda activate ds_env

# 2. Process all posters (uses default paths)
python poster_pipeline.py --sharepoint "path/to/sharepoint_folder" --metadata "path/to/metadata.xlsx"

# 3. Process a single poster
python poster_pipeline.py --sharepoint "test_poster" --metadata "test_poster/AACR26_Selected_Apr13.xlsx" --single "test_poster/160.pdf"

# 4. Custom output directory
python poster_pipeline.py --sharepoint "posters" --metadata "metadata.xlsx" --output "output_posters"

# 5. Reprocess all (ignore previously generated files)
python poster_pipeline.py --sharepoint "posters" --metadata "metadata.xlsx" --no-skip
```

The pipeline will process all PDFs and generate markdown files in `output/`.

---

## Installation & Setup

See [setup.md](setup.md) for complete installation instructions.

**Prerequisites:**
- Miniconda (Python 3.11)
- Required packages: `pdfplumber`, `pymupdf`, `pandas`, `openpyxl`, `anthropic`, `pytesseract`, `pillow`
- API credentials: `ANTHROPIC_AUTH_TOKEN` and `ANTHROPIC_BASE_URL`

**Additional for OCR fallback:**
- Tesseract OCR installed and accessible
- `TESSDATA_PREFIX` configured (auto-detected in conda environments)

---

## Usage

### Command-Line Interface

```bash
python poster_pipeline.py [OPTIONS]
```

| Option | Default | Description |
|--------|---------|-------------|
| `--sharepoint` | *(required)* | Folder containing PDF posters |
| `--metadata` | *(required)* | Excel file with poster metadata |
| `--output` | `output` | Output directory for markdown files |
| `--single` | *(none)* | Process a single PDF file only |
| `--force-ocr` | `False` | Force OCR even if native text extraction works |
| `--no-skip` | `False` | Reprocess files that already exist |
| `--ocr-dpi` | `200` | DPI for Tesseract OCR |
| `--no-detailed-analysis` | `False` | Skip Stage 2 figure analysis (faster) |
| `--recursive` | `False` | Recursively search subfolders for PDF files |

### Python API

```python
from poster_pipeline import PosterPipeline

pipeline = PosterPipeline(
    sharepoint_folder="path/to/posters",
    metadata_excel="path/to/AACR26_Selected_Apr13.xlsx",
    output_dir="output"
)
```

### Processing a Single Poster

```python
from pathlib import Path

pipeline.process_single_poster(
    Path("path/to/160.pdf"),
    skip_existing=False
)
```

### Processing All Posters

```python
# Process all, skip already-processed files
pipeline.process_all_posters(skip_existing=True)

# Reprocess everything
pipeline.process_all_posters(skip_existing=False)

# Fast mode (no detailed figure analysis)
pipeline.process_all_posters(enable_detailed_analysis=False)
```

---

## Pipeline Architecture

The pipeline handles single-page scientific poster PDFs using a multi-stage approach combining native PDF text extraction, OCR fallback, and mandatory Vision AI enhancement.

### Processing Stages

#### 1. Poster Number Extraction & Metadata Matching

- Extracts poster number from filename (e.g., `160.pdf` → "160", `LB197.pdf` → "LB197")
- Requires at least 2 digits to avoid false matches on gene names (B7, CD3)
- Verifies candidate number exists in metadata before accepting
- **Title-based fallback**: If regex extraction fails, matches filename against metadata `Presentation Title` using:
  - Trailing number pattern (e.g., `Some Title - 348.pdf`)
  - Exact substring matching against metadata titles
  - Fuzzy matching (SequenceMatcher, threshold ≥0.75)
- Searches Excel metadata across multiple columns (Presentation Number, Poster Number, Abstract Number, Session Number)
- Returns full metadata row with session info, colleagues, scheduling
- Tracks unrecognized filenames that cannot be matched to any poster number

#### 2. PDF-to-Image Conversion & Encoding

- Converts poster PDF to PIL Image at 300 DPI using PyMuPDF
- Uses `Image.frombytes()` for direct pixel transfer (no PNG roundtrip)
- Resizes to max 2048px dimension (preserving aspect ratio) for API efficiency
- Encodes to base64 JPEG (quality 85) once — reused across all Vision AI calls
- Frees PIL image from memory after encoding

#### 3. Parallel: Structure Extraction + Text Extraction

Runs two operations concurrently via ThreadPoolExecutor:

- **Thread 1 — Structure Extraction (Vision AI)**: Identifies layout type (single-column, multi-column, irregular), section positions, reading order, figure locations
- **Thread 2 — Base Text Extraction**: Attempts native PDF text (pdfplumber), falls back to OCR (Tesseract) if quality is poor

#### 4. Mandatory Vision AI Enhancement (RAG)

- Always runs, regardless of base extraction quality
- Uses extracted text as RAG (Retrieval-Augmented Generation) context
- Corrects reading order issues from multi-column layouts
- Fixes OCR errors (character misrecognition: l vs 1, O vs 0)
- Preserves scientific terminology and formatting
- Result method: `{original_method}_vision_enhanced`

#### 5. Template Artifact Removal

- Detects and removes 20+ printing template patterns
- Cleans Genigraphics instructions, PowerPoint tips, "DO NOT POST" warnings
- Applied after Vision AI enhancement to clean final text

#### 6. Figure Analysis (Two-Stage with Parallel Processing)

- **Stage 1 — Figure Identification**: Identifies all figures, extracts numbers/titles/types, generates initial descriptions, includes caption extraction
- **Stage 2 — Detailed Analysis (Parallel)**: Deep per-figure analysis with up to 5 concurrent API calls; extracts axis labels, statistical values, p-values, data series, color coding, legends
- Fallback: If no figures detected in Stage 1, retries with more explicit prompt

#### 7. Section Extraction + Executive Summary (Combined API Call)

- Single API call extracts structured sections AND generates executive summary
- Uses `---SUMMARY---` delimiter to split combined response
- Sections: Abstract, Objective, Methods, Results, Conclusions
- Executive summary: 2-3 sentence scientific prose
- Robust fallback logic if sections are missing

#### 8. Quality Assessment & Filtering

- Scores 4 dimensions on 0-10 scale (see [Quality Assessment](#processing-log--quality-assessment))
- Applies penalties for template contamination and figure-caption mismatches
- Only generates markdown for GOOD (≥5.5) or EXCELLENT (≥8.0) quality
- FAIR/POOR posters logged but not saved

#### 9. Markdown Generation

- Structured format with YAML frontmatter (quality scores included)
- Sections organized by reading order
- Figures interleaved with detailed analysis
- Full extracted text in collapsible details block

---

## Key Features

### Vision AI with RAG Enhancement
- **Mandatory enhancement**: All posters processed with Vision AI (not optional)
- **RAG-powered**: Uses extracted text as context for better accuracy
- **Multi-column support**: Handles complex poster layouts with correct reading order
- **Error correction**: Fixes OCR character misrecognition against visual content

### Intelligent Figure Analysis
- **Two-stage approach**: Quick identification + deep per-figure analysis
- **Parallel Stage 2**: Up to 5 figures analyzed simultaneously
- **Caption extraction**: Spatially identifies and extracts figure captions
- **Rich output**: Axis labels, statistical values, key findings per figure

### Production Ready
- **Resume capability**: `skip_existing=True` allows interrupted runs to continue
- **Quality filtering**: Only saves high-quality outputs (GOOD or EXCELLENT)
- **Processing log**: Tab-separated log with quality scores for all posters
- **Merck Foundry integration**: Uses internal AI proxy
- **Cached API client**: Single Anthropic client instance reused across all calls

---

## Output Format

### Generated Files

```
output/
├── poster_160.md
├── poster_LB197.md
├── quality_log.txt
└── ...
```

### Filename Convention

```
poster_{poster_number}.md
```

### Markdown Structure

```markdown
---
poster_number: 160
interested_colleagues: John Doe, Jane Smith
session_number: S01
session_title: "Biomarkers Predictive of Therapeutic Benefit"
session_type: Poster Session
day: 04/19/2026
session_start: "4/19/2026 2:00:00 PM"
session_end: "4/19/2026 5:00:00 PM"
location: Section 40
extraction_method: native_vision_enhanced
vision_model: claude-sonnet-4-6
processing_date: 2026-04-21
quality_overall: 8.5/10
quality_assessment: Excellent
---

# Poster Title

**Poster Number**: #160
**Authors**: First Author, PhD; Second Author, MD
**Affiliation**: Institution One; Institution Two

---

## Executive Summary

AI-generated 2-3 sentence summary capturing the research question,
approach, and key findings...

---

## Abstract

Full abstract text extracted from poster...

---

## Methods

Methodology section extracted with proper structure...

---

## Results

### Figure 1: Experimental workflow diagram

**Caption from text**: Complete caption extracted via Vision AI
**Type**: Workflow diagram
**Description**: Detailed description of what the figure shows...
**Key Findings**:
- First major finding from the figure
- Second important observation

**Statistical Notes**: p<0.001, n=50 per group

**Detailed Analysis**:
Axis labels, data series, statistical values...

---

## Conclusions

Conclusion section text...

---

## Quality Assessment

**Overall Quality**: 8.5/10 - Excellent
**Component Scores**:
- Text Extraction: 9.2/10
- Figure Analysis: 9.0/10
- Structure Parsing: 7.5/10
- Caption Quality: 8.0/10

**Quality Issues Detected**: None

---

## Full Extracted Text
<details>
<summary>Click to expand full text</summary>
Complete extracted text for reference...
</details>

---

## Processing Metadata
- **Extraction Method**: native_vision_enhanced
- **Vision Model**: claude-sonnet-4-6
- **Number of Figures**: 5
  - With detailed analysis: 5
- **Text Length**: 8,432 characters, 1,245 words
- **Processing Date**: 2026-04-21 14:23:15
```

---

## Excel Metadata Requirements

The Excel file should contain a sheet named `Full_Program_Copy` with columns such as:

| Column | Description |
|--------|-------------|
| Presentation Number / Poster Number / Abstract Number | Poster identifier (searched across multiple columns) |
| Presentation Title | Used for title-based fuzzy matching when filename is not numeric |
| Interested Colleagues | Names of colleagues who flagged this poster |
| Covered by | Assigned reviewer |
| Session Number | Session identifier |
| Session Title | Session name |
| Session Type Name | e.g., "Poster Session" |
| Day | Presentation date |
| Session Start / Session End | Session time window |
| Location | Physical location |

**Matching logic**: The pipeline extracts a poster number from the filename, then searches across Presentation Number, Poster Number, Abstract Number, and Session Number columns. Falls back to full-row scan if no match found. For title-based filenames, uses substring and fuzzy matching against `Presentation Title` (threshold ≥0.75).

---

## Processing Log & Quality Assessment

### Processing Log

After processing, check `output/quality_log.txt`:

```
TIMESTAMP            POSTER_NUM  QUALITY_SCORE  ASSESSMENT                      STATUS   FILENAME
2026-04-21 14:23:15  160         8.5            Excellent                       SAVED    160.pdf
2026-04-21 14:25:30  LB197       7.2            Good                            SAVED    LB197.pdf
2026-04-21 14:28:01  2956        2.7            Poor - Consider manual review   SKIPPED  2956.pdf
2026-04-21 14:30:12  1887        5.3            Fair                            SKIPPED  1887.pdf
```

### Quality Dimensions

| Dimension | Weight | What it Measures | Score 10 = |
|-----------|--------|------------------|------------|
| Text Extraction | 0.25 | Length, coherence, low fragmentation | >10K chars, low tiny-word ratio |
| Figure Analysis | 0.35 | Figure count, detail level, key findings | 7+ figures with detailed analysis |
| Structure Parsing | 0.25 | Core sections present (Methods, Results, Conclusions) | All core + optional sections |
| Caption Quality | 0.15 | Clean captions without template contamination | All captions clean |

### Quality Penalties

| Penalty | Amount | Trigger |
|---------|--------|---------|
| Template contamination | -0.5 per artifact | Genigraphics, PowerPoint tips detected |
| Template cap | -2.0 max | Multiple template artifacts |
| Figure-caption mismatch | -0.5 per mismatch | Missing or extra figures vs text captions (tolerance ±1) |
| Figure-caption mismatch cap | -2.0 max | Consistency penalty capped to prevent excessive point loss |
| Caption artifacts | -2.0 per contaminated | Template text found in figure captions |

### Quality Thresholds

- **Excellent** (8-10): Comprehensive extraction, rich figure analysis, no issues → **Markdown generated**
- **Good** (5.5-8): High quality, minor issues → **Markdown generated**
- **Fair** (4-5.5): Partial extraction, incomplete sections → **Skipped, logged for review**
- **Poor** (<4): Major extraction failures → **Skipped, manual review required**

### Status Values

- `SAVED`: Successfully processed and markdown generated
- `SKIPPED`: Quality below threshold (FAIR/POOR)
- `FAILED`: Processing error (PDF conversion, API failure)
- `UNRECOGNIZED`: Filename could not be matched to any poster number in metadata

---

## Performance & Scalability

### Processing Time

- **Single poster** (few figures): ~1.5-2 minutes
- **Single poster** (many figures): ~2.5-4 minutes
- **Full batch** (100+ posters): ~3-5 hours estimated
- **Fast mode** (`--no-detailed-analysis`): ~60% faster

### API Usage

Each poster makes approximately:
- 1 API call for structure extraction (parallel with text)
- 1 API call for Vision AI text enhancement (~10K tokens)
- 1 API call for figure identification — Stage 1 (~8K tokens)
- N API calls for detailed figure analysis — Stage 2 (~5K tokens each, parallel)
- 1 API call for section extraction + executive summary combined (~3K tokens)

**Total per poster**: 4+N API calls, ~30K-60K tokens (varies by figure count)

### Optimization Features

- **Cached API client**: Single Anthropic client instance reused across all calls
- **Pre-encoded image**: Base64 JPEG encoded once, passed to all Vision AI calls
- **Parallel Steps 1+2**: Structure extraction and text extraction run concurrently
- **Parallel figure analysis**: Up to 5 concurrent Stage 2 API calls
- **Combined section+summary call**: One API call generates both (saves 1 call per poster)
- **Image.frombytes()**: Direct pixel buffer transfer, no PNG encode/decode roundtrip
- **Max 2048px resize**: Images resized to max dimension before encoding, reduces API payload
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

### Package Import Errors

```bash
# Activate environment
conda activate ds_env

# Install all dependencies
conda install -c conda-forge pdfplumber pymupdf pandas openpyxl pytesseract pillow -y
pip install anthropic

# Verify installations
python -c "import pdfplumber, anthropic, pandas, pytesseract; print('All packages OK')"
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

The pipeline gracefully handles OCR failures — it proceeds with Vision AI only.

### Excel Permission Errors

If `PermissionError` occurs when reading the Excel file:
- Close the file in Excel
- Or use pandas (which handles some lock scenarios): the pipeline uses `pd.read_excel()` which is more resilient

### Low Quality Scores

If many posters are being skipped (FAIR/POOR quality):
- Check PDF quality — ensure they're not corrupted
- Verify PDFs are text-based or have good scan quality
- Try `--force-ocr` for scanned posters
- Review `output/quality_log.txt` for specific issues
- Check for excessive template contamination

---

## Comparison with Talk Pipeline

| Feature | Poster Pipeline | Talk Pipeline |
|---------|----------------|---------------|
| Input format | Single-page PDF (vector/text) | Multi-page PDF (screenshot images) |
| Text extraction | Native + OCR + Vision AI (RAG) | OCR + Vision AI only |
| Pages per document | 1-2 | 13-40+ |
| Processing approach | Single image, multi-pass | Batch slides (5/call) |
| Figure handling | Dedicated 2-stage analysis | Described within slide content |
| Metadata source | `Full_Program_Copy` sheet | `Sheet1` sheet |
| Matching key | Poster number from filename (with title-based fuzzy fallback) | Speaker name + title keywords |
| Abstract availability | Always available (on poster) | 57% of talks |
| Quality dimensions | Text, Figures, Structure, Captions | Content, OCR, Summary, Coverage, Abstract |
| Quality filtering | Yes (FAIR/POOR skipped) | No (all talks saved) |
| Output structure | Section-based (Methods/Results/Conclusions) | Slide-based (Slide 1, 2, 3...) |
| Processing time | 1.5-4 min/poster | 1.5-4 min/talk |

---

## Known Limitations

1. **Language**: Optimized for English scientific posters
2. **Image quality**: Very low-resolution scans may have reduced accuracy
3. **Non-standard formats**: Highly irregular layouts may have reading order issues
4. **Figure types**: 3D molecular structures and complex schematics may have limited analysis
5. **Math equations**: Complex LaTeX equations may not extract perfectly
6. **API dependency**: Requires active Merck Foundry access
7. **Quality gate**: FAIR/POOR posters are not saved — may miss recoverable content

---

## License

Internal tool for Merck Group - Conference poster processing

**Confidential & Proprietary**
