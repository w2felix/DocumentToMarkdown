# Presentation Processing Pipeline

Automated extraction and structuring of corporate and scientific presentations (PPTX + PDF) into searchable markdown. Handles slide decks, meeting materials, agendas, and scientific presentations with native text extraction — no OCR needed.

**Powered by Claude Vision AI (Sonnet 4.6)** for figure analysis, chemical structure detection, and executive summary generation. Vision AI is optional (`--no-vision` for text-only mode).

---

## Table of Contents

- [Quick Start](#quick-start)
- [Installation & Setup](#installation--setup)
- [Usage](#usage)
  - [Command-Line Interface](#command-line-interface)
  - [Python API](#python-api)
- [Pipeline Architecture](#pipeline-architecture)
  - [Processing Stages](#processing-stages)
  - [Key Features](#key-features)
- [Output Format](#output-format)
  - [Naming Schemes](#naming-schemes)
  - [Markdown Structure](#markdown-structure)
- [Classification Detection](#classification-detection)
- [Quality Assessment](#quality-assessment)
- [Performance](#performance)
- [Troubleshooting](#troubleshooting)
- [Known Limitations](#known-limitations)

---

## Quick Start

```bash
# 1. Activate environment
conda activate ds_env

# 2. Process all presentations in a folder
python presentation_pipeline.py --input "path/to/presentations"

# 3. Text-only mode (no API calls)
python presentation_pipeline.py --input "path/to/presentations" --no-vision

# 4. Single file
python presentation_pipeline.py --single "path/to/file.pptx"

# 5. Standardized date-based filenames
python presentation_pipeline.py --input "path/to/presentations" --naming dated

# 6. Include subfolders
python presentation_pipeline.py --input "path/to/presentations" --recursive
```

---

## Installation & Setup

See [setup.md](setup.md) for complete installation instructions.

**Prerequisites:**
- Miniconda (Python 3.11)
- Required packages: `python-pptx`, `pymupdf`, `pdfplumber`, `pillow`, `anthropic`, `pywin32`
- API credentials: `ANTHROPIC_AUTH_TOKEN` and `ANTHROPIC_BASE_URL`
- Optional: Microsoft PowerPoint (enables Vision AI analysis for PPTX files)

```bash
conda activate ds_env
pip install python-pptx pywin32
```

---

## Usage

### Command-Line Interface

```bash
python presentation_pipeline.py [OPTIONS]
```

| Option | Default | Description |
|--------|---------|-------------|
| `--input` | `test_presentation` | Folder containing PPTX/PDF presentations |
| `--output` | `output_presentations` | Output directory for markdown files |
| `--single` | *(none)* | Process a single file only |
| `--recursive` | `False` | Also scan subfolders for presentations |
| `--no-skip` | `False` | Reprocess files that already exist |
| `--no-vision` | `False` | Text-only extraction (no API calls) |
| `--naming` | `default` | Filename scheme: `default`, `dated`, or `classified` |
| `--verbose` | `False` | Enable debug logging |

### Python API

```python
from presentation_pipeline import PresentationPipeline
from pathlib import Path

pipeline = PresentationPipeline(
    input_folder="path/to/presentations",
    output_dir="output_presentations",
    no_vision=False,
    recursive=False,
    naming="dated"
)

# Process all
pipeline.process_all_presentations(skip_existing=True)

# Process single file
pipeline.process_single_presentation(Path("path/to/file.pptx"), skip_existing=False)
```

---

## Pipeline Architecture

### Processing Stages

#### 1. Filename Parsing & Date Extraction

Parses presentation filenames to extract:
- **Date** — supports multiple formats: `YYYYMMDD`, `YYYY_MM_DD_`, `_DD.MM.YYYY`, `DDMonYYYY`, `Month YYYY`
- **Title** — cleaned of date prefixes/suffixes and classification markers
- **Hints** — `[No Structures]` marker, classification keywords

#### 2. Content Extraction

**PPTX (native, no OCR):**
- Core properties: title, author, created/modified dates, revision
- Per-slide: title placeholder, text shapes, paragraphs
- Tables: extracted cell-by-cell, rendered as markdown tables
- Hyperlinks: extracted with anchor text and URL
- Speaker notes: extracted when present
- Embedded media: flagged as non-extractable

**PDF (text-first, vision fallback):**
- Text extraction via PyMuPDF (handles rotated text, subscripts, superscripts correctly)
- Tables extracted via pdfplumber (superior table detection)
- Deduplication removes text that appears in both text and table extractions
- Chart axis data (numeric tick marks) collapsed into `[Chart axis: ...]` markers
- If text is sparse (<50 chars/page avg) or garbled (>20% pages with issues): falls back to Vision AI
- Vision fallback renders pages at 150 DPI, processes in batches of 5

#### 3. Language Detection

Lightweight heuristic applied after content extraction:
- Samples first 3000 characters from up to 20 slides
- Counts frequency of German indicator words (`und`, `der`, `die`, `für`, `wird`, etc.) vs English indicators (`the`, `and`, `for`, `with`, etc.)
- If German ratio > 0.4 → `language: de`, otherwise `language: en`
- Result stored in YAML frontmatter and used to instruct summary generation in English

#### 4. Classification Detection

Three-signal approach (see [Classification Detection](#classification-detection)):
1. Filename patterns
2. Slide content scanning
3. Vision AI visual cues (watermarks, banners)

#### 5. Content Type Classification

Classifies presentations into:
- `external` — project discussions, findings, data, arguments from an external company (partner, CRO, vendor)
- `project` — internal project or concept discussions, findings, data, arguments (discovery, preclinical, clinical)
- `academic` — scientific/technical findings, data, arguments in an academic setting (congress, publication, trial results)
- `operational` — meeting notes, status updates, project plans, department or company updates
- `agenda` — primarily topic/schedule lists
- `template` — mostly empty/placeholder slides

Classification uses keyword scoring across four marker categories (external, academic, project, operational). The highest-scoring category wins, with a minimum threshold of 2 matching markers. Defaults to `operational` when no strong signal is found.

#### 6. Vision AI Analysis (smart gating)

**Decision logic:**

| Condition | Vision AI Action |
|-----------|-----------------|
| `--no-vision` flag | Skip all Vision AI |
| PPTX with no images/charts/diagrams | Skip `analyze_with_vision()` |
| PPTX with images/charts | Run `analyze_with_vision()` (requires PowerPoint for rendering) |
| PDF with sparse/garbled text | Run vision fallback for extraction |
| PDF with good text | Run `analyze_with_vision()` for figures |

When enabled, analyzes representative slides for:
- Figures & charts (type, description, axes, key findings)
- Chemical structures (SMILES notation with confidence)
- Audience classification and topic tagging
- Classification visual cues

**PPTX**: Slides are rendered to images via PowerPoint COM automation (requires PowerPoint installed). If PowerPoint is not available, a warning is logged and Vision AI analysis is skipped for PPTX files (text extraction still works via python-pptx).

**PDF**: Slides are rendered via PyMuPDF at 150 DPI.

#### 6b. Per-Slide Image Enrichment (PPTX only)

For PPTX files with embedded images, the pipeline identifies slides that contain **meaningful visual content** (not logos or icons) and sends their rendered slide images to Vision AI for per-slide descriptions.

**Image filtering criteria** — an embedded image is considered "meaningful" if:
- Blob size ≥ 20 KB (excludes tiny icons and logos)
- Width ≥ 150 px AND height ≥ 100 px (excludes thin banners and decorative elements)

**Processing:**
- Slides are ranked by meaningful image count (descending), then text sparsity (ascending)
- Up to 45 slides are sent for enrichment (batches of 3, up to 5 concurrent)
- Each batch asks Vision AI to describe visual content without repeating slide text
- Results are merged into the slide's `**Visual elements**` field in the output

**What gets described:**
- Scientific figures (axes, data trends, conclusions)
- Screenshots (UI, workflow, tool shown)
- Photos and diagrams (subject matter, key information)
- Conference poster panels (mechanism diagrams, data plots)

**Cost:** Typically 5–15 extra API calls for image-heavy presentations (~$0.10–0.30). Skipped entirely with `--no-vision`.

When enrichment occurs, the `extraction_method` in frontmatter updates to `native_text+vision_enrichment`.

#### 7. Chemical Structure SMILES Refinement

For detected chemical structures with low/medium confidence:
- Re-renders slide at 250 DPI
- Focused SMILES-only extraction pass with surrounding text as context
- Validates SMILES format (bracket/parenthesis balance)

#### 8. Summary Generation (conditional)

Generated for all content types except `agenda` and `template`:
- Executive summary (2-4 paragraphs)
- Key takeaways (4-7 bullet points)
- Always written in English, even when source content is in another language

Skipped for `agenda` and `template` presentations.

#### 9. Structured Data Extraction

- **Action items** — detected from slide titles containing "Next Steps" / "Action Items" / "To-Do"
- **Key metrics** — percentages, p-values, sample sizes, dosages
- **Timelines** — dates and milestones with quarter/year references (1970+)

#### 10. Quality Scoring & Markdown Generation

Weighted 0-10 score with assessment labels, then full markdown output.

---

## Key Features

### Dual Format Support
- **PPTX**: Native text extraction via python-pptx — fast, accurate, no OCR
- **PDF**: PyMuPDF text extraction (direction-aware, handles rotated text and subscripts) + pdfplumber for tables, with Vision AI fallback for image-based slides

### Smart Vision AI Gating
- Vision AI is only invoked when there's visual content that text extraction cannot capture
- **PPTX**: Detects images (`PICTURE`), charts (`CHART`), and diagrams (`IGX_GRAPHIC`, `DIAGRAM`) per slide during extraction. If none are found, `analyze_with_vision()` is skipped entirely
- **PDF**: Triggers Vision AI fallback only when text is sparse (<50 chars/page average) or garbled (>20% pages with reversed/fragmented text)
- Summary generation is always text-based (no images needed) — runs independently of vision gating
- Net effect: a text-only PPTX → 1 API call (summary), not 2–3

### Per-Slide Image Enrichment
- PPTX slides with meaningful embedded images (scientific figures, screenshots, photos) get individual Vision AI descriptions
- Smart filtering excludes logos, icons, and decorative elements using blob size (≥20KB) and dimensions (≥150×100px)
- Descriptions are injected as `**Visual elements**` per slide — capturing information invisible to text extraction
- Up to 45 slides enriched per presentation, processed in parallel batches of 3

### Language Detection
- Lightweight heuristic: counts German vs English indicator words in slide text
- Detected language stored in YAML frontmatter (`language: de` or `language: en`)
- Executive summary and key takeaways are always written in English, regardless of source language
- Slide content is preserved in the original language (direct extraction, not translated)

### Smart Classification
- Detects confidentiality level from filename, content, and visual cues
- Defaults to `confidential` for safety (Merck internal documents)
- Recognizes "Non Con" as explicit public marker

### Conditional Processing
- Skips summary generation for agendas and templates
- Skips SMILES refinement when `[No Structures]` is in filename
- Skips Vision AI analysis for PPTX files with no images/charts/diagrams
- Vision AI is entirely optional (`--no-vision`)

### Chemical Structure Support
- Detects structures and reaction schemes in slides
- Extracts SMILES notation with confidence scoring
- Refines low-confidence structures at higher resolution

### Structured Data Extraction
- Action items with owner/deadline detection
- Key scientific metrics (percentages, p-values, ratios)
- Timeline and milestone extraction
- Hyperlink and reference collection

---

## Output Format

### Naming Schemes

| Scheme | Pattern | Example |
|--------|---------|---------|
| `default` | `presentation_{title}` | `presentation_ru_onc_operations_update_darmstadt.md` |
| `dated` | `{date}_{classification}_{title}` | `2026-04-15_confidential_ru_onc_operations_update_darmstadt.md` |
| `classified` | `{classification}_{date}_{title}` | `confidential_2026-04-15_ru_onc_operations_update_darmstadt.md` |

The `dated` scheme sorts chronologically in file explorers. The `classified` scheme groups by sensitivity level. Files without a detected date use `undated`. Title slugs are limited to 5 words for readable filenames.

### Markdown Structure

```markdown
---
title: "RU ONC Operations Update Darmstadt"
filename: "RU ONC Operations Update Darmstadt_15.04.2026.pptx"
date: "2026-04-15"
author: "Laura Schulz"
last_modified_by: "Laura Schulz"
revision: 1
classification: confidential
audience: internal_team
topics:
  - oncology
  - operations
content_type: operational
language: en
source_format: pptx
num_slides: 16
has_speaker_notes: false
has_chemical_structures: false
has_action_items: true
has_summary: true
extraction_method: native_text+vision_enrichment
vision_model: claude-sonnet-4-6
processing_date: 2026-05-12
quality_overall: 8.3/10
quality_assessment: Excellent
---

# RU ONC Operations Update Darmstadt

**Date**: 2026-04-15  |  **Author**: Laura Schulz  |  **Classification**: CONFIDENTIAL
**Audience**: Internal Team  |  **Topics**: oncology, operations

---

## Executive Summary

[AI-generated 2-4 paragraphs — only for substantive/operational content]

---

## Slide Content

### Slide 1: Title
Text content...
**Visual elements**: Description of embedded images, screenshots, or figures on this slide
**Notes**: speaker notes if available

### Slide 2: Agenda
| Item | Time |
|------|------|
| Topic A | 10:00 |

---

## Figures & Charts

### Figure 1: Enrollment Timeline (Slide 8)
**Type**: line_chart
**Description**: Patient enrollment over time...
**Key Findings**: Enrollment accelerated in Q2

---

## Chemical Structures

### Structure 1: Compound A (Slide 12)
**SMILES** (confidence: high): `CC(=O)Oc1ccccc1C(=O)O`
**Molecular Formula**: C9H8O4  |  **MW**: 180.16

---

## Key Metrics & Data Points

- ORR: 42% (95% CI: 34-51%) (Slide 8)
- n=450 enrolled patients (Slide 5)

---

## Action Items

| Action | Owner | Deadline | Slide |
|--------|-------|----------|-------|
| Submit IND package | J. Smith | 2026-06-15 | 14 |

---

## Timeline & Milestones

- **Q2 2026**: Data readout
- **Q3 2026**: Phase 2 initiation

---

## References & Links

- [Protocol v2.1](https://sharepoint.merck.com/...) (Slide 3)
- Slide 7: embedded video (not extractable)

---

## Key Takeaways

- Takeaway 1
- Takeaway 2

---

## Quality Assessment

**Overall**: 8.3/10 — Excellent

**Component Scores**:
- Text extraction: 9.0/10 (weight 0.30)
- Structure recognition: 8.5/10 (weight 0.20)
- Visual analysis: 7.0/10 (weight 0.15)
- Summary quality: 8.0/10 (weight 0.25)
- Metadata completeness: 10.0/10 (weight 0.10)
```

---

## Classification Detection

The pipeline uses a three-signal approach:

| Signal | Source | Examples |
|--------|--------|----------|
| Filename | Regex on filename | "Confidential", "Non Con", "Secret" |
| Content | Regex on slide text | "For Internal Use Only", "Do Not Distribute" |
| Visual | Vision AI (when enabled) | Watermarks, banners, stamps |

**Resolution:** Highest classification wins. Default is `confidential`.

| Marker | Classification |
|--------|---------------|
| "Non Con", "public" | `public` |
| "Confidential", "Internal Use", "Proprietary" | `confidential` |
| "Secret", "Strictly Confidential" | `secret` |
| No markers found | `confidential` (default) |

---

## Quality Assessment

### Quality Dimensions

| Dimension | Weight | Score 10 = |
|-----------|--------|------------|
| Text extraction | 0.30 | ~600+ chars per slide |
| Structure recognition | 0.20 | All slides have identified titles |
| Visual analysis | 0.15 | All figures described with findings |
| Summary quality | 0.25 | Comprehensive summary + 5+ takeaways |
| Metadata completeness | 0.10 | Title, date, author, classification all present |

### Quality Thresholds

- **Excellent** (8-10): Complete extraction, rich metadata
- **Good** (5.5-8): Good extraction, minor gaps
- **Fair** (4-5.5): Partial extraction, missing sections
- **Poor** (<4): Major extraction failures

All presentations are saved regardless of quality score (unlike the poster pipeline).

---

## Performance

### Processing Time

- **PPTX (no-vision)**: <1 second per file
- **PPTX (with vision, few images)**: ~30-90 seconds (requires PowerPoint installed for slide rendering)
- **PPTX (with vision, image-heavy)**: ~2-4 minutes (includes per-slide image enrichment)
- **PDF (native text)**: 2-5 seconds per file (PyMuPDF text + pdfplumber tables)
- **PDF (with vision)**: ~30-90 seconds (when text is sparse/garbled and Vision AI fallback is triggered)

### API Usage (with Vision enabled)

Each presentation makes approximately:
- 0–1 API call for figure/audience/topic analysis (~4K tokens) — skipped for text-only PPTX with no images/charts
- 0–15 API calls for per-slide image enrichment (~8K tokens each) — only for PPTX slides with meaningful embedded images
- 0-N API calls for SMILES refinement (~1K tokens each)
- 1 API call for executive summary + takeaways (~4K tokens) — only for substantive/operational content

**Smart gating**: A text-only PPTX (no images or charts) uses 1 API call (summary only). A PPTX with a few figures uses 2–3 calls. An image-heavy PPTX (scientific presentations with embedded poster figures) uses 10–18 calls including per-slide enrichment.

**PPTX note**: Full Vision AI analysis requires PowerPoint installed for slide rendering. Without PowerPoint, PPTX still gets text-based summaries (1 API call).

**Text-only mode** (`--no-vision`): Zero API calls — pure local extraction.

---

## Troubleshooting

### python-pptx Not Found

```bash
conda activate ds_env
pip install python-pptx
```

### API Credentials Missing

```powershell
[Environment]::SetEnvironmentVariable("ANTHROPIC_AUTH_TOKEN", "your-token", "User")
[Environment]::SetEnvironmentVariable("ANTHROPIC_BASE_URL", "your-proxy-url", "User")
```

Restart your terminal after setting these.

### PPTX Files Show No Content

Some PPTX files use SmartArt or embedded charts that python-pptx cannot fully extract. Run without `--no-vision` — the pipeline will use PowerPoint to render slides and send them to Vision AI for analysis. Alternatively, convert to PDF first and process the PDF.

### PowerPoint Not Available Warning

If you see "PowerPoint not available for slide rendering", install Microsoft PowerPoint or ensure it's accessible via COM automation. Without it, PPTX files still get text extraction but no figure/structure analysis via Vision AI.

### All Files Classified as "agenda"

In `--no-vision` mode, the content type heuristic is conservative. With Vision AI enabled, the LLM provides more accurate classification.

---

## Known Limitations

1. **PPTX Vision AI requires PowerPoint**: Rendering PPTX slides to images uses PowerPoint COM automation (Windows only). Without PowerPoint installed, PPTX files get text extraction + summary but no figure analysis or chemical structure detection.
2. **SmartArt**: Limited text extraction from complex SmartArt objects
3. **Embedded charts**: Chart data not directly extractable from PPTX (described via Vision AI when PowerPoint is available)
4. **Language**: Supports English and German detection. Other languages default to English processing — summaries will still be in English, but detection may not trigger the translation instruction.
5. **API dependency**: Vision AI features require active Merck Foundry access
6. **Large files**: Presentations with 50+ slides may hit token limits for summary generation
7. **PDF table quality**: pdfplumber table extraction works best for simple grid-style tables; complex nested layouts may extract poorly

---

## License

Internal tool for Merck Group - Presentation processing

**Confidential & Proprietary**
