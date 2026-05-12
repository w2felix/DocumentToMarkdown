# Patent Processing Pipeline

Automated extraction and analysis of pharmaceutical/chemical patent PDFs using selective Vision AI processing. Converts multi-page patent documents (50–300+ pages) into structured, searchable markdown with claims parsing, chemical structure extraction, biological data, and figure analysis.

**Powered by Claude Vision AI (Sonnet 4.6)** with selective page processing — only cover pages and figure/drawing pages use Vision AI; all text-heavy pages use native extraction via pdfplumber with automatic quality assessment and selective Vision AI re-extraction for garbled pages.

**Hybrid OCR** — scanned/image-only patents use Tesseract (local, free) for bulk text extraction, reserving Vision AI only for quality remediation of garbled pages. Reduces API costs by 85–95% while achieving 100% text page coverage.

---

## Table of Contents

- [Quick Start](#quick-start)
- [Installation & Setup](#installation--setup)
- [Usage](#usage)
  - [Command-Line Interface](#command-line-interface)
  - [Python API](#python-api)
  - [Processing a Single Patent](#processing-a-single-patent)
  - [Processing All Patents](#processing-all-patents)
- [Pipeline Architecture](#pipeline-architecture)
  - [Processing Stages](#processing-stages)
  - [Key Features](#key-features)
- [Output Format](#output-format)
  - [Naming Schemes](#naming-schemes)
- [Processing Log & Quality Assessment](#processing-log--quality-assessment)
- [Performance & Scalability](#performance--scalability)
- [Troubleshooting](#troubleshooting)
- [Known Limitations](#known-limitations)

---

## Quick Start

```bash
# 1. Activate environment
conda activate ds_env

# 2. Process a single patent
python patent_pipeline.py --single "test_poster/wo25045758.pdf"

# 3. Process a folder of patents
python patent_pipeline.py --input "path/to/patents/" --output output_patents

# 4. Recursive scan (subfolders)
python patent_pipeline.py --input "path/to/patents/" --recursive

# 5. Text-only mode (no Vision AI, fast)
python patent_pipeline.py --single "patent.pdf" --no-vision

# 6. Claims-only extraction (fastest)
python patent_pipeline.py --single "patent.pdf" --claims-only

# 7. Force reprocessing (default skips existing)
python patent_pipeline.py --input "patents/" --no-skip

# 8. Detailed filenames (includes applicant + title)
python patent_pipeline.py --input "patents/" --naming detailed

# 9. Force Tesseract OCR (no Vision AI for text)
python patent_pipeline.py --single "scanned_patent.pdf" --ocr-engine tesseract
```

The pipeline will process patent PDFs and generate markdown files in `output_patents/`.

---

## Installation & Setup

See [setup.md](setup.md) for complete installation instructions.

**Prerequisites:**
- Miniconda (Python 3.11)
- Required packages: `pdfplumber`, `pymupdf`, `anthropic`, `pillow`
- Tesseract OCR (optional but recommended for scanned patents): `conda install -c conda-forge tesseract pytesseract -y`
- API credentials: `ANTHROPIC_AUTH_TOKEN` and `ANTHROPIC_BASE_URL`

---

## Usage

### Command-Line Interface

```bash
python patent_pipeline.py [OPTIONS]
```

| Option | Default | Description |
|--------|---------|-------------|
| `--input` | *(none)* | Folder containing patent PDFs |
| `--output` | `output_patents` | Output directory for markdown files |
| `--single` | *(none)* | Process a single PDF file only |
| `--recursive` | `False` | Recursively search subfolders for PDF files |
| `--no-skip` | `False` | Reprocess files that already exist (default: skip existing) |
| `--no-vision` | `False` | Skip Vision AI (text-only extraction) |
| `--claims-only` | `False` | Extract only claims section (no Vision AI for figures) |
| `--max-figure-pages` | *(all)* | Limit figure pages for Vision AI analysis |
| `--render-dpi` | `200` | DPI for page rendering |
| `--budget` | `200` | Max API calls per patent (0 = unlimited) |
| `--ocr-engine` | `auto` | OCR engine: `auto`, `tesseract`, or `vision` |
| `--naming` | `default` | Filename scheme: `default`, `detailed`, or `dated` |
| `--verbose` | `False` | Enable debug logging |

### Python API

```python
from patent_pipeline import PatentPipeline
from pathlib import Path

pipeline = PatentPipeline(
    output_dir="output_patents",
    budget=200,
    ocr_engine="auto",       # "auto", "tesseract", or "vision"
    naming="default",         # "default", "detailed", or "dated"
    recursive=False
)
```

### Processing a Single Patent

```python
pipeline.process_single_patent(
    Path("test_poster/wo25045758.pdf"),
    no_vision=False,
    claims_only=False,
    max_figure_pages=None,
    skip_existing=True
)
```

### Processing All Patents

```python
# Process all, skip already-processed files (default)
pipeline.process_all_patents(Path("path/to/patents/"))

# Force reprocessing
pipeline.process_all_patents(Path("path/to/patents/"), skip_existing=False)

# Text-only mode (fast, no API calls for figures)
pipeline.process_all_patents(Path("path/to/patents/"), no_vision=True)
```

---

## Pipeline Architecture

The pipeline handles multi-page patent PDFs (50–300+ pages) using a selective Vision AI approach. Pages are classified by text density — only figure/drawing pages and the cover page are sent to Vision AI. Text-heavy pages use native PDF extraction with automatic quality scoring; pages with garbled text are selectively re-extracted via Vision AI.

### Processing Stages

#### Stage 0: PDF Characterization & Page Classification

- Opens PDF and extracts text from every page with pdfplumber (fast, no API calls)
- Classifies each page by type:
  - `COVER` — first page (bibliographic data)
  - `TEXT_HEAVY` — pages with ≥100 chars extractable text
  - `FIGURE_PAGE` — pages with <100 chars (drawings, structures, graphs)
  - `SEARCH_REPORT` — pages containing "INTERNATIONAL SEARCH REPORT"
- Typical 228-page patent: ~200 text pages, ~23 figure pages, 1 cover, 4 search report

#### Stage 0.5: Scanned Patent OCR (Tesseract Hybrid)

Triggered only when the PDF has no extractable text (image-only/scanned patents):

- **Heuristic page classification** (0 API calls): Renders 5–7 sample pages at 150 DPI, runs Tesseract for quick word-density analysis. Multi-signal classification uses word count, average words/line, and long-line ratio to separate text from figure pages.
- **Bulk OCR with Tesseract** (0 API calls): Renders all text pages at 200 DPI in batches of 20. Processes with `pytesseract --psm 6 --oem 1` (uniform block, LSTM engine). Applies the same `clean_patent_text()` pipeline as native extraction.
- **Quality check + Vision AI remediation** (0–8 API calls): After Tesseract pass, runs quality scoring on all OCR'd pages. Only pages scoring below 0.65 are escalated to Vision AI for re-extraction (typically 2–8 pages out of 300).
- **Claims detection**: Scans all OCR'd text for claims headers. Validates by checking for numbered claims ("1. A/An/The...") after the header to avoid false positives from search report text.

**Fallback**: If Tesseract is unavailable, falls back to full Vision AI OCR (existing behavior).

| OCR Engine | When Used | Cost |
|------------|-----------|------|
| `auto` | Tesseract if available, else Vision AI | 0–8 API calls for text |
| `tesseract` | Tesseract only (fails if not installed) | 0 API calls for text |
| `vision` | Vision AI only (legacy behavior) | 18–21 API calls for text |

#### Stage 1: Bibliographic Data Extraction

- Extracts patent number from filename (most reliable for WIPO documents)
- Parses cover page text with regex for PCT number, IPC classification
- **Always uses Vision AI for cover page**: WIPO multi-column layouts garble with pdfplumber
- Vision AI prompt requests organization names and personal names separately (no addresses)
- Extracts: patent number, title, applicants, inventors, filing/publication/priority dates, IPC codes

#### Stage 1.5: Text Quality Assessment & Selective Vision AI Re-extraction

- Scores every TEXT_HEAVY page on 6 quality dimensions (0.0 = garbled, 1.0 = perfect):
  - **Merged words**: Detects camelCase transitions (lowercase→Uppercase within words)
  - **Average word length**: Very long averages (>9 chars) indicate missing spaces
  - **Long word ratio**: Words >20 chars are almost always merged tokens
  - **Garbled characters**: Non-standard symbol density (°, ±, §, etc.)
  - **Alpha ratio**: Low alphabetic content indicates extraction problems
  - **Tiny word ratio**: Too few short words (1-2 chars) suggests merged text
- **Remediation decisions**:
  - Score ≥ 0.65: Good quality — keep original text
  - Score 0.30–0.65: Vision AI RAG re-extraction (page image + garbled text as context)
  - Score < 0.30: Reclassified as FIGURE_PAGE (completely garbled, likely diagram)
- Re-extracts problematic pages at 300 DPI, batches of 2 pages per API call
- Capped at 30 pages maximum to control API costs (worst-scoring pages prioritized)
- Enhanced text replaces the original pdfplumber text for downstream processing

#### Stage 2: Full Text Extraction & Section Splitting

- Concatenates text from all TEXT_HEAVY pages (including enhanced pages from Stage 1.5)
- **Text repair pipeline** (applied before section splitting):
  - **Encoding error correction**: Fixes 40+ known ligature decomposition errors (`l`→`i` substitution in PDF fonts): "molecuies" → "molecules", "payioads" → "payloads", "coniugate" → "conjugate"
  - **Spacing repair**: Splits at lowercase→Uppercase transitions, adds space after punctuation before uppercase
  - **Merged word splitting**: Safe whitelist of 30+ patterns — preposition merges (`ofthe`, `inthe`), patent claim language (`ofclaim`, `whereinthe`), scientific terms (`invitro`, `invivo`), adverb+verb patterns
  - **Stray character removal**: Removes single-letter lines from column separator artifacts
- Cleans text: strips repetitive page headers (`WO 2025/045758 PCT/EP2024/073667`) and margin line numbers (multiples of 5)
- Identifies section boundaries using regex markers:
  - `TECHNICAL FIELD` / `FIELD OF THE INVENTION`
  - `BACKGROUND` / `PRIOR ART`
  - `SUMMARY` / `SUMMARY OF THE INVENTION`
  - `CLAIMS`
  - `DETAILED DESCRIPTION`
  - `EXAMPLES` / `EXPERIMENTAL`
  - `BRIEF DESCRIPTION OF DRAWINGS`
  - `DEFINITIONS`
- Stores each section with its text content

#### Stage 3: Claims Parsing

- Splits numbered claims by pattern `(\d+)\.\s+`
- Identifies independent vs. dependent claims (`of claim X`, `according to claim X`)
- Categorizes: composition, method, use, pharmaceutical
- Builds multi-level dependency tree (walks up to root for nested dependencies)

#### Stage 4: Selective Vision AI Figure Analysis

- Renders only FIGURE_PAGE classified pages (including reclassified pages from Stage 1.5) to images at 200 DPI
- Opens PDF once per batch (not once per page) for efficiency
- Processes pages in batches of 4 per API call
- Runs up to 3 concurrent batches via ThreadPoolExecutor
- Unified prompt handles both chemical structures and graphs/charts
- **RAG context for figure analysis**:
  - Compound name context extracted from patent text (improves SMILES accuracy)
  - Figure legend text from BRIEF DESCRIPTION section (matched by figure page position)
  - Legend heuristic: estimates figure numbers from page position within figure page list
- Output per figure: type, title, axes, data series, key findings, SMILES (for structures)

#### Stage 5: Key Compound & Biological Data Extraction

- Single text-only API call on examples section
- Feeds first 12K chars (synthesis examples) AND last 8K chars (biological examples)
- Skips to actual synthesis content using regex (`Example \d`, `Synthesis of`)
- Extracts: compound names, roles, yields, MS data, purity, biological results (IC50, in vivo efficacy)
- Structured JSON response parsed into tables

#### Stage 5.5: SMILES Refinement for Chemical Structures

- Filters figures identified as `chemical_structure` with low/medium SMILES confidence
- For each structure needing refinement:
  - Re-renders the specific page at **250 DPI, max 2048px** (higher resolution for chemical detail — bond angles, stereochemistry wedges, small substituent labels)
  - Builds RAG context from:
    - Stage 4 figure description (what was initially recognized)
    - Patent text mentioning the compound (searched by compound number across all sections)
    - Figure legends from the "drawings" section
    - Core scaffold reference from claims (generic formula description)
  - Focused SMILES-only prompt requesting: SMILES, InChI, molecular formula, molecular weight
- **SMILES validation** (no API cost):
  - Balanced brackets `[]` and parentheses `()`
  - Only valid SMILES characters: `[A-Za-z0-9@+\-\[\]()=#$/\\.%:]`
  - Invalid SMILES are rejected (original retained)
- Updates figure records with refined SMILES, new confidence level, and additional fields (InChI, formula, MW)

#### Stage 6: Executive Summary, Protection Scope & Semantic Classification

- Single API call generates three outputs:
  - **Executive summary**: 3-5 sentences describing the patent's innovation and significance
  - **Protection scope**: plain-language description of what the claims cover
  - **Semantic classification** (structured JSON):
    - `therapeutic_area`: primary disease area (e.g., "oncology")
    - `mechanism_of_action`: drug mechanism (e.g., "molecular glue degrader")
    - `target_protein`: specific protein target (e.g., "Cyclin K")
    - `target_class`: broader target class (e.g., "E3 ubiquitin ligase")
    - `drug_modality`: type of therapeutic (e.g., "ADC", "small molecule")
    - `scaffold`: core chemical scaffold name
    - `comparators`: list of mentioned competitor drugs/compounds
- Quality scoring (5 dimensions, weighted):
  - Bibliographic completeness (15%)
  - Claims extraction (30%)
  - Section completeness (25%)
  - Figure analysis (15%)
  - Chemical data extraction (15%) — now includes SMILES confidence ratio

#### Stage 7: Markdown Generation & Output

- Assembles structured markdown with expanded YAML frontmatter
- **Knowledge-base structured output**:
  - Chemical structures: summary table (SMILES, confidence, formula, MW, page) + detailed descriptions with InChI
  - Key compounds: structured table (name, role, yield, MS, purity)
  - Biological results: structured table (assay type, compounds, cell lines, key result, comparison)
  - Figures index: machine-parseable summary table (figure, type, page, title, key finding)
- Outputs to `output_patents/patent_{NUMBER}.md`
- Logs quality to `output_patents/quality_log.txt`

---

## Key Features

### Hybrid OCR (Tesseract + Vision AI)
- **Scanned patent support**: Image-only PDFs (no extractable text) are handled automatically
- **Tesseract for bulk text**: Free, local OCR at 200 DPI for all text pages (~0.3–0.5s/page)
- **Vision AI only for remediation**: Pages scoring below 0.65 quality are re-extracted via API (typically 2–8 out of 300)
- **85–95% cost reduction**: Replaces 18–21 API calls for OCR with 0–8
- **100% text coverage**: All text pages are OCR'd (vs. 36–42 of 300 with Vision AI budget limits)
- **Claims always captured**: Full-text OCR ensures claims section is never missed

### Selective Vision AI
- **Cost-efficient**: Only ~10% of pages use Vision AI (figure pages + cover)
- **Cover page always via Vision AI**: WIPO multi-column layouts cannot be parsed by pdfplumber
- **Batched figure processing**: 4 pages per API call, 3 concurrent workers
- **Full coverage by default**: All figure pages analyzed (no cap unless `--max-figure-pages` set)

### Text Quality & Repair
- **Per-page quality scoring**: 6-dimension heuristic identifies garbled text automatically
- **Selective re-extraction**: Only pages scoring below 0.65 are sent to Vision AI (capped at 30)
- **Encoding error dictionary**: 40+ known ligature decomposition fixes for pharma/chemistry PDFs
- **Safe merged-word splitting**: Whitelist-based approach avoids breaking real scientific terms
- **Stray character cleanup**: Removes column separator artifacts from multi-column layouts
- **RAG-enhanced re-extraction**: Garbled text serves as content guide for Vision AI

### Chemical Structure Analysis
- **Two-pass SMILES extraction**: Initial extraction at 200 DPI + refinement at 250 DPI for low/medium confidence
- **SMILES validation**: Format checking (balanced brackets, valid character set) without chemistry libraries
- **Rich compound data**: SMILES + InChI + molecular formula + molecular weight
- **Compound context RAG**: Text-extracted compound names + scaffold descriptions fed to figure prompts
- **Structured output**: Chemical structures in parseable markdown tables with individual detail blocks

### Semantic Classification
- **Machine-readable metadata**: therapeutic_area, mechanism_of_action, target_protein, drug_modality, scaffold
- **Competitor mapping**: Extracts mentioned comparator drugs for competitive intelligence
- **YAML frontmatter**: All classification fields in structured frontmatter for downstream indexing

### Claims Intelligence
- **Dependency tree**: Multi-level resolution (claims referencing non-root parents)
- **Categorization**: composition, method, use, pharmaceutical
- **AI Protection Scope**: Plain-language summary of what the patent protects
- **Independent/dependent split**: Automatic identification and hierarchy

### Production Ready
- **Skip existing by default**: Interrupted runs resume automatically; use `--no-skip` to reprocess
- **Recursive scanning**: `--recursive` searches subfolders for PDFs
- **Configurable naming**: `--naming` controls output filename scheme (default, detailed, dated)
- **Budget control**: `--budget` caps API calls per patent (default: 200)
- **OCR engine choice**: `--ocr-engine` selects auto/tesseract/vision
- **Text-only mode**: `--no-vision` for fast extraction without API costs
- **Claims-only mode**: `--claims-only` for rapid claims extraction
- **Quality logging**: Tab-separated log with scores for all patents
- **Merck Foundry integration**: Uses internal AI proxy
- **Cached API client**: Single Anthropic client instance reused across all calls

---

## Output Format

### Generated Files

```
output_patents/
├── patent_WO25045758A1.md
├── patent_EP1234567B1.md
├── quality_log.txt
└── ...
```

### Naming Schemes

| Scheme | `--naming` | Pattern | Example |
|--------|-----------|---------|---------|
| Default | `default` | `patent_{ID}` | `patent_WO25045758A1.md` |
| Detailed | `detailed` | `patent_{ID}_{applicant}_{title}` | `patent_WO25045758A1_MABLINK_antibody_drug_conjugates_based_on.md` |
| Dated | `dated` | `{pubdate}_{ID}_{title}` | `2025-03-06_WO25045758A1_antibody_drug_conjugates_based_on.md` |

- **Default** — compact, easy to find by patent number
- **Detailed** — includes applicant (first word) and title (first 5 words) for at-a-glance identification
- **Dated** — sorts chronologically in file explorers; useful for monitoring new publications

Patent number is extracted from the PDF filename (e.g., `wo25045758.pdf` → `WO25045758A1`). Title and applicant are taken from extracted bibliographic data after processing.

### Markdown Structure

```markdown
---
patent_number: WO25045758A1
title: "ANTIBODY-DRUG CONJUGATES BASED ON MOLECULAR GLUE DEGRADERS"
applicants:
  - "MABLINK BIOSCIENCE"
  - "UNIVERSITE CLAUDE BERNARD LYON 1"
inventors:
  - "Benoît JOSEPH"
  - "Warren VIRICEL"
filing_date: 23 August 2024
publication_date: 06 March 2025
priority_date: 25 August 2023
ipc_classification: A61K47/68
pct_number: PCT/EP2024/073667
total_pages: 228
total_claims: 15
independent_claims: 7
extraction_method: native_vision_enhanced
ocr_engine: tesseract_hybrid
tesseract_pages: 287
vision_model: claude-sonnet-4-6
processing_date: 2026-05-08
figure_pages_analyzed: 23
text_pages_enhanced: 5
figure_pages_reclassified: 2
quality_overall: 9.5/10
quality_assessment: Excellent
therapeutic_area: "oncology"
mechanism_of_action: "molecular glue degrader / targeted protein degradation"
target_protein: "Cyclin K"
target_class: "E3 ubiquitin ligase (CRL4-CRBN)"
drug_modality: "ADC (antibody-drug conjugate)"
scaffold: "pyrazolo[1,5-a][1,3,5]triazine"
comparators:
  - "Kadcyla (T-DM1)"
  - "Enhertu (T-DXd)"
compound_count: 10
biological_assay_count: 10
chemical_structures_count: 5
smiles_high_confidence: 3
---

# Patent Title

## Executive Summary

AI-generated 3-5 sentence summary of the patent's innovation,
key compounds, mechanism of action, and therapeutic applications...

---

## Bibliographic Data

| Field | Value |
|-------|-------|
| Patent Number | WO 2025/045758 A1 |
| Applicants | MABLINK BIOSCIENCE; ... |
| Inventors | Benoît JOSEPH; ... |
| Filing Date | 23 August 2024 |
| IPC Classification | A61K47/68 |

---

## Technical Field

Extracted section text...

---

## Background

Extracted section text...

---

## Summary of Invention

Extracted section text...

---

## Claims

### Protection Scope

Plain-language AI summary of what the patent protects...

### Independent Claims

Full text of independent claims with category labels...

### Claim Dependency Tree

- **Claim 1** (composition): Main claim text...
  - Claim 5: Dependent claim...
    - Claim 8: Sub-dependent claim...
- **Claim 12** (method): Method claim text...

---

## Key Chemical Structures

| ID | Label | SMILES | Confidence | Formula | MW | Page |
|----|-------|--------|------------|---------|---:|------|
| 2384 | Compound 2384 | `Brc1cn2nc...` | high | C23H28BrN7O | 498.4 | 156 |
| 2478 | Compound 2478 | `CC(=O)N...` | medium | C25H30N6O2 | 446.5 | 158 |

### Compound 2384
**Description**: Pyrazolotriazine derivative with bromine substituent...
**SMILES** (confidence: high): `Brc1cn2nc...`
**InChI**: `InChI=1S/C23H28BrN7O/c...`
**Formula**: C23H28BrN7O | **MW**: 498.4
**Key Data**: Core payload for ADC conjugation...

---

## Key Compounds

| Compound | Role | Yield | MS [M+H]+ | Purity |
|----------|------|------:|-----------|--------|
| 2384 | payload | 65% | 493.2 | 98.2% |
| TRA-VA-2384 | ADC | — | — | — |

---

## Biological Results

| Assay Type | Compounds | Cell Lines/Models | Key Result | Comparison |
|-----------|-----------|-------------------|------------|------------|
| In vitro cytotoxicity | TRA-VA-2384 | BT-474, NCI-N87 | IC50 = 0.3 nM | Kadcyla: 3.0 nM (10x) |
| In vivo efficacy | TRA-VA-2384 | BT-474 xenograft (SCID) | Complete regression at 10 mg/kg | — |

---

## Figures & Drawings

### FIG. 44
**Type**: Dose Response Curve
**X-axis**: Drug (nM) | **Y-axis**: Cell survival (%)
**Data Series**: 2384, 2420
**Key Findings**: IC50 values, fold-differences...

---

## Figures Index

| Figure | Type | Page | Title | Key Finding |
|--------|------|-----:|-------|-------------|
| FIG. 1 | chemical structure | 135 | Compound 2384 | Core payload structure |
| FIG. 44 | dose response curve | 156 | BT-474 cytotoxicity | IC50 = 0.3 nM |
| FIG. 45 | kaplan meier | 157 | In vivo survival | Complete regression |

---

## Quality Assessment

**Overall Quality**: 9.5/10 - Excellent
**Component Scores**:
- Bibliographic Data: 10.0/10
- Claims Extraction: 10/10
- Section Completeness: 8.0/10
- Figure Analysis: 10/10
- Chemical Data: 10/10

---

## Processing Metadata

- **Extraction Method**: native_vision_enhanced
- **Vision Model**: claude-sonnet-4-6
- **Total Pages**: 228
- **Figure Pages Analyzed**: 23
- **Text Pages Vision-Enhanced**: 5
- **Text Pages Reclassified as Figure**: 2
- **Total Figures Described**: 27
- **Text Length**: 249,027 characters
- **Processing Date**: 2026-05-08
```

---

## Processing Log & Quality Assessment

### Processing Log

After processing, check `output_patents/quality_log.txt`:

```
TIMESTAMP            PATENT_NUMBER   QUALITY  ASSESSMENT  STATUS  FILENAME
2026-05-08 16:02:39  WO25045758A1    9.5      Excellent   SAVED   wo25045758.pdf
```

### Quality Dimensions

| Dimension | Weight | What it Measures | Score 10 = |
|-----------|--------|------------------|------------|
| Bibliographic Data | 0.15 | Title, applicants, inventors, dates, IPC | All fields present |
| Claims Extraction | 0.30 | Claim count, independent/dependent split, dependency tree | 10+ claims with tree |
| Section Completeness | 0.25 | Core sections identified (Tech Field, Background, Summary, Claims, Examples) | All sections found |
| Figure Analysis | 0.15 | Figures described, key findings, data series | 20+ figures with findings |
| Chemical Data | 0.15 | Key compounds + biological results + SMILES confidence ratio | 5+ compounds, 5+ bio results, high-confidence SMILES |

### SMILES Confidence Scoring

The chemical data quality score now incorporates SMILES accuracy:
- Compounds present: up to 4 points
- Biological results present: up to 4 points
- SMILES confidence bonus: up to 2 points (proportional to % of chemical structures with "high" confidence)

### Quality Thresholds

- **Excellent** (8-10): Complete extraction, rich figure analysis, full claims tree → **Markdown generated**
- **Good** (5.5-8): High quality, most sections captured, good data → **Markdown generated**
- **Fair** (4-5.5): Partial extraction, missing sections → **Skipped, logged for review**
- **Poor** (<4): Major extraction failures → **Skipped, manual review required**

### Status Values

- `SAVED`: Successfully processed and markdown generated
- `SKIPPED`: Quality below threshold (FAIR/POOR)
- `FAILED`: Processing error (PDF conversion, API failure)

---

## Performance & Scalability

### Processing Time

- **Single patent** (228 pages, 23 figure pages): ~2-3 minutes
- **Scanned patent** (300 pages, Tesseract OCR): ~3-5 minutes (90-150s OCR + figure analysis)
- **Text-only mode** (`--no-vision`): ~15 seconds
- **Claims-only mode** (`--claims-only`): ~5 seconds

### API Usage

**Text-native patents** (extractable text):
- 1 API call for cover page Vision AI (~2K tokens)
- 0–15 API calls for text page re-extraction (Stage 1.5: garbled text only)
- 6–25 API calls for figure batches (4 pages/batch, 3 concurrent, ~8K tokens each)
- 1 API call for key data extraction (~4K tokens)
- 1–5 API calls for SMILES refinement (low/medium confidence structures)
- 1 API call for executive summary + protection scope + classification (~3K tokens)

**Scanned patents** (Tesseract hybrid):
- 0 API calls for text extraction (all handled by Tesseract)
- 0–8 API calls for quality remediation (only if Tesseract output is garbled)
- 6–25 API calls for figure batches (same as above)
- 1 API call for key data + 1 API call for summaries

**Total per patent**: ~10-48 API calls (text-native), ~8-35 API calls (scanned with Tesseract)

**Example**: 228-page text-native patent → ~18 API calls, ~3 minutes
**Example**: 300-page scanned patent → ~12 API calls, ~4 minutes (Tesseract handles all text)

### DPI Strategy

| Content Type | DPI | Max Dimension | Purpose |
|-------------|-----|---------------|---------|
| Tesseract OCR (scanned pages) | 200 | — | Bulk text extraction, free local processing |
| Heuristic page sampling | 150 | — | Quick classification of text vs. figure pages |
| Figure pages (charts/graphs) | 200 | 1568px | Large text, simple geometry — standard resolution sufficient |
| Text page re-extraction | 300 | 2048px | Needs fine detail for small text in patent body |
| Vision AI text enhancement | 250 | 1568px | Remediation of garbled Tesseract output |
| Chemical structure refinement | 250 | 2048px | Bond angles, stereochemistry wedges, small substituent labels |

### Optimization Features

- **Tesseract bulk OCR**: Scanned patents processed locally for free (0 API calls for text)
- **Selective Vision AI**: Only ~10% of pages sent to API (figures + cover)
- **Quality-gated text enhancement**: Only garbled pages re-extracted (not all text pages)
- **Budget enforcement**: All API-calling stages tracked against per-patent budget cap
- **Cached API client**: Single Anthropic client instance reused across all calls
- **Batch figure processing**: 4 pages per API call (reduces call count 4x)
- **Concurrent batches**: Up to 3 simultaneous API calls (~2x speedup)
- **Single PDF open per batch**: Opens fitz document once per batch, not per page
- **Pre-encoded images**: All batch pages rendered and encoded before API call
- **Memory-managed OCR**: Tesseract processes pages in batches of 20, freeing images after each batch
- **Text repair pipeline**: Encoding fixes + spacing repair handles ~80% of issues without API calls
- **Filename-based patent number**: Avoids garbled OCR on cover page
- **Combined summary call**: Executive summary + protection scope + classification in one API call
- **SMILES validation**: Format checking rejects invalid structures without API retry

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
conda install -c conda-forge pdfplumber pymupdf pillow -y
pip install anthropic

# Verify installations
python -c "import pdfplumber, anthropic, fitz; print('All packages OK')"
```

### Tesseract Not Found

If you see "Tesseract not available — will use Vision AI for OCR":
```bash
# Install Tesseract via conda
conda install -c conda-forge tesseract pytesseract -y

# Verify installation
python -c "import pytesseract; print(pytesseract.get_tesseract_version())"
```

The pipeline falls back to Vision AI OCR automatically if Tesseract is unavailable. Use `--ocr-engine vision` to skip Tesseract intentionally.

### Low Quality Scores

If patents score FAIR or POOR:
- Check PDF quality — scanned patents require Tesseract (install it or use `--ocr-engine vision`)
- Verify the patent has standard section headers (WIPO/EPO format)
- Try `--verbose` to see which stages are failing
- Check `output_patents/quality_log.txt` for specific scores
- Non-English patents may have reduced section detection

### Cover Page Extraction Issues

If applicants/inventors are missing or garbled:
- Ensure Vision AI is enabled (not using `--no-vision`)
- Check API connectivity — cover page extraction requires one API call
- Non-WIPO formats (USPTO, EPO) may have different cover layouts

### Text Quality Issues (Merged Words, Encoding Errors)

If the output text has merged words or encoding errors:
- Check `text_pages_enhanced` in the YAML frontmatter — if 0, quality scoring found no issues
- The encoding fix dictionary covers common ligature decomposition errors (`l`→`i`)
- The merged-word whitelist handles ~30 known patent text patterns
- For patents with unusual fonts, the pipeline relies on Stage 1.5 Vision AI re-extraction
- Very large patents (300+ pages) may hit the 30-page re-extraction cap

### No Biological Results

If biological results are 0 despite the patent having examples:
- Check that the patent has an "Examples" or "Experimental" section header
- The pipeline looks for `Example \d` or `Synthesis of` patterns to find real data
- Very short examples sections (<1K chars) may not trigger extraction

### Low SMILES Confidence

If all SMILES are "low" or "medium" confidence:
- Complex fused ring systems and stereochemistry are inherently difficult for Vision AI
- Stage 5.5 refinement at 250 DPI should improve some structures
- Validate critical SMILES externally with RDKit if accuracy is required
- SMILES that fail format validation are rejected and marked accordingly

---

## Known Limitations

1. **WIPO/EPO format optimized**: Section detection regex patterns are tuned for international patent formats; USPTO-specific formatting may need adaptation
2. **pdfplumber text quality**: Multi-column layouts on cover pages always garble — Vision AI is mandatory for applicants/inventors
3. **Chemical SMILES accuracy**: Vision AI-generated SMILES are approximate; complex structures with stereochemistry may have errors despite two-pass refinement
4. **Figure page threshold**: Pages with <100 chars are classified as figures — borderline pages (chemical formula tables) may be misclassified
5. **Language**: Optimized for English-language patents
6. **Tesseract OCR quality**: Scanned patents with unusual fonts, low scan resolution, or complex layouts may produce lower-quality text than native PDF extraction
7. **API dependency**: Requires active Merck Foundry access for Vision AI features (figures, cover page, summaries)
8. **Large patents**: Patents with 100+ figure pages may take 5-10 minutes and 48+ API calls
9. **Text re-extraction cap**: Maximum 30 text pages can be Vision AI enhanced per patent (worst-scoring prioritized)
10. **Encoding fix coverage**: The ligature decomposition dictionary covers common pharma/chemistry terms but may miss rare vocabulary

---

## License

Internal tool for Merck Group - Patent document processing

**Confidential & Proprietary**
