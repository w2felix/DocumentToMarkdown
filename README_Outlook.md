# Outlook Email Processing Pipeline

Ingests emails from a specified Outlook folder, groups them by conversation thread, and produces structured markdown. Attachments are extracted and routed to the appropriate pipeline (PDF, PPTX, DOCX, XLSX) for processing.

**No OAuth or admin rights required** — uses `win32com` to communicate with the locally running Outlook desktop app.

---

## Table of Contents

- [Quick Start](#quick-start)
- [Installation & Setup](#installation--setup)
- [Usage](#usage)
  - [Command-Line Interface](#command-line-interface)
  - [Examples](#examples)
- [Pipeline Architecture](#pipeline-architecture)
  - [Processing Stages](#processing-stages)
  - [Threading Logic](#threading-logic)
  - [Incremental Processing](#incremental-processing)
- [Output Format](#output-format)
  - [Directory Structure](#directory-structure)
  - [Markdown Structure](#markdown-structure)
- [Attachment Routing](#attachment-routing)
- [Troubleshooting](#troubleshooting)
- [Known Limitations](#known-limitations)

---

## Quick Start

```bash
# 1. Activate environment
conda activate ds_env

# 2. Process all emails in an Outlook folder
python outlook_pipeline.py --folder "Inbox/CI Reports"

# 3. Limit to 20 most recent emails
python outlook_pipeline.py --folder "Inbox" --limit 20

# 4. Skip attachment processing (email text only)
python outlook_pipeline.py --folder "Inbox/Newsletter" --no-attachments

# 5. Reprocess everything (ignore state file)
python outlook_pipeline.py --folder "Inbox/Projects" --no-skip
```

---

## Installation & Setup

See [setup.md](setup.md) for complete installation instructions.

**Prerequisites:**
- Miniconda (Python 3.11)
- Required packages: `pywin32`, `python-docx`, `pandas`, `openpyxl`
- Microsoft Outlook (desktop app, running and logged in)

**Additional package for DOCX attachment extraction:**

```bash
pip install python-docx
```

No API credentials are needed for the outlook pipeline itself (email body processing is pure text extraction). API credentials are only needed if attachments are routed to AI-powered pipelines (paper, presentation).

---

## Usage

### Command-Line Interface

```bash
python outlook_pipeline.py --folder FOLDER [options]
```

| Flag | Description | Default |
|------|-------------|---------|
| `--folder` | Outlook folder path (required) | — |
| `--output` | Output directory for markdown | `output_outlook` |
| `--no-attachments` | Skip attachment extraction and processing | Off |
| `--no-skip` | Reprocess threads even if already in state | Off |
| `--limit` | Max emails to retrieve (0 = all) | 0 |
| `--verbose` | Enable debug logging | Off |

### Examples

```bash
# Process a subfolder
python outlook_pipeline.py --folder "Inbox/ADC Updates"

# German Outlook ("Posteingang" treated same as "Inbox")
python outlook_pipeline.py --folder "Posteingang/Projekte"

# Deep subfolder
python outlook_pipeline.py --folder "Inbox/Research/Oncology/Pipeline"

# Custom output location
python outlook_pipeline.py --folder "Inbox" --output "C:/Notes/email_exports"

# Fast scan: text only, limited emails
python outlook_pipeline.py --folder "Inbox" --limit 50 --no-attachments
```

---

## Pipeline Architecture

### Processing Stages

```
Outlook COM ──► Retrieve emails ──► Group by thread ──► Check state
                                                              │
                    ┌─────────────────────────────────────────┘
                    ▼
              New threads only
                    │
        ┌───────────┼───────────────┐
        ▼           ▼               ▼
  Extract body  Extract atts   Generate thread.md
        │           │
        │      Route to pipelines
        │           │
        └───────────┴──► Save state ──► Log
```

### Threading Logic

Emails are grouped into conversation threads using a two-tier strategy:

1. **Primary**: Outlook's `ConversationID` property — reliable, survives subject changes
2. **Fallback**: Normalized subject matching — strips `Re:`, `Fwd:`, `FW:`, `AW:`, `WG:` prefixes and groups by the cleaned subject line

Within each thread, emails are sorted chronologically (oldest first).

### Incremental Processing

The pipeline maintains a JSON state file (`processed_state.json`) tracking every processed email by its Outlook `EntryID`. On subsequent runs:

- Threads where **all emails** are already in state → skipped entirely
- Threads with **new emails** → regenerated (full thread.md rewrite to maintain chronological coherence)
- Use `--no-skip` to force reprocessing of all threads

---

## Output Format

### Directory Structure

```
output_outlook/
├── processed_state.json                    # Incremental state tracking
├── processing_log.tsv                      # Run history
├── thread_project_alpha_update/            # One subfolder per thread
│   ├── thread.md                           # The conversation markdown
│   ├── attachment_report.md                # Processed PDF (via paper pipeline)
│   ├── attachment_slides.md                # Processed PPTX (via presentation pipeline)
│   └── image001.png                        # Inline image (kept raw)
├── thread_meeting_notes_q2/
│   ├── thread.md
│   ├── attachment_budget_overview.md       # Processed XLSX (tables)
│   └── budget_overview.xlsx                # Original file preserved
└── thread_newsletter_weekly/
    └── thread.md
```

### Markdown Structure

Each `thread.md` contains:

```markdown
---
title: "Re: Project Alpha Update"
thread_id: "DD0DDE70BD2D4F578122128C7CD61F50"
folder: "Inbox/CI Reports"
participants:
  - "Alice Smith"
  - "Bob Jones"
date_range: "2026-05-01 to 2026-05-15"
email_count: 4
attachments:
  - name: "report.pdf"
    processed: true
    output: "attachment_report.md"
processing_date: "2026-06-01"
---

# Re: Project Alpha Update

**Thread**: 4 emails | **Period**: 2026-05-01 to 2026-05-15 | **Folder**: Inbox/CI Reports

---

## Email 1 — 2026-05-01 09:15

**From**: Alice Smith <alice@company.com>
**To**: Bob Jones, Carol White
**Subject**: Project Alpha Update

[email body text]

---

## Email 2 — 2026-05-03 14:22

**From**: Bob Jones
**To**: Alice Smith
**Cc**: Carol White
**Subject**: Re: Project Alpha Update

[email body text]

**Attachments**: [report.pdf](attachment_report.md) | [slides.pptx](attachment_slides.md)
```

---

## Attachment Routing

| Extension | Handler | Output |
|-----------|---------|--------|
| `.pdf` | Paper Pipeline (AI) | Structured markdown with sections, summary |
| `.pptx` | Presentation Pipeline (AI) | Slide-by-slide extraction |
| `.docx` | python-docx (local) | Headings + paragraphs as markdown |
| `.xlsx` / `.xls` | pandas (local) | Markdown tables per sheet |
| `.png`, `.jpg`, `.gif` | Saved raw | Referenced in thread.md |
| Other | Saved raw | Noted as unprocessed |

**Deduplication**: If the same attachment (by filename + size) appears in multiple emails within a thread, it is extracted only once.

**Size limit**: Attachments larger than 50 MB are skipped with a warning.

---

## Troubleshooting

### "Cannot connect to Outlook"

Outlook must be running and logged in. The pipeline talks to the running instance via COM — it cannot start Outlook itself.

```
Error: Cannot connect to Outlook. Is Outlook running?
```

**Fix**: Open Outlook, wait for it to finish syncing, then retry.

### "Subfolder not found"

The folder path must match exactly (case-sensitive for subfolders). The error message lists available subfolders at the level where navigation failed.

```
Error: Subfolder 'Reports' not found. Available folders: ['CI Reports', 'Newsletters', 'Archive']
```

### Attachments fail to save

Ensure the output directory path is not excessively long (Windows MAX_PATH limit). Use `--output` with a short path if needed:

```bash
python outlook_pipeline.py --folder "Inbox" --output "C:/tmp/outlook"
```

### pywin32 not installed

```bash
pip install pywin32
```

If already installed but not working, ensure your conda environment is activated:

```bash
conda activate ds_env
python -c "import win32com.client; print('OK')"
```

---

## Known Limitations

- **Outlook must be running** — the pipeline cannot access offline .pst/.ost files directly
- **Single-threaded COM** — cannot process multiple folders in parallel (Outlook COM is apartment-threaded)
- **Sort order with --limit** — when using `--limit`, the subset of emails returned may vary between runs; without `--limit`, all emails are processed deterministically
- **HTML email conversion** — complex HTML emails (newsletters with heavy CSS) are converted to plain text via regex stripping; some formatting may be lost
- **No calendar/meeting items** — only `MailItem` objects (Class=43) are processed; meeting requests, calendar items, and contacts are skipped
- **Exchange X.500 addresses** — internal Exchange addresses are resolved to display names where possible, but some edge cases may show raw addresses
