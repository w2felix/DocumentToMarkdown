# Outlook Email Processing Pipeline

← [Back to main README](README.md)

Ingests emails from a specified Outlook folder, groups them by conversation thread, and produces structured markdown. Attachments are extracted and routed to the appropriate pipeline (PDF, PPTX, DOCX, XLSX) for processing. Detected email signatures are parsed into a persistent, change-tracked contact book (`people/`).

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
- Required packages: `pywin32`, `python-docx`, `pandas`, `openpyxl`, `pyyaml`
- Microsoft Outlook (desktop app, running and logged in)

```bash
pip install python-docx
pip install pyyaml   # for reading/writing contact YAML files in people/
```

**API credentials:** not needed for email body processing (pure text extraction). Required if attachments are routed to AI-powered sub-pipelines (paper, presentation). See [setup.md](setup.md) for credential setup.

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
        ┌───────────┼─────────────────────┐
        ▼           ▼                     ▼
  Extract body  Extract atts        Generate thread.md
  (strip banner,    │               (strip signatures
   quoted chains)   │                from body text)
                Route to                  │
                pipelines            Parse signatures
                    │               ──► Update people/
                    │
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

The output directory mirrors the Outlook folder hierarchy, so threads from different folders stay grouped and easy to navigate. The `people/` directory is global and shared across all folder runs.

```
output_outlook/
├── processed_state.json                    # Incremental state tracking (shared)
├── processing_log.tsv                      # Run history (shared)
├── people/                                 # Global contact book (shared across folders)
│   ├── carsten_schweer.yml                 # One YAML file per person
│   ├── alice_smith.yml
│   └── ...
└── Posteingang/                            # Mirrors Outlook folder hierarchy
    └── Trans. Projects/
        └── CARIS JRC/
            ├── thread_caris_fragen.../     # One subfolder per thread
            │   └── thread.md
            └── thread_jrc_17.../
                ├── thread.md
                ├── attachment_report.md    # Processed PDF (via paper pipeline)
                └── slides.pptx             # Original file preserved

# Another folder run populates its own branch:
# output_outlook/Inbox/CI Reports/thread_.../
```

**Contact files** (`people/*.yml`) accumulate over time. Each re-run adds to existing contacts and appends change history when fields differ:

```yaml
name: Carsten Schweer
email: carsten.schweer@merckgroup.com
phone: "+49 6151 72 26695"
mobile: "+4915114544754"
title: Alliance Management
department: Biopharma | Global Business Development and Alliance Management
organization: Merck
location: "Frankfurter Str. 250, 64293 Darmstadt, Germany"
website: merckgroup.com
first_seen: "2026-05-07"
last_seen: "2026-06-01"
changes:
  - date: "2026-09-01"
    field: title
    old: Alliance Management
    new: Senior Director Alliance Management
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

## Body Cleaning

Before the email body is written to `thread.md`, several cleanup passes run automatically:

| Pass | What it removes |
|------|-----------------|
| **Security banners** | Proofpoint/Merck injected `[EXTERNAL]` warning blocks (`ZjQcmQRYFpfpt*` markers) |
| **Quoted chains** | `From: … Sent: … To: …` reply-history blocks that repeat earlier emails |
| **Signatures** | Everything from the sign-off phrase onwards is stripped from the body and parsed into `people/` instead |

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
