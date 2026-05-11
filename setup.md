# Setup Guide

Complete installation instructions for the DocumentToMarkdown pipelines (poster, patent, talk).

---

## Step 1: Install Miniconda

### Download

1. Go to https://docs.conda.io/en/latest/miniconda.html
2. Download **Miniconda3 Windows 64-bit**
3. Run installer

### Installation

- Install for: Just Me (recommended)
- Location: Default (`C:\Users\YourName\miniconda3`)
- Add to PATH: ☑ (optional but recommended)

### Verify

```bash
conda --version
# Should output: conda 23.x.x or newer
```

---

## Step 2: Create Conda Environment

```bash
# Create environment with Python 3.11
conda create -n ds_env python=3.11 -y

# Activate
conda activate ds_env

# Verify
python --version
# Should output: Python 3.11.x
```

---

## Step 3: Install Packages

```bash
conda activate ds_env

# PDF processing
conda install -c conda-forge pdfplumber pymupdf pillow -y

# Data handling
conda install pandas openpyxl -y

# OCR support (used by poster and talk pipelines)
conda install -c conda-forge tesseract pytesseract -y

# Anthropic API (Claude Vision)
pip install anthropic
```

### Verify Installation

```bash
python -c "import pdfplumber, fitz, pandas, openpyxl, anthropic; print('All packages OK')"

tesseract --version
# Should output: tesseract 5.x.x
```

---

## Step 4: Configure API Credentials

All three pipelines use the Anthropic API (Claude) for vision analysis. Set credentials as Windows User environment variables:

```powershell
# Set authentication token (PowerShell)
[Environment]::SetEnvironmentVariable("ANTHROPIC_AUTH_TOKEN", "your-token", "User")

# Set base URL (e.g. corporate proxy)
[Environment]::SetEnvironmentVariable("ANTHROPIC_BASE_URL", "https://your-proxy-url/api/proxy/anthropic", "User")

# Verify
[Environment]::GetEnvironmentVariable('ANTHROPIC_AUTH_TOKEN', 'User')
[Environment]::GetEnvironmentVariable('ANTHROPIC_BASE_URL', 'User')
```

**IMPORTANT**: Restart your terminal after setting variables.

Alternatively, if using the Anthropic API directly (no proxy), set `ANTHROPIC_API_KEY` instead:

```powershell
[Environment]::SetEnvironmentVariable("ANTHROPIC_API_KEY", "sk-ant-...", "User")
```

---

## Step 5: Run a Pipeline

```bash
# Navigate to project
cd DocumentToMarkdown

# Activate environment
conda activate ds_env

# Process posters
python poster_pipeline.py --sharepoint "path/to/poster_pdfs" --metadata "abstracts.xlsx"

# Process patents
python patent_pipeline.py --input "path/to/patent_pdfs"

# Process talks
python talk_pipeline.py --talks "path/to/talk_pdfs" --metadata "abstracts.xlsx"
```

See the [README](README.md) for full CLI options, or the pipeline-specific guides for architecture details:
- [Poster Pipeline](README_Posters.md)
- [Patent Pipeline](README_Patents.md)
- [Talk Pipeline](README_Talks.md)

---

## Troubleshooting

### Conda Not Found

Restart terminal or add to PATH manually:

```bash
set PATH=%PATH%;C:\Users\YourName\miniconda3\Scripts
```

### Package Import Errors

```bash
# Verify environment is active
conda activate ds_env
where python
# Should show: C:\Users\YourName\miniconda3\envs\ds_env\python.exe

# Reinstall a package
conda install -c conda-forge <package-name> -y
```

### Tesseract Not Found

```powershell
# Add to PATH
[Environment]::SetEnvironmentVariable("Path", $env:Path + ";C:\Program Files\Tesseract-OCR", "User")

# Restart terminal
```

### API Credentials Not Working

```powershell
# Check current values
[Environment]::GetEnvironmentVariable('ANTHROPIC_AUTH_TOKEN', 'User')

# Re-set if empty
[Environment]::SetEnvironmentVariable("ANTHROPIC_AUTH_TOKEN", "your-token", "User")
[Environment]::SetEnvironmentVariable("ANTHROPIC_BASE_URL", "https://your-proxy-url/api/proxy/anthropic", "User")

# MUST restart terminal after changes
```

---

## Installation Checklist

- [ ] Install Miniconda
- [ ] Create `ds_env` environment (Python 3.11)
- [ ] Install pdfplumber, pymupdf, pandas, openpyxl, anthropic, pytesseract
- [ ] Set `ANTHROPIC_AUTH_TOKEN` and `ANTHROPIC_BASE_URL` (or `ANTHROPIC_API_KEY`)
- [ ] Restart terminal
- [ ] Run a pipeline to verify everything works

---

## Package Reference

| Package | Purpose | Install Command |
|---------|---------|-----------------|
| pdfplumber | Native PDF text extraction | `conda install -c conda-forge pdfplumber` |
| pymupdf (fitz) | PDF rendering to images | `conda install -c conda-forge pymupdf` |
| pillow | Image processing | `conda install pillow` |
| pandas | Metadata handling | `conda install pandas` |
| openpyxl | Excel file support | `conda install openpyxl` |
| pytesseract | OCR interface (Tesseract) | `conda install -c conda-forge pytesseract` |
| anthropic | Claude Vision API client | `pip install anthropic` |
