# Setup Guide - Poster Processing Pipeline

Complete installation instructions from scratch.

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

### Core Packages

```bash
conda activate ds_env

# PDF processing
conda install -c conda-forge pdfplumber pymupdf pillow -y

# Data handling
conda install pandas openpyxl -y

# OCR support
conda install -c conda-forge tesseract pytesseract -y

# Anthropic API
pip install anthropic
```

### Verify Installation

```bash
python -c "import pdfplumber, fitz, pandas, openpyxl, anthropic; print('✓ All packages OK')"

tesseract --version
# Should output: tesseract 5.x.x
```

---

## Step 4: Configure Credentials

```powershell
# Set authentication token (PowerShell)
[Environment]::SetEnvironmentVariable("ANTHROPIC_AUTH_TOKEN", "your-token", "User")

# Set base URL (Foundry proxy)
[Environment]::SetEnvironmentVariable("ANTHROPIC_BASE_URL", "https://palantir.mcloud.merckgroup.com/language-model-service/api/proxy/anthropic", "User")

# Verify
[Environment]::GetEnvironmentVariable('ANTHROPIC_AUTH_TOKEN', 'User')
[Environment]::GetEnvironmentVariable('ANTHROPIC_BASE_URL', 'User')
```

**IMPORTANT**: Restart terminal after setting variables.

---

## Step 5: Run the Pipeline

```bash
# Navigate to project
cd PPTXToMarkdown

# Activate environment
conda activate ds_env

# Process posters
python poster_pipeline.py --sharepoint "path/to/sharepoint_folder" --metadata "path/to/metadata.xlsx"
```

---

## Troubleshooting

### Conda Not Found

Restart terminal or add to PATH manually:

```bash
set PATH=%PATH%;C:\Users\YourName\miniconda3\Scripts
```

### Package Import Errors

```bash
# Verify environment
conda activate ds_env
where python
# Should show: C:\Users\YourName\miniconda3\envs\ds_env\python.exe

# Reinstall package
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
# Check
[Environment]::GetEnvironmentVariable('ANTHROPIC_AUTH_TOKEN', 'User')

# Set again if empty
[Environment]::SetEnvironmentVariable("ANTHROPIC_AUTH_TOKEN", "your-token", "User")
[Environment]::SetEnvironmentVariable("ANTHROPIC_BASE_URL", "https://palantir.mcloud.merckgroup.com/language-model-service/api/proxy/anthropic", "User")

# MUST restart terminal
```

---

## Installation Checklist

- ✅ Install Miniconda
- ✅ Create `ds_env` environment (Python 3.11)
- ✅ Install pdfplumber, pymupdf, pandas, openpyxl, anthropic, pytesseract
- ✅ Set ANTHROPIC_AUTH_TOKEN and ANTHROPIC_BASE_URL
- ✅ Restart terminal
- ✅ Run pipeline: `python poster_pipeline.py --sharepoint "folder" --metadata "file.xlsx"`

**You're ready!**

---

## Package Reference

| Package | Purpose | Install Command |
|---------|---------|-----------------|
| pdfplumber | Native PDF text extraction | `conda install -c conda-forge pdfplumber` |
| pymupdf (fitz) | PDF manipulation, rendering | `conda install -c conda-forge pymupdf` |
| pillow | Image processing | `conda install pillow` |
| pandas | Data handling | `conda install pandas` |
| openpyxl | Excel file support | `conda install openpyxl` |
| pytesseract | OCR interface | `conda install -c conda-forge pytesseract` |
| anthropic | Claude Vision API | `pip install anthropic` |
