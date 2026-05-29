"""
Scientific Poster Processing Pipeline with Vision Analysis
Converts biomedical scientific posters (PDFs) to structured markdown with AI-powered figure analysis

Enhanced with:
- Mandatory Vision AI for figure analysis and text correction
- Two-pass OCR (Tesseract + Vision AI)
- Excel metadata validation and tracking
- Figure-to-text position matching
- Multi-page poster support
- Quality validation and scoring
"""

import os
import re
import json
import base64
from pathlib import Path
from typing import List, Dict, Tuple, Optional, Any
from io import BytesIO
import logging
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from difflib import SequenceMatcher

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Load credentials and shared auth
from pipeline_security import validate_path, validate_output_path, check_excel_safe, sanitize_filename
from auth import get_anthropic_client as _get_shared_client


class PosterPipeline:
    """Main pipeline for processing scientific posters with vision analysis

    Vision AI is now MANDATORY for all processing (figure analysis + text correction)
    """

    # Class-level constants
    VISION_MODEL = "claude-sonnet-4-6"
    MAX_PARALLEL_FIGURE_WORKERS = 5
    MIN_TEXT_LENGTH_THRESHOLD = 100
    MIN_SECTION_LENGTH = 50
    QUALITY_EXCELLENT_THRESHOLD = 8.0
    QUALITY_GOOD_THRESHOLD = 5.5  # Lowered from 6.0 to recover marginal posters
    QUALITY_FAIR_THRESHOLD = 4.0
    MAX_PDF_SIZE_MB = 200
    MAX_TEMPLATE_PENALTY = 2.0
    TEMPLATE_PENALTY_PER_ARTIFACT = 0.5
    CAPTION_PENALTY_PER_CONTAMINATION = 2.0
    CONSISTENCY_PENALTY_PER_MISMATCH = 0.5
    MAX_CONSISTENCY_PENALTY = 2.0  # Cap consistency penalty to prevent excessive point loss
    DEFAULT_VISION_DPI = 300  # Optimized: 300 DPI sufficient after resize to 2048px
    STRUCTURE_VISION_DPI = 300
    DEFAULT_OCR_DPI = 200

    NAMING_SCHEMES = {
        'default': 'poster_{poster_num}',
        'standardized': 'Poster_{author}_{conference}_{year}_{title_slug}',
    }

    DEFAULT_SHEET = 'Full_Program_Copy'

    DEFAULT_COLUMN_MAP = {
        'poster_number': ['Presentation Number', 'Poster Number', 'Abstract Number', 'Session Number'],
        'title': 'Presentation Title',
        'authors': 'authors',
        'institution': 'institution',
        'session_number': 'Session Number',
        'session_title': 'Session Title',
        'session_type': 'Session Type Name',
        'day': 'Day',
        'session_start': 'Session Start',
        'session_end': 'Session End',
        'location': 'Location',
        'covered_by': 'Covered by',
        'interested': 'Interested Colleagues',
    }

    def __init__(self, sharepoint_folder: str, metadata_excel: str = None, output_dir: str = "output",
                 recursive: bool = False, naming: str = 'default', conference: str = None, year: str = None,
                 sheet: str = None, column_overrides: dict = None):
        self.sharepoint_folder = Path(sharepoint_folder)
        self.metadata_excel = Path(metadata_excel) if metadata_excel else None
        self.output_dir = Path(output_dir)
        self.recursive = recursive
        self.naming = naming
        self.conference = conference
        self.year = year

        self.sheet = sheet or self.DEFAULT_SHEET
        self.column_map = dict(self.DEFAULT_COLUMN_MAP)
        if column_overrides:
            self.column_map.update(column_overrides)

        # Create output directory
        self.output_dir.mkdir(exist_ok=True)

        # Initialize quality log file
        self.quality_log_path = self.output_dir / "quality_log.txt"
        self.failed_posters = []  # Track posters that fail quality threshold
        self.unrecognized_posters = []  # Track posters with unresolvable filenames

        # Cached client and existing files
        self._client = None
        self._existing_files = set(f.stem for f in self.output_dir.glob("*.md"))

        # Setup TESSDATA_PREFIX once
        conda_prefix = os.environ.get('CONDA_PREFIX')
        if conda_prefix and not os.environ.get('TESSDATA_PREFIX'):
            tessdata_dir = os.path.join(conda_prefix, 'Library', 'share', 'tessdata')
            if os.path.exists(tessdata_dir):
                os.environ['TESSDATA_PREFIX'] = tessdata_dir

        # Validate Excel metadata file (or proceed without it)
        if self.metadata_excel:
            self._validate_metadata_file()
        else:
            self._metadata_df = None
            logger.info("Unfortunately, no additional metadata is available for the poster extraction. "
                        "Poster extraction proceeds without additional information.")

    def list_pdf_files(self) -> List[Path]:
        """List all PDF files in the SharePoint folder"""
        try:
            if self.recursive:
                pdf_files = list(self.sharepoint_folder.rglob("*.pdf"))
                logger.info(f"Found {len(pdf_files)} PDF files recursively in {self.sharepoint_folder}")
            else:
                pdf_files = list(self.sharepoint_folder.glob("*.pdf"))
                logger.info(f"Found {len(pdf_files)} PDF files in {self.sharepoint_folder}")
            return pdf_files
        except Exception as e:
            logger.error(f"Error listing PDF files: {e}")
            return []

    def _validate_metadata_file(self):
        """Validate that Excel metadata file exists and load it (cached)."""
        if not self.metadata_excel.exists():
            raise FileNotFoundError(f"Excel metadata file not found: {self.metadata_excel}")

        # File size safety check
        if not check_excel_safe(self.metadata_excel):
            raise ValueError(f"Excel file failed safety check (too large): {self.metadata_excel}")

        try:
            import pandas as pd

            xl = pd.ExcelFile(self.metadata_excel)
            if self.sheet not in xl.sheet_names:
                raise ValueError(
                    f"Sheet '{self.sheet}' not found in {self.metadata_excel.name}. "
                    f"Available sheets: {', '.join(xl.sheet_names)}"
                )

            self._metadata_df = pd.read_excel(self.metadata_excel, sheet_name=self.sheet)

            if self._metadata_df.empty:
                logger.warning("Excel metadata file is empty")
                self._poster_lookup = {}
            else:
                logger.info(f"✓ Excel metadata loaded: {len(self._metadata_df)} rows from sheet '{self.sheet}'")
                available = set(self._metadata_df.columns)
                for role, col_name in self.column_map.items():
                    names = col_name if isinstance(col_name, list) else [col_name]
                    found = [n for n in names if n in available]
                    missing = [n for n in names if n not in available]
                    if missing and not found:
                        logger.info(f"  Column '{role}': configured as {missing} — not in spreadsheet (will be skipped)")

                # Build O(1) poster number lookup
                self._poster_lookup = {}
                possible_columns = self.column_map['poster_number']
                if isinstance(possible_columns, str):
                    possible_columns = [possible_columns]
                for col_name in possible_columns:
                    if col_name in self._metadata_df.columns:
                        for idx, row in self._metadata_df.iterrows():
                            poster_num = str(row[col_name]).strip()
                            if poster_num and poster_num != 'nan' and poster_num not in self._poster_lookup:
                                self._poster_lookup[poster_num] = row.to_dict()
                logger.info(f"  Built poster lookup index: {len(self._poster_lookup)} entries")

        except ImportError:
            raise ImportError("pandas not installed. Install with: conda install pandas openpyxl")
        except ValueError:
            raise
        except Exception as e:
            raise ValueError(f"Error validating Excel metadata: {e}")

    def load_metadata(self) -> Optional[object]:
        """Return cached metadata DataFrame (loaded during __init__)."""
        return getattr(self, '_metadata_df', None)

    def extract_poster_number(self, pdf_path: Path) -> Optional[str]:
        """Extract poster number from filename (extracts leading number/code before any separator)

        Supports various filename formats:
        - 160.pdf -> "160"
        - LB197.pdf -> "LB197"
        - CT070.pdf -> "CT070"
        - 7733.pdf -> "7733"
        - 348 - Title Here.pdf -> "348"
        - LB197 - Some Title.pdf -> "LB197"
        - Title-based filenames -> matched against metadata
        """
        if pdf_path.suffix.lower() != '.pdf':
            return None

        filename = pdf_path.stem  # Get filename without extension

        # Extract poster number from start of filename.
        # Require at least 2 digits to avoid matching gene names (B7, CD3) and short drug codes.
        match = re.match(r'^([A-Z]{0,4}\d{2,}[A-Z]*)', filename, re.IGNORECASE)
        if match:
            candidate = match.group(1)
            if self._verify_poster_number_in_metadata(candidate):
                return candidate
            logger.info(f"Regex matched '{candidate}' but not found in metadata, trying title match...")

        # Fallback: try title-based matching against metadata
        resolved = self._resolve_poster_number_from_title(filename)
        if resolved:
            logger.info(f"Resolved poster number '{resolved}' from title-based filename: {filename[:60]}...")
            return resolved

        logger.warning(f"Could not resolve poster number from filename: {filename[:80]}...")
        return None

    def _verify_poster_number_in_metadata(self, candidate: str) -> bool:
        """Check if a candidate poster number exists in the metadata."""
        if self._metadata_df is None or self._metadata_df.empty:
            return True
        poster_num_cols = self.column_map['poster_number']
        if isinstance(poster_num_cols, str):
            poster_num_cols = [poster_num_cols]
        for col_name in poster_num_cols:
            if col_name in self._metadata_df.columns:
                if any(self._metadata_df[col_name].astype(str).str.strip() == str(candidate)):
                    return True
        return False

    def _build_title_lookup_cache(self) -> List[Tuple[str, str]]:
        """Build cached lookup of (cleaned_title, presentation_number) from metadata."""
        if self._metadata_df is None or self._metadata_df.empty:
            return []

        title_col = self.column_map['title']
        poster_num_cols = self.column_map['poster_number']
        if isinstance(poster_num_cols, str):
            poster_num_cols = [poster_num_cols]

        html_tag_re = re.compile(r'<[^>]+>')
        cache = []
        for _, row in self._metadata_df.iterrows():
            title = str(row.get(title_col, ''))
            pres_num = ''
            for col in poster_num_cols:
                val = str(row.get(col, ''))
                if val and val != 'nan':
                    pres_num = val
                    break
            if not title or title == 'nan' or not pres_num or pres_num == 'nan':
                continue
            cleaned = html_tag_re.sub('', title).strip()
            if cleaned:
                cache.append((cleaned, pres_num.strip()))

        logger.info(f"Built title lookup cache: {len(cache)} entries")
        return cache

    def _resolve_poster_number_from_title(self, filename_stem: str) -> Optional[str]:
        """Resolve poster number by matching filename against metadata titles.

        Used as fallback when standard regex-based extraction fails (title-based filenames).
        """
        if self._metadata_df is None or self._metadata_df.empty:
            return None

        if not hasattr(self, '_title_lookup_cache'):
            self._title_lookup_cache = self._build_title_lookup_cache()

        if not self._title_lookup_cache:
            return None

        # Phase A: Check trailing " - NUMBER" pattern
        trailing_match = re.search(r'\s+-\s+(\d+)$', filename_stem)
        if trailing_match:
            candidate_num = trailing_match.group(1)
            for _, pres_num in self._title_lookup_cache:
                if pres_num == candidate_num:
                    logger.info(f"  Matched via trailing number: {candidate_num}")
                    return candidate_num

        # Prepare filename for title matching
        if trailing_match:
            search_title = filename_stem[:trailing_match.start()].strip()
        else:
            search_title = filename_stem.strip()

        search_title_lower = search_title.lower()

        # Phase B: Exact substring match
        substring_matches = []
        for cleaned_title, pres_num in self._title_lookup_cache:
            title_lower = cleaned_title.lower()
            if search_title_lower in title_lower or title_lower in search_title_lower:
                substring_matches.append((cleaned_title, pres_num))

        if len(substring_matches) == 1:
            logger.info(f"  Matched via substring: {substring_matches[0][1]}")
            return substring_matches[0][1]
        elif len(substring_matches) > 1:
            logger.warning(f"  Multiple substring matches ({len(substring_matches)}) for: {search_title[:50]}...")

        # Phase C: Fuzzy match
        if len(search_title) < 10:
            return None

        best_ratio = 0.0
        best_match = None
        for cleaned_title, pres_num in self._title_lookup_cache:
            ratio = SequenceMatcher(None, search_title_lower, cleaned_title.lower()).ratio()
            if ratio > best_ratio:
                best_ratio = ratio
                best_match = (cleaned_title, pres_num)

        if best_ratio >= 0.75 and best_match:
            logger.info(f"  Fuzzy matched (ratio={best_ratio:.3f}): {best_match[1]}")
            return best_match[1]

        logger.warning(f"  No title match found (best ratio={best_ratio:.3f})")
        return None

    def check_if_processed(self, poster_num: str) -> bool:
        """Check if poster has already been processed (uses cached filenames)."""
        suffix = f"_{poster_num}"
        return any(f.endswith(suffix) for f in self._existing_files)

    def extract_text_native(self, pdf_path: Path) -> Tuple[Optional[str], str]:
        """Extract text using native PDF text extraction"""
        try:
            import pdfplumber
            text = ""
            with pdfplumber.open(pdf_path) as pdf:
                logger.info(f"PDF has {len(pdf.pages)} page(s)")
                for page_num, page in enumerate(pdf.pages):
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + "\n\n"
                        logger.debug(f"Page {page_num+1}: extracted {len(page_text)} chars")

            logger.info(f"Total extracted: {len(text)} characters, {len(text.split())} words")

            if self.is_readable_text(text):
                logger.info(f"✓ Native text extraction successful")
                return text, "native"
            else:
                logger.warning(f"✗ Text quality check failed - may need OCR")
                return text, "native_low_quality"

        except ImportError:
            logger.error("pdfplumber not installed")
            return None, "error"
        except Exception as e:
            logger.error(f"Error extracting text: {e}")
            return None, "error"

    def extract_text_ocr(self, pdf_path: Path, dpi: int = 200) -> Tuple[Optional[str], str]:
        """Extract text using OCR (fallback for scanned PDFs)"""
        try:
            from pdf2image import convert_from_path
            import pytesseract
            from PIL import Image

            Image.MAX_IMAGE_PIXELS = 300_000_000
            logger.info(f"Running OCR at {dpi} DPI...")

            images = convert_from_path(str(pdf_path), dpi=dpi)
            logger.info(f"Converted to {len(images)} image(s)")

            text = ""
            for i, image in enumerate(images):
                width, height = image.size
                megapixels = (width * height) / 1_000_000
                logger.info(f"Processing page {i+1}/{len(images)} ({width}x{height}, {megapixels:.1f}MP)")

                page_text = pytesseract.image_to_string(image)
                text += page_text + "\n\n"
                logger.info(f"  Extracted {len(page_text)} characters from page {i+1}")

            logger.info(f"OCR extraction complete: {len(text)} chars, {len(text.split())} words")
            return text, "ocr"

        except ImportError:
            logger.error("OCR libraries not installed")
            return None, "error"
        except Exception as e:
            logger.error(f"Error during OCR: {e}")
            return None, "error"

    def is_readable_text(self, text: str) -> bool:
        """Check if extracted text is meaningful and not fragmented

        Enhanced checks:
        - Minimum length and word count
        - Alphabetic character ratio
        - Fragmentation detection (too many tiny words)
        - Reading order issues
        """
        if not text or len(text.strip()) < 100:
            return False

        words = text.split()
        word_count = len(words)

        if word_count < 20:
            return False

        alpha_chars = sum(c.isalpha() for c in text)
        total_chars = len(text)
        alpha_ratio = alpha_chars / total_chars if total_chars > 0 else 0

        if alpha_ratio < 0.3:
            return False

        # Check for fragmentation (too many 1-2 character words)
        tiny_words = sum(1 for w in words if len(w) <= 2)
        fragmentation_ratio = tiny_words / word_count if word_count > 0 else 0

        if fragmentation_ratio > 0.4:  # More than 40% tiny words indicates fragmentation
            logger.warning(f"Text appears fragmented: {fragmentation_ratio:.1%} tiny words")
            return False

        # Check average word length (scientific text should have reasonable average)
        avg_word_length = sum(len(w) for w in words) / word_count if word_count > 0 else 0
        if avg_word_length < 3:  # Too short average indicates poor extraction
            logger.warning(f"Average word length too short: {avg_word_length:.1f}")
            return False

        return True

    def convert_pdf_to_image(self, pdf_path: Path, dpi: int = 300, page_num: int = 0):
        """Convert PDF page to high-resolution image for Vision API

        Args:
            pdf_path: Path to PDF file
            dpi: Resolution for conversion (default 300)
            page_num: Page number to convert (0-indexed, default 0)

        Returns:
            PIL Image or None
        """
        try:
            import fitz  # PyMuPDF
            from PIL import Image

            Image.MAX_IMAGE_PIXELS = 300_000_000

            doc = fitz.open(str(pdf_path))
            if len(doc) == 0:
                return None

            if page_num >= len(doc):
                logger.error(f"Page {page_num} does not exist (PDF has {len(doc)} pages)")
                return None

            page = doc[page_num]
            zoom = dpi / 72
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat)

            page_image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            megapixels = (pix.width * pix.height) / 1_000_000
            logger.info(f"✓ Converted page {page_num+1} to {pix.width}x{pix.height} ({megapixels:.1f}MP) at {dpi} DPI")

            doc.close()
            return page_image

        except Exception as e:
            logger.error(f"Error converting PDF to image: {e}")
            return None

    def convert_pdf_to_images_all_pages(self, pdf_path: Path, dpi: int = 300) -> List[Any]:
        """Convert all PDF pages to images

        Args:
            pdf_path: Path to PDF file
            dpi: Resolution for conversion

        Returns:
            List of PIL Images
        """
        try:
            import fitz  # PyMuPDF
            from PIL import Image

            Image.MAX_IMAGE_PIXELS = 300_000_000

            doc = fitz.open(str(pdf_path))
            num_pages = len(doc)
            logger.info(f"Converting all {num_pages} pages to images at {dpi} DPI...")

            images = []
            zoom = dpi / 72
            mat = fitz.Matrix(zoom, zoom)

            for page_num in range(num_pages):
                page = doc[page_num]
                pix = page.get_pixmap(matrix=mat)
                page_image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                images.append(page_image)
                logger.info(f"  Page {page_num + 1}/{num_pages}: {pix.width}x{pix.height}")

            doc.close()
            logger.info(f"✓ Converted {len(images)} pages")
            return images

        except Exception as e:
            logger.error(f"Error converting PDF pages to images: {e}")
            return []

    def remove_template_artifacts(self, text: str) -> str:
        """Remove common poster template instructions and artifacts

        Patterns to remove:
        - Genigraphics branding and instructions
        - PowerPoint template guidance
        - Generic "DO NOT POST" warnings
        - Conference submission placeholders

        Args:
            text: Raw extracted text

        Returns:
            Cleaned text with template artifacts removed
        """
        if not text:
            return text

        # Define template patterns to remove
        template_patterns = [
            r'Genigraphics®[^\n]*(?:\n[^\n]*){0,10}?(?=\n[A-Z][a-z]|\Z)',  # Genigraphics blocks
            r'Change Color Theme:.*?(?=\n[A-Z])',
            r'This poster template is.*?ratio\.',
            r'The various elements included in this.*?needs\.',
            r'Always check with your conference.*?requirements\.',
            r'Image Quality:.*?(?=\n[A-Z])',
            r'You can place digital photos.*?(?=\n[A-Z])',
            r'standard copy & paste\. For best results,',
            r'be sure to preview your graphics at',
            r'Printing Your Poster:.*?(?=\n[A-Z])',
            r'Once your poster file is ready, visit',
            r'www\.genigraphics\.com.*?(?=\n)',
            r'DO NOT POST[^\n]*',
            r'PLEASE DO NOT DISTRIBUTE[^\n]*',
            r'Every order receives a free design review',
            r'anyone in the industry; dating back to',
            r'when we helped Microsoft.*?software\.',
            r'return to that after trying some of the',
            r'Please note that graphics from websites.*?printing\.',
            r'output from PowerPoint.*?longer than',
            r'has been producing.*?(?=\n)',
        ]

        cleaned_text = text
        removed_count = 0

        for pattern in template_patterns:
            matches = re.findall(pattern, cleaned_text, flags=re.IGNORECASE | re.DOTALL)
            if matches:
                removed_count += len(matches)
                cleaned_text = re.sub(pattern, '', cleaned_text, flags=re.IGNORECASE | re.DOTALL)

        # Remove excessive whitespace
        cleaned_text = re.sub(r'\n\s*\n\s*\n+', '\n\n', cleaned_text)

        removed_chars = len(text) - len(cleaned_text)
        if removed_chars > 0:
            logger.info(f"✓ Removed {removed_chars} chars of template artifacts ({removed_count} patterns matched)")

        return cleaned_text

    def extract_text_vision_with_rag(self, pdf_path: Path, reference_text: str, original_method: str, poster_image=None, encoded_image: Tuple[str, str] = None) -> Tuple[Optional[str], str]:
        """MANDATORY Vision AI pass using extracted text as RAG context

        This method always runs to enhance text extraction quality using Vision AI,
        regardless of the quality of the original extraction. Uses the extracted text
        as RAG (Retrieval-Augmented Generation) context to guide corrections.

        Args:
            pdf_path: Path to PDF file
            reference_text: Text from native/OCR extraction (RAG context)
            original_method: Method used for initial extraction (e.g., "native", "ocr")
            poster_image: Optional pre-converted PIL Image (for performance)
            encoded_image: Optional pre-encoded (base64, media_type) tuple

        Returns:
            Tuple of (enhanced_text, method_name)
        """
        try:
            client = self.get_anthropic_client()
            if not client:
                logger.error("API credentials not found for Vision AI enhancement")
                return None, "error"

            # Use pre-encoded image if available
            if encoded_image:
                img_base64, media_type = encoded_image
            else:
                # Use cached image if provided, otherwise convert
                if poster_image is not None:
                    image = poster_image
                else:
                    image = self.convert_pdf_to_image(pdf_path, dpi=self.DEFAULT_VISION_DPI)
                    if image is None:
                        logger.error("Failed to convert PDF to image for Vision AI enhancement")
                        return None, "error"

                img_base64, media_type = self.encode_image_to_base64(image, format="JPEG")
                if img_base64 is None:
                    logger.error("Failed to encode image for Vision AI enhancement")
                    return None, "error"

            # Build enhanced prompt with RAG context
            prompt = f"""You are analyzing a scientific poster. I've extracted text using {original_method}, but it may have:
- Reading order issues (columns mixed up, text from different sections interleaved)
- OCR errors (character misrecognition: l vs 1, O vs 0, etc.)
- Template contamination (printing company instructions like "Genigraphics®", "Change Color Theme", "DO NOT POST")
- Missing or misplaced figure captions
- Fragmented sentences or words

REFERENCE TEXT (RAG Context - use as guide but verify against image):
---
{reference_text[:8000]}
---

Your task:
1. Read the poster image carefully to see the actual layout
2. Use the reference text as a guide (it contains the content but may have errors)
3. Correct reading order issues - respect poster layout:
   - Title at top
   - Abstract/Introduction typically top-left
   - Methods, Results in middle
   - Conclusions/Discussion at bottom
   - Figures should stay near their sections
4. Fix OCR errors by comparing reference text with visual content
5. Remove ALL template/printing instructions (Genigraphics®, PowerPoint tips, "DO NOT POST", etc.)
6. Ensure figure captions are correctly extracted and positioned with their figures
7. Preserve ALL scientific content: author names, affiliations, references, data, statistics
8. Maintain proper paragraph structure and section headers

Output ONLY the corrected, clean text in proper reading order. No commentary or explanations."""

            # Call Vision API with RAG (using extracted text as retrieval context)
            logger.info(f"🔍 Running MANDATORY Vision AI RAG enhancement (using {len(reference_text)} chars {original_method} text as context)...")

            message = client.messages.create(
                model=self.VISION_MODEL,
                max_tokens=4096,
                messages=[{
                    "role": "user",
                    "content": [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": media_type,
                                "data": img_base64,
                            },
                        },
                        {
                            "type": "text",
                            "text": prompt
                        }
                    ],
                }]
            )

            enhanced_text = message.content[0].text

            # Calculate enhancement metrics
            char_diff = len(enhanced_text) - len(reference_text)
            char_change_pct = (char_diff / len(reference_text) * 100) if len(reference_text) > 0 else 0

            logger.info(f"✓ Vision AI RAG enhancement complete: {len(enhanced_text)} chars")
            logger.info(f"  Input tokens: {message.usage.input_tokens}, Output tokens: {message.usage.output_tokens}")

            if char_diff > 0:
                logger.info(f"  Enhanced: +{char_diff} chars (+{char_change_pct:.1f}%) - added content/corrections")
            elif char_diff < 0:
                logger.info(f"  Cleaned: {char_diff} chars ({char_change_pct:.1f}%) - removed artifacts/duplicates")
            else:
                logger.info(f"  Optimized: maintained {len(enhanced_text)} chars - corrected reading order")

            # Word count comparison for better quality indication
            ref_words = len(reference_text.split())
            enhanced_words = len(enhanced_text.split())
            word_diff = enhanced_words - ref_words

            if abs(word_diff) > 10:
                if word_diff > 0:
                    logger.info(f"  Content gain: +{word_diff} words recovered")
                else:
                    logger.info(f"  Cleanup: {abs(word_diff)} words removed (likely artifacts)")

            return enhanced_text, f"{original_method}_vision_enhanced"

        except Exception as e:
            logger.error(f"Error during Vision AI enhancement: {e}")
            return None, "error"

    def encode_image_to_base64(self, image, format="JPEG", max_dimension=2048):
        """Convert PIL Image to base64 string for API."""
        try:
            from PIL import Image

            width, height = image.size
            if max(width, height) > max_dimension:
                scale = max_dimension / max(width, height)
                new_width = int(width * scale)
                new_height = int(height * scale)
                image = image.resize((new_width, new_height), Image.Resampling.LANCZOS)

            if image.mode == 'RGBA':
                image = image.convert('RGB')

            buffered = BytesIO()
            image.save(buffered, format=format, quality=85 if format == "JPEG" else None)
            img_bytes = buffered.getvalue()
            img_base64 = base64.b64encode(img_bytes).decode('utf-8')

            media_type = f"image/{format.lower()}"
            return img_base64, media_type

        except Exception as e:
            logger.error(f"Error encoding image: {e}")
            return None, None

    def get_anthropic_client(self):
        """Get configured Anthropic client instance (cached via shared auth module)."""
        if self._client is not None:
            return self._client
        try:
            self._client = _get_shared_client()
        except RuntimeError:
            logger.error("Cannot create Anthropic client - API credentials not found")
            return None
        return self._client

    def extract_poster_structure_vision(self, pdf_path: Path, extracted_text: str = None, poster_image=None, encoded_image: Tuple[str, str] = None) -> Dict:
        """Use Vision AI to identify poster structure and extract metadata

        Args:
            pdf_path: Path to PDF file
            extracted_text: Optional extracted text (unused, kept for backward compatibility)
            poster_image: Optional pre-converted PIL Image (for performance)
            encoded_image: Optional pre-encoded (base64, media_type) tuple

        Returns:
            Dictionary with poster structure info (title, authors, sections, layout)
        """
        try:
            client = self.get_anthropic_client()
            if not client:
                logger.warning("API credentials not found for structure extraction")
                return {}

            if encoded_image:
                img_base64, media_type = encoded_image
            else:
                if poster_image is not None:
                    image = poster_image
                else:
                    image = self.convert_pdf_to_image(pdf_path, dpi=self.STRUCTURE_VISION_DPI)
                    if image is None:
                        return {}

                img_base64, media_type = self.encode_image_to_base64(image, format="JPEG")
                if img_base64 is None:
                    return {}

            prompt = """Analyze the structure and metadata of this scientific poster.

Extract:
1. Poster title (exact text)
2. Authors and affiliations
3. All major section headers visible as larger/bold/colored text (use exact names as shown)
4. Layout structure (single-column, two-column, three-column, irregular)
5. Reading order for sections (top-to-bottom, left-to-right)
6. Number of figures/charts visible
7. Approximate location of each section and figure positions

Format as JSON:
{
    "title": "...",
    "authors": "...",
    "affiliation": "...",
    "layout_type": "two-column",
    "sections": [
        {"name": "Title", "location": "top-center", "order": 1},
        {"name": "Abstract", "location": "top-left", "order": 2},
        {"name": "Methods", "location": "middle-left", "order": 3},
        {"name": "Results", "location": "middle-right", "order": 4},
        {"name": "Conclusions", "location": "bottom", "order": 5}
    ],
    "sections_present": ["Abstract", "Methods", "Results", "Conclusions"],
    "reading_order": ["Title", "Abstract", "Background", "Methods", "Results", "Discussion", "Conclusions"],
    "num_figures": 5,
    "figure_positions": ["right-column", "bottom-left"],
    "layout_notes": "Two-column layout with figures on right"
}

Be specific about section names as they appear on the poster (e.g., "BACKGROUND" not "Introduction" if that's what's shown)."""

            logger.info("Analyzing poster structure with Vision AI...")

            message = client.messages.create(
                model=self.VISION_MODEL,
                max_tokens=2048,
                messages=[{
                    "role": "user",
                    "content": [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": media_type,
                                "data": img_base64,
                            },
                        },
                        {
                            "type": "text",
                            "text": prompt
                        }
                    ],
                }]
            )

            response_text = message.content[0].text
            structure = self.parse_json_response(response_text)

            if structure:
                logger.info(f"✓ Extracted poster structure")
                logger.info(f"  Title: {structure.get('title', 'N/A')[:60]}...")
                logger.info(f"  Sections: {', '.join(structure.get('sections_present', []))}")
            else:
                logger.warning("Could not parse structure from Vision response")
                structure = {}

            return structure

        except Exception as e:
            logger.error(f"Error extracting structure with vision: {e}")
            return {}

    def build_vision_context(self, extracted_text: str, structure: Dict = None) -> str:
        """Build optimized context for Vision AI figure analysis

        Selects most relevant parts of text instead of just first 3000 chars

        Args:
            extracted_text: Full OCR text
            structure: Optional structure metadata

        Returns:
            Formatted context string for Vision prompt
        """
        context_parts = []

        # Extract title (first line or from structure)
        if structure and structure.get('title'):
            title = structure['title']
        else:
            lines = extracted_text.split('\n')
            title = lines[0].strip() if lines else ''

        if title:
            context_parts.append(f"POSTER TITLE:\n{title}\n")

        # Extract abstract (first 500 chars or identified section)
        abstract = ""
        if 'ABSTRACT' in extracted_text.upper():
            abstract_match = re.search(r'ABSTRACT[:\n]+(.*?)(?=\n[A-Z]{3,}|\Z)',
                                       extracted_text, re.IGNORECASE | re.DOTALL)
            if abstract_match:
                abstract = abstract_match.group(1).strip()[:500]

        if abstract:
            context_parts.append(f"ABSTRACT:\n{abstract}\n")

        # Extract figure captions
        captions = self.extract_figure_captions(extracted_text)
        if captions:
            context_parts.append("FIGURE CAPTIONS FROM TEXT:")
            for fig_num, caption in sorted(captions.items()):
                context_parts.append(f"  Figure {fig_num}: {caption}")
            context_parts.append("")

        # Extract methods summary (first 300 chars)
        if 'METHOD' in extracted_text.upper():
            methods_match = re.search(r'METHOD[Ss]?[:\n]+(.*?)(?=\n[A-Z]{3,}|\Z)',
                                      extracted_text, re.IGNORECASE | re.DOTALL)
            if methods_match:
                methods = methods_match.group(1).strip()[:300]
                context_parts.append(f"METHODS (summary):\n{methods}\n")

        context = '\n'.join(context_parts)

        # If still too short, add beginning of full text
        if len(context) < 1000:
            context += f"\n\nADDITIONAL CONTEXT:\n{extracted_text[:2000]}\n"

        return context

    def analyze_figures_with_vision(self, pdf_path: Path, poster_num: str, extracted_text: str = None, poster_image=None, encoded_image: Tuple[str, str] = None) -> List[Dict]:
        """Use Claude Vision API to identify and analyze figures

        Args:
            pdf_path: Path to PDF file
            poster_num: Poster number/identifier
            extracted_text: Optional extracted text from poster for context
            poster_image: Optional pre-converted PIL Image (for performance)
            encoded_image: Optional pre-encoded (base64, media_type) tuple
        """
        try:
            client = self.get_anthropic_client()
            if not client:
                logger.error("API credentials not found")
                return []

            if encoded_image:
                img_base64, media_type = encoded_image
            else:
                if poster_image is not None:
                    image = poster_image
                else:
                    image = self.convert_pdf_to_image(pdf_path, dpi=self.DEFAULT_VISION_DPI)
                    if image is None:
                        logger.error("Failed to convert PDF to image")
                        return []

                img_base64, media_type = self.encode_image_to_base64(image, format="JPEG")
                if img_base64 is None:
                    logger.error("Failed to encode image")
                    return []

            # Build prompt with optimized text context
            base_prompt = """Analyze this scientific poster and extract detailed information about all figures/charts.

For each distinct figure or chart panel, provide:
1. Figure number/label (as it appears on the poster)
2. Figure title/caption (if visible on the figure)
3. Detailed description of what the figure shows
4. Key findings or results illustrated
5. Figure type (e.g., bar chart, line plot, scatter plot, heatmap, microscopy image, etc.)
6. Any statistical significance markers or p-values
7. Position/location on the poster (optional, if relevant)"""

            if extracted_text:
                # Use optimized context builder
                context = self.build_vision_context(extracted_text)
                context_prompt = f"""

CONTEXT FROM EXTRACTED TEXT:
---
{context}
---

Use this context to better understand the research question, figure captions, and statistical details. Match figure descriptions with their captions."""
                full_prompt = base_prompt + context_prompt
                logger.info("Including optimized text context for vision analysis")
            else:
                full_prompt = base_prompt
                logger.info("Analyzing without text context (text not available)")

            full_prompt += """

Format as JSON:
{
    "poster_title": "...",
    "figures": [
        {
            "figure_number": "1",
            "title": "...",
            "caption": "Complete caption text as shown near/below the figure",
            "description": "...",
            "key_findings": ["...", "..."],
            "figure_type": "...",
            "statistical_notes": "..."
        }
    ]
}

IMPORTANT: For each figure, also extract its complete caption text (usually found below or near the figure, starting with "Figure X:" or "Fig X:"). Include the full caption in the "caption" field."""

            # Call Claude Vision API
            logger.info("Analyzing poster with Claude Vision API...")

            message = client.messages.create(
                model=self.VISION_MODEL,
                max_tokens=8192,
                messages=[{
                    "role": "user",
                    "content": [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": media_type,
                                "data": img_base64,
                            },
                        },
                        {
                            "type": "text",
                            "text": full_prompt
                        }
                    ],
                }]
            )

            response_text = message.content[0].text
            logger.info(f"✓ Received analysis from Claude")
            logger.info(f"Input tokens: {message.usage.input_tokens}, Output tokens: {message.usage.output_tokens}")

            if message.stop_reason == "max_tokens":
                logger.warning("⚠ Figure analysis response was truncated (hit max_tokens limit) - attempting recovery")

            # Parse JSON response
            figures_data = self.parse_json_response(response_text)
            if figures_data and 'figures' in figures_data:
                logger.info(f"✓ Identified {len(figures_data['figures'])} figures")
                return figures_data['figures']
            else:
                logger.warning("Could not parse figure data from response")
                return []

        except ImportError:
            logger.error("anthropic package not installed. Install with: conda install -c conda-forge anthropic")
            return []
        except Exception as e:
            logger.error(f"Error analyzing figures with vision: {e}")
            return []

    def enhance_figure_with_detailed_analysis(self, image_or_path_or_encoded, figure: Dict,
                                             extracted_text: str = None) -> Dict:
        """Stage 2: Enhance figure with detailed Vision AI analysis

        Performs focused analysis on individual figure region for more detail

        Args:
            image_or_path_or_encoded: Pre-encoded (base64, media_type) tuple, PIL Image, or Path to PDF
            figure: Figure dict from Stage 1 with basic info
            extracted_text: Optional text for additional context

        Returns:
            Enhanced figure dictionary
        """
        try:
            client = self.get_anthropic_client()
            if not client:
                return figure

            # Accept pre-encoded image, PIL Image, or path
            if isinstance(image_or_path_or_encoded, tuple) and len(image_or_path_or_encoded) == 2:
                img_base64, media_type = image_or_path_or_encoded
            else:
                from PIL import Image as PILImage
                if isinstance(image_or_path_or_encoded, PILImage.Image):
                    image = image_or_path_or_encoded
                else:
                    image = self.convert_pdf_to_image(image_or_path_or_encoded, dpi=self.DEFAULT_VISION_DPI)
                    if image is None:
                        return figure

                img_base64, media_type = self.encode_image_to_base64(image, format="JPEG")
                if img_base64 is None:
                    return figure

            fig_num = figure.get('figure_number', '?')
            fig_title = figure.get('title', '')
            fig_type = figure.get('figure_type', '')

            # Build focused prompt
            prompt = f"""Focus on Figure {fig_num} in this poster ({fig_type}: "{fig_title}").

Provide DETAILED analysis:
1. Axis labels and units (for charts/plots)
2. All visible data series or groups
3. Exact statistical values if readable (means, error bars, p-values)
4. Color coding and legend interpretation
5. Sample sizes if shown
6. Any annotations or callouts
7. What experimental comparison is being shown
8. What is the main conclusion from this figure

Be specific and quantitative where possible."""

            if extracted_text:
                # Find context around this figure
                fig_pattern = rf'Fig(?:ure)?\s*{re.escape(str(fig_num))}'
                match = re.search(fig_pattern, extracted_text, re.IGNORECASE)
                if match:
                    context_start = max(0, match.start() - 300)
                    context_end = min(len(extracted_text), match.end() + 300)
                    fig_context = extracted_text[context_start:context_end]
                    prompt += f"\n\nTEXT CONTEXT:\n{fig_context}"

            prompt += "\n\nProvide response as plain text (not JSON)."

            logger.info(f"  Detailed analysis of Figure {fig_num}...")

            message = client.messages.create(
                model=self.VISION_MODEL,
                max_tokens=2048,
                messages=[{
                    "role": "user",
                    "content": [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": media_type,
                                "data": img_base64,
                            },
                        },
                        {
                            "type": "text",
                            "text": prompt
                        }
                    ],
                }]
            )

            detailed_analysis = message.content[0].text
            figure['detailed_analysis'] = detailed_analysis
            logger.info(f"    ✓ Added detailed analysis ({len(detailed_analysis)} chars)")

            return figure

        except Exception as e:
            logger.error(f"Error in detailed figure analysis: {e}")
            return figure

    def analyze_figures_with_vision_fallback(self, pdf_path: Path, poster_num: str, extracted_text: str = None, encoded_image: Tuple[str, str] = None) -> List[Dict]:
        """Fallback figure detection with more explicit prompt

        Used when standard detection returns 0 figures

        Args:
            pdf_path: Path to PDF file
            poster_num: Poster number/identifier
            extracted_text: Optional extracted text from poster for context
            encoded_image: Optional pre-encoded (base64, media_type) tuple

        Returns:
            List of figure dictionaries
        """
        try:
            client = self.get_anthropic_client()

            if encoded_image:
                img_base64, media_type = encoded_image
            else:
                image = self.convert_pdf_to_image(pdf_path, dpi=self.DEFAULT_VISION_DPI)
                if image is None:
                    return []

                img_base64, media_type = self.encode_image_to_base64(image, format="JPEG")
                if img_base64 is None:
                    return []

            # More explicit prompt focusing on counting
            prompt = """Look very carefully at this scientific poster image.

IMPORTANT: Count ALL visible charts, graphs, plots, diagrams, images, or data visualizations.

Even if they are small, even if they are part of a larger panel, count each one.

How many distinct figures, charts, or graphs do you see? Please count carefully and then describe each one briefly.

Format your response as:
FIGURE_COUNT: [number]

Then for each figure:
Figure [number]: [brief description]
"""

            response = client.messages.create(
                model=self.VISION_MODEL,
                max_tokens=2048,
                messages=[{
                    "role": "user",
                    "content": [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": media_type,
                                "data": img_base64
                            }
                        },
                        {"type": "text", "text": prompt}
                    ]
                }]
            )

            result_text = response.content[0].text
            logger.info(f"Fallback detection response: {result_text[:500]}")

            # Parse response for figures
            figures = []
            count_match = re.search(r'FIGURE_COUNT:\s*(\d+)', result_text)
            if count_match:
                count = int(count_match.group(1))
                logger.info(f"Fallback detected {count} figures")

                # Extract figure descriptions
                figure_pattern = re.compile(r'Figure\s+(\d+):\s*([^\n]+)', re.IGNORECASE)
                for match in figure_pattern.finditer(result_text):
                    fig_num = match.group(1)
                    description = match.group(2).strip()
                    figures.append({
                        'number': fig_num,
                        'description': description,
                        'type': 'Unknown',
                        'key_findings': []
                    })

            return figures

        except Exception as e:
            logger.error(f"Fallback figure detection failed: {e}")
            return []

    def analyze_figures_two_stage(self, pdf_path: Path, poster_num: str,
                                  extracted_text: str = None, poster_image=None, encoded_image: Tuple[str, str] = None) -> List[Dict]:
        """Two-stage Vision AI figure analysis with parallel Stage 2 processing

        Stage 1: Identify all figures and get basic descriptions
        Stage 2: Detailed analysis of each figure individually (PARALLEL)

        Args:
            pdf_path: Path to PDF
            poster_num: Poster identifier
            extracted_text: Optional text for context
            poster_image: Optional pre-converted PIL Image (for performance)
            encoded_image: Optional pre-encoded (base64, media_type) tuple

        Returns:
            List of figure dictionaries with detailed analysis
        """
        logger.info("Starting two-stage figure analysis...")

        # Stage 1: Overview (use pre-encoded image)
        figures = self.analyze_figures_with_vision(pdf_path, poster_num, extracted_text, poster_image=poster_image, encoded_image=encoded_image)

        # Fallback: If no figures detected, try again with more explicit prompt
        if not figures:
            logger.warning("No figures identified in Stage 1 - trying fallback detection")
            figures = self.analyze_figures_with_vision_fallback(pdf_path, poster_num, extracted_text, encoded_image=encoded_image)

        if not figures:
            logger.warning("No figures identified even after fallback attempt")
            return []

        logger.info(f"Stage 1 complete: {len(figures)} figures identified")

        # Stage 2: Detailed analysis of each figure (PARALLEL)
        logger.info(f"Stage 2: Detailed analysis of {len(figures)} figures (parallel processing)...")

        max_workers = min(self.MAX_PARALLEL_FIGURE_WORKERS, len(figures))

        enhanced_figures = [None] * len(figures)  # Preserve order

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            # Pass pre-encoded image to avoid redundant encoding in threads
            image_arg = encoded_image if encoded_image else (poster_image if poster_image else pdf_path)
            future_to_index = {}
            for i, fig in enumerate(figures):
                future = executor.submit(
                    self.enhance_figure_with_detailed_analysis,
                    image_arg, fig, extracted_text
                )
                future_to_index[future] = i

            # Collect results as they complete
            completed = 0
            failed = 0
            for future in as_completed(future_to_index):
                idx = future_to_index[future]
                try:
                    enhanced_fig = future.result()
                    if enhanced_fig is None:
                        logger.warning(f"  ⚠ Figure {idx+1} returned None - using Stage 1 data")
                        enhanced_figures[idx] = figures[idx]
                        failed += 1
                    else:
                        enhanced_figures[idx] = enhanced_fig
                    completed += 1
                    logger.info(f"  ✓ Figure {idx+1}/{len(figures)} complete ({completed}/{len(figures)} done)")
                except Exception as e:
                    logger.error(f"  ✗ Figure {idx+1} detailed analysis failed: {type(e).__name__}: {str(e)}")
                    # Use original figure if detailed analysis fails
                    enhanced_figures[idx] = figures[idx]
                    failed += 1
                    completed += 1

        if failed > 0:
            logger.warning(f"⚠ {failed}/{len(figures)} figures failed Stage 2 - using Stage 1 data as fallback")

        logger.info(f"✓ Two-stage analysis complete (parallel processing)")
        return enhanced_figures

    def extract_figure_captions_vision(self, pdf_path: Path, figures_list: List[Dict], extracted_text: str, poster_image=None, encoded_image: Tuple[str, str] = None) -> Dict[str, str]:
        """Use Vision AI to extract clean, accurate figure captions

        Args:
            pdf_path: Path to PDF
            figures_list: List of identified figures from Stage 1
            extracted_text: Full extracted text (for cross-reference)
            poster_image: Optional pre-converted PIL Image (for performance)
            encoded_image: Optional pre-encoded (base64, media_type) tuple

        Returns:
            Dictionary mapping figure numbers to clean captions
        """
        if not figures_list:
            return {}

        try:
            client = self.get_anthropic_client()
            if not client:
                logger.warning("Cannot extract captions via Vision - using text-based extraction")
                return self.extract_figure_captions(extracted_text)

            if encoded_image:
                img_base64, media_type = encoded_image
            else:
                if poster_image is not None:
                    image = poster_image
                else:
                    image = self.convert_pdf_to_image(pdf_path, dpi=self.DEFAULT_VISION_DPI)
                    if image is None:
                        logger.warning("Failed to convert PDF for caption extraction")
                        return self.extract_figure_captions(extracted_text)

                img_base64, media_type = self.encode_image_to_base64(image, format="JPEG")
                if img_base64 is None:
                    logger.warning("Failed to encode image for caption extraction")
                    return self.extract_figure_captions(extracted_text)

            # Build caption extraction prompt
            figure_summary = "\n".join([f"Figure {fig.get('figure_number', fig.get('number', '?'))}: {fig.get('description', 'Unknown')}"
                                        for fig in figures_list])

            prompt = f"""I've identified {len(figures_list)} figures in this poster:
{figure_summary}

For EACH figure, extract the complete caption text that appears near or below it.
Captions typically start with "Figure X:" or "Fig X:" or similar.

REFERENCE TEXT (may contain caption fragments):
---
{extracted_text[:3000]}
---

Your task:
1. For each figure, find its caption on the poster image
2. Extract the COMPLETE caption text
3. Remove any template artifacts or printing instructions
4. Preserve scientific terminology, statistics, and formatting

Output format (one line per figure):
Figure 1: [complete caption text, cleaned of artifacts]
Figure 2: [complete caption text, cleaned of artifacts]
...

Extract ONLY the caption text that describes what the figure shows, not the figure content itself.
"""

            logger.info("🔍 Extracting figure captions via Vision AI...")

            response = client.messages.create(
                model=self.VISION_MODEL,
                max_tokens=2048,
                messages=[{
                    "role": "user",
                    "content": [
                        {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": img_base64}},
                        {"type": "text", "text": prompt}
                    ]
                }]
            )

            # Parse response into caption dictionary
            captions = {}
            caption_text = response.content[0].text

            # Match "Figure X: caption" patterns
            caption_pattern = re.compile(r'Figure\s+(\d+|[A-Z]):\s*(.+?)(?=\nFigure|\Z)', re.IGNORECASE | re.DOTALL)

            for match in caption_pattern.finditer(caption_text):
                fig_num = match.group(1)
                caption = match.group(2).strip()
                # Clean caption of template artifacts
                caption = self.remove_template_artifacts(caption)
                # Remove excessive whitespace
                caption = re.sub(r'\s+', ' ', caption).strip()
                captions[fig_num] = caption

            logger.info(f"✓ Extracted {len(captions)} figure captions via Vision AI")

            return captions

        except Exception as e:
            logger.error(f"Error extracting captions via Vision: {e}")
            # Fallback to text-based extraction
            return self.extract_figure_captions(extracted_text)

    def parse_structure_with_vision(self, pdf_path: Path, extracted_text: str, poster_image=None, encoded_image: Tuple[str, str] = None) -> Optional[Dict[str, Any]]:
        """Use Vision AI to identify poster structure and section headers

        Handles:
        - Non-standard section names (BACKGROUND vs Introduction)
        - Visual styling (font size, bold, color)
        - Multi-column layouts
        - Section boundaries

        Args:
            pdf_path: Path to PDF
            extracted_text: Extracted text for content identification
            poster_image: Optional pre-converted PIL Image (for performance)
            encoded_image: Optional pre-encoded (base64, media_type) tuple

        Returns:
            Dictionary with structure information or None if failed
        """
        try:
            client = self.get_anthropic_client()
            if not client:
                logger.warning("Vision structure parsing unavailable - using text-based parsing")
                return None

            if encoded_image:
                img_base64, media_type = encoded_image
            else:
                if poster_image is not None:
                    image = poster_image
                else:
                    image = self.convert_pdf_to_image(pdf_path, dpi=self.STRUCTURE_VISION_DPI)
                    if image is None:
                        return None

                img_base64, media_type = self.encode_image_to_base64(image, format="JPEG")
                if img_base64 is None:
                    return None

            prompt = f"""Analyze the structure of this scientific poster.

REFERENCE TEXT (for section content identification):
---
{extracted_text[:5000]}
---

Identify:
1. All major section headers visible as larger/bold/colored text
2. Section names (may be non-standard: BACKGROUND, HYPOTHESIS, OBJECTIVES, etc.)
3. Layout structure (single column, two-column, three-column, irregular)
4. Approximate location of each section (top/middle/bottom, left/center/right)
5. Reading order for sections

Output JSON format:
{{
  "layout_type": "multi-column-2" or "single-column" or "three-column",
  "sections": [
    {{"name": "Title", "location": "top-center", "order": 1}},
    {{"name": "Abstract", "location": "top-left", "order": 2}},
    {{"name": "Methods", "location": "middle-left", "order": 3}},
    {{"name": "Results", "location": "middle-right", "order": 4}}
  ],
  "reading_order": ["Title", "Abstract", "Background", "Methods", "Results", "Discussion", "Conclusions"],
  "figure_positions": ["right-column", "bottom-left", "bottom-right"]
}}

Be specific about section names as they appear on the poster (e.g., "BACKGROUND" not "Introduction" if that's what's shown).
"""

            logger.info("🔍 Analyzing poster structure via Vision AI...")

            response = client.messages.create(
                model=self.VISION_MODEL,
                max_tokens=1024,
                messages=[{
                    "role": "user",
                    "content": [
                        {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": img_base64}},
                        {"type": "text", "text": prompt}
                    ]
                }]
            )

            # Parse JSON response
            structure_data = self.parse_json_response(response.content[0].text)

            if structure_data:
                layout = structure_data.get('layout_type', 'unknown')
                sections = structure_data.get('sections', [])
                logger.info(f"✓ Vision-based structure: {layout}, {len(sections)} sections identified")
                return structure_data
            else:
                logger.warning("Failed to parse Vision structure response")
                return None

        except Exception as e:
            logger.error(f"Error in Vision structure parsing: {e}")
            return None

    def parse_json_response(self, response_text: str):
        """Parse JSON from Claude's response, with truncation recovery."""
        json_str = response_text

        # Extract from markdown code block if present
        if "```json" in json_str:
            json_start = json_str.find("```json") + 7
            json_end = json_str.find("```", json_start)
            json_str = json_str[json_start:json_end].strip() if json_end > json_start else json_str[json_start:].strip()
        elif "```" in json_str:
            json_start = json_str.find("```") + 3
            json_end = json_str.find("```", json_start)
            json_str = json_str[json_start:json_end].strip() if json_end > json_start else json_str[json_start:].strip()

        try:
            return json.loads(json_str)
        except json.JSONDecodeError:
            pass

        # Truncation recovery: find the last complete figure object in a "figures" array
        figures_key = '"figures"'
        if figures_key in json_str:
            # Find the last complete figure object by looking for "}," or "}" before truncation
            last_complete = json_str.rfind('},')
            last_single = json_str.rfind('}')

            # Try closing at last "}," (end of a complete array element)
            if last_complete > json_str.find(figures_key):
                candidate = json_str[:last_complete + 1] + ']}'
                try:
                    result = json.loads(candidate)
                    logger.info(f"✓ Recovered truncated JSON (figures array closed after last complete entry)")
                    return result
                except json.JSONDecodeError:
                    pass

            # Try closing at last "}" (might be the last figure without trailing comma)
            if last_single > json_str.find(figures_key):
                candidate = json_str[:last_single + 1] + ']}'
                try:
                    result = json.loads(candidate)
                    logger.info(f"✓ Recovered truncated JSON (closed at last brace)")
                    return result
                except json.JSONDecodeError:
                    pass

        # Generic recovery: find last valid closing brace
        last_brace = json_str.rfind('}')
        if last_brace > 0:
            try:
                result = json.loads(json_str[:last_brace + 1])
                logger.info(f"✓ Recovered truncated JSON (generic brace recovery)")
                return result
            except json.JSONDecodeError:
                pass

        return None

    def extract_figure_captions(self, text: str) -> Dict[str, str]:
        """Extract figure captions from OCR text

        Looks for patterns like:
        - "Figure 1: Caption text"
        - "Fig. 2: Caption text"
        - "Fig 3. Caption text"
        - "Fig.1 //** Caption text"  (handles special poster format)

        Extracts captions up to next figure mention or paragraph break.

        Returns:
            Dictionary mapping figure number to caption text
        """
        captions = {}

        # Enhanced patterns to match various figure caption formats
        # Capture everything until next figure mention or double newline
        patterns = [
            # Standard formats with flexible caption capture
            r'Figure\s+(\d+)[:\.]?\s*//\*\*\s*([^\n]+(?:\n(?!Fig|Figure|\n)[^\n]+)*)',
            r'Fig\.\s*(\d+)[:\.]?\s*//\*\*\s*([^\n]+(?:\n(?!Fig|Figure|\n)[^\n]+)*)',
            r'Fig\s+(\d+)[:\.]?\s*//\*\*\s*([^\n]+(?:\n(?!Fig|Figure|\n)[^\n]+)*)',
            r'Figure\s+(\d+)[:\.]?\s*([^\n]+(?:\n(?!Fig|Figure|\n)[^\n]+)*)',
            r'Fig\.\s*(\d+)[:\.]?\s*([^\n]+(?:\n(?!Fig|Figure|\n)[^\n]+)*)',
            r'Fig\s+(\d+)[:\.]?\s*([^\n]+(?:\n(?!Fig|Figure|\n)[^\n]+)*)',
        ]

        for pattern in patterns:
            matches = re.finditer(pattern, text, re.IGNORECASE | re.MULTILINE)
            for match in matches:
                fig_num = match.group(1)
                caption = match.group(2).strip()

                # Clean up caption: remove excessive whitespace, limit to reasonable length
                caption = re.sub(r'\s+', ' ', caption)  # Normalize whitespace

                # Take first 500 chars if caption is too long (while preserving sentence boundaries)
                if len(caption) > 500:
                    # Try to break at sentence boundary
                    truncate_pos = caption.rfind('.', 0, 500)
                    if truncate_pos > 300:  # Reasonable sentence found
                        caption = caption[:truncate_pos + 1]
                    else:
                        caption = caption[:500] + '...'

                # Keep longest caption found for each figure number
                if fig_num not in captions or len(caption) > len(captions.get(fig_num, '')):
                    captions[fig_num] = caption

        if captions:
            logger.info(f"✓ Extracted {len(captions)} figure captions from text")

        return captions

    def find_figure_positions(self, text: str, figures: List[Dict]) -> Dict[str, Dict]:
        """Find positions where each figure is mentioned in text

        Args:
            text: Full extracted text
            figures: List of figure dictionaries with 'figure_number' field

        Returns:
            Dictionary mapping figure_number to position info
        """
        positions = {}

        for fig in figures:
            fig_num = str(fig.get('figure_number', ''))
            if not fig_num:
                continue

            # Try multiple patterns
            patterns = [
                rf'\bFig\.\s*{re.escape(fig_num)}\b',
                rf'\bFigure\s+{re.escape(fig_num)}\b',
                rf'\bFig\s+{re.escape(fig_num)}\b',
            ]

            best_match = None
            for pattern in patterns:
                match = re.search(pattern, text, re.IGNORECASE)
                if match:
                    if best_match is None or match.start() < best_match['position']:
                        # Extract context (3 sentences before and after)
                        context_start = max(0, match.start() - 200)
                        context_end = min(len(text), match.end() + 200)
                        context = text[context_start:context_end].strip()

                        best_match = {
                            'position': match.start(),
                            'context': context,
                            'match_text': match.group()
                        }

            if best_match:
                positions[fig_num] = best_match

        if positions:
            logger.info(f"✓ Found {len(positions)} figure mentions in text")

        return positions

    def parse_poster_structure(self, text: str, vision_structure: Dict = None) -> Dict:
        """Parse poster text into structured sections

        Args:
            text: Full extracted text
            vision_structure: Optional structure info from Vision AI

        Returns:
            Dictionary with parsed sections
        """
        sections = {
            'title': '',
            'authors': '',
            'institution': '',
            'abstract': '',
            'objective': '',
            'results': '',
            'conclusions': '',
            'methods': ''
        }

        # Try Vision AI-assisted parsing first (more reliable)
        if vision_structure and text:
            logger.info("Using Vision AI-assisted section parsing...")
            vision_sections = self.extract_sections_with_vision(text, vision_structure)
            if vision_sections:
                sections.update(vision_sections)
                return sections
            else:
                logger.warning("Vision AI section extraction failed, falling back to regex")

        # Fallback to regex-based extraction with enhanced pattern matching
        text_upper = text.upper()
        lines = text.split('\n')
        if lines:
            sections['title'] = lines[0].strip()

        # Enhanced section detection with multiple pattern variations
        # Abstract
        if re.search(r'\b(ABSTRACT|SUMMARY)\b', text_upper):
            sections['abstract'] = self.extract_section_enhanced(text, ['ABSTRACT', 'SUMMARY'])

        # Objective/Background/Introduction
        if re.search(r'\b(OBJECTIVE|BACKGROUND|INTRODUCTION|AIM|HYPOTHESIS)\b', text_upper):
            sections['objective'] = self.extract_section_enhanced(text, ['OBJECTIVE', 'BACKGROUND', 'INTRODUCTION', 'AIM', 'HYPOTHESIS'])

        # Results/Findings
        if re.search(r'\b(RESULTS?|FINDINGS?|OBSERVATIONS?|DATA|OUTCOMES?)\b', text_upper):
            sections['results'] = self.extract_section_enhanced(text, ['RESULTS', 'FINDINGS', 'OBSERVATIONS', 'DATA', 'OUTCOMES'])

        # Conclusions/Discussion
        if re.search(r'\b(CONCLUSIONS?|DISCUSSION|SUMMARY)\b', text_upper):
            sections['conclusions'] = self.extract_section_enhanced(text, ['CONCLUSION', 'DISCUSSION', 'SUMMARY'])

        # Methods - Enhanced with common variations
        if re.search(r'\b(METHODS?|MATERIALS?|EXPERIMENTAL|PROCEDURES?|METHODOLOGY|STUDY DESIGN|APPROACH)\b', text_upper):
            sections['methods'] = self.extract_section_enhanced(text, [
                'METHOD', 'MATERIAL', 'EXPERIMENTAL', 'PROCEDURE',
                'METHODOLOGY', 'STUDY DESIGN', 'APPROACH',
                'MATERIALS AND METHODS', 'EXPERIMENTAL DESIGN',
                'MATERIALS & METHODS'
            ])

        return sections

    def extract_sections_with_vision(self, text: str, vision_structure: Dict) -> Dict:
        """Use Vision AI to extract section content from text

        More reliable than regex parsing, especially for posters with non-standard layouts

        Args:
            text: Full extracted text
            vision_structure: Structure info from Vision AI (with title, authors, sections)

        Returns:
            Dictionary with extracted sections
        """
        try:
            client = self.get_anthropic_client()

            # Build prompt with section information from Vision AI
            sections_list = vision_structure.get('sections', [])
            if not sections_list:
                return None

            # Convert section dicts to strings (sections may be [{"name": "Title", ...}] format)
            section_names = []
            for section in sections_list:
                if isinstance(section, dict):
                    section_names.append(section.get('name', str(section)))
                else:
                    section_names.append(str(section))

            prompt = f"""You are analyzing extracted text from a scientific poster. The poster has these sections: {', '.join(section_names)}

Please extract the content for each section from the text below. Return ONLY a JSON object with these keys:
- abstract: The abstract text (or empty string if not found - NOTE: abstracts are often not present on posters)
- objective: The objective/background/introduction text (or empty string if not found)
- methods: The methods/methodology text (or empty string if not found)
- results: The results/findings text (or empty string if not found)
- conclusions: The conclusions/discussion text (or empty string if not found)

Guidelines:
1. Extract the actual content, not section headers
2. Include all text that belongs to each section
3. If a section is not present, use empty string "" - this is normal for posters
4. Abstracts are OPTIONAL on posters - the poster itself serves as the abstract
5. Focus on extracting the core scientific content: methods, results, conclusions
6. Preserve scientific terminology and numbers accurately
7. Remove duplicate section headers if present

Extracted text:
{text[:15000]}
"""

            response = client.messages.create(
                model=self.VISION_MODEL,
                max_tokens=4096,
                messages=[{"role": "user", "content": prompt}]
            )

            result_text = response.content[0].text.strip()

            # Parse JSON response
            import json
            import re
            # Extract JSON from response (handle markdown code blocks)
            if '```json' in result_text:
                result_text = result_text.split('```json')[1].split('```')[0].strip()
            elif '```' in result_text:
                result_text = result_text.split('```')[1].split('```')[0].strip()

            # Attempt to fix common JSON issues before parsing
            try:
                sections = json.loads(result_text)
            except json.JSONDecodeError as e:
                logger.warning(f"JSON parse error: {e}. Attempting to fix...")
                # Try to fix unterminated strings by finding the last valid closing brace
                # Find the last complete JSON object
                last_brace = result_text.rfind('}')
                if last_brace > 0:
                    # Try parsing up to the last closing brace
                    truncated = result_text[:last_brace+1]
                    try:
                        sections = json.loads(truncated)
                        logger.info("✓ Recovered JSON by truncating at last closing brace")
                    except json.JSONDecodeError:
                        # If that fails, give up and return None
                        raise
                else:
                    raise

            logger.info(f"✓ Vision AI extracted sections: {', '.join([k for k, v in sections.items() if v])}")
            return sections

        except json.JSONDecodeError as e:
            logger.error(f"Error in Vision structure parsing: {e}")
            return None
        except Exception as e:
            logger.error(f"Vision AI section extraction failed: {e}")
            return None

    def extract_section(self, text: str, section_name: str) -> str:
        """Extract text for a specific section

        Filters out duplicate standalone section headers that are OCR artifacts
        """
        pattern = re.compile(f'{section_name}s?[:\n]+(.*?)(?=\n[A-Z]{{3,}}|\Z)', re.IGNORECASE | re.DOTALL)
        match = pattern.search(text)
        if match:
            section_text = match.group(1).strip()

            # Remove standalone section header lines (OCR artifacts)
            # These are lines that are ONLY an all-caps section name, possibly with punctuation
            section_text = self._filter_duplicate_headers(section_text)

            return section_text
        return ""

    def extract_section_enhanced(self, text: str, section_names: List[str]) -> str:
        """Extract text for a section using multiple possible header variations

        Tries multiple patterns to match common section header variations
        (e.g., "Methods", "Materials and Methods", "Methodology", etc.)

        Args:
            text: Full poster text
            section_names: List of possible section header names to try

        Returns:
            Extracted section text or empty string
        """
        # Try each possible section name
        for section_name in section_names:
            # Build flexible pattern that matches various formats:
            # - "METHODS:", "Methods:", "methods:"
            # - "METHODS\n", "Methods\n"
            # - "Materials and Methods:", etc.
            pattern = re.compile(
                f'\\b{re.escape(section_name)}s?[:\n]+(.{{20,}}?)(?=\n[A-Z]{{3,}}[:\n]|\Z)',
                re.IGNORECASE | re.DOTALL
            )
            match = pattern.search(text)
            if match:
                section_text = match.group(1).strip()

                # Ensure we have substantial content (not just noise)
                if len(section_text) >= self.MIN_SECTION_LENGTH:
                    # Remove duplicate headers
                    section_text = self._filter_duplicate_headers(section_text)
                    logger.debug(f"  Matched section header: '{section_name}' ({len(section_text)} chars)")
                    return section_text

        return ""

    def _filter_duplicate_headers(self, text: str) -> str:
        """Filter out standalone section header lines that are OCR duplicates

        Removes lines that are just section names in ALL CAPS (e.g., "OBJECTIVE", "RESULTS")
        These are OCR artifacts that duplicate the structural sections we add in markdown

        Args:
            text: Section text to filter

        Returns:
            Filtered text with duplicate headers removed
        """
        common_headers = [
            'ABSTRACT', 'OBJECTIVE', 'OBJECTIVES', 'BACKGROUND', 'INTRODUCTION',
            'METHODS', 'METHODOLOGY', 'MATERIALS AND METHODS',
            'RESULTS', 'FINDINGS', 'DISCUSSION', 'CONCLUSIONS', 'CONCLUSION',
            'REFERENCES', 'ACKNOWLEDGMENTS', 'ACKNOWLEDGEMENTS'
        ]

        lines = text.split('\n')
        filtered_lines = []

        for line in lines:
            line_stripped = line.strip()
            # Check if this line is ONLY a section header (all caps, possibly with colon)
            line_upper = line_stripped.rstrip(':').rstrip('.').strip()

            # Skip if it's a standalone section header
            if line_upper in common_headers and len(line_stripped) < 30:
                continue  # Skip this duplicate header line

            filtered_lines.append(line)

        return '\n'.join(filtered_lines)

    def log_quality(self, poster_num: str, quality_scores: Dict, saved: bool, filename: str = ""):
        """Log quality information to quality log file with timestamp

        Args:
            poster_num: Poster number/ID
            quality_scores: Quality score dictionary
            saved: Whether markdown was saved (False if quality too low)
            filename: Original PDF filename (without path)
        """
        from datetime import datetime

        status = "SAVED" if saved else "SKIPPED"
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"{timestamp}\t{poster_num}\t{quality_scores['overall']}\t{quality_scores['assessment']}\t{status}\t{filename}\n"

        with open(self.quality_log_path, 'a', encoding='utf-8') as f:
            f.write(log_entry)

    def calculate_quality_score(self, text: str, figures: List[Dict], sections: Dict,
                               extraction_method: str) -> Dict[str, Any]:
        """Calculate quality scores for the extracted content with enhanced checks

        Returns:
            Dictionary with quality metrics and overall score
        """
        scores = {}

        # Text extraction quality (0-10)
        text_score = 0
        if text and len(text) > 100:
            text_score = min(10, len(text) / 1000)  # Length-based

            # Check for fragmentation
            words = text.split()
            if words:
                tiny_words_ratio = sum(1 for w in words if len(w) <= 2) / len(words)
                if tiny_words_ratio < 0.3:
                    text_score += 2
                avg_word_len = sum(len(w) for w in words) / len(words)
                if avg_word_len > 4:
                    text_score += 2

            # Bonus for Vision OCR
            if 'vision' in extraction_method:
                text_score += 1

        scores['text_extraction'] = min(10, text_score)

        # Figure analysis quality (0-10)
        figure_score = 0
        if figures:
            figure_score = min(7, len(figures) * 1.5)  # Base on count

            # Check for detailed descriptions
            has_detailed = any('detailed_analysis' in fig for fig in figures)
            if has_detailed:
                figure_score += 2

            # Check for key findings
            has_findings = any(fig.get('key_findings') for fig in figures)
            if has_findings:
                figure_score += 1

        scores['figure_analysis'] = min(10, figure_score)

        # Structure parsing quality (0-10)
        # NOTE: Abstracts are not always present on posters (poster itself is the abstract)
        # Focus on core scientific sections: methods, results, conclusions
        structure_score = 0
        core_sections = ['methods', 'results', 'conclusions']
        optional_sections = ['abstract', 'objective', 'background']

        # Core sections: 3 points each (max 9 points)
        for section in core_sections:
            if sections.get(section) and len(sections[section]) > 50:
                structure_score += 3.0

        # Optional sections: bonus points if present (max 3 additional points)
        for section in optional_sections:
            if sections.get(section) and len(sections[section]) > 50:
                structure_score += 0.5

        scores['structure_parsing'] = min(10, structure_score)

        # NEW: Template contamination check
        template_penalty = 0
        contamination_patterns = [
            'Genigraphics®',
            'Change Color Theme',
            'DO NOT POST',
            'PLEASE DO NOT DISTRIBUTE',
            'www.genigraphics.com',
            'dating back to',
            'anyone in the industry',
            'return to that after trying',
            'PowerPoint® software'
        ]
        contamination_count = sum(1 for pattern in contamination_patterns if pattern.lower() in text.lower())

        if contamination_count > 0:
            template_penalty = min(2.0, contamination_count * 0.5)  # -0.5 per artifact, max -2.0
            logger.warning(f"⚠ Template contamination detected: {contamination_count} artifacts (-{template_penalty} points)")

        scores['template_contamination'] = contamination_count

        # NEW: Caption coherence check
        caption_score = 10.0
        caption_issues = 0
        if figures:
            for fig in figures:
                # Check caption field
                caption = fig.get('caption', fig.get('caption_from_text', ''))
                if caption and len(caption) > 100:  # Long captions should make sense
                    # Check for template contamination in caption
                    if any(pattern.lower() in caption.lower() for pattern in contamination_patterns):
                        caption_score -= 2.0
                        caption_issues += 1
                        logger.warning(f"⚠ Figure {fig.get('figure_number', '?')} caption contains template artifacts")

        caption_score = max(0, caption_score)
        scores['caption_quality'] = caption_score

        # NEW: Figure-caption consistency check (count unique figure numbers, not total mentions)
        vision_figure_count = len(figures)
        text_figure_numbers = set(re.findall(r'Figure\s+(\d+)[:\.]', text, re.IGNORECASE))
        text_caption_count = len(text_figure_numbers)

        consistency_penalty = 0
        if vision_figure_count > 0 and text_caption_count > 0:
            # Allow some tolerance (±1 figure)
            if abs(vision_figure_count - text_caption_count) > 1:
                consistency_penalty = abs(vision_figure_count - text_caption_count) * self.CONSISTENCY_PENALTY_PER_MISMATCH
                # Cap penalty to prevent excessive point loss
                consistency_penalty = min(self.MAX_CONSISTENCY_PENALTY, consistency_penalty)
                logger.warning(f"⚠ Figure-caption mismatch: Vision={vision_figure_count}, Text captions={text_caption_count} (-{consistency_penalty} points)")

        scores['figure_caption_consistency'] = consistency_penalty

        # Calculate overall score with penalties
        base_weights = {'text_extraction': 0.25, 'figure_analysis': 0.35, 'structure_parsing': 0.25, 'caption_quality': 0.15}
        base_overall = sum(scores.get(k, 0) * base_weights.get(k, 0) for k in base_weights if k in scores)

        # Apply penalties
        overall = base_overall - template_penalty - consistency_penalty
        overall = max(0, min(10, overall))  # Clamp to 0-10
        scores['overall'] = round(overall, 1)

        # Store penalties for transparency
        scores['template_penalty'] = template_penalty
        scores['consistency_penalty'] = consistency_penalty

        # Quality assessment (updated thresholds)
        if overall >= 8:
            scores['assessment'] = 'Excellent'
        elif overall >= self.QUALITY_GOOD_THRESHOLD:  # Now 5.5 instead of hardcoded 6
            scores['assessment'] = 'Good'
        elif overall >= 4:
            scores['assessment'] = 'Fair'
        else:
            scores['assessment'] = 'Poor - Consider manual review'

        # Log quality details with enhanced debugging
        logger.info(f"  Quality breakdown:")
        logger.info(f"    Text: {scores['text_extraction']:.1f}/10 (weight 0.25) = {scores['text_extraction'] * 0.25:.2f}")
        logger.info(f"    Figures: {scores['figure_analysis']:.1f}/10 (weight 0.35) = {scores['figure_analysis'] * 0.35:.2f}")
        logger.info(f"    Structure: {scores['structure_parsing']:.1f}/10 (weight 0.25) = {scores['structure_parsing'] * 0.25:.2f}")
        logger.info(f"    Captions: {scores['caption_quality']:.1f}/10 (weight 0.15) = {scores['caption_quality'] * 0.15:.2f}")
        logger.info(f"    Base total: {base_overall:.2f}/10")
        if template_penalty > 0:
            logger.info(f"    Template penalty: -{template_penalty:.2f}")
        if consistency_penalty > 0:
            logger.info(f"    Consistency penalty: -{consistency_penalty:.2f}")
        logger.info(f"    FINAL SCORE: {overall:.1f}/10 (threshold: {self.QUALITY_GOOD_THRESHOLD})")

        return scores

    def validate_output_quality(self, text: str, figures: List[Dict], sections: Dict) -> List[str]:
        """Validate output quality and return list of issues

        Returns:
            List of quality issue strings (empty if no issues)
        """
        issues = []

        # Check text length
        if not text or len(text) < 500:
            issues.append("Very short text extraction (< 500 chars)")

        # Check fragmentation
        if text:
            words = text.split()
            if words:
                tiny_ratio = sum(1 for w in words if len(w) <= 2) / len(words)
                if tiny_ratio > 0.4:
                    issues.append("High text fragmentation detected")

        # Check for missing critical sections
        # NOTE: Abstracts are optional on posters (poster itself is the abstract)
        # Only warn about truly critical missing sections

        if not sections.get('results') or len(sections['results']) < 50:
            # Results are critical - should be present
            issues.append("Results section missing or very short")

        if not sections.get('methods') or len(sections['methods']) < 50:
            # Methods help understand the work
            if not sections.get('objective') and not sections.get('background'):
                issues.append("Methods and context sections (objective/background) missing")

        # Check figures
        if not figures:
            issues.append("No figures analyzed")
        elif len(figures) < 2:
            issues.append("Very few figures identified (expected more for typical poster)")

        # Check for figure descriptions
        if figures:
            missing_desc = [f for f in figures if not f.get('description') or len(f['description']) < 20]
            if missing_desc:
                issues.append(f"{len(missing_desc)} figures have inadequate descriptions")

        return issues

    def extract_sections_and_summary(self, text: str, figures: List[Dict],
                                     structure: Dict = None) -> Tuple[Dict, str]:
        """Extract poster sections and executive summary in a single API call.

        Combines what was previously two separate calls (section extraction + summary generation).

        Returns:
            Tuple of (sections_dict, executive_summary_string)
        """
        sections = {
            'title': '', 'authors': '', 'institution': '',
            'abstract': '', 'objective': '', 'results': '',
            'conclusions': '', 'methods': ''
        }

        client = self.get_anthropic_client()
        if not client:
            # Fallback to regex parsing
            sections = self.parse_poster_structure(text)
            return sections, ""

        # Build figure context
        figure_findings = ""
        if figures:
            findings = []
            for fig in figures[:5]:
                for f in fig.get('key_findings', [])[:2]:
                    findings.append(f)
            if findings:
                figure_findings = "\nKey figure findings: " + "; ".join(findings[:8])

        section_names_hint = ""
        if structure and structure.get('sections'):
            names = []
            for s in structure['sections']:
                if isinstance(s, dict):
                    names.append(s.get('name', ''))
                else:
                    names.append(str(s))
            section_names_hint = f"\nIdentified sections on poster: {', '.join(names)}"

        prompt = f"""Analyze this scientific poster text and produce TWO outputs separated by "---SUMMARY---":

SECTION 1: Extract section content as JSON with these keys (empty string if not found):
- abstract, objective, methods, results, conclusions
{section_names_hint}

SECTION 2: A 2-3 sentence executive summary capturing the main research question, approach, and key finding.
{figure_findings}

Text:
---
{text[:15000]}
---

Output format:
```json
{{"abstract": "...", "objective": "...", "methods": "...", "results": "...", "conclusions": "..."}}
```

---SUMMARY---

[2-3 sentence executive summary here]"""

        try:
            message = client.messages.create(
                model=self.VISION_MODEL,
                max_tokens=4096,
                messages=[{"role": "user", "content": prompt}]
            )

            response = message.content[0].text
            logger.info(f"Sections + summary: {len(response)} chars "
                       f"(in:{message.usage.input_tokens}, out:{message.usage.output_tokens})")

            if "---SUMMARY---" in response:
                parts = response.split("---SUMMARY---", 1)
                sections_text = parts[0].strip()
                executive_summary = parts[1].strip()
            else:
                sections_text = response
                executive_summary = ""

            # Parse sections JSON
            parsed = self.parse_json_response(sections_text)
            if parsed:
                for key in ['abstract', 'objective', 'methods', 'results', 'conclusions']:
                    if key in parsed and parsed[key]:
                        sections[key] = parsed[key]
                logger.info(f"✓ Extracted sections: {', '.join(k for k, v in sections.items() if v)}")
            else:
                # Fallback to regex
                logger.warning("JSON parse failed for sections, using regex fallback")
                sections = self.parse_poster_structure(text)

            if executive_summary:
                logger.info(f"✓ Executive summary: {len(executive_summary)} chars")

            return sections, executive_summary

        except Exception as e:
            logger.error(f"Error in combined sections+summary: {e}")
            sections = self.parse_poster_structure(text)
            return sections, ""

    def generate_markdown(self, poster_num: str, metadata_row: Optional[Dict],
                         text: str, figures: List[Dict], extraction_method: str,
                         pdf_path: Path = None, structure: Dict = None,
                         executive_summary: str = None, quality_scores: Dict = None) -> str:
        """Generate enhanced structured markdown from poster content

        Args:
            poster_num: Poster number
            metadata_row: Row from Excel metadata
            text: Extracted text
            figures: List of figure dictionaries with Vision AI analysis
            extraction_method: Method used for text extraction
            pdf_path: Optional PDF path for additional processing
            structure: Optional poster structure from Vision AI
            executive_summary: Optional executive summary
            quality_scores: Optional quality scores

        Returns:
            Formatted markdown string
        """

        sections = self.parse_poster_structure(text)

        # Extract figure captions and positions
        figure_captions = self.extract_figure_captions(text)
        figure_positions = self.find_figure_positions(text, figures) if figures else {}

        md = []

        # YAML-style metadata header
        title = sections.get('title', '')
        if metadata_row:
            title_col = self.column_map['title']
            meta_title = metadata_row.get(title_col) or metadata_row.get('title')
            if meta_title and str(meta_title) != 'nan':
                title = re.sub(r'<[^>]+>', '', str(meta_title)).strip()
        if structure and structure.get('title'):
            title = structure['title']

        md.append("---\n")
        md.append(f"poster_number: {poster_num}\n")
        if metadata_row:
            interested_col = self.column_map['interested']
            if interested_col != self.DEFAULT_COLUMN_MAP['interested']:
                cols_to_try = [interested_col] if isinstance(interested_col, str) else interested_col
            else:
                cols_to_try = ['Interested Colleagues', 'interested_people', 'Interested People', 'interested', 'contacts']
            for col in cols_to_try:
                if col in metadata_row and metadata_row[col] and str(metadata_row[col]) != 'nan':
                    md.append(f"interested_colleagues: {metadata_row[col]}\n")
                    break

            cov_col = self.column_map['covered_by']
            if cov_col in metadata_row and metadata_row[cov_col] and str(metadata_row[cov_col]) != 'nan':
                md.append(f"covered_by: {metadata_row[cov_col]}\n")

            session_fields = [
                ('session_number', 'session_number'),
                ('session_title', 'session_title'),
                ('session_type', 'session_type'),
                ('day', 'day'),
                ('session_start', 'session_start'),
                ('session_end', 'session_end'),
                ('location', 'location'),
            ]
            for yaml_key, role in session_fields:
                col = self.column_map[role]
                if col in metadata_row and metadata_row[col] and str(metadata_row[col]) != 'nan':
                    md.append(f"{yaml_key}: {metadata_row[col]}\n")

        md.append(f"extraction_method: {extraction_method}\n")
        md.append(f"vision_model: {self.VISION_MODEL}\n")
        md.append(f"processing_date: {datetime.now().strftime('%Y-%m-%d')}\n")

        if quality_scores:
            md.append(f"quality_overall: {quality_scores['overall']}/10\n")
            md.append(f"quality_assessment: {quality_scores['assessment']}\n")

        md.append("---\n\n")

        # Title and header
        md.append(f"# {title}\n\n")

        md.append(f"**Poster Number**: #{poster_num}\n\n")

        authors = None
        affiliation = None
        if metadata_row:
            auth_col = self.column_map['authors']
            inst_col = self.column_map['institution']
            if auth_col in metadata_row and metadata_row[auth_col] and str(metadata_row[auth_col]) != 'nan':
                authors = metadata_row[auth_col]
            if inst_col in metadata_row and metadata_row[inst_col] and str(metadata_row[inst_col]) != 'nan':
                affiliation = metadata_row[inst_col]
        if not authors and structure and structure.get('authors'):
            authors = structure['authors']
        if not affiliation and structure and structure.get('affiliation'):
            affiliation = structure['affiliation']

        if authors:
            md.append(f"**Authors**: {authors}\n\n")
        if affiliation:
            md.append(f"**Affiliation**: {affiliation}\n\n")

        md.append("---\n\n")

        # Executive Summary
        if executive_summary:
            md.append("## Executive Summary\n\n")
            md.append(f"{executive_summary}\n\n")
            md.append("---\n\n")

        # Abstract
        if sections['abstract']:
            md.append("## Abstract\n\n")
            md.append(f"{sections['abstract']}\n\n")
            md.append("---\n\n")

        # Objective
        if sections['objective']:
            md.append("## Objective\n\n")
            md.append(f"{sections['objective']}\n\n")
            md.append("---\n\n")

        # Methods
        if sections['methods']:
            md.append("## Methods\n\n")
            md.append(f"{sections['methods']}\n\n")
            md.append("---\n\n")

        # Results with integrated figures
        md.append("## Results\n\n")

        # Add detailed figure analyses (main content)
        if figures:
            for fig in figures:
                md.append(self._format_figure_markdown(fig, figure_captions))
                md.append("---\n\n")
        elif sections.get('results'):
            # No figures, just show results text
            md.append(f"{sections['results']}\n\n")
            md.append("---\n\n")

        # Conclusions
        if sections['conclusions']:
            md.append("## Conclusions\n\n")
            md.append(f"{sections['conclusions']}\n\n")
            md.append("---\n\n")

        # Quality Validation
        if quality_scores:
            md.append("## Quality Assessment\n\n")
            md.append(f"**Overall Quality**: {quality_scores['overall']}/10 - {quality_scores['assessment']}\n\n")
            md.append("**Component Scores**:\n")
            md.append(f"- Text Extraction: {quality_scores['text_extraction']}/10\n")
            md.append(f"- Figure Analysis: {quality_scores['figure_analysis']}/10\n")
            md.append(f"- Structure Parsing: {quality_scores['structure_parsing']}/10\n\n")

            # Add quality issues if any
            issues = self.validate_output_quality(text, figures, sections)
            if issues:
                md.append("**Quality Issues Detected**:\n")
                for issue in issues:
                    md.append(f"- {issue}\n")
                md.append("\n")

            md.append("---\n\n")

        # Full text
        md.append("## Full Extracted Text\n\n")
        md.append("<details>\n<summary>Click to expand full text</summary>\n\n")
        md.append(f"```\n{text}\n```\n\n")
        md.append("</details>\n\n")

        # Processing Metadata
        md.append("---\n\n")
        md.append("## Processing Metadata\n\n")
        md.append(f"- **Extraction Method**: {extraction_method}\n")
        md.append(f"- **Vision Model**: {self.VISION_MODEL}\n")
        if figures:
            md.append(f"- **Number of Figures**: {len(figures)}\n")
            has_detailed = sum(1 for f in figures if 'detailed_analysis' in f)
            if has_detailed:
                md.append(f"  - With detailed analysis: {has_detailed}\n")
        md.append(f"- **Text Length**: {len(text)} characters, {len(text.split())} words\n")
        md.append(f"- **Processing Date**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")

        return ''.join(md)

    def _interleave_figures_with_text(self, results_text: str, figures: List[Dict],
                                      figure_positions: Dict[str, Dict],
                                      figure_captions: Dict[str, str],
                                      full_text: str) -> str:
        """Interleave figures with Results text at their mention positions

        Splits the Results text and inserts figure descriptions where mentioned

        Args:
            results_text: Results section text
            figures: List of figure dictionaries
            figure_positions: Dictionary mapping figure numbers to position info
            figure_captions: Extracted figure captions
            full_text: Full poster text (for better position matching)

        Returns:
            Markdown string with figures interleaved in text
        """
        # Build list of (position, content_type, content) tuples
        content_items = []

        # Add all figure positions
        for fig in figures:
            fig_num_raw = fig.get('figure_number', '?')
            fig_num = re.sub(r'^Figure\s+', '', str(fig_num_raw), flags=re.IGNORECASE)

            if fig_num in figure_positions:
                pos = figure_positions[fig_num]['position']
                content_items.append((pos, 'figure', fig))

        # Sort by position
        content_items.sort(key=lambda x: x[0])

        # Split text into segments between figures
        # Find all figure mention positions in results_text
        figure_mentions = []
        for fig in figures:
            fig_num_raw = fig.get('figure_number', '?')
            fig_num = re.sub(r'^Figure\s+', '', str(fig_num_raw), flags=re.IGNORECASE)

            patterns = [
                rf'\bFig\.\s*{re.escape(fig_num)}\b',
                rf'\bFigure\s+{re.escape(fig_num)}\b',
                rf'\bFig\s+{re.escape(fig_num)}\b',
            ]

            for pattern in patterns:
                for match in re.finditer(pattern, results_text, re.IGNORECASE):
                    figure_mentions.append({
                        'fig_num': fig_num,
                        'position': match.start(),
                        'match_end': match.end()
                    })

        # Sort mentions by position
        figure_mentions.sort(key=lambda x: x['position'])

        # Build interleaved content
        md_parts = []
        last_pos = 0

        for mention in figure_mentions:
            fig_num = mention['fig_num']

            # Add text before this figure mention
            text_segment = results_text[last_pos:mention['position']].strip()
            if text_segment:
                md_parts.append(f"{text_segment}\n\n")

            # Add the figure
            fig = next((f for f in figures if re.sub(r'^Figure\s+', '', str(f.get('figure_number', '')), flags=re.IGNORECASE) == fig_num), None)
            if fig:
                md_parts.append(self._format_figure_markdown(fig, figure_captions))
                md_parts.append("---\n\n")

            # Move position forward past the figure mention
            last_pos = mention['match_end']

        # Add remaining text after last figure
        remaining_text = results_text[last_pos:].strip()
        if remaining_text:
            md_parts.append(f"{remaining_text}\n\n")
            md_parts.append("---\n\n")

        # Handle figures not mentioned in text
        mentioned_nums = [m['fig_num'] for m in figure_mentions]
        unmentioned_figs = [
            f for f in figures
            if re.sub(r'^Figure\s+', '', str(f.get('figure_number', '')), flags=re.IGNORECASE) not in mentioned_nums
        ]

        if unmentioned_figs:
            md_parts.append("### Additional Figures\n\n")
            for fig in unmentioned_figs:
                md_parts.append(self._format_figure_markdown(fig, figure_captions))
                md_parts.append("---\n\n")

        return ''.join(md_parts)

    def _format_figure_markdown(self, fig: Dict, figure_captions: Dict[str, str]) -> str:
        """Format a single figure as markdown

        Args:
            fig: Figure dictionary
            figure_captions: Dictionary of extracted captions

        Returns:
            Formatted markdown string
        """
        md = []

        fig_num_raw = fig.get('figure_number', '?')
        # Extract just the number if Vision AI returned "Figure 1" instead of "1"
        fig_num = re.sub(r'^Figure\s+', '', str(fig_num_raw), flags=re.IGNORECASE)
        fig_title = fig.get('title', 'Untitled')
        fig_desc = fig.get('description', '')
        fig_type = fig.get('figure_type', 'Unknown')
        key_findings = fig.get('key_findings', [])
        stat_notes = fig.get('statistical_notes', '')
        detailed = fig.get('detailed_analysis', '')

        # Figure header with anchor
        md.append(f"<a name=\"figure-{fig_num}\"></a>\n")
        md.append(f"### Figure {fig_num}: {fig_title}\n\n")

        # Add OCR caption if different from Vision title
        ocr_caption = figure_captions.get(str(fig_num), '')
        if ocr_caption and ocr_caption.lower() != fig_title.lower():
            md.append(f"**Caption from text**: {ocr_caption}\n\n")

        md.append(f"**Type**: {fig_type}\n\n")
        md.append(f"**Description**: {fig_desc}\n\n")

        if key_findings:
            md.append(f"**Key Findings**:\n")
            for finding in key_findings:
                md.append(f"- {finding}\n")
            md.append("\n")

        if stat_notes:
            md.append(f"**Statistical Notes**: {stat_notes}\n\n")

        if detailed:
            md.append(f"**Detailed Analysis**:\n\n{detailed}\n\n")

        return ''.join(md)

    def _extract_first_author_lastname(self, authors_str: str) -> str:
        """Extract last name of first author from an authors string."""
        if not authors_str or str(authors_str) == 'nan':
            return "Unknown"
        first_author = authors_str.split(',')[0].strip()
        parts = first_author.split()
        if not parts:
            return "Unknown"
        lastname = parts[-1].rstrip('¹²³⁴⁵⁶⁷⁸⁹⁰*†‡§')
        return re.sub(r'[^a-zA-Z]', '', lastname) or "Unknown"

    def _sanitize_filename_part(self, text: str) -> str:
        """Remove special characters and join words with underscores."""
        clean = re.sub(r'[^a-zA-Z0-9\s]', '', text)
        return '_'.join(clean.split())

    def generate_filename(self, poster_num: str, metadata_row: Optional[Dict], structure: Dict = None) -> str:
        """Generate output filename based on naming scheme."""
        if self.naming == 'default':
            return f"poster_{poster_num}.md"

        # Standardized naming: Poster_{Author}_{Conference}_{Year}_{Title}.md
        authors_str = None
        auth_col = self.column_map['authors']
        if metadata_row and metadata_row.get(auth_col) and str(metadata_row.get(auth_col)) != 'nan':
            authors_str = metadata_row[auth_col]
        if not authors_str and structure and structure.get('authors'):
            authors_str = structure['authors']
        author = self._extract_first_author_lastname(authors_str)

        conference = self._sanitize_filename_part(self.conference) if self.conference else "NoConference"

        year = self.year
        if not year and metadata_row:
            for col in [self.column_map['day'], self.column_map['session_start']]:
                val = metadata_row.get(col)
                if val and str(val) != 'nan':
                    year_match = re.search(r'20\d{2}', str(val))
                    if year_match:
                        year = year_match.group(0)
                        break
        if not year:
            year = "NoYear"

        title = None
        if structure and structure.get('title'):
            title = structure['title']
        elif metadata_row:
            title_col = self.column_map['title']
            meta_title = metadata_row.get(title_col) or metadata_row.get('title')
            if meta_title and str(meta_title) != 'nan':
                title = re.sub(r'<[^>]+>', '', str(meta_title)).strip()
        if not title:
            title = poster_num

        title_clean = re.sub(r'[^a-zA-Z0-9\s]', '', title)
        title_words = title_clean.split()[:5]
        title_slug = '_'.join(w for w in title_words) or poster_num

        return f"Poster_{author}_{conference}_{year}_{title_slug}.md"

    def process_single_poster(self, pdf_path: Path, metadata_df=None,
                             force_ocr: bool = False, skip_existing: bool = True,
                             ocr_dpi: int = 200, enable_detailed_analysis: bool = True) -> bool:
        """Process a single poster PDF with enhanced pipeline

        Args:
            pdf_path: Path to PDF file
            metadata_df: DataFrame with metadata
            force_ocr: Force OCR even if native text available
            skip_existing: Skip if already processed
            ocr_dpi: DPI for Tesseract OCR
            enable_detailed_analysis: Enable two-stage figure analysis

        Returns:
            True if successful, False otherwise
        """

        logger.info(f"\n{'='*60}")
        logger.info(f"Processing: {pdf_path.name}")
        logger.info(f"{'='*60}")

        # File size check
        file_size_mb = pdf_path.stat().st_size / (1024 * 1024)
        if file_size_mb > self.MAX_PDF_SIZE_MB:
            logger.error(f"  File too large: {file_size_mb:.1f}MB (max {self.MAX_PDF_SIZE_MB}MB) — skipping")
            return False

        # Extract poster number
        poster_num = self.extract_poster_number(pdf_path)
        if not poster_num:
            logger.warning(f"Could not extract poster number from {pdf_path.name}")
            self.unrecognized_posters.append(pdf_path.name)
            return False

        logger.info(f"File '{pdf_path.name}' -> Poster {poster_num}")

        # Check if already processed (by poster number)
        if skip_existing and self.check_if_processed(poster_num):
            logger.info(f"Poster {poster_num} already processed - skipping")
            return True

        # Get metadata (if available) using O(1) lookup
        metadata_row = None
        if metadata_df is not None:
            if hasattr(self, '_poster_lookup') and self._poster_lookup:
                metadata_row = self._poster_lookup.get(str(poster_num))
                if metadata_row:
                    logger.info(f"✓ Found metadata for poster {poster_num}")
                else:
                    logger.warning(f"⚠ Poster {poster_num} not found in metadata lookup")
            else:
                # Fallback if lookup wasn't built (e.g., no metadata file)
                possible_columns = self.column_map['poster_number']
                if isinstance(possible_columns, str):
                    possible_columns = [possible_columns]
                for col_name in possible_columns:
                    if col_name in metadata_df.columns:
                        matches = metadata_df[metadata_df[col_name].astype(str).str.strip() == str(poster_num)]
                        if not matches.empty:
                            metadata_row = matches.iloc[0].to_dict()
                            logger.info(f"✓ Found metadata for poster {poster_num} in '{col_name}' column")
                            break
                if not metadata_row:
                    logger.warning(f"⚠ Poster {poster_num} not found in Excel metadata")
        else:
            logger.warning("⚠ No metadata DataFrame provided")

        # PERFORMANCE OPTIMIZATION: Convert PDF to image once and encode to base64
        logger.info("\n--- IMAGE CONVERSION & ENCODING (One-time) ---")
        poster_image = self.convert_pdf_to_image(pdf_path, dpi=self.DEFAULT_VISION_DPI)

        if not poster_image:
            logger.error("Failed to convert PDF to image")
            return False

        logger.info(f"✓ Image converted: {self.DEFAULT_VISION_DPI} DPI ({poster_image.size[0]}x{poster_image.size[1]})")

        # Encode once — reused by all Vision AI calls
        encoded_image = self.encode_image_to_base64(poster_image, format="JPEG")
        if encoded_image[0] is None:
            logger.error("Failed to encode image to base64")
            return False
        logger.info(f"✓ Image encoded to base64 ({len(encoded_image[0]) / 1024:.0f} KB)")
        del poster_image  # Free PIL image memory

        # STEPS 1+2 (parallel): Structure extraction (API) + Base text extraction (local)
        logger.info("\n--- STEPS 1+2: Structure + Text Extraction (parallel) ---")

        def _extract_base_text():
            ref_text = None
            meth = "unknown"
            if force_ocr:
                logger.info("Force OCR mode enabled")
                ref_text, meth = self.extract_text_ocr(pdf_path, dpi=ocr_dpi)
            else:
                ref_text, meth = self.extract_text_native(pdf_path)
                if meth == "native_low_quality" or not ref_text or len(ref_text.strip()) < 100:
                    logger.warning("Native extraction insufficient - attempting OCR...")
                    ocr_text, ocr_method = self.extract_text_ocr(pdf_path, dpi=ocr_dpi)
                    if ocr_text and len(ocr_text) > len(ref_text or ""):
                        ref_text, meth = ocr_text, ocr_method
                        logger.info(f"✓ OCR produced better results: {len(ref_text)} chars")
            return ref_text, meth

        with ThreadPoolExecutor(max_workers=2) as executor:
            structure_future = executor.submit(self.extract_poster_structure_vision, pdf_path, encoded_image=encoded_image)
            text_future = executor.submit(_extract_base_text)

            try:
                structure = structure_future.result()
            except Exception as e:
                logger.error(f"Structure extraction failed: {e}")
                structure = {}

            try:
                reference_text, method = text_future.result()
            except Exception as e:
                logger.error(f"Text extraction failed: {e}")
                return False

        if not reference_text or len(reference_text.strip()) == 0:
            logger.error("Failed to extract any base text (native/OCR)")
            return False

        logger.info(f"✓ Base extraction: {len(reference_text)} chars using {method}")

        # PHASE 2: MANDATORY Vision AI RAG enhancement
        # RAG = Retrieval-Augmented Generation: Uses extracted text as retrieval context
        # to guide Vision AI in correcting reading order, OCR errors, and removing artifacts
        logger.info("Phase 2.2: MANDATORY Vision AI RAG enhancement...")
        logger.info(f"  RAG context: {len(reference_text)} chars from {method} extraction")

        enhanced_text, enhanced_method = self.extract_text_vision_with_rag(
            pdf_path,
            reference_text=reference_text,
            original_method=method,
            encoded_image=encoded_image
        )

        if enhanced_text and len(enhanced_text.strip()) > 0:
            text, method = enhanced_text, enhanced_method
            # Detailed characterization already logged in extract_text_vision_with_rag()
        else:
            logger.warning("⚠ Vision AI enhancement failed - using base extraction")
            text, method = reference_text, f"{method}_no_vision"

        # PHASE 3: Template artifact removal
        logger.info("Phase 2.3: Cleaning template artifacts...")
        cleaned_text = self.remove_template_artifacts(text)
        if len(cleaned_text) < len(text):
            text = cleaned_text
            logger.info(f"✓ Text cleaned: {len(text)} chars remain")

        if not text or len(text.strip()) == 0:
            logger.error(f"Failed to extract any text from {pdf_path.name}")
            return False

        logger.info(f"✓ Text extraction complete: {len(text)} chars, final method: {method}")

        # STEP 3: Figure analysis (MANDATORY Vision AI)
        logger.info("\n--- STEP 3: Figure Analysis ---")

        if enable_detailed_analysis:
            figures = self.analyze_figures_two_stage(pdf_path, poster_num, extracted_text=text, encoded_image=encoded_image)
        else:
            figures = self.analyze_figures_with_vision(pdf_path, poster_num, extracted_text=text, encoded_image=encoded_image)

        if not figures:
            logger.warning("⚠ No figures identified by Vision AI")
        else:
            logger.info(f"✓ {len(figures)} figures analyzed (captions included from Stage 1)")

        # STEP 4: Section extraction + Executive summary (combined)
        logger.info("\n--- STEP 4: Sections & Executive Summary ---")
        sections, executive_summary = self.extract_sections_and_summary(text, figures, structure)

        # STEP 5: Quality assessment
        logger.info("\n--- STEP 5: Quality Assessment ---")
        quality_scores = self.calculate_quality_score(text, figures, sections, method)
        logger.info(f"✓ Quality score: {quality_scores['overall']}/10 - {quality_scores['assessment']}")

        issues = self.validate_output_quality(text, figures, sections)
        if issues:
            logger.warning("Quality issues detected:")
            for issue in issues:
                logger.warning(f"  - {issue}")

        # Check quality threshold - skip markdown generation if POOR or FAIR
        if quality_scores['assessment'] in ['Poor - Consider manual review', 'Fair']:
            logger.warning(f"⚠ Quality {quality_scores['assessment']} - SKIPPING markdown generation")
            logger.warning(f"   Poster {poster_num} requires manual review or better extraction")

            # Log to quality log
            self.log_quality(poster_num, quality_scores, saved=False, filename=pdf_path.name)

            # Track failed poster
            self.failed_posters.append({
                'poster_num': poster_num,
                'quality': quality_scores['overall'],
                'assessment': quality_scores['assessment'],
                'reason': 'Quality too low'
            })

            return False

        # STEP 6: Generate markdown
        logger.info("\n--- STEP 6: Markdown Generation ---")
        markdown = self.generate_markdown(
            poster_num=poster_num,
            metadata_row=metadata_row,
            text=text,
            figures=figures,
            extraction_method=method,
            pdf_path=pdf_path,
            structure=structure,
            executive_summary=executive_summary,
            quality_scores=quality_scores
        )

        # Save output
        output_filename = self.generate_filename(poster_num, metadata_row, structure=structure)
        output_path = self.output_dir / output_filename

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(markdown)
        self._existing_files.add(output_path.stem)

        logger.info(f"\n✓ Successfully saved to: {output_path}")
        logger.info(f"  - Text: {len(text)} chars ({len(text.split())} words)")
        logger.info(f"  - Figures: {len(figures)}")
        logger.info(f"  - Method: {method}")
        logger.info(f"  - Quality: {quality_scores['overall']}/10")

        # Log to quality log
        self.log_quality(poster_num, quality_scores, saved=True, filename=pdf_path.name)

        return True

    def process_all_posters(self, force_ocr: bool = False, skip_existing: bool = True,
                           ocr_dpi: int = 200, enable_detailed_analysis: bool = True):
        """Process all posters in the SharePoint folder

        Args:
            force_ocr: Force OCR for all files
            skip_existing: Skip already processed files
            ocr_dpi: DPI for Tesseract OCR
            enable_detailed_analysis: Enable two-stage figure analysis
        """

        logger.info("="*70)
        logger.info("ENHANCED POSTER PROCESSING PIPELINE")
        logger.info("="*70)
        logger.info(f"SharePoint folder: {self.sharepoint_folder}")
        logger.info(f"Output directory: {self.output_dir}")
        logger.info(f"Vision AI: MANDATORY ({self.VISION_MODEL})")
        logger.info(f"OCR DPI: {ocr_dpi}")
        logger.info(f"Detailed figure analysis: {'ENABLED' if enable_detailed_analysis else 'DISABLED'}")
        logger.info("="*70)

        # Initialize quality log file with header (only if file doesn't exist or is empty)
        if not self.quality_log_path.exists() or self.quality_log_path.stat().st_size == 0:
            with open(self.quality_log_path, 'w', encoding='utf-8') as f:
                f.write("TIMESTAMP\tPOSTER_NUM\tQUALITY_SCORE\tASSESSMENT\tSTATUS\tFILENAME\n")
            logger.info(f"Created new quality log: {self.quality_log_path}")
        else:
            logger.info(f"Appending to existing quality log: {self.quality_log_path}")

        # Load metadata (optional - pipeline continues without it)
        metadata_df = self.load_metadata()

        # Get all PDFs
        pdf_files = self.list_pdf_files()

        if not pdf_files:
            logger.error("No PDF files found")
            return

        # Process each poster
        success_count = 0
        fail_count = 0

        for i, pdf_path in enumerate(pdf_files):
            logger.info(f"\n{'='*70}")
            logger.info(f"POSTER {i+1}/{len(pdf_files)}")
            logger.info(f"{'='*70}")

            try:
                success = self.process_single_poster(
                    pdf_path,
                    metadata_df,
                    force_ocr=force_ocr,
                    skip_existing=skip_existing,
                    ocr_dpi=ocr_dpi,
                    enable_detailed_analysis=enable_detailed_analysis
                )
                if success:
                    success_count += 1
                else:
                    fail_count += 1
            except Exception as e:
                logger.error(f"Unexpected error processing {pdf_path.name}: {e}")
                import traceback
                logger.error(traceback.format_exc())
                fail_count += 1

        # Summary
        logger.info(f"\n{'='*70}")
        logger.info("PROCESSING COMPLETE")
        logger.info(f"{'='*70}")
        logger.info(f"Total files: {len(pdf_files)}")
        logger.info(f"Successful: {success_count}")
        logger.info(f"Failed: {fail_count}")
        logger.info(f"Success rate: {success_count/len(pdf_files)*100:.1f}%")
        logger.info(f"Output directory: {self.output_dir}")

        # Report failed posters (low quality)
        if self.failed_posters:
            logger.info(f"\n{'='*70}")
            logger.info(f"⚠ QUALITY ISSUES - {len(self.failed_posters)} POSTER(S) SKIPPED")
            logger.info(f"{'='*70}")
            logger.info("The following posters did not meet quality threshold (FAIR/POOR):")
            logger.info("These require manual review or better extraction:")
            logger.info("")
            for failed in self.failed_posters:
                logger.info(f"  • Poster {failed['poster_num']}: Quality {failed['quality']}/10 ({failed['assessment']})")
            logger.info("")
            logger.info(f"See {self.quality_log_path} for complete quality log")
            logger.info(f"{'='*70}")

        # Report unrecognized posters (filename could not be matched to metadata)
        if self.unrecognized_posters:
            logger.info(f"\n{'='*70}")
            logger.info(f"UNRECOGNIZED POSTERS - {len(self.unrecognized_posters)} FILE(S) NOT MATCHED TO METADATA")
            logger.info(f"{'='*70}")
            logger.info("The following files could not be matched to any poster number:")
            logger.info("")
            for filename in self.unrecognized_posters:
                logger.info(f"  • {filename}")
            logger.info("")
            logger.info(f"{'='*70}")

            # Append to quality log
            with open(self.quality_log_path, 'a', encoding='utf-8') as f:
                for filename in self.unrecognized_posters:
                    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    f.write(f"{timestamp}\t{filename}\t-\tUNRECOGNIZED\tSkipped\n")

        logger.info(f"{'='*70}")


def main():
    """Main entry point for enhanced poster processing pipeline

    Vision AI is now MANDATORY for all processing.
    """
    import argparse

    parser = argparse.ArgumentParser(
        description='Process scientific posters to structured markdown with AI',
        epilog='Note: Vision AI (Claude) is required for all processing.'
    )
    parser.add_argument('--input', '--sharepoint', dest='sharepoint', type=str, required=True,
                       help='Path to folder containing poster PDFs')
    parser.add_argument('--metadata', type=str, required=False, default=None,
                       help='Path to Excel metadata file (optional, enriches output with session info)')
    parser.add_argument('--output', type=str, default='output',
                       help='Output directory for markdown files (default: output)')
    parser.add_argument('--force-ocr', action='store_true',
                       help='Force OCR even if native text extraction works')
    parser.add_argument('--no-skip', action='store_true',
                       help='Reprocess files that already exist')
    parser.add_argument('--single', type=str,
                       help='Process single PDF file (provide full path)')
    parser.add_argument('--ocr-dpi', type=int, default=200,
                       help='DPI for Tesseract OCR (default: 200, higher=better quality but slower)')
    parser.add_argument('--no-detailed-analysis', action='store_true',
                       help='Disable two-stage detailed figure analysis (faster but less detail)')
    parser.add_argument('--recursive', action='store_true',
                       help='Recursively search subfolders for PDF files')
    parser.add_argument('--naming', default='default',
                       choices=['default', 'standardized'],
                       help="Output filename scheme: "
                            "'default' = poster_{NUM}; "
                            "'standardized' = Poster_{Author}_{Conference}_{Year}_{Title}")
    parser.add_argument('--conference', type=str, default=None,
                       help='Conference name for standardized naming (e.g., AACR, ASCO)')
    parser.add_argument('--year', type=str, default=None,
                       help='Year for standardized naming (e.g., 2026)')
    parser.add_argument('--verbose', action='store_true',
                       help='Enable verbose/debug logging')

    col_group = parser.add_argument_group('Metadata Column Configuration',
        'Map logical roles to actual Excel column names. Only specify columns that differ from defaults.')
    col_group.add_argument('--sheet', type=str, default=None,
                           help='Excel sheet name (default: Full_Program_Copy)')
    col_group.add_argument('--col-poster-number', type=str, default=None,
                           help='Comma-separated columns for poster number matching '
                                '(default: "Presentation Number,Poster Number,Abstract Number,Session Number")')
    col_group.add_argument('--col-title', type=str, default=None,
                           help='Title column (default: "Presentation Title")')
    col_group.add_argument('--col-authors', type=str, default=None,
                           help='Authors column (default: "authors")')
    col_group.add_argument('--col-institution', type=str, default=None,
                           help='Institution/affiliation column (default: "institution")')
    col_group.add_argument('--col-session-number', type=str, default=None,
                           help='Session number column (default: "Session Number")')
    col_group.add_argument('--col-session-title', type=str, default=None,
                           help='Session title column (default: "Session Title")')
    col_group.add_argument('--col-session-type', type=str, default=None,
                           help='Session type column (default: "Session Type Name")')
    col_group.add_argument('--col-day', type=str, default=None,
                           help='Day/date column (default: "Day")')
    col_group.add_argument('--col-session-start', type=str, default=None,
                           help='Session start time column (default: "Session Start")')
    col_group.add_argument('--col-session-end', type=str, default=None,
                           help='Session end time column (default: "Session End")')
    col_group.add_argument('--col-location', type=str, default=None,
                           help='Location column (default: "Location")')
    col_group.add_argument('--col-covered-by', type=str, default=None,
                           help='Coverage person column (default: "Covered by")')
    col_group.add_argument('--col-interested', type=str, default=None,
                           help='Interested people column (default: "Interested Colleagues")')

    args = parser.parse_args()

    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    # Validate paths
    try:
        validate_path(args.sharepoint, must_exist=True, allow_dir=True, allow_file=False)
        validate_output_path(args.output)
        if args.metadata:
            validate_path(args.metadata, must_exist=True, allow_file=True, allow_dir=False)
        if args.single:
            validate_path(args.single, must_exist=True, allow_file=True, allow_dir=False)
    except (ValueError, FileNotFoundError) as e:
        logger.error(f"Path validation failed: {e}")
        return

    # Build column overrides from CLI args
    column_overrides = {}
    cli_to_role = {
        'col_poster_number': 'poster_number',
        'col_title': 'title',
        'col_authors': 'authors',
        'col_institution': 'institution',
        'col_session_number': 'session_number',
        'col_session_title': 'session_title',
        'col_session_type': 'session_type',
        'col_day': 'day',
        'col_session_start': 'session_start',
        'col_session_end': 'session_end',
        'col_location': 'location',
        'col_covered_by': 'covered_by',
        'col_interested': 'interested',
    }
    for cli_attr, role in cli_to_role.items():
        val = getattr(args, cli_attr, None)
        if val is not None:
            if role == 'poster_number':
                column_overrides[role] = [v.strip() for v in val.split(',')]
            else:
                column_overrides[role] = val

    try:
        # Initialize pipeline
        logger.info("Initializing Enhanced Poster Pipeline...")
        pipeline = PosterPipeline(
            sharepoint_folder=args.sharepoint,
            metadata_excel=args.metadata,
            output_dir=args.output,
            recursive=args.recursive,
            naming=args.naming,
            conference=args.conference,
            year=args.year,
            sheet=args.sheet,
            column_overrides=column_overrides if column_overrides else None,
        )

        # Process
        if args.single:
            logger.info(f"Single file mode: {args.single}")
            pdf_path = Path(args.single)
            if not pdf_path.exists():
                logger.error(f"File not found: {pdf_path}")
                return

            # Single mode: always overwrite existing files
            pipeline.process_single_poster(
                pdf_path,
                pipeline.load_metadata(),
                force_ocr=args.force_ocr,
                skip_existing=False,  # Always process in single mode
                ocr_dpi=args.ocr_dpi,
                enable_detailed_analysis=not args.no_detailed_analysis
            )
        else:
            logger.info("Batch processing mode")
            # Batch mode: skip existing by default (use --no-skip to reprocess)
            pipeline.process_all_posters(
                force_ocr=args.force_ocr,
                skip_existing=not args.no_skip,
                ocr_dpi=args.ocr_dpi,
                enable_detailed_analysis=not args.no_detailed_analysis
            )

    except Exception as e:
        logger.error(f"Pipeline failed: {e}")
        import traceback
        logger.error(traceback.format_exc())
        raise


if __name__ == '__main__':
    main()
