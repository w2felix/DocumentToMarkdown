"""
Scientific Paper Processing Pipeline
Converts scientific research paper PDFs to structured markdown with AI-powered analysis

Features:
- Multi-page PDF handling with page classification
- Text-first extraction (pdfplumber), Vision AI only for metadata + figures
- Automatic metadata extraction (title, authors, DOI, journal, year)
- IMRaD section detection (regex-first, AI fallback)
- Selective Vision AI for figure-heavy pages (batched)
- Table extraction via pdfplumber
- Quality scoring adapted for research papers
- Standardized naming: FirstAuthor_Year_TitleSlug.md
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

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

from pipeline_security import validate_path, validate_output_path, sanitize_filename
from auth import get_anthropic_client as _get_shared_client


class PaperPipeline:
    """Pipeline for processing scientific paper PDFs into structured markdown"""

    VISION_MODEL = "claude-sonnet-4-6"
    PAGES_PER_VISION_BATCH = 4
    RENDER_DPI = 200
    MAX_IMAGE_DIMENSION = 1568
    DEFAULT_OCR_DPI = 200
    MIN_TEXT_FOR_TEXT_PAGE = 100
    MAX_VISION_PAGES = 8

    QUALITY_EXCELLENT_THRESHOLD = 8.0
    QUALITY_GOOD_THRESHOLD = 5.5
    QUALITY_FAIR_THRESHOLD = 4.0

    MAX_PDF_SIZE_MB = 200
    DEFAULT_BUDGET_CAP = 50

    NAMING_SCHEMES = {
        'default': '{first_author}_{year}_{title_slug}',
        'detailed': '{first_author}_et_al_{year}_{journal}_{title_slug}',
        'doi': '{doi_suffix}',
    }

    SECTION_PATTERNS = {
        'abstract': re.compile(
            r'^\s*(?:ABSTRACT|Abstract)\s*\n', re.MULTILINE),
        'introduction': re.compile(
            r'^\s*(?:\d+\.?\s*)?(?:INTRODUCTION|Introduction|BACKGROUND|Background)\s*\n', re.MULTILINE),
        'methods': re.compile(
            r'^\s*(?:\d+\.?\s*)?(?:METHODS?|MATERIALS?\s+AND\s+METHODS?|EXPERIMENTAL(?:\s+PROCEDURES?)?|METHODOLOGY|STUDY\s+DESIGN|Methods?|Materials?\s+and\s+Methods?|Online\s+methods?)\s*\n', re.MULTILINE),
        'results': re.compile(
            r'^\s*(?:\d+\.?\s*)?(?:RESULTS?|FINDINGS|Results?)\s*\n', re.MULTILINE),
        'discussion': re.compile(
            r'^\s*(?:\d+\.?\s*)?(?:DISCUSSION|Discussion|INTERPRETATION)\s*\n', re.MULTILINE),
        'conclusions': re.compile(
            r'^\s*(?:\d+\.?\s*)?(?:CONCLUSIONS?|CONCLUDING\s+REMARKS?|SUMMARY|Conclusions?|Summary)\s*\n', re.MULTILINE),
        'references': re.compile(
            r'^\s*(?:REFERENCES|BIBLIOGRAPHY|LITERATURE\s+CITED|References|Bibliography)\s*\n', re.MULTILINE),
        'acknowledgements': re.compile(
            r'^\s*(?:ACKNOWLEDGEMENTS?|ACKNOWLEDGMENTS?|FUNDING|Acknowledgements?)\s*\n', re.MULTILINE),
    }

    # Relaxed patterns for papers with inline section headers (e.g., Nature articles)
    SECTION_PATTERNS_RELAXED = {
        'abstract': re.compile(
            r'(?:^|\n)\s*(?:ABSTRACT|Abstract)[.\s]', re.MULTILINE),
        'introduction': re.compile(
            r'(?:^|\n)\s*(?:\d+\.?\s*)?(?:INTRODUCTION|Introduction|BACKGROUND)[.\s]', re.MULTILINE),
        'methods': re.compile(
            r'(?:^|\n)\s*(?:\d+\.?\s*)?(?:METHODS?|Materials?\s+and\s+methods?|Online\s+methods?|EXPERIMENTAL)[.\s]', re.MULTILINE),
        'results': re.compile(
            r'(?:^|\n)\s*(?:\d+\.?\s*)?(?:RESULTS?|Results?)[.\s]', re.MULTILINE),
        'discussion': re.compile(
            r'(?:^|\n)\s*(?:\d+\.?\s*)?(?:DISCUSSION|Discussion)[.\s]', re.MULTILINE),
        'conclusions': re.compile(
            r'(?:^|\n)\s*(?:\d+\.?\s*)?(?:CONCLUSIONS?|Conclusions?)[.\s]', re.MULTILINE),
        'references': re.compile(
            r'(?:^|\n)\s*(?:REFERENCES|References)\s*\n', re.MULTILINE),
    }

    FILENAME_PATTERN = re.compile(r'^(.+?)\s+et\s+al\.\s*[-–—]\s*(\d{4})\s*[-–—]\s*(.+)$')
    FILENAME_PATTERN_SIMPLE = re.compile(r'^(.+?)\s*[-–—]\s*(\d{4})\s*[-–—]\s*(.+)$')

    DOI_PATTERN = re.compile(r'(?:doi|DOI|https?://doi\.org/)[\s:]*(10\.\d{4,}/[^\s,;]+)')
    YEAR_PATTERN = re.compile(r'\b(19[89]\d|20[012]\d)\b')

    def __init__(self, input_folder: str, output_dir: str = "output_papers",
                 recursive: bool = False, naming: str = 'default',
                 max_vision_pages: int = 8, budget: int = 50):
        self.input_folder = Path(input_folder)
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.recursive = recursive
        self.naming = naming
        self.max_vision_pages = max_vision_pages

        self._client = None
        self._budget_cap = budget if budget > 0 else None
        self._api_call_count = 0

        self._existing_files = set()
        if self.output_dir.exists():
            self._existing_files = {f.stem for f in self.output_dir.glob('*.md')}

        self.quality_log_path = self.output_dir / 'quality_log.tsv'

    @classmethod
    def from_file(cls, pdf_path, *, max_vision_pages: int = 8,
                  budget: int = 50, **kwargs):
        """Create an instance for processing a single file (no output dir needed).

        Used by the doc2md public API for programmatic access.
        """
        pdf_path = Path(pdf_path)
        import tempfile
        tmpdir = tempfile.mkdtemp(prefix="doc2md_")
        instance = cls(
            input_folder=str(pdf_path.parent),
            output_dir=tmpdir,
            max_vision_pages=max_vision_pages,
            budget=budget,
            **kwargs,
        )
        instance._tmpdir = tmpdir
        return instance

    # ─── API Client ───────────────────────────────────────────────────────────

    def get_anthropic_client(self):
        if self._client is not None:
            return self._client
        try:
            self._client = _get_shared_client()
        except RuntimeError:
            logger.error("Cannot create Anthropic client - no credentials found")
            return None
        return self._client

    def _track_api_call(self, count: int = 1) -> bool:
        if self._budget_cap and (self._api_call_count + count) > self._budget_cap:
            logger.warning(f"  BUDGET EXCEEDED: {self._api_call_count + count}/{self._budget_cap} API calls")
            return False
        self._api_call_count += count
        return True

    def _budget_remaining(self) -> int:
        if not self._budget_cap:
            return 999
        return max(0, self._budget_cap - self._api_call_count)

    # ─── Image Handling ───────────────────────────────────────────────────────

    def render_pages_to_images(self, pdf_path: Path, page_numbers: List[int],
                               dpi: int = None) -> Dict[int, Any]:
        if dpi is None:
            dpi = self.RENDER_DPI
        images = {}
        try:
            import fitz
            from PIL import Image
            Image.MAX_IMAGE_PIXELS = 300_000_000
            doc = fitz.open(str(pdf_path))
            try:
                zoom = dpi / 72
                mat = fitz.Matrix(zoom, zoom)
                for page_num in page_numbers:
                    if page_num < len(doc):
                        pix = doc[page_num].get_pixmap(matrix=mat)
                        images[page_num] = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            finally:
                doc.close()
        except Exception as e:
            logger.error(f"Error rendering pages: {e}")
            for img in images.values():
                try:
                    img.close()
                except Exception:
                    pass
            images.clear()
        return images

    def encode_image_to_base64(self, image, format="JPEG",
                               max_dimension=None, quality: int = 85) -> Tuple[Optional[str], Optional[str]]:
        if max_dimension is None:
            max_dimension = self.MAX_IMAGE_DIMENSION
        try:
            from PIL import Image
            width, height = image.size
            if max(width, height) > max_dimension:
                scale = max_dimension / max(width, height)
                image = image.resize((int(width * scale), int(height * scale)), Image.Resampling.LANCZOS)
            if image.mode == 'RGBA':
                image = image.convert('RGB')
            buffered = BytesIO()
            image.save(buffered, format=format,
                       quality=quality if format == "JPEG" else None)
            img_base64 = base64.b64encode(buffered.getvalue()).decode('utf-8')
            return img_base64, f"image/{format.lower()}"
        except Exception as e:
            logger.error(f"Error encoding image: {e}")
            return None, None

    # ─── JSON Parsing ─────────────────────────────────────────────────────────

    def parse_json_response(self, response_text: str) -> Optional[Any]:
        try:
            return json.loads(response_text)
        except json.JSONDecodeError:
            pass
        json_match = re.search(r'```(?:json)?\s*\n?(.*?)\n?```', response_text, re.DOTALL)
        if json_match:
            try:
                return json.loads(json_match.group(1))
            except json.JSONDecodeError:
                pass
        for start_char, end_char in [('[', ']'), ('{', '}')]:
            start = response_text.find(start_char)
            if start >= 0:
                depth = 0
                for i, c in enumerate(response_text[start:], start):
                    if c == start_char:
                        depth += 1
                    elif c == end_char:
                        depth -= 1
                        if depth == 0:
                            try:
                                return json.loads(response_text[start:i+1])
                            except json.JSONDecodeError:
                                break
        return None

    # ─── PDF Characterization ─────────────────────────────────────────────────

    def characterize_pdf(self, pdf_path: Path) -> Tuple[List[Dict], str]:
        """Extract text from all pages and classify each page.

        Returns (page_data, extraction_method) where page_data is a list of dicts:
            {page_num, text, char_count, classification}
        """
        import pdfplumber

        page_data = []
        method = "native"

        try:
            with pdfplumber.open(str(pdf_path)) as pdf:
                total_pages = len(pdf.pages)
                for i, page in enumerate(pdf.pages):
                    text = page.extract_text() or ""
                    char_count = len(text.strip())

                    classification = self._classify_page(text, i, total_pages)
                    page_data.append({
                        'page_num': i,
                        'text': text,
                        'char_count': char_count,
                        'classification': classification,
                    })
        except Exception as e:
            logger.error(f"Error reading PDF with pdfplumber: {e}")
            method = "error"

        if not page_data or all(p['char_count'] < self.MIN_TEXT_FOR_TEXT_PAGE for p in page_data):
            logger.info("  Native text extraction failed or low quality, trying OCR...")
            ocr_data = self._ocr_all_pages(pdf_path)
            if ocr_data:
                page_data = ocr_data
                method = "ocr"

        return page_data, method

    def _classify_page(self, text: str, page_num: int, total_pages: int) -> str:
        char_count = len(text.strip())

        if page_num == 0:
            return 'TITLE_PAGE'

        if char_count < self.MIN_TEXT_FOR_TEXT_PAGE:
            return 'FIGURE_PAGE'

        text_lower = text.lower()
        if self.SECTION_PATTERNS['references'].search(text):
            ref_pos = self.SECTION_PATTERNS['references'].search(text).start()
            if ref_pos < len(text) * 0.3:
                return 'REFERENCES_PAGE'

        if page_num >= total_pages - 2 and 'supplementary' in text_lower:
            return 'SUPPLEMENTARY_PAGE'

        return 'TEXT_PAGE'

    def _ocr_all_pages(self, pdf_path: Path) -> Optional[List[Dict]]:
        try:
            import fitz
            from PIL import Image
            import pytesseract
            Image.MAX_IMAGE_PIXELS = 300_000_000

            page_data = []
            doc = fitz.open(str(pdf_path))
            try:
                total_pages = len(doc)
                zoom = self.DEFAULT_OCR_DPI / 72
                mat = fitz.Matrix(zoom, zoom)
                for i in range(total_pages):
                    pix = doc[i].get_pixmap(matrix=mat)
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    text = pytesseract.image_to_string(img)
                    char_count = len(text.strip())
                    classification = self._classify_page(text, i, total_pages)
                    page_data.append({
                        'page_num': i,
                        'text': text,
                        'char_count': char_count,
                        'classification': classification,
                    })
                    img.close()
            finally:
                doc.close()
            return page_data if page_data else None
        except Exception as e:
            logger.error(f"OCR failed: {e}")
            return None

    # ─── Full Text Assembly ───────────────────────────────────────────────────

    def assemble_full_text(self, page_data: List[Dict]) -> str:
        """Combine text from all pages into a single document."""
        parts = []
        for p in page_data:
            if p['text'].strip():
                parts.append(p['text'].strip())
        return '\n\n'.join(parts)

    # ─── Metadata Extraction ─────────────────────────────────────────────────

    def extract_metadata_from_filename(self, pdf_path: Path) -> Dict[str, Any]:
        """Parse metadata from filename pattern: 'Author et al. - YEAR - Title.pdf'"""
        stem = pdf_path.stem
        metadata = {}

        match = self.FILENAME_PATTERN.match(stem)
        if match:
            authors_str = match.group(1).strip()
            metadata['first_author'] = authors_str.split()[0] if authors_str else None
            metadata['authors_hint'] = authors_str
            metadata['year'] = match.group(2)
            metadata['title_hint'] = match.group(3).strip()
            return metadata

        match = self.FILENAME_PATTERN_SIMPLE.match(stem)
        if match:
            metadata['first_author'] = match.group(1).strip().split()[0]
            metadata['year'] = match.group(2)
            metadata['title_hint'] = match.group(3).strip()
            return metadata

        return metadata

    def extract_metadata_from_text(self, full_text: str, page_data: List[Dict]) -> Dict[str, Any]:
        """Extract metadata using regex/heuristics from the first pages of text."""
        metadata = {}
        first_page = page_data[0]['text'] if page_data else ""
        first_pages_text = '\n'.join(p['text'] for p in page_data[:2])

        # DOI
        doi_match = self.DOI_PATTERN.search(first_pages_text)
        if doi_match:
            doi = doi_match.group(1).rstrip('.)')
            metadata['doi'] = doi

        # Year - look for publication year patterns
        year_patterns = [
            re.compile(r'Published(?:\s+online)?[:\s]+\d{1,2}\s+\w+\s+(20[012]\d)'),
            re.compile(r'Received[:\s]+\d{1,2}\s+\w+\s+(20[012]\d)'),
            re.compile(r'\((\d{4})\)\s*$', re.MULTILINE),
            re.compile(r'©\s*(20[012]\d)'),
        ]
        for yp in year_patterns:
            ym = yp.search(first_pages_text[:3000])
            if ym:
                metadata['year'] = ym.group(1)
                break
        if 'year' not in metadata:
            year_matches = self.YEAR_PATTERN.findall(first_pages_text[:2000])
            if year_matches:
                metadata['year'] = year_matches[0]

        # Title - heuristic: first substantial line(s) on page 1 that aren't
        # a journal header, DOI, or author line
        title = self._extract_title_from_first_page(first_page)
        if title:
            metadata['title'] = title

        # Authors - look for line with multiple names after title/DOI
        authors = self._extract_authors_from_first_page(first_page)
        if authors:
            metadata['authors'] = authors

        # Journal - common patterns (search broader text since journal name
        # often appears in headers/footers on later pages)
        full_search = '\n'.join(p['text'] for p in page_data[:5]) if len(page_data) > 2 else first_pages_text
        journal_patterns = [
            re.compile(r'(Nature|Science|Cell|PNAS|The Lancet|BMJ|JAMA|PLoS\s+\w+|Nucleic Acids Research|Bioinformatics|Cancer Research|Nature\s+\w+)(?:\s*\||\s+Vol)', re.IGNORECASE),
            re.compile(r'(?:Published in|Journal:?)\s+([A-Z][A-Za-z\s&]+?)(?:\s*[,|]|\s+\d)', re.IGNORECASE),
        ]
        for jp in journal_patterns:
            jm = jp.search(full_search)
            if jm:
                metadata['journal'] = jm.group(1).strip()
                break

        return metadata

    def _extract_title_from_first_page(self, first_page: str) -> Optional[str]:
        """Extract title from first page text using heuristics."""
        lines = first_page.split('\n')
        skip_patterns = re.compile(
            r'^(Article|Letter|Review|Report|Research|Original|Open\s+access|'
            r'https?://|doi:|Received|Accepted|Published|Check for|©|\d+\s*$)',
            re.IGNORECASE
        )

        title_lines = []
        collecting = False
        for line in lines:
            line = line.strip()
            if not line:
                if collecting and title_lines:
                    break
                continue
            if skip_patterns.match(line):
                if collecting and title_lines:
                    break
                continue
            if len(line) < 10:
                if collecting and title_lines:
                    break
                continue
            # Stop if we hit what looks like author names (multiple commas, affiliations)
            if re.search(r'\d{1,2}[,\s]+\d', line) and ',' in line:
                break
            if re.search(r'✉|correspondence|@', line, re.IGNORECASE):
                break
            # Likely a title line if it's substantial text
            if len(line) > 15 and not re.match(r'^\d+\.', line):
                title_lines.append(line)
                collecting = True
                if len(' '.join(title_lines)) > 200:
                    break

        title = ' '.join(title_lines).strip()
        if title and len(title) > 15:
            return title
        return None

    def _extract_authors_from_first_page(self, first_page: str) -> List[str]:
        """Extract author names from first page."""
        lines = first_page.split('\n')

        # Strategy: find lines after DOI/title that contain multiple names with
        # superscript numbers, commas, and affiliation markers
        author_lines = []
        found_doi_or_title = False
        for line in lines:
            if re.search(r'doi\.org|https?://', line, re.IGNORECASE):
                found_doi_or_title = True
                # Authors might be on same line as DOI
                doi_end = re.search(r'(10\.\d{4,}/[^\s]+)', line)
                if doi_end:
                    remainder = line[doi_end.end():].strip()
                    if remainder and re.search(r'[A-Z][a-z]', remainder):
                        author_lines.append(remainder)
                continue
            if not found_doi_or_title:
                continue
            # Stop at date/metadata lines
            if re.match(r'^(Received|Accepted|Published|Open access|Check for)', line.strip()):
                break
            if line.strip() and re.search(r'[A-Z][a-z]+', line):
                author_lines.append(line.strip())
            if len(author_lines) >= 3:
                break

        if not author_lines:
            return []

        raw_authors = ' '.join(author_lines)
        # Remove superscript numbers and affiliation markers
        cleaned = re.sub(r'[\d]+', '', raw_authors)
        # Remove special chars like ✉, *, †
        cleaned = re.sub(r'[✉*†‡§¶]', '', cleaned)
        # Remove & and replace with comma
        cleaned = cleaned.replace(' & ', ', ')
        # Split by comma and clean
        authors = []
        for part in cleaned.split(','):
            name = part.strip()
            # Must have at least 2 words and look like a name
            if name and len(name) > 3 and re.search(r'^[A-Z]', name):
                words = name.split()
                if len(words) >= 2 and all(len(w) > 0 for w in words):
                    authors.append(name)

        return authors[:20]

    def extract_metadata_vision(self, pdf_path: Path, page_data: List[Dict]) -> Dict[str, Any]:
        """Use Vision AI on page 1 to extract structured metadata."""
        if not self._track_api_call():
            return {}

        client = self.get_anthropic_client()
        if not client:
            return {}

        images = self.render_pages_to_images(pdf_path, [0])
        if 0 not in images:
            return {}

        encoded, media_type = self.encode_image_to_base64(images[0])
        images[0].close()
        if not encoded:
            return {}

        first_page_text = page_data[0]['text'] if page_data else ""

        prompt = """Analyze this scientific paper's first page and extract metadata. Return ONLY valid JSON:

{
  "title": "exact paper title",
  "authors": ["Full Name 1", "Full Name 2", "..."],
  "affiliations": ["Institution 1", "..."],
  "journal": "journal name if visible",
  "doi": "DOI if visible (e.g., 10.1038/...)",
  "year": "publication year",
  "volume": "volume number if visible",
  "pages": "page range if visible",
  "keywords": ["keyword1", "keyword2"],
  "abstract": "full abstract text if visible on this page",
  "correspondence_author": "corresponding author name and email if visible"
}

Use null for fields not visible on this page. Extract the COMPLETE abstract if it's on this page."""

        try:
            response = client.messages.create(
                model=self.VISION_MODEL,
                max_tokens=4096,
                messages=[{
                    "role": "user",
                    "content": [
                        {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": encoded}},
                        {"type": "text", "text": f"Reference text from PDF extraction:\n{first_page_text[:3000]}\n\n{prompt}"}
                    ]
                }]
            )
            result = self.parse_json_response(response.content[0].text)
            if result:
                return {k: v for k, v in result.items() if v is not None}
        except Exception as e:
            logger.error(f"Vision metadata extraction failed: {e}")

        return {}

    def merge_metadata(self, vision_meta: Dict, text_meta: Dict,
                       filename_meta: Dict) -> Dict[str, Any]:
        """Merge metadata from all sources. Priority: Vision > Text > Filename."""
        merged = {}

        merged['title'] = (vision_meta.get('title') or
                          text_meta.get('title') or
                          filename_meta.get('title_hint') or 'Untitled')
        merged['authors'] = vision_meta.get('authors') or text_meta.get('authors') or []
        merged['affiliations'] = vision_meta.get('affiliations') or []
        merged['journal'] = vision_meta.get('journal') or text_meta.get('journal')
        merged['doi'] = vision_meta.get('doi') or text_meta.get('doi')
        merged['year'] = (vision_meta.get('year') or
                         text_meta.get('year') or
                         filename_meta.get('year'))
        merged['volume'] = vision_meta.get('volume')
        merged['pages'] = vision_meta.get('pages')
        merged['keywords'] = vision_meta.get('keywords') or []
        merged['abstract'] = vision_meta.get('abstract')
        merged['correspondence_author'] = vision_meta.get('correspondence_author')

        if not merged['authors'] and filename_meta.get('authors_hint'):
            merged['authors'] = [filename_meta['authors_hint']]

        # First author: prefer filename (reliable when pattern matches) > vision > text
        first_author = None
        if filename_meta.get('first_author'):
            first_author = filename_meta['first_author']
        elif vision_meta.get('authors'):
            first_name = vision_meta['authors'][0]
            parts = first_name.replace(',', ' ').split()
            first_author = parts[-1] if parts else first_name
        elif merged['authors']:
            first_name = merged['authors'][0]
            parts = first_name.replace(',', ' ').split()
            first_author = parts[-1] if parts else first_name
        merged['first_author'] = first_author or 'Unknown'

        return merged

    def _extract_abstract_from_first_page(self, page_data: List[Dict]) -> Optional[str]:
        """Try to extract abstract from first page text as a continuous paragraph."""
        if not page_data:
            return None
        first_page = page_data[0]['text']

        # Many papers have the abstract as a paragraph after metadata
        # Look for a long paragraph (>200 chars) that starts with a capital letter
        # and appears after DOI/authors but before section headers
        lines = first_page.split('\n')
        paragraph_lines = []
        in_paragraph = False
        past_metadata = False

        for line in lines:
            stripped = line.strip()
            # Skip until we're past the metadata (DOI, dates, etc.)
            if not past_metadata:
                if re.search(r'doi|Published|Accepted|Open access|Check for updates', stripped, re.IGNORECASE):
                    past_metadata = True
                continue

            if not stripped:
                if in_paragraph and len(' '.join(paragraph_lines)) > 200:
                    break
                if in_paragraph:
                    paragraph_lines = []
                    in_paragraph = False
                continue

            # Skip short lines that look like headers/metadata
            if len(stripped) < 30 and re.match(r'^[A-Z][a-z]+$', stripped):
                continue

            paragraph_lines.append(stripped)
            in_paragraph = True

        abstract = ' '.join(paragraph_lines).strip()
        if len(abstract) > 200:
            return abstract
        return None

    # ─── Section Splitting ────────────────────────────────────────────────────

    def split_into_sections(self, full_text: str) -> Dict[str, str]:
        """Split paper text into sections using regex IMRaD detection."""
        sections = self._find_sections_with_patterns(full_text, self.SECTION_PATTERNS)

        # If strict patterns find <3 sections, try relaxed patterns
        if len(sections) < 3:
            relaxed = self._find_sections_with_patterns(full_text, self.SECTION_PATTERNS_RELAXED)
            if len(relaxed) > len(sections):
                sections = relaxed

        return sections

    def _find_sections_with_patterns(self, full_text: str, patterns: Dict) -> Dict[str, str]:
        """Find sections using a given set of patterns."""
        section_positions = []

        for section_name, pattern in patterns.items():
            for match in pattern.finditer(full_text):
                section_positions.append((match.start(), match.end(), section_name))

        section_positions.sort(key=lambda x: x[0])

        # Deduplicate: keep only first match per section name
        seen = set()
        unique_positions = []
        for pos in section_positions:
            if pos[2] not in seen:
                seen.add(pos[2])
                unique_positions.append(pos)
        section_positions = unique_positions

        sections = {}
        for i, (start, header_end, name) in enumerate(section_positions):
            if i + 1 < len(section_positions):
                end = section_positions[i + 1][0]
            else:
                end = len(full_text)

            content = full_text[header_end:end].strip()
            if len(content) > 50:
                sections[name] = content

        return sections

    def split_sections_with_ai(self, full_text: str, metadata: Dict) -> Dict[str, str]:
        """Use AI to identify and extract paper sections when regex fails."""
        if not self._track_api_call():
            return {}

        client = self.get_anthropic_client()
        if not client:
            return {}

        text_sample = full_text[:15000]

        prompt = f"""Analyze this scientific paper text and identify the main sections.
The paper is titled: "{metadata.get('title', 'Unknown')}"

Return ONLY valid JSON with these keys (use null if section not found):
{{
  "abstract": "full abstract text",
  "introduction": "full introduction text",
  "methods": "full methods/materials section text",
  "results": "full results text",
  "discussion": "full discussion text",
  "conclusions": "full conclusions text"
}}

Important:
- Extract the COMPLETE text for each section, not just summaries
- If sections are combined (e.g., "Results and Discussion"), put the content in both keys
- Omit the references section

Paper text:
{text_sample}"""

        try:
            response = client.messages.create(
                model=self.VISION_MODEL,
                max_tokens=8192,
                messages=[{"role": "user", "content": prompt}]
            )
            result = self.parse_json_response(response.content[0].text)
            if result:
                return {k: v for k, v in result.items() if v and len(str(v)) > 50}
        except Exception as e:
            logger.error(f"AI section splitting failed: {e}")

        return {}

    # ─── Figure Analysis ──────────────────────────────────────────────────────

    def identify_figure_pages(self, page_data: List[Dict]) -> List[int]:
        """Identify pages that likely contain figures (low text, or figure captions nearby)."""
        figure_pages = []
        for p in page_data:
            if p['classification'] == 'FIGURE_PAGE':
                figure_pages.append(p['page_num'])
            elif 'figure' in p['text'].lower()[:200] or 'fig.' in p['text'].lower()[:200]:
                if p['char_count'] < 500:
                    figure_pages.append(p['page_num'])

        return figure_pages[:self.max_vision_pages]

    def extract_figure_captions_from_text(self, full_text: str) -> List[Dict]:
        """Extract figure captions from text using regex."""
        captions = []
        pattern = re.compile(
            r'(?:Figure|Fig\.?)\s*(\d+)[.:]\s*(.+?)(?=(?:Figure|Fig\.?)\s*\d+[.:]|\n\n|\Z)',
            re.DOTALL | re.IGNORECASE
        )
        for match in pattern.finditer(full_text):
            fig_num = match.group(1)
            caption = match.group(2).strip()[:500]
            caption = re.sub(r'\s+', ' ', caption)
            captions.append({
                'figure_number': int(fig_num),
                'caption': caption,
            })
        return captions

    def analyze_figures_batch(self, pdf_path: Path, figure_pages: List[int],
                              full_text: str, *,
                              include_images: bool = False,
                              dpi: int = None) -> List[Dict]:
        """Analyze figure pages with Vision AI in batches.

        Args:
            include_images: If True, each result dict includes 'image_bytes' (JPEG).
            dpi: Override render resolution (default: self.RENDER_DPI = 200).

        Returns list of dicts with keys: figure_number, title, figure_type,
        description, key_findings, statistical_notes, relevance, page_num.
        When include_images=True, also includes 'image_bytes'.
        """
        if not figure_pages:
            return []

        render_dpi = dpi or self.RENDER_DPI
        all_figures = []
        text_captions = self.extract_figure_captions_from_text(full_text)
        caption_context = "\n".join(
            f"Figure {c['figure_number']}: {c['caption'][:200]}" for c in text_captions
        )

        batches = [figure_pages[i:i + self.PAGES_PER_VISION_BATCH]
                   for i in range(0, len(figure_pages), self.PAGES_PER_VISION_BATCH)]

        for batch in batches:
            if not self._track_api_call():
                break

            images = self.render_pages_to_images(pdf_path, batch, dpi=render_dpi)
            if not images:
                continue

            # Optionally retain raw JPEG bytes for the caller
            batch_image_bytes = {}
            content_blocks = []
            for page_num in batch:
                if page_num in images:
                    if include_images:
                        from io import BytesIO as _BytesIO
                        buf = _BytesIO()
                        img = images[page_num]
                        if img.mode == 'RGBA':
                            img = img.convert('RGB')
                        img.save(buf, format="JPEG", quality=85)
                        batch_image_bytes[page_num] = buf.getvalue()

                    encoded, media_type = self.encode_image_to_base64(images[page_num])
                    images[page_num].close()
                    if encoded:
                        content_blocks.append({
                            "type": "image",
                            "source": {"type": "base64", "media_type": media_type, "data": encoded}
                        })

            if not content_blocks:
                continue

            prompt = f"""Analyze all figures and charts visible in these pages from a scientific paper.

Known figure captions from the text:
{caption_context}

For each figure, return a JSON array:
[
  {{
    "figure_number": 1,
    "title": "figure title or caption",
    "figure_type": "bar chart / line plot / scatter plot / heatmap / diagram / microscopy / gel / flow chart / other",
    "description": "what the figure shows (2-3 sentences)",
    "key_findings": ["finding 1", "finding 2"],
    "statistical_notes": "any p-values, significance markers, sample sizes",
    "relevance": "HIGH / MEDIUM / LOW"
  }}
]

Relevance scoring rubric:
  HIGH = overview or summary figure, key clinical outcome (survival, ORR, response), essential mechanistic diagram, landscape/pipeline overview, multi-panel summary
  MEDIUM = individual experimental result, supporting methods, secondary validation, single-assay result
  LOW = technical detail only, minor supplementary-level content, formatting/legend page

Return ONLY the JSON array."""

            content_blocks.append({"type": "text", "text": prompt})

            try:
                client = self.get_anthropic_client()
                response = client.messages.create(
                    model=self.VISION_MODEL,
                    max_tokens=4096,
                    messages=[{"role": "user", "content": content_blocks}]
                )
                result = self.parse_json_response(response.content[0].text)
                items = result if isinstance(result, list) else [result] if isinstance(result, dict) else []
                for idx, item in enumerate(items):
                    item.setdefault('relevance', 'MEDIUM')
                    item['page_num'] = batch[min(idx, len(batch) - 1)]
                    if include_images and item['page_num'] in batch_image_bytes:
                        item['image_bytes'] = batch_image_bytes[item['page_num']]
                all_figures.extend(items)
            except Exception as e:
                logger.error(f"Figure analysis batch failed: {e}")

        return all_figures

    # ─── Table Extraction ─────────────────────────────────────────────────────

    def extract_tables(self, pdf_path: Path, page_data: List[Dict]) -> List[Dict]:
        """Extract tables using pdfplumber's table detection."""
        import pdfplumber

        tables = []
        table_captions = self._extract_table_captions(
            self.assemble_full_text(page_data))

        try:
            with pdfplumber.open(str(pdf_path)) as pdf:
                for i, page in enumerate(pdf.pages):
                    page_tables = page.extract_tables()
                    if not page_tables:
                        continue
                    for t_idx, table_data in enumerate(page_tables):
                        if not table_data or len(table_data) < 2:
                            continue
                        header = table_data[0] if table_data[0] else []
                        rows = table_data[1:]
                        table_num = len(tables) + 1

                        caption = ""
                        for tc in table_captions:
                            if tc['table_number'] == table_num:
                                caption = tc['caption']
                                break

                        tables.append({
                            'table_number': table_num,
                            'page': i,
                            'caption': caption,
                            'header': [str(h) if h else '' for h in header],
                            'rows': [[str(cell) if cell else '' for cell in row] for row in rows],
                        })
        except Exception as e:
            logger.error(f"Table extraction failed: {e}")

        return tables

    def _extract_table_captions(self, full_text: str) -> List[Dict]:
        """Extract table captions from text."""
        captions = []
        pattern = re.compile(
            r'(?:Table)\s*(\d+)[.:]\s*(.+?)(?=(?:Table)\s*\d+[.:]|\n\n|\Z)',
            re.DOTALL | re.IGNORECASE
        )
        for match in pattern.finditer(full_text):
            captions.append({
                'table_number': int(match.group(1)),
                'caption': re.sub(r'\s+', ' ', match.group(2).strip()[:300]),
            })
        return captions

    # ─── Executive Summary ────────────────────────────────────────────────────

    def generate_executive_summary(self, sections: Dict[str, str],
                                    metadata: Dict) -> str:
        """Generate a concise executive summary using AI."""
        if not self._track_api_call():
            return ""

        client = self.get_anthropic_client()
        if not client:
            return ""

        abstract = sections.get('abstract', '')
        conclusions = sections.get('conclusions', '')
        results_snippet = sections.get('results', '')[:2000]

        prompt = f"""Write a 2-3 sentence executive summary of this scientific paper.

Title: {metadata.get('title', 'Unknown')}
Authors: {', '.join(metadata.get('authors', [])[:3])}

Abstract: {abstract[:1000]}

Key Results: {results_snippet}

Conclusions: {conclusions[:1000]}

Write a concise summary highlighting the main finding, method, and significance. No more than 3 sentences."""

        try:
            response = client.messages.create(
                model=self.VISION_MODEL,
                max_tokens=512,
                messages=[{"role": "user", "content": prompt}]
            )
            return response.content[0].text.strip()
        except Exception as e:
            logger.error(f"Executive summary generation failed: {e}")
            return ""

    # ─── Quality Scoring ──────────────────────────────────────────────────────

    def calculate_quality_score(self, full_text: str, metadata: Dict,
                                sections: Dict, figures: List[Dict],
                                tables: List[Dict]) -> Dict[str, Any]:
        """Calculate quality score adapted for research papers."""
        scores = {}

        # Text extraction (0-10)
        text_len = len(full_text)
        if text_len > 20000:
            text_score = 10.0
        elif text_len > 10000:
            text_score = 8.0
        elif text_len > 5000:
            text_score = 6.0
        elif text_len > 1000:
            text_score = 4.0
        else:
            text_score = 2.0
        scores['text_extraction'] = text_score

        # Metadata completeness (0-10)
        meta_score = 0.0
        if metadata.get('title') and metadata['title'] != 'Untitled':
            meta_score += 2.0
        if metadata.get('authors'):
            meta_score += 2.0
        if metadata.get('year'):
            meta_score += 2.0
        if metadata.get('journal'):
            meta_score += 2.0
        if metadata.get('doi'):
            meta_score += 2.0
        scores['metadata_completeness'] = meta_score

        # Structure parsing (0-10)
        core_sections = ['methods', 'results', 'discussion']
        optional_sections = ['abstract', 'introduction', 'conclusions']
        struct_score = 0.0
        for s in core_sections:
            if s in sections and len(sections[s]) > 100:
                struct_score += 2.5
        for s in optional_sections:
            if s in sections and len(sections[s]) > 50:
                struct_score += 0.83
        scores['structure_parsing'] = min(10.0, struct_score)

        # Figure analysis (0-10)
        if figures:
            fig_score = min(10.0, len(figures) * 2.0)
            has_descriptions = sum(1 for f in figures if f.get('description'))
            if has_descriptions >= len(figures) * 0.7:
                fig_score = min(10.0, fig_score + 2.0)
        else:
            fig_score = 5.0  # Many papers have inline figures that aren't separate pages
        scores['figure_analysis'] = fig_score

        # Table extraction (0-10)
        if tables:
            table_score = min(10.0, len(tables) * 2.5)
        else:
            table_score = 5.0  # Not all papers have tables
        scores['table_extraction'] = table_score

        # Weighted overall
        weights = {
            'text_extraction': 0.30,
            'metadata_completeness': 0.20,
            'structure_parsing': 0.25,
            'figure_analysis': 0.15,
            'table_extraction': 0.10,
        }
        overall = sum(scores[k] * weights[k] for k in weights)
        overall = round(min(10.0, max(0.0, overall)), 1)

        if overall >= self.QUALITY_EXCELLENT_THRESHOLD:
            assessment = "Excellent"
        elif overall >= self.QUALITY_GOOD_THRESHOLD:
            assessment = "Good"
        elif overall >= self.QUALITY_FAIR_THRESHOLD:
            assessment = "Fair"
        else:
            assessment = "Poor"

        return {
            'overall': overall,
            'assessment': assessment,
            'components': scores,
        }

    # ─── Filename Generation ──────────────────────────────────────────────────

    def _make_title_slug(self, title: str, max_words: int = 6) -> str:
        """Create a filename-safe slug from the title."""
        words = re.findall(r'[A-Za-z0-9]+', title)
        stop_words = {'a', 'an', 'the', 'of', 'in', 'on', 'at', 'to', 'for',
                      'and', 'or', 'but', 'is', 'are', 'was', 'were', 'with', 'by'}
        meaningful = [w for w in words if w.lower() not in stop_words]
        if not meaningful:
            meaningful = words
        slug = '_'.join(meaningful[:max_words])
        return slug if slug else 'untitled'

    def generate_filename(self, metadata: Dict, pdf_path: Path) -> str:
        """Generate output filename based on naming scheme."""
        first_author = metadata.get('first_author', 'Unknown')
        year = metadata.get('year', 'NoYear')
        title = metadata.get('title', pdf_path.stem)

        if self.naming == 'detailed':
            journal = metadata.get('journal', 'Unknown')
            journal_slug = re.sub(r'[^A-Za-z0-9]', '', journal)[:15]
            title_slug = self._make_title_slug(title)
            filename = f"{first_author}_et_al_{year}_{journal_slug}_{title_slug}.md"
        elif self.naming == 'doi':
            doi = metadata.get('doi', '')
            if doi:
                doi_suffix = doi.split('/')[-1] if '/' in doi else doi
                doi_suffix = re.sub(r'[^A-Za-z0-9_-]', '_', doi_suffix)
                filename = f"{doi_suffix}.md"
            else:
                title_slug = self._make_title_slug(title)
                filename = f"{first_author}_{year}_{title_slug}.md"
        else:  # default
            title_slug = self._make_title_slug(title)
            filename = f"{first_author}_{year}_{title_slug}.md"

        return sanitize_filename(filename)

    # ─── Markdown Generation ──────────────────────────────────────────────────

    def generate_markdown(self, metadata: Dict, sections: Dict,
                          figures: List[Dict], tables: List[Dict],
                          summary: str, quality: Dict,
                          extraction_method: str, page_count: int,
                          full_text: str) -> str:
        """Generate the final markdown document."""
        lines = []

        # YAML frontmatter
        lines.append('---')
        lines.append(f'title: "{self._escape_yaml(metadata.get("title", "Untitled"))}"')
        if metadata.get('authors'):
            lines.append('authors:')
            for author in metadata['authors']:
                lines.append(f'  - "{self._escape_yaml(author)}"')
        lines.append(f'first_author: "{self._escape_yaml(metadata.get("first_author", "Unknown"))}"')
        if metadata.get('year'):
            lines.append(f'year: {metadata["year"]}')
        if metadata.get('journal'):
            lines.append(f'journal: "{self._escape_yaml(metadata["journal"])}"')
        if metadata.get('doi'):
            lines.append(f'doi: "{metadata["doi"]}"')
        if metadata.get('volume'):
            lines.append(f'volume: "{metadata["volume"]}"')
        if metadata.get('pages'):
            lines.append(f'pages: "{metadata["pages"]}"')
        if metadata.get('keywords'):
            lines.append('keywords:')
            for kw in metadata['keywords']:
                lines.append(f'  - "{self._escape_yaml(kw)}"')
        lines.append(f'extraction_method: {extraction_method}')
        lines.append(f'vision_model: {self.VISION_MODEL}')
        lines.append(f'processing_date: {datetime.now().strftime("%Y-%m-%d")}')
        lines.append(f'quality_overall: {quality["overall"]}')
        lines.append(f'quality_assessment: {quality["assessment"]}')
        lines.append(f'total_pages: {page_count}')
        lines.append(f'figures_analyzed: {len(figures)}')
        lines.append(f'tables_extracted: {len(tables)}')
        lines.append('---')
        lines.append('')

        # Title and metadata header
        lines.append(f'# {metadata.get("title", "Untitled")}')
        lines.append('')

        author_str = ', '.join(metadata.get('authors', ['Unknown']))
        lines.append(f'**Authors**: {author_str}')
        lines.append('')

        journal_parts = []
        if metadata.get('journal'):
            journal_parts.append(metadata['journal'])
        if metadata.get('volume'):
            journal_parts.append(f'Volume {metadata["volume"]}')
        if metadata.get('pages'):
            journal_parts.append(f'Pages {metadata["pages"]}')
        if metadata.get('year'):
            journal_parts.append(metadata['year'])
        if journal_parts:
            lines.append(f'**Journal**: {" | ".join(journal_parts)}')
            lines.append('')

        if metadata.get('doi'):
            lines.append(f'**DOI**: [{metadata["doi"]}](https://doi.org/{metadata["doi"]})')
            lines.append('')

        if metadata.get('affiliations'):
            lines.append(f'**Affiliations**: {"; ".join(metadata["affiliations"])}')
            lines.append('')

        lines.append('---')
        lines.append('')

        # Executive Summary
        if summary:
            lines.append('## Executive Summary')
            lines.append('')
            lines.append(summary)
            lines.append('')
            lines.append('---')
            lines.append('')

        # Content Sections
        section_order = ['abstract', 'introduction', 'methods', 'results',
                         'discussion', 'conclusions']
        section_titles = {
            'abstract': 'Abstract',
            'introduction': 'Introduction',
            'methods': 'Methods',
            'results': 'Results',
            'discussion': 'Discussion',
            'conclusions': 'Conclusions',
        }

        for section_name in section_order:
            if section_name in sections:
                lines.append(f'## {section_titles[section_name]}')
                lines.append('')
                lines.append(sections[section_name])
                lines.append('')

                # Insert figures after results
                if section_name == 'results' and figures:
                    lines.append('')
                    for fig in figures:
                        lines.append(self._format_figure_markdown(fig))
                        lines.append('')

                # Insert tables after results if no figures there
                if section_name == 'results' and tables and not figures:
                    for table in tables:
                        lines.append(self._format_table_markdown(table))
                        lines.append('')

                lines.append('---')
                lines.append('')

        # Standalone figures section if not placed in results
        if figures and 'results' not in sections:
            lines.append('## Figures')
            lines.append('')
            for fig in figures:
                lines.append(self._format_figure_markdown(fig))
                lines.append('')
            lines.append('---')
            lines.append('')

        # Standalone tables if not placed
        if tables and ('results' not in sections or figures):
            lines.append('## Tables')
            lines.append('')
            for table in tables:
                lines.append(self._format_table_markdown(table))
                lines.append('')
            lines.append('---')
            lines.append('')

        # References (collapsible)
        if 'references' in sections:
            lines.append('## References')
            lines.append('')
            lines.append('<details>')
            lines.append('<summary>Click to expand references</summary>')
            lines.append('')
            lines.append(sections['references'])
            lines.append('')
            lines.append('</details>')
            lines.append('')
            lines.append('---')
            lines.append('')

        # Quality Assessment
        lines.append('## Quality Assessment')
        lines.append('')
        lines.append(f'**Overall Quality**: {quality["overall"]}/10 - {quality["assessment"]}')
        lines.append('')
        lines.append('**Component Scores**:')
        for comp, score in quality['components'].items():
            comp_label = comp.replace('_', ' ').title()
            lines.append(f'- {comp_label}: {score:.1f}/10')
        lines.append('')
        lines.append('---')
        lines.append('')

        # Processing Metadata
        lines.append('## Processing Metadata')
        lines.append('')
        lines.append(f'- **Extraction Method**: {extraction_method}')
        lines.append(f'- **Vision Model**: {self.VISION_MODEL}')
        lines.append(f'- **Total Pages**: {page_count}')
        lines.append(f'- **Text Length**: {len(full_text):,} characters')
        lines.append(f'- **Figures Analyzed**: {len(figures)}')
        lines.append(f'- **Tables Extracted**: {len(tables)}')
        lines.append(f'- **API Calls Used**: {self._api_call_count}')
        lines.append(f'- **Processing Date**: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
        lines.append('')

        return '\n'.join(lines)

    def _format_figure_markdown(self, fig: Dict) -> str:
        """Format a single figure as markdown."""
        lines = []
        fig_num = fig.get('figure_number', '?')
        title = fig.get('title', 'Untitled')
        lines.append(f'### Figure {fig_num}: {title}')
        lines.append('')
        if fig.get('figure_type'):
            lines.append(f'**Type**: {fig["figure_type"]}')
        if fig.get('description'):
            lines.append(f'**Description**: {fig["description"]}')
        if fig.get('key_findings'):
            lines.append('**Key Findings**:')
            for finding in fig['key_findings']:
                lines.append(f'- {finding}')
        if fig.get('statistical_notes'):
            lines.append(f'**Statistical Notes**: {fig["statistical_notes"]}')
        lines.append('')
        return '\n'.join(lines)

    def _format_table_markdown(self, table: Dict) -> str:
        """Format a table as markdown."""
        lines = []
        table_num = table.get('table_number', '?')
        caption = table.get('caption', '')
        lines.append(f'### Table {table_num}' + (f': {caption}' if caption else ''))
        lines.append('')

        header = table.get('header', [])
        rows = table.get('rows', [])

        if header:
            lines.append('| ' + ' | '.join(header) + ' |')
            lines.append('|' + '|'.join(['---'] * len(header)) + '|')
            for row in rows[:50]:  # Cap at 50 rows
                padded = row + [''] * (len(header) - len(row))
                lines.append('| ' + ' | '.join(padded[:len(header)]) + ' |')
        lines.append('')
        return '\n'.join(lines)

    def _escape_yaml(self, text: str) -> str:
        """Escape special characters for YAML strings."""
        if not text:
            return ''
        return text.replace('"', '\\"').replace('\n', ' ')

    # ─── Quality Log ──────────────────────────────────────────────────────────

    def _init_quality_log(self):
        """Initialize quality log file with header."""
        if not self.quality_log_path.exists():
            with open(self.quality_log_path, 'w', encoding='utf-8') as f:
                f.write("TIMESTAMP\tFILENAME\tQUALITY_SCORE\tASSESSMENT\tSTATUS\tOUTPUT_FILE\n")

    def _log_quality(self, pdf_name: str, quality: Dict, status: str, output_file: str = ""):
        """Append quality entry to log."""
        with open(self.quality_log_path, 'a', encoding='utf-8') as f:
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            f.write(f"{timestamp}\t{pdf_name}\t{quality['overall']}\t"
                    f"{quality['assessment']}\t{status}\t{output_file}\n")

    # ─── Single Paper Processing ──────────────────────────────────────────────

    def process_single_paper(self, pdf_path: Path, skip_existing: bool = True,
                              force_ocr: bool = False, no_vision: bool = False) -> bool:
        """Process a single paper PDF to markdown.

        Returns True if successful and saved, False otherwise.
        """
        self._api_call_count = 0
        pdf_name = pdf_path.name
        logger.info(f"\n{'='*60}")
        logger.info(f"Processing: {pdf_name}")
        logger.info(f"{'='*60}")

        # File size check
        file_size_mb = pdf_path.stat().st_size / (1024 * 1024)
        if file_size_mb > self.MAX_PDF_SIZE_MB:
            logger.error(f"  File too large: {file_size_mb:.1f}MB (max {self.MAX_PDF_SIZE_MB}MB) — skipping")
            return False

        # Step 1: PDF Characterization
        logger.info("  [1/7] Characterizing PDF...")
        page_data, extraction_method = self.characterize_pdf(pdf_path)
        if not page_data:
            logger.error(f"  FAILED: Could not read PDF - {pdf_name}")
            return False

        page_count = len(page_data)
        logger.info(f"  Pages: {page_count}, Method: {extraction_method}")

        if force_ocr and extraction_method == "native":
            logger.info("  Force OCR requested, re-extracting...")
            ocr_data = self._ocr_all_pages(pdf_path)
            if ocr_data:
                page_data = ocr_data
                extraction_method = "ocr"

        # Step 2: Full text assembly
        full_text = self.assemble_full_text(page_data)
        if len(full_text) < 200:
            logger.warning(f"  Very short text ({len(full_text)} chars) - may be scan-only PDF")

        # Step 3: Metadata extraction
        logger.info("  [2/7] Extracting metadata...")
        filename_meta = self.extract_metadata_from_filename(pdf_path)
        text_meta = self.extract_metadata_from_text(full_text, page_data)

        if no_vision:
            vision_meta = {}
        else:
            vision_meta = self.extract_metadata_vision(pdf_path, page_data)

        metadata = self.merge_metadata(vision_meta, text_meta, filename_meta)
        logger.info(f"  Title: {metadata.get('title', 'Unknown')[:60]}")
        logger.info(f"  Authors: {metadata.get('first_author', '?')} et al.")

        # Check skip-existing after we have metadata for filename
        output_filename = self.generate_filename(metadata, pdf_path)
        output_stem = Path(output_filename).stem
        if skip_existing and output_stem in self._existing_files:
            logger.info(f"  SKIPPED (already exists): {output_filename}")
            return True

        # Step 4: Section splitting
        logger.info("  [3/7] Splitting into sections...")
        sections = self.split_into_sections(full_text)

        if metadata.get('abstract') and 'abstract' not in sections:
            sections['abstract'] = metadata['abstract']

        # Try extracting abstract from first page if still missing
        if 'abstract' not in sections:
            abstract = self._extract_abstract_from_first_page(page_data)
            if abstract:
                sections['abstract'] = abstract

        if len(sections) < 3 and not no_vision:
            logger.info("  Regex found <3 sections, using AI fallback...")
            ai_sections = self.split_sections_with_ai(full_text, metadata)
            if ai_sections:
                for k, v in ai_sections.items():
                    if k not in sections or len(v) > len(sections.get(k, '')):
                        sections[k] = v

        logger.info(f"  Sections found: {list(sections.keys())}")

        # Step 5: Figure and table analysis
        logger.info("  [4/7] Analyzing figures and tables...")
        figures = []
        if not no_vision:
            figure_pages = self.identify_figure_pages(page_data)
            if figure_pages:
                logger.info(f"  Figure pages identified: {figure_pages}")
                figures = self.analyze_figures_batch(pdf_path, figure_pages, full_text)
        else:
            text_captions = self.extract_figure_captions_from_text(full_text)
            figures = [{'figure_number': c['figure_number'], 'title': c['caption'][:100],
                       'description': c['caption']} for c in text_captions]

        tables = self.extract_tables(pdf_path, page_data)
        logger.info(f"  Figures: {len(figures)}, Tables: {len(tables)}")

        # Step 6: Executive summary
        logger.info("  [5/7] Generating executive summary...")
        if no_vision:
            summary = ""
        else:
            summary = self.generate_executive_summary(sections, metadata)

        # Step 7: Quality assessment
        logger.info("  [6/7] Calculating quality score...")
        quality = self.calculate_quality_score(full_text, metadata, sections, figures, tables)
        logger.info(f"  Quality: {quality['overall']}/10 - {quality['assessment']}")

        if quality['overall'] < self.QUALITY_GOOD_THRESHOLD:
            logger.warning(f"  LOW QUALITY ({quality['assessment']}) - saving anyway for review")

        # Step 8: Generate and save markdown
        logger.info("  [7/7] Generating markdown...")
        if no_vision:
            extraction_method = f"{extraction_method}_text_only"
        else:
            extraction_method = f"{extraction_method}_with_vision"

        markdown = self.generate_markdown(
            metadata, sections, figures, tables, summary,
            quality, extraction_method, page_count, full_text
        )

        output_path = self.output_dir / output_filename
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(markdown)

        self._existing_files.add(output_stem)
        self._log_quality(pdf_name, quality, "SAVED", output_filename)

        logger.info(f"  SAVED: {output_filename} ({quality['overall']}/10, {self._api_call_count} API calls)")
        return True

    # ─── Batch Processing ─────────────────────────────────────────────────────

    def process_all_papers(self, skip_existing: bool = True,
                            force_ocr: bool = False, no_vision: bool = False):
        """Process all PDF files in the input folder."""
        self._init_quality_log()

        if self.recursive:
            pdf_files = sorted(self.input_folder.rglob('*.pdf'))
        else:
            pdf_files = sorted(self.input_folder.glob('*.pdf'))

        if not pdf_files:
            logger.warning(f"No PDF files found in: {self.input_folder}")
            return

        logger.info(f"Found {len(pdf_files)} PDF files to process")
        logger.info(f"Output directory: {self.output_dir}")
        logger.info(f"Vision AI: {'disabled' if no_vision else 'enabled'}")
        logger.info(f"Budget: {self._budget_cap or 'unlimited'} calls per paper")
        logger.info("")

        success_count = 0
        fail_count = 0

        for i, pdf_path in enumerate(pdf_files, 1):
            logger.info(f"\n[{i}/{len(pdf_files)}] {pdf_path.name}")
            try:
                result = self.process_single_paper(
                    pdf_path, skip_existing=skip_existing,
                    force_ocr=force_ocr, no_vision=no_vision
                )
                if result:
                    success_count += 1
                else:
                    fail_count += 1
            except Exception as e:
                logger.error(f"  EXCEPTION: {e}")
                fail_count += 1

        logger.info(f"\n{'='*60}")
        logger.info(f"BATCH COMPLETE: {success_count} succeeded, {fail_count} failed "
                   f"out of {len(pdf_files)} papers")
        logger.info(f"Quality log: {self.quality_log_path}")
        logger.info(f"{'='*60}")


def main():
    """Main entry point for paper processing pipeline."""
    import argparse

    parser = argparse.ArgumentParser(
        description='Process scientific paper PDFs to structured markdown with AI',
        epilog='Uses selective Vision AI for metadata extraction and figure analysis.'
    )
    parser.add_argument('--input', type=str, required=True,
                       help='Path to folder containing paper PDFs')
    parser.add_argument('--output', type=str, default='output_papers',
                       help='Output directory for markdown files (default: output_papers)')
    parser.add_argument('--single', type=str,
                       help='Process single PDF file (provide full path)')
    parser.add_argument('--recursive', action='store_true',
                       help='Recursively search subfolders for PDF files')
    parser.add_argument('--skip-existing', action='store_true', default=True,
                       help='Skip papers that already have output files (default: True)')
    parser.add_argument('--no-skip', action='store_true',
                       help='Reprocess papers even if output files exist')
    parser.add_argument('--force-ocr', action='store_true',
                       help='Force OCR even if native text extraction works')
    parser.add_argument('--naming', default='default',
                       choices=['default', 'detailed', 'doi'],
                       help="Output filename scheme: "
                            "'default' = Author_Year_Title; "
                            "'detailed' = Author_et_al_Year_Journal_Title; "
                            "'doi' = DOI-based")
    parser.add_argument('--max-vision-pages', type=int, default=8,
                       help='Max figure pages to analyze with Vision AI (default: 8)')
    parser.add_argument('--budget', type=int, default=50,
                       help='Max API calls per paper (default: 50, 0=unlimited)')
    parser.add_argument('--no-vision', action='store_true',
                       help='Skip all Vision AI calls (text-only extraction)')
    parser.add_argument('--ocr-dpi', type=int, default=200,
                       help='DPI for OCR if needed (default: 200)')
    parser.add_argument('--verbose', action='store_true',
                       help='Enable verbose/debug logging')

    args = parser.parse_args()

    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    # Validate paths
    try:
        validate_path(args.input, must_exist=True, allow_dir=True, allow_file=False)
        validate_output_path(args.output)
        if args.single:
            validate_path(args.single, must_exist=True, allow_file=True, allow_dir=False)
    except (ValueError, FileNotFoundError) as e:
        logger.error(f"Path validation failed: {e}")
        return

    try:
        pipeline = PaperPipeline(
            input_folder=args.input,
            output_dir=args.output,
            recursive=args.recursive,
            naming=args.naming,
            max_vision_pages=args.max_vision_pages,
            budget=args.budget,
        )

        if args.single:
            logger.info(f"Single file mode: {args.single}")
            pdf_path = Path(args.single)
            if not pdf_path.exists():
                logger.error(f"File not found: {pdf_path}")
                return
            pipeline.process_single_paper(
                pdf_path,
                skip_existing=not args.no_skip,
                force_ocr=args.force_ocr,
                no_vision=args.no_vision,
            )
        else:
            pipeline.process_all_papers(
                skip_existing=not args.no_skip,
                force_ocr=args.force_ocr,
                no_vision=args.no_vision,
            )

    except Exception as e:
        logger.error(f"Pipeline failed: {e}")
        raise


if __name__ == '__main__':
    main()
