"""
Patent Processing Pipeline with Selective Vision AI Analysis
Converts pharmaceutical/chemical patent PDFs to structured markdown

Features:
- Multi-page PDF handling with page classification
- Selective Vision AI for figure/drawing pages only
- Vision AI for cover page (reliable applicant/inventor extraction)
- Claims parsing with dependency tree
- Chemical structure extraction with SMILES attempts
- AI-generated executive summary and protection scope
- Quality scoring adapted for patent documents
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

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Load credentials from Windows User environment if not in current environment
if not os.environ.get('ANTHROPIC_AUTH_TOKEN') or not os.environ.get('ANTHROPIC_BASE_URL'):
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, 'Environment', 0, winreg.KEY_READ)
        try:
            token, _ = winreg.QueryValueEx(key, 'ANTHROPIC_AUTH_TOKEN')
            os.environ['ANTHROPIC_AUTH_TOKEN'] = token
            logger.info("Loaded ANTHROPIC_AUTH_TOKEN from Windows User environment")
        except:
            pass
        try:
            url, _ = winreg.QueryValueEx(key, 'ANTHROPIC_BASE_URL')
            os.environ['ANTHROPIC_BASE_URL'] = url
            logger.info("Loaded ANTHROPIC_BASE_URL from Windows User environment")
        except:
            pass
        winreg.CloseKey(key)
    except:
        pass


class PatentPipeline:
    """Pipeline for processing pharmaceutical patent PDFs into structured markdown"""

    VISION_MODEL = "claude-sonnet-4-6"
    PAGES_PER_VISION_BATCH = 4
    MAX_CONCURRENT_BATCHES = 3
    RENDER_DPI = 200
    MAX_IMAGE_DIMENSION = 1568
    MIN_TEXT_FOR_TEXT_PAGE = 100
    QUALITY_EXCELLENT_THRESHOLD = 8.0
    QUALITY_GOOD_THRESHOLD = 5.5
    QUALITY_FAIR_THRESHOLD = 4.0

    # Figure page limits
    DEFAULT_MAX_FIGURE_PAGES = 50
    MAX_SMILES_REFINEMENT = 50

    # Budget / cost controls
    DEFAULT_BUDGET_CAP = 200

    # Chemical structure refinement
    CHEMICAL_STRUCTURE_DPI = 250
    MAX_IMAGE_DIMENSION_CHEMICAL = 2048
    SMILES_VALID_CHARS = re.compile(r'^[A-Za-z0-9@+\-\[\]()=#$/\\.%:]+$')

    # Text quality thresholds for Vision AI triggering
    TEXT_QUALITY_GOOD_THRESHOLD = 0.65
    TEXT_QUALITY_GARBLED_THRESHOLD = 0.30
    MAX_VISION_TEXT_PAGES = 30
    VISION_DPI_TEXT_ENHANCEMENT = 250
    MAX_IMAGE_DIMENSION_TEXT = 1568
    PAGES_PER_VISION_TEXT_BATCH = 4

    # Tesseract OCR configuration (local, free, fast)
    TESSERACT_DPI = 200
    TESSERACT_BATCH_SIZE = 20
    TESSERACT_CONFIG = '--psm 6 --oem 1'
    TESSERACT_MIN_CHARS = 50

    # Known encoding errors in pharma/scientific PDF text (ligature decomposition)
    ENCODING_FIXES = {
        'molecuies': 'molecules',
        'molecuie': 'molecule',
        'moiecuie': 'molecule',
        'moiecuies': 'molecules',
        'moiecuiar': 'molecular',
        'payioads': 'payloads',
        'payioad': 'payload',
        'poiyubiquitination': 'polyubiquitination',
        'poiy': 'poly',
        'ceii': 'cell',
        'ceils': 'cells',
        'coniugate': 'conjugate',
        'coniugates': 'conjugates',
        'coniugated': 'conjugated',
        'coniugation': 'conjugation',
        'Bioconiugate': 'Bioconjugate',
        'surpnsingly': 'surprisingly',
        'particuiar': 'particular',
        'particuiariy': 'particularly',
        'cieavabie': 'cleavable',
        'cieavable': 'cleavable',
        'avaiiable': 'available',
        'availabie': 'available',
        'preferentiaiiy': 'preferentially',
        'preferentiaily': 'preferentially',
        'seif': 'self',
        'haif': 'half',
        'haif-iife': 'half-life',
        'fiuorescent': 'fluorescent',
        'fiuoro': 'fluoro',
        'Olin': 'Clin',
        'aikyiene': 'alkylene',
        'aIkyI': 'alkyl',
        'aikyi': 'alkyl',
        'cycioaikyl': 'cycloalkyl',
        'cycloaikyl': 'cycloalkyl',
        'cycioalkyl': 'cycloalkyl',
        'cycloaikyiene': 'cycloalkylene',
        'heteroaikyiene': 'heteroalkylene',
        'heteroaikyi': 'heteroalkyl',
        'heteroaryiene': 'heteroarylene',
        'heterocycioalkyl': 'heterocycloalkyl',
        'heterocycloaikyiene': 'heterocycloalkylene',
        'heterocyciyi': 'heterocyclyl',
        'heterocyciyiene': 'heterocyclylene',
        'pharmaceuticaiiy': 'pharmaceutically',
        'pharmaceuticaily': 'pharmaceutically',
        'optionaiiy': 'optionally',
        'optionaily': 'optionally',
        'independentiy': 'independently',
        'independentIy': 'independently',
        'denvative': 'derivative',
        'pyrazoio': 'pyrazolo',
    }

    PATENT_SECTION_PATTERNS = {
        'technical_field': r'(?:Technical\s+[Ff]ield|TECHNICAL\s+FIELD|Field\s+of\s+(?:the\s+)?[Ii]nvention|FIELD\s+OF\s+(?:THE\s+)?INVENTION)',
        'background': r'(?:Background|BACKGROUND|Prior\s+[Aa]rt|PRIOR\s+ART|Background\s+of\s+the\s+[Ii]nvention)',
        'summary': r'(?:Summary|SUMMARY|Summary\s+of\s+(?:the\s+)?[Ii]nvention|SUMMARY\s+OF\s+(?:THE\s+)?INVENTION)',
        'claims': r'(?:Claims|CLAIMS|[Ww]hat\s+[Ii]s\s+[Cc]laimed|WHAT\s+IS\s+CLAIMED)',
        'detailed_description': r'(?:Detailed\s+[Dd]escription|DETAILED\s+DESCRIPTION|Description\s+of\s+(?:the\s+)?[Ee]mbodiments)',
        'examples': r'(?:Examples?\s|EXAMPLES?\s|Experimental|EXPERIMENTAL)',
        'drawings': r'(?:Brief\s+[Dd]escription\s+of\s+(?:the\s+)?[Dd]rawings|BRIEF\s+DESCRIPTION\s+OF\s+(?:THE\s+)?DRAWINGS|Legends?\s+of\s+(?:the\s+)?[Ff]igures)',
        'definitions': r'(?:Definitions\n|DEFINITIONS\n)',
    }

    # Patent page header pattern (repeats on every page)
    PAGE_HEADER_PATTERN = re.compile(
        r'^(?:WO|EP|US)\s*\d{4}/?\d+\s*(?:PCT/[A-Z]{2}\d{4}/\d+)?\s*$',
        re.MULTILINE
    )
    # Margin line numbers (5, 10, 15, 20, 25, 30, 35 at line starts)
    # Patent margins use multiples of 5; match these specifically to avoid stripping real numbers
    MARGIN_NUMBER_PATTERN = re.compile(r'^(5|10|15|20|25|30|35|40|45|50)\s+(?=[A-Za-z("])', re.MULTILINE)

    CLAIM_DEPENDENCY_PATTERN = re.compile(
        r'(?:of|according to|as (?:defined|claimed) in)\s+(?:any\s+(?:one\s+)?of\s+)?claims?\s+(\d+)',
        re.IGNORECASE
    )

    NAMING_SCHEMES = {
        'default': 'patent_{patent_id}',
        'detailed': 'patent_{patent_id}_{applicant}_{title_slug}',
        'dated': '{pub_date}_{patent_id}_{title_slug}',
    }

    def __init__(self, output_dir: str = "output_patents", budget: int = 200,
                 ocr_engine: str = "auto", naming: str = "default",
                 recursive: bool = False):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)
        self.quality_log_path = self.output_dir / "quality_log.txt"
        self._client = None
        self._existing_files = set(f.stem for f in self.output_dir.glob("*.md"))
        self._budget_cap = budget if budget > 0 else None
        self._api_call_count = 0
        self._ocr_engine = ocr_engine
        self._naming = naming
        self._recursive = recursive
        self._tesseract_available = self._check_tesseract_available()

    def _check_tesseract_available(self) -> bool:
        """Check if Tesseract binary is installed and accessible."""
        if self._ocr_engine == 'vision':
            return False
        try:
            import pytesseract
            pytesseract.get_tesseract_version()
            return True
        except Exception:
            if self._ocr_engine == 'tesseract':
                logger.error("Tesseract not found but --ocr-engine=tesseract was specified")
            else:
                logger.info("Tesseract not available — will use Vision AI for OCR")
            return False

    def _generate_output_filename(self, patent_id: str, bib: Dict) -> str:
        """Generate output filename based on naming scheme."""
        if self._naming == 'default':
            return f"patent_{patent_id}"

        # Build template variables
        title = bib.get('title', '')
        title_slug = re.sub(r'[^\w\s]', '', title.lower()).split()[:5]
        title_slug = '_'.join(title_slug) if title_slug else 'untitled'

        applicants = bib.get('applicants', [])
        applicant = ''
        if applicants:
            # First applicant, take first meaningful word (skip "THE", "A")
            raw = str(applicants[0]).strip()
            words = [w for w in raw.split() if w.upper() not in ('THE', 'A', 'AN', 'INC.', 'INC', 'LTD', 'LTD.', 'CO.', 'CORP.', 'CORPORATION')]
            applicant = re.sub(r'[^\w]', '', words[0]) if words else 'unknown'

        pub_date = bib.get('publication_date', '')
        # Normalize date to YYYY-MM-DD format
        date_match = re.search(r'(\d{1,2})\s+(\w+)\s+(\d{4})', pub_date)
        if date_match:
            day, month_str, year = date_match.groups()
            months = {'january': '01', 'february': '02', 'march': '03', 'april': '04',
                      'may': '05', 'june': '06', 'july': '07', 'august': '08',
                      'september': '09', 'october': '10', 'november': '11', 'december': '12'}
            month_num = months.get(month_str.lower(), '00')
            pub_date_fmt = f"{year}-{month_num}-{day.zfill(2)}"
        else:
            pub_date_fmt = pub_date.replace(' ', '-') if pub_date else 'nodate'

        template = self.NAMING_SCHEMES.get(self._naming, self.NAMING_SCHEMES['default'])
        filename = template.format(
            patent_id=patent_id,
            applicant=applicant,
            title_slug=title_slug,
            pub_date=pub_date_fmt
        )
        # Sanitize: remove any remaining invalid filename characters
        filename = re.sub(r'[<>:"/\\|?*]', '', filename)
        filename = re.sub(r'\s+', '_', filename)
        return filename

    # ─── Shared Utilities ──────────────────────────────────────────────────

    def get_api_config(self) -> Tuple[Optional[str], Optional[str]]:
        auth_token = os.environ.get('ANTHROPIC_AUTH_TOKEN')
        base_url = os.environ.get('ANTHROPIC_BASE_URL')
        if auth_token and base_url:
            return auth_token, base_url
        api_key = os.environ.get('ANTHROPIC_API_KEY')
        if api_key:
            return api_key, None
        return None, None

    def get_anthropic_client(self):
        if self._client is not None:
            return self._client
        from anthropic import Anthropic
        api_key, base_url = self.get_api_config()
        if not api_key:
            logger.error("Cannot create Anthropic client - API credentials not found")
            return None
        if base_url:
            self._client = Anthropic(api_key=api_key, base_url=base_url)
        else:
            self._client = Anthropic(api_key=api_key)
        return self._client

    def _track_api_call(self, count: int = 1) -> bool:
        """Track API calls and check budget. Returns False if budget exceeded."""
        self._api_call_count += count
        if self._budget_cap and self._api_call_count > self._budget_cap:
            logger.warning(f"  BUDGET EXCEEDED: {self._api_call_count}/{self._budget_cap} API calls")
            return False
        return True

    def _budget_remaining(self) -> int:
        """Return remaining API calls before budget is hit, or 999 if unlimited."""
        if not self._budget_cap:
            return 999
        return max(0, self._budget_cap - self._api_call_count)

    def _estimate_and_log_cost(self, page_data: List[Dict], no_vision: bool) -> Dict[str, int]:
        """Estimate API calls needed and log. Returns per-stage estimates."""
        figure_pages = sum(1 for p in page_data if p['classification'] == 'FIGURE_PAGE')
        text_heavy = sum(1 for p in page_data if p['classification'] == 'TEXT_HEAVY')
        effective_fig = min(figure_pages, self.DEFAULT_MAX_FIGURE_PAGES)

        est = {
            'cover_vision': 0 if no_vision else 1,
            'text_enhancement': 0,
            'figure_analysis': 0 if no_vision else -(-effective_fig // self.PAGES_PER_VISION_BATCH),
            'key_data': 0 if no_vision else 1,
            'smiles_refinement': 0,
            'summaries': 0 if no_vision else 1,
        }
        est['total'] = sum(est.values())

        logger.info(f"  API call estimate: ~{est['total']} calls "
                   f"(figures: {est['figure_analysis']}, other: {est['total'] - est['figure_analysis']})")
        if self._budget_cap:
            logger.info(f"  Budget: {self._budget_cap} calls "
                       f"(remaining: {self._budget_remaining()})")
        return est

    def encode_image_to_base64(self, image, format="JPEG", max_dimension=None,
                              quality: int = 85) -> Tuple[Optional[str], Optional[str]]:
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

    def render_pages_to_images(self, pdf_path: Path, page_numbers: List[int], dpi: int = None) -> Dict[int, Any]:
        """Render multiple pages from a single PDF open — avoids repeated open/close."""
        if dpi is None:
            dpi = self.RENDER_DPI
        images = {}
        try:
            import fitz
            from PIL import Image
            Image.MAX_IMAGE_PIXELS = None
            doc = fitz.open(str(pdf_path))
            zoom = dpi / 72
            mat = fitz.Matrix(zoom, zoom)
            for page_num in page_numbers:
                if page_num < len(doc):
                    pix = doc[page_num].get_pixmap(matrix=mat)
                    images[page_num] = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            doc.close()
        except Exception as e:
            logger.error(f"Error rendering pages: {e}")
        return images

    def parse_json_response(self, response_text: str) -> Optional[Any]:
        try:
            return json.loads(response_text)
        except json.JSONDecodeError:
            pass
        # Try extracting from code block
        json_match = re.search(r'```(?:json)?\s*\n?(.*?)\n?```', response_text, re.DOTALL)
        if json_match:
            try:
                return json.loads(json_match.group(1))
            except json.JSONDecodeError:
                pass
        # Try finding array or object
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

    def clean_patent_text(self, text: str) -> str:
        """Remove patent page headers, margin numbers, and repair common text extraction errors."""
        text = self.PAGE_HEADER_PATTERN.sub('', text)
        text = self.MARGIN_NUMBER_PATTERN.sub('', text)
        text = self._repair_encoding_errors(text)
        text = self._repair_text_spacing(text)
        text = self._remove_stray_characters(text)
        text = re.sub(r'\n{3,}', '\n\n', text)
        return text.strip()

    # Safe whitelist for specific merged-word patterns that won't break real words.
    # Uses \b (word boundary) to match only standalone tokens or at word edges.
    _SAFE_MERGE_FIXES = [
        # "of" at start of word followed by non-f + 5+ lowercase (safe: no English words match)
        (re.compile(r'\bof([^f\s][a-z]{5,})'), r'of \1'),
        # Common preposition + "the" merges (standalone words only)
        (re.compile(r'\bofthe\b'), 'of the'),
        (re.compile(r'\binthe\b'), 'in the'),
        (re.compile(r'\btothe\b'), 'to the'),
        (re.compile(r'\bforthe\b'), 'for the'),
        (re.compile(r'\bfromthe\b'), 'from the'),
        (re.compile(r'\bwiththe\b'), 'with the'),
        (re.compile(r'\bbythe\b'), 'by the'),
        (re.compile(r'\bonthe\b'), 'on the'),
        (re.compile(r'\batthe\b'), 'at the'),
        (re.compile(r'\basthe\b'), 'as the'),
        (re.compile(r'\bandthe\b'), 'and the'),
        (re.compile(r'\bisthe\b'), 'is the'),
        (re.compile(r'\bthatthe\b'), 'that the'),
        # Words ending with "the" + rest (standalone tokens)
        (re.compile(r'\breachesthe\b'), 'reaches the'),
        (re.compile(r'\bexplainsthe\b'), 'explains the'),
        # "their" merges at word start (no English words start with "their" as prefix)
        (re.compile(r'\btheir([a-z]{4,})'), r'their \1'),
        # Multi-word phrases
        (re.compile(r'\binparticular\b'), 'in particular'),
        (re.compile(r'\binpartwhy\b'), 'in part why'),
        (re.compile(r'\binpart\b'), 'in part'),
        (re.compile(r'\binorder\b'), 'in order'),
        (re.compile(r'\bsuchthat\b'), 'such that'),
        (re.compile(r'\binvitro\b'), 'in vitro'),
        (re.compile(r'\binvivo\b'), 'in vivo'),
        (re.compile(r'\baswell\b'), 'as well'),
        (re.compile(r'\binwhich\b'), 'in which'),
        (re.compile(r'\bofwhich\b'), 'of which'),
        # Patent claim language merges (explicit for short words not caught by general 'of' rule)
        (re.compile(r'\bofclaim\b'), 'of claim'),
        (re.compile(r'\bofformula\b'), 'of formula'),
        (re.compile(r'\btoclaim\b'), 'to claim'),
        (re.compile(r'\baccordingto\b'), 'according to'),
        (re.compile(r'\bselectedfrom\b'), 'selected from'),
        (re.compile(r'\bwhereinthe\b'), 'wherein the'),
        # Scientific text common merges
        (re.compile(r'\bnewclass\b'), 'new class'),
        (re.compile(r'\bareaclass\b'), 'are a class'),
        (re.compile(r'\baverysmall\b'), 'a very small'),
        # "ly" suffix followed by next word (adverb+verb/adj merges)
        (re.compile(r'\b([a-z]{4,}ly)(targeting|binding|acting|reducing|promoting|explaining)\b'), r'\1 \2'),
        (re.compile(r'\b([a-z]{4,}ly)(expressed|described|developed|produced|achieved|reported)\b'), r'\1 \2'),
    ]

    def _repair_text_spacing(self, text: str) -> str:
        """Insert missing spaces at word boundaries caused by PDF extraction."""
        # Split at lowercase→Uppercase when followed by 3+ lowercase chars
        # Avoids splitting: pH, mRNA, IgG, CamelCase compound names like BamHI
        text = re.sub(r'([a-z])([A-Z][a-z]{2,})', r'\1 \2', text)

        # Space after closing paren before uppercase: ").This" → "). This"
        text = re.sub(r'(\))([A-Z])', r'\1 \2', text)

        # Space after period before uppercase+lowercase: "drug.The" → "drug. The"
        # Avoids splitting abbreviations like "U.S." or "Fig.1"
        text = re.sub(r'([a-z]\.)([A-Z][a-z]{2,})', r'\1 \2', text)

        # Space after comma before uppercase: "drugs,The" → "drugs, The"
        text = re.sub(r',([A-Z][a-z]{2,})', r', \1', text)

        # Safe whitelist fixes for specific merged-word patterns
        for pattern, replacement in self._SAFE_MERGE_FIXES:
            text = pattern.sub(replacement, text)

        return text

    def _repair_encoding_errors(self, text: str) -> str:
        """Fix known PDF font encoding/ligature errors common in pharma patents."""
        for bad, good in self.ENCODING_FIXES.items():
            text = re.sub(re.escape(bad), good, text)
        return text

    def _remove_stray_characters(self, text: str) -> str:
        """Remove stray single characters on their own lines (column separator artifacts)."""
        # Lines that are just a single letter/symbol (not part of a list like "a." or "1.")
        text = re.sub(r'^\s*[A-Z|l1]\s*$', '', text, flags=re.MULTILINE)
        # Lines that are 1-3 non-word characters (stray punctuation/symbols)
        text = re.sub(r'^\s*[^\w\s]{1,3}\s*$', '', text, flags=re.MULTILINE)
        return text

    # ─── Text Quality Assessment ─────────────────────────────────────────────

    def score_page_text_quality(self, text: str) -> Dict[str, float]:
        """Score a single page's text quality (0.0=garbled, 1.0=perfect).

        Detects both camelCase merges (lowercase→Uppercase) and all-lowercase
        merged words (high avg word length, many very long words).
        """
        if not text or len(text.strip()) < 50:
            return {'overall': 0.0, 'merged_words': 0.0, 'avg_word_length': 0.0,
                    'long_words': 0.0, 'garbled_chars': 0.0, 'alpha_ratio': 0.0,
                    'tiny_words': 0.0}

        words = text.split()
        word_count = len(words)
        if word_count < 5:
            return {'overall': 0.0, 'merged_words': 0.0, 'avg_word_length': 0.0,
                    'long_words': 0.0, 'garbled_chars': 0.0, 'alpha_ratio': 0.0,
                    'tiny_words': 0.0}

        # 1. Merged word detection: lowercase→Uppercase transitions within words
        merge_count = len(re.findall(r'[a-z][A-Z]', text))
        merged_score = max(0.0, 1.0 - (merge_count / max(word_count, 1)) * 5)

        # 2. Average word length (very long = merged words)
        avg_len = sum(len(w) for w in words) / word_count
        if avg_len <= 7:
            avg_len_score = 1.0
        elif avg_len <= 9:
            avg_len_score = 1.0 - (avg_len - 7) * 0.2
        else:
            avg_len_score = max(0.0, 0.6 - (avg_len - 9) / 10)

        # 3. Long word ratio: words >20 chars are almost always merged
        alpha_words = [w for w in words if len(w) > 2 and w[0].isalpha()]
        long_words = sum(1 for w in alpha_words if len(w) > 20)
        long_word_ratio = long_words / max(len(alpha_words), 1)
        long_word_score = max(0.0, 1.0 - long_word_ratio * 10)

        # 4. Garbled characters: non-standard symbols density
        garbled_chars = sum(1 for c in text if c in '~`^°±§¶©®™«»¿¡')
        garbled_ratio = garbled_chars / max(len(text), 1)
        garbled_score = max(0.0, 1.0 - garbled_ratio * 50)

        # 5. Alpha ratio
        alpha_chars = sum(1 for c in text if c.isalpha())
        alpha_ratio = alpha_chars / max(len(text), 1)
        alpha_score = min(1.0, alpha_ratio / 0.6)

        # 6. Tiny word ratio (1-2 char words) — too few short words = merged text
        tiny_words = sum(1 for w in words if len(w) <= 2)
        tiny_ratio = tiny_words / word_count
        tiny_score = max(0.0, 1.0 - tiny_ratio * 2.5)

        # Weighted overall — avg_word_length and long_words catch all-lowercase merges
        overall = (merged_score * 0.20 + avg_len_score * 0.25 +
                   long_word_score * 0.15 + garbled_score * 0.15 +
                   alpha_score * 0.10 + tiny_score * 0.15)
        overall = max(0.0, min(1.0, overall))

        return {
            'overall': round(overall, 3),
            'merged_words': round(merged_score, 3),
            'avg_word_length': round(avg_len_score, 3),
            'long_words': round(long_word_score, 3),
            'garbled_chars': round(garbled_score, 3),
            'alpha_ratio': round(alpha_score, 3),
            'tiny_words': round(tiny_score, 3),
        }

    def assess_all_pages_quality(self, page_data: List[Dict]) -> List[Dict]:
        """Score all TEXT_HEAVY pages and identify those needing Vision AI re-extraction.

        Modifies page_data in-place (adds quality scores, reclassifies garbled pages).
        Returns list of pages needing Vision AI RAG enhancement (sorted worst-first).
        """
        logger.info("Stage 1.5: Assessing text quality on all TEXT_HEAVY pages...")
        pages_good = 0
        pages_vision = []
        pages_reclassified = 0

        for page in page_data:
            if page['classification'] != 'TEXT_HEAVY':
                continue

            scores = self.score_page_text_quality(page['text'])
            page['text_quality_score'] = scores['overall']
            page['text_quality_details'] = scores

            if scores['overall'] >= self.TEXT_QUALITY_GOOD_THRESHOLD:
                page['remediation'] = 'none'
                pages_good += 1
            elif scores['overall'] < self.TEXT_QUALITY_GARBLED_THRESHOLD:
                page['remediation'] = 'reclassify_figure'
                page['classification'] = 'FIGURE_PAGE'
                pages_reclassified += 1
            else:
                page['remediation'] = 'vision_rag'
                pages_vision.append(page)

        # Cap vision pages (take worst-scoring first)
        pages_vision.sort(key=lambda p: p['text_quality_score'])
        if len(pages_vision) > self.MAX_VISION_TEXT_PAGES:
            for page in pages_vision[self.MAX_VISION_TEXT_PAGES:]:
                page['remediation'] = 'none'
            pages_vision = pages_vision[:self.MAX_VISION_TEXT_PAGES]

        logger.info(f"  Text quality: {pages_good} good, "
                    f"{len(pages_vision)} need Vision RAG, "
                    f"{pages_reclassified} reclassified as figure pages")

        if pages_vision:
            worst = pages_vision[0]
            logger.info(f"  Worst text page: p{worst['page_num']+1} "
                        f"(score={worst['text_quality_score']:.3f})")

        return pages_vision

    # ─── Selective Vision AI Text Enhancement ─────────────────────────────────

    def enhance_text_pages_vision(self, pdf_path: Path, pages_to_enhance: List[Dict]) -> Dict[int, str]:
        """Re-extract text for problematic pages using Vision AI with RAG context.

        Uses garbled pdfplumber text as RAG context combined with page image.
        Batches 2 pages per API call with parallel execution.

        Returns:
            Dict mapping page_num -> enhanced_text
        """
        if not pages_to_enhance:
            return {}

        client = self.get_anthropic_client()
        if not client:
            logger.warning("  Cannot enhance text pages - no API client")
            return {}

        logger.info(f"  Re-extracting {len(pages_to_enhance)} pages with Vision AI (RAG)...")

        # Build batches of 2 pages each
        batches = []
        for i in range(0, len(pages_to_enhance), self.PAGES_PER_VISION_TEXT_BATCH):
            batches.append(pages_to_enhance[i:i + self.PAGES_PER_VISION_TEXT_BATCH])

        enhanced = {}
        with ThreadPoolExecutor(max_workers=self.MAX_CONCURRENT_BATCHES) as executor:
            futures = {}
            for batch_idx, batch in enumerate(batches):
                future = executor.submit(self._enhance_text_batch, pdf_path, batch, batch_idx)
                futures[future] = batch_idx

            for future in as_completed(futures):
                batch_idx = futures[future]
                try:
                    result = future.result()
                    enhanced.update(result)
                    logger.info(f"  Text batch {batch_idx+1}/{len(batches)} complete "
                                f"({len(result)} pages enhanced)")
                except Exception as e:
                    logger.error(f"  Text batch {batch_idx+1} failed: {e}")

        logger.info(f"  Vision AI text enhancement complete: {len(enhanced)} pages improved")
        return enhanced

    def _enhance_text_batch(self, pdf_path: Path, batch: List[Dict], batch_idx: int) -> Dict[int, str]:
        """Enhance a batch of text pages with Vision AI."""
        client = self.get_anthropic_client()
        if not client:
            return {}

        page_numbers = [p['page_num'] for p in batch]
        images = self.render_pages_to_images(pdf_path, page_numbers, dpi=self.VISION_DPI_TEXT_ENHANCEMENT)

        content = []
        valid_pages = []
        for page in batch:
            page_num = page['page_num']
            if page_num not in images:
                continue
            img_base64, media_type = self.encode_image_to_base64(
                images[page_num], max_dimension=self.MAX_IMAGE_DIMENSION_TEXT)
            del images[page_num]
            if not img_base64:
                continue

            content.append({
                "type": "image",
                "source": {"type": "base64", "media_type": media_type, "data": img_base64}
            })

            garbled_text = page['text'][:4000]
            content.append({
                "type": "text",
                "text": f"[PAGE {page_num + 1} - GARBLED REFERENCE TEXT (use as content guide):\n"
                        f"{garbled_text}\n"
                        f"--- END REFERENCE ---]"
            })
            valid_pages.append(page_num)

        if not content:
            return {}

        prompt = self._build_text_enhancement_prompt(valid_pages)
        content.append({"type": "text", "text": prompt})

        try:
            response = client.messages.create(
                model=self.VISION_MODEL,
                max_tokens=4096 * len(valid_pages),
                messages=[{"role": "user", "content": content}]
            )
            response_text = response.content[0].text
            logger.info(f"    Batch {batch_idx+1}: {len(response_text)} chars "
                        f"(in:{response.usage.input_tokens}, out:{response.usage.output_tokens})")
            return self._parse_enhanced_text_response(response_text, valid_pages)
        except Exception as e:
            logger.error(f"    Text enhancement API call failed: {e}")
            return {}

    def _build_text_enhancement_prompt(self, page_numbers: List[int]) -> str:
        """Build the Vision AI prompt for text re-extraction."""
        pages_str = ', '.join(str(p + 1) for p in page_numbers)
        return f"""Re-extract the text from these patent document pages ({pages_str}). The PDF text extraction produced garbled output due to font encoding issues.

For each page, I provided the garbled reference text as a content guide — it contains the correct content but with errors including:
- Words merged without spaces: "potencyrequirementexplains" should be "potency requirement explains"
- Font encoding errors where 'l' becomes 'i': "molecuies" → "molecules", "payioads" → "payloads"
- Missing spaces after periods/parentheses: ").This" → "). This"
- Stray single characters from column separators

Your task:
1. Read each page image carefully
2. Use the garbled text as a content guide (it tells you WHAT is on the page)
3. Produce clean, properly-spaced text with correct spelling
4. Preserve ALL scientific terminology, chemical names, compound numbers, formulas
5. Maintain paragraph structure, section headers, and numbered lists
6. If a page contains chemical structure DIAGRAMS (graphical molecular drawings, not text formulas), write: [CHEMICAL STRUCTURE: Figure label - brief description of the structure]
7. Do NOT add commentary or content not visible on the page

Output the clean text for each page, separated by:
===PAGE N===
(where N is the page number)

Start with ===PAGE {page_numbers[0] + 1}==="""

    def _parse_enhanced_text_response(self, response_text: str, page_numbers: List[int]) -> Dict[int, str]:
        """Parse the Vision AI response back into per-page text."""
        result = {}

        if len(page_numbers) == 1:
            # Single page - strip any page marker and use whole response
            cleaned = re.sub(r'^===PAGE\s+\d+===\s*\n?', '', response_text).strip()
            result[page_numbers[0]] = cleaned
            return result

        # Multi-page: split on page markers
        parts = re.split(r'===PAGE\s+(\d+)===', response_text)
        # parts alternates: [preamble, page_num, text, page_num, text, ...]
        for i in range(1, len(parts) - 1, 2):
            try:
                page_num_from_response = int(parts[i]) - 1  # Convert to 0-indexed
                page_text = parts[i + 1].strip()
                if page_num_from_response in page_numbers:
                    result[page_num_from_response] = page_text
            except (ValueError, IndexError):
                continue

        # Fallback: if parsing failed but we have content, assign to first page
        if not result and response_text.strip():
            result[page_numbers[0]] = response_text.strip()

        return result

    # ─── Stage 0: PDF Characterization & Page Classification ───────────────

    def characterize_pdf(self, pdf_path: Path) -> Dict:
        import pdfplumber

        logger.info("Stage 0: Characterizing PDF and classifying pages...")
        pdf = pdfplumber.open(str(pdf_path))
        total_pages = len(pdf.pages)
        logger.info(f"  Total pages: {total_pages}")

        page_data = []
        for i in range(total_pages):
            text = pdf.pages[i].extract_text() or ""
            page_data.append({
                'page_num': i,
                'text': text,
                'text_length': len(text),
                'classification': self._classify_page(text, i, total_pages)
            })

        pdf.close()

        classifications = {}
        for pd_item in page_data:
            cls = pd_item['classification']
            classifications[cls] = classifications.get(cls, 0) + 1

        logger.info(f"  Page classifications: {classifications}")
        return {
            'total_pages': total_pages,
            'page_data': page_data,
            'classifications': classifications
        }

    def _classify_page(self, text: str, page_num: int, total_pages: int) -> str:
        if page_num == 0:
            return 'COVER'

        if page_num >= total_pages - 10:
            if re.search(r'INTERNATIONAL\s+SEARCH\s+REPORT|Patent\s+family\s+members', text, re.IGNORECASE):
                return 'SEARCH_REPORT'

        text_length = len(text.strip())
        if text_length < self.MIN_TEXT_FOR_TEXT_PAGE:
            return 'FIGURE_PAGE'

        return 'TEXT_HEAVY'

    # ─── Stage 0.5: Image-Only PDF Detection & OCR ────────────────────────

    def _detect_image_only_pdf(self, pdf_info: Dict) -> bool:
        """Detect if PDF is image-only (scanned) with no extractable text."""
        classifications = pdf_info['classifications']
        total = pdf_info['total_pages']
        text_heavy = classifications.get('TEXT_HEAVY', 0)
        return text_heavy == 0 and total > 10

    def _sample_page_types_vision(self, pdf_path: Path, page_data: List[Dict]) -> Dict[str, List[int]]:
        """Sample evenly-spaced pages to classify as text vs figure using Vision AI."""
        total = len(page_data)
        sample_indices = [total // 6, 2 * total // 6, 3 * total // 6,
                          4 * total // 6, 5 * total // 6]
        sample_indices = [i for i in sample_indices if 0 < i < total]

        client = self.get_anthropic_client()
        if not client:
            return {'text_pages': list(range(1, total)), 'figure_pages': []}

        images = self.render_pages_to_images(pdf_path, sample_indices, dpi=150)
        content = []
        valid_indices = []
        for idx in sample_indices:
            if idx not in images:
                continue
            img_base64, media_type = self.encode_image_to_base64(images[idx], max_dimension=1024)
            del images[idx]
            if not img_base64:
                continue
            content.append({
                "type": "image",
                "source": {"type": "base64", "media_type": media_type, "data": img_base64}
            })
            content.append({"type": "text", "text": f"[Page {idx + 1}]"})
            valid_indices.append(idx)

        if not content:
            return {'text_pages': list(range(1, total)), 'figure_pages': []}

        content.append({"type": "text", "text": f"""Classify each page shown above.
For each page, determine if it is:
- "text": A page primarily containing typed/printed text (patent description, claims, definitions)
- "figure": A page primarily containing diagrams, chemical structures, graphs, or drawings

Return ONLY a JSON array like: [{{"page": 55, "type": "text"}}, {{"page": 111, "type": "figure"}}]"""})

        try:
            self._track_api_call()
            response = client.messages.create(
                model=self.VISION_MODEL,
                max_tokens=1024,
                messages=[{"role": "user", "content": content}]
            )
            results = self.parse_json_response(response.content[0].text)
            if results:
                return self._infer_page_boundaries(results, valid_indices, total)
        except Exception as e:
            logger.error(f"  Page type sampling failed: {e}")

        return {'text_pages': list(range(1, total)), 'figure_pages': []}

    def _infer_page_boundaries(self, sample_results: List[Dict],
                               sampled_indices: List[int], total_pages: int) -> Dict[str, List[int]]:
        """Infer text/figure page boundary from sample classification results.

        Patents typically have text first, then figures at the end.
        Find the transition point from the samples.
        """
        type_map = {}
        for entry in sample_results:
            page_num = entry.get('page', 0) - 1
            page_type = entry.get('type', 'text')
            type_map[page_num] = page_type

        last_text_sample = 0
        first_figure_sample = total_pages
        for idx in sorted(sampled_indices):
            if type_map.get(idx, 'text') == 'text':
                last_text_sample = idx
            else:
                first_figure_sample = min(first_figure_sample, idx)

        if first_figure_sample <= last_text_sample:
            boundary = (last_text_sample + first_figure_sample) // 2
        else:
            boundary = first_figure_sample

        text_pages = list(range(1, boundary))
        figure_pages = list(range(boundary, total_pages))

        logger.info(f"  Inferred boundary: pages 1-{boundary} = text, "
                   f"pages {boundary+1}-{total_pages} = figures")
        return {'text_pages': text_pages, 'figure_pages': figure_pages}

    def _classify_pages_heuristic(self, pdf_path: Path, page_data: List[Dict]) -> Dict[str, List[int]]:
        """Classify pages as text vs figure using local Tesseract (no API calls).

        Samples evenly-spaced pages, runs quick Tesseract OCR at 150 DPI,
        and classifies based on word count and line density (not just word count,
        since chemical structure pages also contain labels/formulas).
        """
        import pytesseract

        total = len(page_data)
        n_samples = min(10, max(5, total // 40))
        step = total // (n_samples + 1)
        sample_indices = [step * (i + 1) for i in range(n_samples)]
        sample_indices = [i for i in sample_indices if 0 < i < total]

        logger.info(f"  Heuristic classification: sampling {len(sample_indices)} pages at 150 DPI...")
        images = self.render_pages_to_images(pdf_path, sample_indices, dpi=150)

        sample_results = []
        for idx in sample_indices:
            if idx not in images:
                continue
            try:
                text = pytesseract.image_to_string(images[idx], config='--psm 6 --oem 1')
                page_type = self._classify_page_from_ocr_text(text)
                sample_results.append({'page': idx + 1, 'type': page_type})
            except Exception:
                sample_results.append({'page': idx + 1, 'type': 'text'})
            del images[idx]

        if not sample_results:
            return {'text_pages': list(range(1, total)), 'figure_pages': []}

        return self._infer_page_boundaries(sample_results, sample_indices, total)

    def _classify_page_from_ocr_text(self, text: str) -> str:
        """Classify a single page as text or figure based on OCR output characteristics.

        Text pages have dense paragraphs (many words, long lines, high line count).
        Figure pages have sparse labels, short lines, and low word count.
        Chemical structure pages fall in between — they have some words (labels,
        compound names) but lack the paragraph density of text pages.
        """
        lines = [l for l in text.strip().split('\n') if l.strip()]
        words = text.split()
        word_count = len(words)
        line_count = len(lines)

        if word_count < 15:
            return 'figure'

        # Average words per line — text pages have ~8-15 words/line (paragraphs),
        # figure pages have ~2-4 words/line (labels, short captions)
        avg_words_per_line = word_count / max(1, line_count)

        # Long lines (>40 chars) indicate paragraph text
        long_lines = sum(1 for l in lines if len(l.strip()) > 40)
        long_line_ratio = long_lines / max(1, line_count)

        if word_count > 100 and avg_words_per_line > 5 and long_line_ratio > 0.3:
            return 'text'

        if word_count < 50 or avg_words_per_line < 3:
            return 'figure'

        # Ambiguous — lean toward text (safer: Tesseract OCR handles it fine,
        # and figure analysis can still process these pages later)
        return 'text'

    def _ocr_text_pages_tesseract(self, pdf_path: Path, page_data: List[Dict],
                                   text_page_indices: List[int]) -> int:
        """OCR text pages from scanned PDF using local Tesseract (free, fast).

        Renders pages in batches at 200 DPI via PyMuPDF, runs pytesseract per page.
        Processes all pages — no budget constraint since it's free.
        """
        import pytesseract

        logger.info(f"  Tesseract OCR: processing {len(text_page_indices)} pages "
                   f"(batch size {self.TESSERACT_BATCH_SIZE}, {self.TESSERACT_DPI} DPI)...")

        ocr_count = 0
        total_batches = -(-len(text_page_indices) // self.TESSERACT_BATCH_SIZE)

        for batch_num, batch_start in enumerate(range(0, len(text_page_indices), self.TESSERACT_BATCH_SIZE)):
            batch_indices = text_page_indices[batch_start:batch_start + self.TESSERACT_BATCH_SIZE]
            images = self.render_pages_to_images(pdf_path, batch_indices, dpi=self.TESSERACT_DPI)

            for page_num in batch_indices:
                if page_num not in images:
                    continue
                try:
                    text = pytesseract.image_to_string(
                        images[page_num], config=self.TESSERACT_CONFIG)
                    text = self.clean_patent_text(text)
                    if text and len(text.strip()) > self.TESSERACT_MIN_CHARS:
                        page_data[page_num]['text'] = text
                        page_data[page_num]['text_length'] = len(text)
                        page_data[page_num]['classification'] = 'TEXT_HEAVY'
                        page_data[page_num]['ocr_method'] = 'tesseract'
                        ocr_count += 1
                except Exception as e:
                    logger.debug(f"  Tesseract failed on page {page_num + 1}: {e}")
                del images[page_num]

            if (batch_num + 1) % 5 == 0 or batch_num == total_batches - 1:
                logger.info(f"  Tesseract progress: batch {batch_num + 1}/{total_batches}, "
                           f"{ocr_count} pages extracted")

        logger.info(f"  Tesseract OCR complete: {ocr_count}/{len(text_page_indices)} pages recovered")
        return ocr_count

    def _find_claims_pages(self, page_data: List[Dict]) -> List[int]:
        """Identify which pages contain the claims section from OCR'd text."""
        claims_pattern = re.compile(self.PATENT_SECTION_PATTERNS['claims'])
        claim_number_pattern = re.compile(r'^\s*(\d+)\.\s+(?:A|An|The|Wherein|Said)\s', re.MULTILINE)

        claims_pages = []
        for i, page in enumerate(page_data):
            if page['classification'] != 'TEXT_HEAVY' or not page.get('text'):
                continue
            text = page['text'][:2000]
            if claims_pattern.search(text) or len(claim_number_pattern.findall(text)) >= 3:
                claims_pages.append(i)

        return claims_pages

    def _ocr_text_pages(self, pdf_path: Path, page_data: List[Dict],
                        text_page_indices: List[int]) -> int:
        """OCR text pages from image-only PDF using Vision AI.

        Reuses the enhance_text_pages_vision pattern but without garbled-text RAG context.
        """
        max_ocr = min(len(text_page_indices), self.MAX_VISION_TEXT_PAGES * 3)
        if len(text_page_indices) > max_ocr:
            logger.info(f"  Capping OCR from {len(text_page_indices)} to {max_ocr} pages")
            text_page_indices = text_page_indices[:max_ocr]

        budget_for_ocr = self._budget_remaining() - 20
        max_by_budget = budget_for_ocr * self.PAGES_PER_VISION_TEXT_BATCH
        if max_by_budget < len(text_page_indices):
            logger.info(f"  Budget limits OCR to {max_by_budget} pages")
            text_page_indices = text_page_indices[:max(10, max_by_budget)]

        logger.info(f"  OCR: extracting text from {len(text_page_indices)} pages via Vision AI...")

        client = self.get_anthropic_client()
        if not client:
            return 0

        batches = []
        for i in range(0, len(text_page_indices), self.PAGES_PER_VISION_TEXT_BATCH):
            batch_indices = text_page_indices[i:i + self.PAGES_PER_VISION_TEXT_BATCH]
            batches.append(batch_indices)

        ocr_count = 0
        with ThreadPoolExecutor(max_workers=self.MAX_CONCURRENT_BATCHES) as executor:
            futures = {}
            for batch_idx, batch_indices in enumerate(batches):
                future = executor.submit(self._ocr_batch, pdf_path, batch_indices, batch_idx)
                futures[future] = batch_idx

            for future in as_completed(futures):
                batch_idx = futures[future]
                try:
                    result = future.result()
                    for page_num, text in result.items():
                        if text and len(text.strip()) > 50:
                            page_data[page_num]['text'] = text
                            page_data[page_num]['text_length'] = len(text)
                            page_data[page_num]['classification'] = 'TEXT_HEAVY'
                            page_data[page_num]['ocr_extracted'] = True
                            ocr_count += 1
                    if result:
                        logger.info(f"  OCR batch {batch_idx+1}/{len(batches)}: "
                                   f"{len(result)} pages extracted")
                except Exception as e:
                    logger.error(f"  OCR batch {batch_idx+1} failed: {e}")

        logger.info(f"  OCR complete: {ocr_count}/{len(text_page_indices)} pages recovered text")
        return ocr_count

    def _ocr_batch(self, pdf_path: Path, page_indices: List[int], batch_idx: int) -> Dict[int, str]:
        """OCR a batch of pages using Vision AI (no garbled-text context)."""
        if not self._track_api_call():
            return {}

        client = self.get_anthropic_client()
        if not client:
            return {}

        images = self.render_pages_to_images(pdf_path, page_indices,
                                             dpi=self.VISION_DPI_TEXT_ENHANCEMENT)
        content = []
        valid_pages = []
        for page_num in page_indices:
            if page_num not in images:
                continue
            img_base64, media_type = self.encode_image_to_base64(
                images[page_num], max_dimension=self.MAX_IMAGE_DIMENSION_TEXT)
            del images[page_num]
            if not img_base64:
                continue
            content.append({
                "type": "image",
                "source": {"type": "base64", "media_type": media_type, "data": img_base64}
            })
            valid_pages.append(page_num)

        if not content:
            return {}

        pages_str = ', '.join(str(p + 1) for p in valid_pages)
        prompt = f"""Extract ALL text verbatim from these patent document pages (pages {pages_str}).

These are scanned/image-based patent pages. Rules:
1. Extract ALL visible text exactly as printed — every word, paragraph number, claim number
2. Preserve paragraph structure, section headers ([0001], [0002]...), numbered claims, and lists
3. Preserve ALL scientific terminology, chemical names, compound numbers, formulas exactly as written
4. If a page contains chemical structure DIAGRAMS (graphical molecular drawings), write ONLY: [CHEMICAL STRUCTURE: brief description]
5. Do NOT add ANY commentary, questions, summaries, or text not visible on the page
6. Do NOT say "Would you like..." or "I can help..." — output ONLY the extracted text
7. Extract the COMPLETE text — do not truncate, summarize, or skip any content

Output the text for each page, separated by:
===PAGE N===
(where N is the page number)

Start with ===PAGE {valid_pages[0] + 1}==="""

        content.append({"type": "text", "text": prompt})

        try:
            response = client.messages.create(
                model=self.VISION_MODEL,
                max_tokens=4096 * len(valid_pages),
                messages=[{"role": "user", "content": content}]
            )
            return self._parse_enhanced_text_response(response.content[0].text, valid_pages)
        except Exception as e:
            logger.error(f"  OCR batch {batch_idx} API call failed: {e}")
            return {}

    # ─── Stage 1: Bibliographic Data Extraction ────────────────────────────

    def extract_bibliographic_data(self, cover_text: str, pdf_path: Path) -> Dict:
        logger.info("Stage 1: Extracting bibliographic data...")
        bib = {}

        # Patent number from filename (most reliable)
        # Handles: wo2023037268, wo-2023037268-a1, WO-2023-037268-A1, wo25045758
        stem = pdf_path.stem.lower()
        wo_file_match = re.match(r'wo[-_]?(\d{4})[-_]?(\d{3,})(?:[-_]([a-z]\d?))?$', stem)
        ep_file_match = re.match(r'ep[-_]?(\d{6,})(?:[-_]([a-z]\d?))?$', stem)
        us_file_match = re.match(r'us[-_]?(\d{6,})(?:[-_]([a-z]\d?))?$', stem)

        if wo_file_match:
            num = wo_file_match.group(1) + wo_file_match.group(2)
            kind = (wo_file_match.group(3) or 'a1').upper()
            bib['patent_number'] = f"WO{num}{kind}"
            bib['patent_number_formatted'] = f"WO {num[:4]}/{num[4:]} {kind}"
        elif ep_file_match:
            num = ep_file_match.group(1)
            kind = (ep_file_match.group(2) or '').upper()
            bib['patent_number'] = f"EP{num}{kind}"
            bib['patent_number_formatted'] = f"EP {num} {kind}".strip()
        elif us_file_match:
            num = us_file_match.group(1)
            kind = (us_file_match.group(2) or '').upper()
            bib['patent_number'] = f"US{num}{kind}"
            bib['patent_number_formatted'] = f"US {num} {kind}".strip()
        else:
            # Fallback: try cover text
            pub_match = re.search(r'WO\s*(\d{4})[/]?(\d+)\s*([A-Z]\d?)?', cover_text)
            if pub_match:
                bib['patent_number'] = f"WO{pub_match.group(1)}{pub_match.group(2)}{pub_match.group(3) or 'A1'}"
                bib['patent_number_formatted'] = f"WO {pub_match.group(1)}/{pub_match.group(2)} {pub_match.group(3) or 'A1'}"

        # PCT number (usually survives OCR)
        pct_match = re.search(r'PCT/([A-Z]{2}\d{4}/\d+)', cover_text)
        if pct_match:
            bib['pct_number'] = f"PCT/{pct_match.group(1)}"

        # Filing date
        filing_match = re.search(r'Filing\s*Date[:\s]*(\d{1,2}\s+\w+\s*\d{4})', cover_text, re.IGNORECASE)
        if filing_match:
            bib['filing_date'] = filing_match.group(1).strip()

        # Publication date
        pub_date_match = re.search(r'Publication\s*Date[:\s]*(\d{1,2}\s+\w+\s*\d{4})', cover_text, re.IGNORECASE)
        if pub_date_match:
            bib['publication_date'] = pub_date_match.group(1).strip()

        # IPC classification (survives OCR well)
        ipc_codes = re.findall(r'[A-H]\d{2}[A-Z]\d+/\d+', cover_text)
        if ipc_codes:
            bib['ipc_classification'] = list(set(ipc_codes))

        found_fields = [k for k in bib if bib[k]]
        logger.info(f"  From text/filename: {found_fields}")
        return bib

    def extract_bibliographic_vision(self, pdf_path: Path) -> Dict:
        """Use Vision AI to extract clean bibliographic data from cover page.
        WIPO/EPO cover pages have multi-column layouts that garble with pdfplumber."""
        logger.info("  Using Vision AI for cover page (applicants, inventors, title)...")
        client = self.get_anthropic_client()
        if not client:
            return {}
        if not self._track_api_call():
            return {}

        images = self.render_pages_to_images(pdf_path, [0])
        if 0 not in images:
            return {}

        img_base64, media_type = self.encode_image_to_base64(images[0])
        del images[0]
        if not img_base64:
            return {}

        prompt = """Extract the following from this patent cover page. Return ONLY a JSON object:
{
  "title": "the full patent title (usually in bold near the top after the bibliographic fields)",
  "applicants": ["list of applicant organization names only - no addresses"],
  "inventors": ["list of inventor full names only - no addresses"],
  "filing_date": "filing date",
  "publication_date": "publication date",
  "priority_date": "earliest priority date"
}

Important:
- For applicants: extract ONLY organization names (e.g. "MabLink Bioscience", "Universite Claude Bernard Lyon 1"), NOT addresses
- For inventors: extract ONLY personal names (e.g. "Benoit Joseph", "Guy Fornet"), NOT addresses
- Return ONLY valid JSON, no other text."""

        try:
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
            result = self.parse_json_response(response.content[0].text)
            if result:
                logger.info(f"  Vision AI extracted: {list(result.keys())}")
                return result
        except Exception as e:
            logger.error(f"  Vision AI bibliographic extraction failed: {e}")

        return {}

    # ─── Stage 2: Full Text Extraction & Section Splitting ─────────────────

    def extract_full_text(self, page_data: List[Dict]) -> str:
        logger.info("Stage 2: Extracting full text and splitting into sections...")
        text_pages = [p for p in page_data if p['classification'] in ('TEXT_HEAVY', 'COVER', 'SEARCH_REPORT')]
        raw_text = "\n\n".join(p['text'] for p in text_pages if p['text'])
        full_text = self.clean_patent_text(raw_text)
        logger.info(f"  Total extracted text: {len(full_text)} chars from {len(text_pages)} text pages")
        return full_text

    def split_into_sections(self, full_text: str) -> Dict[str, str]:
        sections = {}
        section_positions = []

        for section_name, pattern in self.PATENT_SECTION_PATTERNS.items():
            matches = list(re.finditer(pattern, full_text))
            if not matches:
                continue
            # For claims: use the match that's followed by numbered claims (1. A compound...)
            # to avoid matching "Claims" in search reports or cross-references
            if section_name == 'claims':
                best = self._find_real_claims_header(matches, full_text)
                if best:
                    section_positions.append({
                        'name': section_name,
                        'start': best.start(),
                        'header_end': best.end()
                    })
            else:
                section_positions.append({
                    'name': section_name,
                    'start': matches[0].start(),
                    'header_end': matches[0].end()
                })

        section_positions.sort(key=lambda x: x['start'])

        for i, pos in enumerate(section_positions):
            start = pos['header_end']
            end = section_positions[i + 1]['start'] if i + 1 < len(section_positions) else len(full_text)
            section_text = full_text[start:end].strip()
            if section_text:
                sections[pos['name']] = section_text

        found = list(sections.keys())
        logger.info(f"  Sections identified: {found}")
        return sections

    def _find_real_claims_header(self, matches: list, full_text: str):
        """Find the claims header that's followed by actual numbered claims.

        Patents have "Claims" mentioned in search reports, cross-references, and
        table of contents. The real claims section is followed by "1. A/An..." pattern.
        """
        claim_start_pattern = re.compile(
            r'\n\s*1\.\s+(?:A|An|The|What)\s', re.IGNORECASE)

        for match in matches:
            # Look at the 500 chars after this "Claims" header
            after = full_text[match.end():match.end() + 500]
            if claim_start_pattern.search(after):
                return match

        # Fallback: if no match has numbered claims after it, look for standalone
        # "Claims" or "CLAIMS" on its own line (likely a section header)
        standalone = re.compile(r'^\s*(?:Claims|CLAIMS)\s*$', re.MULTILINE)
        for match in matches:
            line_start = full_text.rfind('\n', 0, match.start()) + 1
            line_end = full_text.find('\n', match.end())
            if line_end == -1:
                line_end = len(full_text)
            line = full_text[line_start:line_end].strip()
            if standalone.match(line):
                return match

        # Last resort: use the last match (claims are typically near the end)
        return matches[-1] if matches else None

    # ─── Stage 3: Claims Parsing ──────────────────────────────────────────

    def parse_claims(self, claims_text: str) -> Dict:
        logger.info("Stage 3: Parsing claims...")
        if not claims_text:
            logger.warning("  No claims text provided")
            return {'total_claims': 0, 'independent_claims': [], 'dependent_claims': [], 'claims': []}

        claims = []
        claim_splits = re.split(r'\n(?=\d+\.\s)', claims_text)

        for chunk in claim_splits:
            chunk = chunk.strip()
            if not chunk:
                continue
            num_match = re.match(r'^(\d+)\.\s+(.*)', chunk, re.DOTALL)
            if num_match:
                claim_num = int(num_match.group(1))
                claim_text = num_match.group(2).strip()

                dep_match = self.CLAIM_DEPENDENCY_PATTERN.search(claim_text)
                is_independent = dep_match is None
                depends_on = int(dep_match.group(1)) if dep_match else None
                category = self._classify_claim(claim_text)

                claims.append({
                    'number': claim_num,
                    'text': claim_text,
                    'is_independent': is_independent,
                    'depends_on': depends_on,
                    'category': category
                })

        independent = [c for c in claims if c['is_independent']]
        dependent = [c for c in claims if not c['is_independent']]

        logger.info(f"  Total claims: {len(claims)} ({len(independent)} independent, {len(dependent)} dependent)")
        if claims:
            logger.info(f"  Claim categories: {set(c['category'] for c in claims)}")

        return {
            'total_claims': len(claims),
            'independent_claims': independent,
            'dependent_claims': dependent,
            'claims': claims,
            'claim_tree': self._build_claim_tree(claims)
        }

    def _classify_claim(self, claim_text: str) -> str:
        first_100 = claim_text[:100].lower()
        if re.match(r'(?:a|an|the)\s+(?:compound|composition|conjugate|antibody|polypeptide|nucleic acid|vector|cell|adc)\b', first_100):
            return 'composition'
        elif re.match(r'(?:a|an|the)\s+(?:method|process)\s', first_100):
            return 'method'
        elif re.match(r'(?:use\s+of\b|.{0,30}\bfor\s+use\s+)', first_100):
            return 'use'
        elif re.match(r'(?:a|an|the)\s+pharmaceutical', first_100):
            return 'pharmaceutical'
        return 'other'

    def _build_claim_tree(self, claims: List[Dict]) -> Dict:
        """Build claim tree — handles multi-level dependencies."""
        tree = {}
        claims_by_num = {c['number']: c for c in claims}

        # Find root claims (independent)
        for c in claims:
            if c['is_independent']:
                tree[c['number']] = {'claim': c, 'children': []}

        # Attach dependent claims to their parent (may be another dependent)
        for c in claims:
            if not c['is_independent'] and c['depends_on']:
                parent_num = c['depends_on']
                # Walk up to find root
                root_num = parent_num
                visited = set()
                while root_num not in tree and root_num in claims_by_num:
                    if root_num in visited:
                        break
                    visited.add(root_num)
                    parent_claim = claims_by_num[root_num]
                    root_num = parent_claim.get('depends_on', root_num)
                    if parent_claim['is_independent']:
                        root_num = parent_claim['number']
                        break

                if root_num in tree:
                    tree[root_num]['children'].append(c)

        return tree

    # ─── Stage 4: Selective Vision AI Analysis ─────────────────────────────

    def analyze_figure_pages(self, pdf_path: Path, page_data: List[Dict],
                            sections: Dict[str, str],
                            max_pages: Optional[int] = None) -> List[Dict]:
        figure_pages = [p for p in page_data if p['classification'] == 'FIGURE_PAGE']

        effective_max = max_pages if max_pages is not None else self.DEFAULT_MAX_FIGURE_PAGES
        if effective_max > 0 and len(figure_pages) > effective_max:
            logger.warning(f"  {len(figure_pages)} figure pages exceed cap of {effective_max} — "
                          f"truncating (use --max-figure-pages 0 for unlimited)")
            figure_pages = figure_pages[:effective_max]

        if not figure_pages:
            logger.info("  No figure pages to analyze")
            return []

        logger.info(f"Stage 4: Analyzing {len(figure_pages)} figure pages with Vision AI...")

        # Determine which pages are likely chemical structures vs assay figures
        # by checking where they fall relative to section boundaries
        drawings_text = sections.get('drawings', '')
        compound_names = self._extract_compound_names_from_text(sections)

        batches = []
        batch_fig_indices = []
        for i in range(0, len(figure_pages), self.PAGES_PER_VISION_BATCH):
            batch = figure_pages[i:i + self.PAGES_PER_VISION_BATCH]
            batches.append(batch)
            batch_fig_indices.append(list(range(i, min(i + self.PAGES_PER_VISION_BATCH, len(figure_pages)))))

        logger.info(f"  Created {len(batches)} batches of up to {self.PAGES_PER_VISION_BATCH} pages")

        all_figures = []
        with ThreadPoolExecutor(max_workers=self.MAX_CONCURRENT_BATCHES) as executor:
            futures = {}
            for batch_idx, batch in enumerate(batches):
                future = executor.submit(
                    self._analyze_figure_batch, pdf_path, batch, batch_idx,
                    compound_names, drawings_text, batch_fig_indices[batch_idx]
                )
                futures[future] = batch_idx

            for future in as_completed(futures):
                batch_idx = futures[future]
                try:
                    result = future.result()
                    if result:
                        all_figures.extend(result)
                    logger.info(f"  Batch {batch_idx + 1}/{len(batches)} complete")
                except Exception as e:
                    logger.error(f"  Batch {batch_idx + 1} failed: {e}")

        # Sort by page number for consistent output
        all_figures.sort(key=lambda f: f.get('page', 0))
        logger.info(f"  Total figures described: {len(all_figures)}")
        return all_figures

    def _extract_compound_names_from_text(self, sections: Dict[str, str]) -> List[str]:
        """Extract key compound identifiers from the patent text for context in figure prompts."""
        examples = sections.get('examples', '') + sections.get('detailed_description', '')
        compounds = set()
        for m in re.finditer(r'(?:compound|Compound)\s+(\d{3,4}[a-z]?)', examples):
            compounds.add(m.group(1))
        for m in re.finditer(r'(?:ADC|adc)[\s-]+(\w+[\s-]\d+)', examples):
            compounds.add(m.group(1))
        return sorted(compounds)[:20]

    def _analyze_figure_batch(self, pdf_path: Path, batch: List[Dict], batch_idx: int,
                             compound_names: List[str], drawings_text: str,
                             figure_page_indices: List[int] = None) -> List[Dict]:
        if not self._track_api_call():
            return []
        client = self.get_anthropic_client()
        if not client:
            return []

        page_numbers = [p['page_num'] for p in batch]
        images = self.render_pages_to_images(pdf_path, page_numbers)

        content = []
        rendered_pages = []
        for page_num in page_numbers:
            if page_num not in images:
                continue
            img_base64, media_type = self.encode_image_to_base64(images[page_num])
            del images[page_num]
            if not img_base64:
                continue
            content.append({
                "type": "image",
                "source": {"type": "base64", "media_type": media_type, "data": img_base64}
            })
            rendered_pages.append(page_num + 1)

        if not content:
            return []

        prompt = self._build_figure_prompt(rendered_pages, compound_names, drawings_text,
                                           figure_page_indices)
        content.append({"type": "text", "text": prompt})

        try:
            response = client.messages.create(
                model=self.VISION_MODEL,
                max_tokens=4096,
                messages=[{"role": "user", "content": content}]
            )
            return self._parse_figure_response(response.content[0].text, rendered_pages)
        except Exception as e:
            logger.error(f"  Vision API call failed for batch {batch_idx}: {e}")
            return []

    def _build_figure_prompt(self, page_numbers: List[int], compound_names: List[str],
                            drawings_text: str, figure_page_indices: List[int] = None) -> str:
        """Single unified prompt for all figure types — let the model classify."""
        compound_context = ""
        if compound_names:
            compound_context = f"\nKnown compounds in this patent: {', '.join(compound_names[:15])}\n"

        legend_context = ""
        if drawings_text:
            fig_nums_to_search = set()
            if figure_page_indices:
                # Figure numbers are sequential (1, 2, 3...) across all figure pages.
                # Estimate which figures are on these pages by their position in the
                # figure page sequence.
                n_fig_pages = len(figure_page_indices)
                for batch_pos, idx in enumerate(figure_page_indices):
                    # Estimate ~1-2 figures per page on average
                    fig_start = max(1, batch_pos + 1)
                    fig_nums_to_search.add(fig_start)
                    fig_nums_to_search.add(fig_start + 1)
                # Also always search for the first 10 figures (most important)
                for i in range(1, min(11, n_fig_pages + 5)):
                    fig_nums_to_search.add(i)
            else:
                for i in range(1, 60):
                    fig_nums_to_search.add(i)

            for fig_num in sorted(fig_nums_to_search):
                pattern = rf'FIG[S.]?\s*{fig_num}\b[^\n]*(?:\n(?![A-Z]{{2,}})[^\n]*)?'
                match = re.search(pattern, drawings_text, re.IGNORECASE)
                if match:
                    legend_context += f"  - {match.group(0).strip()[:400]}\n"
            if legend_context:
                legend_context = f"\nFigure legends from patent text:\n{legend_context}\n"

        return f"""Analyze ALL figures, charts, and chemical structures on these patent pages (pages {page_numbers}).
{compound_context}{legend_context}
For EACH distinct figure or structure, provide a JSON object with:
- "page": page number
- "label": figure/structure label (e.g. "FIG. 1", "Compound 2384", "Formula (II)")
- "type": one of: dose_response_curve, bar_chart, kaplan_meier, tumor_growth, kinase_map, western_blot, chemical_structure, reaction_scheme, generic_structure, other
- "title": title or cell line name if visible
- "description": detailed description (for structures: backbone, functional groups, key substituents)
- "x_axis": x-axis label with units (null if not a chart)
- "y_axis": y-axis label with units (null if not a chart)
- "data_series": list of treatment groups/compounds shown (null if not applicable)
- "key_findings": key data points, IC50 values, trends, statistical annotations
- "smiles": SMILES notation attempt for chemical structures (null otherwise). For complex structures, provide the core scaffold SMILES.
- "smiles_confidence": "high" (simple known scaffold), "medium" (moderate complexity), "low" (complex/uncertain), or null

Return ONLY a valid JSON array. No other text."""

    def _parse_figure_response(self, response_text: str, page_numbers: List[int]) -> List[Dict]:
        result = self.parse_json_response(response_text)
        if isinstance(result, list):
            return result
        # If parsing fails, return a minimal entry
        return [{'page': page_numbers[0], 'label': 'Unknown', 'type': 'other',
                 'description': response_text[:500], 'key_findings': ''}]

    # ─── Stage 5: Key Data Extraction ──────────────────────────────────────

    def extract_key_data(self, sections: Dict[str, str]) -> Dict:
        logger.info("Stage 5: Extracting key compound and biological data...")
        client = self.get_anthropic_client()
        if not client:
            return {}
        if not self._track_api_call():
            return {}

        # Get synthesis data from examples section
        examples_text = sections.get('examples', '')
        if not examples_text or len(examples_text) < 200:
            examples_text = sections.get('detailed_description', '')
        if not examples_text or len(examples_text) < 200:
            logger.info("  Insufficient text for data extraction")
            return {}

        # Skip definitions/preamble, find actual experimental data
        example_start = re.search(r'(?:Example\s+\d|EXAMPLE\s+\d|Synthesis\s+of|1\.\s+Synthesis)', examples_text)
        if example_start and example_start.start() > 200:
            examples_text = examples_text[example_start.start():]

        # Take first 12K chars for synthesis + last 8K chars for biological results
        synthesis_text = examples_text[:12000]
        bio_text = examples_text[-8000:] if len(examples_text) > 12000 else ""

        combined_text = synthesis_text
        if bio_text and bio_text != synthesis_text:
            combined_text += "\n\n[...BIOLOGICAL EXAMPLES SECTION...]\n\n" + bio_text

        prompt = f"""From this patent experimental section, extract key scientific data.

TEXT:
{combined_text}

Extract and return as JSON:
{{
  "key_compounds": [
    {{
      "name": "compound name/number (e.g. 'Compound 2384' or 'DL-VA-2384')",
      "role": "payload|linker|ADC|drug-linker|intermediate",
      "yield": "percent yield if mentioned",
      "ms_data": "MS m/z value if available (e.g. '[M+H]+ = 493.2')",
      "purity": "HPLC purity if available"
    }}
  ],
  "biological_results": [
    {{
      "assay_type": "in_vitro_cytotoxicity|in_vivo_efficacy|tolerability|bystander|mechanistic|other",
      "compounds_tested": ["compound names/numbers"],
      "cell_lines_or_models": ["cell line names or animal models"],
      "key_result": "IC50 values, TGI%, tumor regression, or key metric",
      "comparison": "comparison to control/reference compound if any"
    }}
  ],
  "key_findings_summary": "2-3 sentence summary of the most important experimental results (potency, selectivity, in vivo efficacy)"
}}

Focus on FINAL compounds and ADCs (not intermediates unless they are the payload).
For biological results, include IC50 values and comparisons between compounds.
Max 10 compounds, max 10 biological results. Return ONLY valid JSON."""

        try:
            response = client.messages.create(
                model=self.VISION_MODEL,
                max_tokens=4096,
                messages=[{"role": "user", "content": prompt}]
            )
            result = self.parse_json_response(response.content[0].text)
            if result:
                compounds = result.get('key_compounds', [])
                bio_results = result.get('biological_results', [])
                logger.info(f"  Extracted {len(compounds)} key compounds, {len(bio_results)} biological results")
                return result
        except Exception as e:
            logger.error(f"  Data extraction API call failed: {e}")

        return {}

    # ─── Stage 5.5: Chemical Structure SMILES Refinement ────────────────────

    def validate_smiles(self, smiles: str) -> bool:
        """Basic SMILES format validation (no chemistry library needed)."""
        if not smiles or len(smiles) < 3:
            return False
        if not self.SMILES_VALID_CHARS.match(smiles):
            return False
        if smiles.count('(') != smiles.count(')'):
            return False
        if smiles.count('[') != smiles.count(']'):
            return False
        return True

    def refine_chemical_structures(self, pdf_path: Path, figures: List[Dict],
                                    sections: Dict[str, str]) -> List[Dict]:
        """Re-extract SMILES for chemical structures with medium/null confidence.

        Batched by page and parallelized. Skips structures already at 'low' confidence
        (refinement rarely upgrades these). Capped at MAX_SMILES_REFINEMENT structures.
        """
        chem_figs = [f for f in figures
                     if f.get('type') == 'chemical_structure'
                     and f.get('smiles_confidence') in ('medium', None)
                     and f.get('page')]

        if not chem_figs:
            logger.info("Stage 5.5: No chemical structures need SMILES refinement")
            return figures

        if len(chem_figs) > self.MAX_SMILES_REFINEMENT:
            logger.info(f"Stage 5.5: Capping SMILES refinement from {len(chem_figs)} "
                       f"to {self.MAX_SMILES_REFINEMENT} structures")
            chem_figs = chem_figs[:self.MAX_SMILES_REFINEMENT]

        logger.info(f"Stage 5.5: Refining SMILES for {len(chem_figs)} structures (batched by page)...")

        client = self.get_anthropic_client()
        if not client:
            return figures

        drawings_text = sections.get('drawings', '')
        claims_text = sections.get('claims', '')[:4000]
        scaffold_context = ""
        scaffold_match = re.search(r'formula\s*\(I+\)[^.]*\.', claims_text, re.IGNORECASE)
        if scaffold_match:
            scaffold_context = f"\nCore scaffold from claims: {scaffold_match.group(0)[:500]}\n"

        # Group structures by page for batched API calls
        from collections import defaultdict
        page_groups = defaultdict(list)
        for fig in chem_figs:
            page_groups[fig['page'] - 1].append(fig)

        pages_needed = sorted(page_groups.keys())
        images = self.render_pages_to_images(pdf_path, pages_needed,
                                              dpi=self.CHEMICAL_STRUCTURE_DPI)

        logger.info(f"  {len(chem_figs)} structures across {len(page_groups)} pages "
                   f"→ {len(page_groups)} API calls (was {len(chem_figs)} sequential)")

        refined_count = 0
        with ThreadPoolExecutor(max_workers=self.MAX_CONCURRENT_BATCHES) as executor:
            futures = {}
            for page_0, figs_on_page in page_groups.items():
                if page_0 not in images:
                    continue
                future = executor.submit(
                    self._refine_smiles_batch, images[page_0], page_0,
                    figs_on_page, scaffold_context, drawings_text, sections
                )
                futures[future] = page_0

            for future in as_completed(futures):
                page_0 = futures[future]
                try:
                    count = future.result()
                    refined_count += count
                except Exception as e:
                    logger.error(f"  SMILES batch for page {page_0+1} failed: {e}")

        for page_num in list(images.keys()):
            del images[page_num]

        logger.info(f"  SMILES refinement complete: {refined_count}/{len(chem_figs)} improved")

        for fig in figures:
            if fig.get('smiles') and not fig.get('smiles_validated'):
                fig['smiles_validated'] = self.validate_smiles(fig['smiles'])

        return figures

    def _refine_smiles_batch(self, page_image, page_0: int, figs_on_page: List[Dict],
                             scaffold_context: str, drawings_text: str,
                             sections: Dict[str, str]) -> int:
        """Refine SMILES for all structures on a single page in one API call."""
        if not self._track_api_call():
            return 0

        client = self.get_anthropic_client()
        if not client:
            return 0

        img_base64, media_type = self.encode_image_to_base64(
            page_image, max_dimension=self.MAX_IMAGE_DIMENSION_CHEMICAL)
        if not img_base64:
            return 0

        # Build context for all structures on this page
        structures_desc = []
        for fig in figs_on_page:
            label = fig.get('label', 'Unknown')
            description = fig.get('description', '')[:200]
            structures_desc.append(f'- "{label}": {description}')
        structures_list = '\n'.join(structures_desc)

        legend = ""
        fig_match = re.search(rf'FIG[S.]?\s*{page_0 + 1}\b[^\n]*',
                              drawings_text, re.IGNORECASE)
        if fig_match:
            legend = f"\nFigure legend: {fig_match.group(0).strip()[:300]}\n"

        prompt = f"""Analyze the chemical structures on page {page_0 + 1}.

Structures to identify:
{structures_list}
{scaffold_context}{legend}
For EACH structure listed above, provide a JSON array with one object per structure:
[{{
  "label": "structure label exactly as listed above",
  "smiles": "Canonical SMILES with stereochemistry where visible",
  "smiles_confidence": "high"|"medium"|"low",
  "molecular_formula": "e.g. C23H28BrN7O or null",
  "molecular_weight": "approximate MW or null"
}}]

Return ONLY the JSON array. No other text."""

        content = [
            {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": img_base64}},
            {"type": "text", "text": prompt}
        ]

        try:
            response = client.messages.create(
                model=self.VISION_MODEL,
                max_tokens=1024 * max(1, len(figs_on_page)),
                messages=[{"role": "user", "content": content}]
            )
            results = self.parse_json_response(response.content[0].text)
            if not isinstance(results, list):
                results = [results] if results else []

            refined = 0
            results_by_label = {r.get('label', ''): r for r in results if r}

            for fig in figs_on_page:
                label = fig.get('label', 'Unknown')
                result = results_by_label.get(label)
                if not result:
                    for r in results:
                        if r and label.lower() in r.get('label', '').lower():
                            result = r
                            break
                if not result or not result.get('smiles'):
                    continue
                new_smiles = result['smiles']
                if self.validate_smiles(new_smiles):
                    old_conf = fig.get('smiles_confidence', 'none')
                    new_conf = result.get('smiles_confidence', 'medium')
                    fig['smiles'] = new_smiles
                    fig['smiles_confidence'] = new_conf
                    if result.get('molecular_formula'):
                        fig['molecular_formula'] = result['molecular_formula']
                    if result.get('molecular_weight'):
                        fig['molecular_weight'] = result['molecular_weight']
                    refined += 1
                    logger.info(f"    {label}: SMILES refined ({old_conf} -> {new_conf})")
                else:
                    logger.warning(f"    {label}: Refined SMILES failed validation")

            return refined
        except Exception as e:
            logger.error(f"  SMILES batch page {page_0+1} failed: {e}")
            return 0

    # ─── Stage 6: Executive Summary & Protection Scope ─────────────────────

    def generate_summaries(self, sections: Dict[str, str], claims_data: Dict, bib: Dict) -> Dict[str, str]:
        logger.info("Stage 6: Generating executive summary and protection scope...")
        client = self.get_anthropic_client()
        if not client:
            return {}
        if not self._track_api_call():
            return {}

        summary_text = sections.get('summary', '')[:3000]
        technical_field = sections.get('technical_field', '')[:500]

        independent_claims_text = ""
        for claim in claims_data.get('independent_claims', [])[:5]:
            independent_claims_text += f"\nClaim {claim['number']} ({claim['category']}): {claim['text'][:300]}...\n"

        title = bib.get('title', 'Unknown')
        patent_num = bib.get('patent_number', 'Unknown')

        prompt = f"""Based on this patent information, provide the following:

PATENT: {patent_num} - {title}

TECHNICAL FIELD:
{technical_field}

SUMMARY OF INVENTION:
{summary_text[:2000]}

INDEPENDENT CLAIMS:
{independent_claims_text}

Please provide:
1. EXECUTIVE SUMMARY: A 3-5 sentence summary of what this patent discloses. Cover: the problem being solved, the key innovation, the types of compounds/compositions covered, and the therapeutic application.

2. PROTECTION SCOPE: A 3-5 sentence plain-language summary of what this patent protects/claims. Describe in practical terms what a competitor cannot do without infringing. Be specific about the chemical/biological space covered.

3. CLASSIFICATION: Extract these structured fields as a JSON object:
- "therapeutic_area": primary disease area (e.g. "oncology", "immunology", "neurology")
- "mechanism_of_action": the drug's mechanism (e.g. "molecular glue degrader", "kinase inhibitor")
- "target_protein": specific protein target if mentioned (e.g. "Cyclin K", "EGFR")
- "target_class": broader target class (e.g. "E3 ubiquitin ligase", "receptor tyrosine kinase")
- "drug_modality": type of therapeutic (e.g. "ADC", "small molecule", "bispecific antibody")
- "scaffold": core chemical scaffold name (e.g. "pyrazolo[1,5-a][1,3,5]triazine")
- "comparators": list of mentioned competitor drugs/compounds (e.g. ["Kadcyla", "Enhertu"])

Format your response EXACTLY as:
---EXECUTIVE_SUMMARY---
[your executive summary here]
---PROTECTION_SCOPE---
[your protection scope summary here]
---CLASSIFICATION---
[JSON object with the structured fields]"""

        try:
            response = client.messages.create(
                model=self.VISION_MODEL,
                max_tokens=2048,
                messages=[{"role": "user", "content": prompt}]
            )
            text = response.content[0].text
            result = {}

            exec_match = re.search(r'---EXECUTIVE_SUMMARY---\s*(.*?)(?:---PROTECTION_SCOPE---|$)', text, re.DOTALL)
            if exec_match:
                result['executive_summary'] = exec_match.group(1).strip()

            scope_match = re.search(r'---PROTECTION_SCOPE---\s*(.*?)(?:---CLASSIFICATION---|$)', text, re.DOTALL)
            if scope_match:
                result['protection_scope'] = scope_match.group(1).strip()

            class_match = re.search(r'---CLASSIFICATION---\s*(.*?)$', text, re.DOTALL)
            if class_match:
                classification = self.parse_json_response(class_match.group(1).strip())
                if classification:
                    result['classification'] = classification

            if not result.get('executive_summary'):
                result['executive_summary'] = text.strip()

            logger.info(f"  Executive summary: {len(result.get('executive_summary', ''))} chars")
            logger.info(f"  Protection scope: {len(result.get('protection_scope', ''))} chars")
            if result.get('classification'):
                logger.info(f"  Classification: {result['classification'].get('therapeutic_area', 'N/A')} / {result['classification'].get('mechanism_of_action', 'N/A')}")
            return result

        except Exception as e:
            logger.error(f"  Summary generation failed: {e}")
            return {}

    # ─── Quality Scoring ───────────────────────────────────────────────────

    def calculate_quality_score(self, bib: Dict, claims_data: Dict, sections: Dict,
                                figures: List[Dict], key_data: Dict) -> Dict:
        scores = {}

        # Bibliographic: patent_number, title, applicants, inventors, filing_date, publication_date, ipc
        bib_fields = ['patent_number', 'title', 'applicants', 'inventors',
                      'filing_date', 'publication_date', 'ipc_classification']
        bib_present = sum(1 for f in bib_fields if bib.get(f))
        scores['bibliographic'] = round(bib_present / len(bib_fields) * 10, 1)

        # Claims: found + independent parsed + tree built
        claims_score = 0
        if claims_data.get('total_claims', 0) > 0:
            claims_score += 4
            if claims_data.get('independent_claims'):
                claims_score += 3
            if claims_data.get('claim_tree'):
                claims_score += 3
        scores['claims'] = min(10, claims_score)

        # Sections
        expected_sections = ['technical_field', 'background', 'summary', 'detailed_description', 'examples']
        sections_found = sum(1 for s in expected_sections if sections.get(s))
        scores['sections'] = round(sections_found / len(expected_sections) * 10, 1)

        # Figures
        figure_score = 0
        if figures:
            figure_score = min(7, len(figures) * 0.3)
            has_quality = sum(1 for f in figures if f.get('key_findings') or f.get('description'))
            if has_quality > 0:
                figure_score += min(3, has_quality * 0.3)
        scores['figures'] = min(10, round(figure_score, 1))

        # Chemical data (quantity + quality)
        chem_score = 0
        compounds = key_data.get('key_compounds', [])
        bio_results = key_data.get('biological_results', [])
        if compounds:
            chem_score += min(4, len(compounds) * 0.6)
        if bio_results:
            chem_score += min(4, len(bio_results) * 0.8)
        # SMILES confidence bonus
        chem_figs = [f for f in figures if f.get('type') == 'chemical_structure']
        if chem_figs:
            high_conf = sum(1 for f in chem_figs if f.get('smiles_confidence') == 'high')
            smiles_ratio = high_conf / len(chem_figs)
            chem_score += 2 * smiles_ratio
        elif compounds:
            chem_score += 1
        scores['chemical_data'] = min(10, round(chem_score, 1))

        weights = {
            'bibliographic': 0.15,
            'claims': 0.30,
            'sections': 0.25,
            'figures': 0.15,
            'chemical_data': 0.15
        }
        overall = sum(scores.get(k, 0) * w for k, w in weights.items())
        scores['overall'] = round(min(10, overall), 1)

        if overall >= self.QUALITY_EXCELLENT_THRESHOLD:
            scores['assessment'] = 'Excellent'
        elif overall >= self.QUALITY_GOOD_THRESHOLD:
            scores['assessment'] = 'Good'
        elif overall >= self.QUALITY_FAIR_THRESHOLD:
            scores['assessment'] = 'Fair'
        else:
            scores['assessment'] = 'Poor'

        return scores

    def log_quality(self, patent_id: str, quality_scores: Dict, saved: bool, filename: str = ""):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        status = "SAVED" if saved else "SKIPPED"
        with open(self.quality_log_path, 'a', encoding='utf-8') as f:
            f.write(f"{timestamp}\t{patent_id}\t{quality_scores['overall']}\t{quality_scores['assessment']}\t{status}\t{filename}\n")

    # ─── Stage 7: Markdown Generation ──────────────────────────────────────

    def generate_markdown(self, bib: Dict, sections: Dict, claims_data: Dict,
                         figures: List[Dict], key_data: Dict, summaries: Dict,
                         quality_scores: Dict, full_text_length: int,
                         total_pages: int, figure_pages_analyzed: int,
                         text_enhanced_count: int = 0, reclassified_count: int = 0,
                         is_scanned_pdf: bool = False, ocr_page_count: int = 0,
                         total_api_calls: int = 0,
                         ocr_engine_used: Optional[str] = None) -> str:
        md = []

        # YAML frontmatter
        md.append("---\n")
        md.append(f"patent_number: {bib.get('patent_number', 'unknown')}\n")
        if bib.get('title'):
            safe_title = bib['title'].replace('"', '\\"').replace('\n', ' ')
            md.append(f'title: "{safe_title}"\n')
        if bib.get('applicants'):
            md.append("applicants:\n")
            for a in bib['applicants'][:8]:
                safe_a = str(a).replace('"', '').replace('\n', ' ').strip()
                if safe_a:
                    md.append(f'  - "{safe_a}"\n')
        if bib.get('inventors'):
            md.append("inventors:\n")
            for inv in bib['inventors'][:15]:
                safe_inv = str(inv).replace('"', '').replace('\n', ' ').strip()
                if safe_inv:
                    md.append(f'  - "{safe_inv}"\n')
        if bib.get('filing_date'):
            md.append(f"filing_date: {bib['filing_date']}\n")
        if bib.get('publication_date'):
            md.append(f"publication_date: {bib['publication_date']}\n")
        if bib.get('priority_date'):
            md.append(f"priority_date: {bib['priority_date']}\n")
        elif bib.get('priority_data'):
            md.append(f"priority_date: {bib['priority_data'][0]['date']}\n")
        if bib.get('ipc_classification'):
            md.append(f"ipc_classification: {', '.join(bib['ipc_classification'])}\n")
        if bib.get('pct_number'):
            md.append(f"pct_number: {bib['pct_number']}\n")
        md.append(f"total_pages: {total_pages}\n")
        md.append(f"total_claims: {claims_data.get('total_claims', 0)}\n")
        md.append(f"independent_claims: {len(claims_data.get('independent_claims', []))}\n")
        if is_scanned_pdf:
            if ocr_engine_used and 'tesseract' in ocr_engine_used:
                method = "tesseract_hybrid" if ocr_engine_used == 'tesseract_hybrid' else "tesseract_ocr"
            else:
                method = "vision_ocr"
        elif text_enhanced_count > 0:
            method = "native_vision_enhanced"
        else:
            method = "native_vision_selective"
        md.append(f"extraction_method: {method}\n")
        md.append(f"is_scanned_pdf: {str(is_scanned_pdf).lower()}\n")
        if ocr_engine_used:
            md.append(f"ocr_engine: {ocr_engine_used}\n")
        md.append(f"vision_model: {self.VISION_MODEL}\n")
        md.append(f"processing_date: {datetime.now().strftime('%Y-%m-%d')}\n")
        md.append(f"figure_pages_analyzed: {figure_pages_analyzed}\n")
        if ocr_page_count > 0:
            md.append(f"ocr_pages_recovered: {ocr_page_count}\n")
        if text_enhanced_count > 0:
            md.append(f"text_pages_enhanced: {text_enhanced_count}\n")
        if reclassified_count > 0:
            md.append(f"figure_pages_reclassified: {reclassified_count}\n")
        if total_api_calls > 0:
            md.append(f"total_api_calls: {total_api_calls}\n")
        md.append(f"quality_overall: {quality_scores['overall']}/10\n")
        md.append(f"quality_assessment: {quality_scores['assessment']}\n")
        # Semantic classification fields from Stage 6
        classification = summaries.get('classification', {})
        if classification:
            if classification.get('therapeutic_area'):
                md.append(f"therapeutic_area: \"{classification['therapeutic_area']}\"\n")
            if classification.get('mechanism_of_action'):
                md.append(f"mechanism_of_action: \"{classification['mechanism_of_action']}\"\n")
            if classification.get('target_protein'):
                md.append(f"target_protein: \"{classification['target_protein']}\"\n")
            if classification.get('target_class'):
                md.append(f"target_class: \"{classification['target_class']}\"\n")
            if classification.get('drug_modality'):
                md.append(f"drug_modality: \"{classification['drug_modality']}\"\n")
            if classification.get('scaffold'):
                md.append(f"scaffold: \"{classification['scaffold']}\"\n")
            if classification.get('comparators'):
                comps = classification['comparators']
                if isinstance(comps, list):
                    md.append("comparators:\n")
                    for c in comps[:10]:
                        md.append(f"  - \"{c}\"\n")
        # Compound statistics
        compounds = key_data.get('key_compounds', [])
        bio_results = key_data.get('biological_results', [])
        chem_figs = [f for f in figures if f.get('type') == 'chemical_structure']
        high_conf = sum(1 for f in chem_figs if f.get('smiles_confidence') == 'high')
        md.append(f"compound_count: {len(compounds)}\n")
        md.append(f"biological_assay_count: {len(bio_results)}\n")
        md.append(f"chemical_structures_count: {len(chem_figs)}\n")
        md.append(f"smiles_high_confidence: {high_conf}\n")
        md.append("---\n\n")

        # Title
        title = bib.get('title', 'Untitled Patent').replace('\n', ' ')
        md.append(f"# {title}\n\n")

        # Executive Summary
        if summaries.get('executive_summary'):
            md.append("## Executive Summary\n\n")
            md.append(f"{summaries['executive_summary']}\n\n---\n\n")

        # Bibliographic Data table
        md.append("## Bibliographic Data\n\n")
        md.append("| Field | Value |\n|-------|-------|\n")
        md.append(f"| Patent Number | {bib.get('patent_number_formatted', bib.get('patent_number', 'N/A'))} |\n")
        if bib.get('pct_number'):
            md.append(f"| PCT Number | {bib['pct_number']} |\n")
        if bib.get('filing_date'):
            md.append(f"| Filing Date | {bib['filing_date']} |\n")
        if bib.get('publication_date'):
            md.append(f"| Publication Date | {bib['publication_date']} |\n")
        if bib.get('priority_date'):
            md.append(f"| Priority Date | {bib['priority_date']} |\n")
        if bib.get('ipc_classification'):
            md.append(f"| IPC Classification | {', '.join(bib['ipc_classification'])} |\n")
        if bib.get('applicants'):
            md.append(f"| Applicants | {'; '.join(str(a) for a in bib['applicants'][:5])} |\n")
        if bib.get('inventors'):
            md.append(f"| Inventors | {'; '.join(str(i) for i in bib['inventors'][:10])} |\n")
        md.append("\n---\n\n")

        # Technical Field
        if sections.get('technical_field'):
            md.append("## Technical Field\n\n")
            md.append(f"{sections['technical_field'][:2000]}\n\n---\n\n")

        # Background
        if sections.get('background'):
            md.append("## Background\n\n")
            md.append(f"{sections['background'][:3000]}\n\n---\n\n")

        # Summary of Invention
        if sections.get('summary'):
            md.append("## Summary of Invention\n\n")
            md.append(f"{sections['summary'][:5000]}\n\n---\n\n")

        # Claims
        md.append("## Claims\n\n")
        if summaries.get('protection_scope'):
            md.append("### Protection Scope\n\n")
            md.append(f"{summaries['protection_scope']}\n\n")

        if claims_data.get('independent_claims'):
            md.append("### Independent Claims\n\n")
            for claim in claims_data['independent_claims']:
                md.append(f"**Claim {claim['number']}** ({claim['category']}): {claim['text'][:500]}")
                if len(claim['text']) > 500:
                    md.append("...")
                md.append("\n\n")

        if claims_data.get('claim_tree'):
            md.append("### Claim Dependency Tree\n\n")
            for root_num, node in sorted(claims_data['claim_tree'].items()):
                claim = node['claim']
                md.append(f"- **Claim {claim['number']}** ({claim['category']}): {claim['text'][:100]}...\n")
                for child in node['children'][:10]:
                    md.append(f"  - Claim {child['number']}: {child['text'][:80]}...\n")
            md.append("\n")

        md.append(f"\n*Total claims: {claims_data.get('total_claims', 0)}*\n\n---\n\n")

        # Key Chemical Structures (from figures) — structured table + detail
        chem_figures = [f for f in figures if f.get('type') in ('chemical_structure', 'reaction_scheme', 'generic_structure')]
        if chem_figures:
            md.append("## Key Chemical Structures\n\n")
            # Summary table for parseability
            md.append("| ID | Label | SMILES | Confidence | Formula | MW | Page |\n")
            md.append("|----|-------|--------|------------|---------|----:|------|\n")
            for fig in chem_figures[:20]:
                label = str(fig.get('label', 'Structure') or 'Structure')
                smiles_raw = fig.get('smiles', '')
                smiles = str(smiles_raw) if str(smiles_raw or '').lower() not in ('null', 'none', '', '[]') else '-'
                conf = str(fig.get('smiles_confidence', '-') or '-')
                formula = str(fig.get('molecular_formula', '-') or '-')
                mw = str(fig.get('molecular_weight', '-') or '-')
                page = fig.get('page', '-')
                md.append(f"| {label} | {label} | `{smiles}` | {conf} | {formula} | {mw} | {page} |\n")
            md.append("\n")
            # Detailed descriptions below the table
            for fig in chem_figures[:20]:
                label = str(fig.get('label', 'Structure') or 'Structure')
                md.append(f"### {label}\n")
                if fig.get('description'):
                    md.append(f"**Description**: {str(fig['description'])}\n\n")
                if fig.get('smiles') and str(fig['smiles']).lower() not in ('null', 'none', '', '[]'):
                    confidence = str(fig.get('smiles_confidence', 'unknown') or 'unknown')
                    md.append(f"**SMILES** (confidence: {confidence}): `{fig['smiles']}`\n\n")
                if fig.get('inchi'):
                    md.append(f"**InChI**: `{fig['inchi']}`\n\n")
                if fig.get('molecular_formula'):
                    md.append(f"**Formula**: {fig['molecular_formula']}")
                    if fig.get('molecular_weight'):
                        md.append(f" | **MW**: {fig['molecular_weight']}")
                    md.append("\n\n")
                if fig.get('key_findings'):
                    md.append(f"**Key Data**: {str(fig['key_findings'])}\n\n")
                md.append(f"*Page {fig.get('page', 'N/A')}*\n\n")
            md.append("---\n\n")

        # Key Compounds (synthesis data)
        if key_data.get('key_compounds'):
            md.append("## Key Compounds\n\n")
            md.append("| Compound | Role | Yield | MS [M+H]+ | Purity |\n")
            md.append("|----------|------|------:|-----------|--------|\n")
            for comp in key_data['key_compounds'][:15]:
                name = comp.get('name', 'N/A')
                role = comp.get('role', 'N/A')
                yield_val = comp.get('yield', '-') or '-'
                ms = comp.get('ms_data', '-') or '-'
                purity = comp.get('purity', '-') or '-'
                md.append(f"| {name} | {role} | {yield_val} | {ms} | {purity} |\n")
            md.append("\n---\n\n")

        # Biological Results — structured table
        if key_data.get('biological_results'):
            md.append("## Biological Results\n\n")
            md.append("| Assay Type | Compounds | Cell Lines/Models | Key Result | Comparison |\n")
            md.append("|-----------|-----------|-------------------|------------|------------|\n")
            for res in key_data['biological_results'][:10]:
                assay = str(res.get('assay_type', '-') or '-')
                compounds_val = res.get('compounds_tested', ['-'])
                compounds = ', '.join(compounds_val) if isinstance(compounds_val, list) else str(compounds_val or '-')
                models_val = res.get('cell_lines_or_models', res.get('cell_lines', ['-']))
                models = ', '.join(models_val) if isinstance(models_val, list) else str(models_val or '-')
                key_result = str(res.get('key_result', '-') or '-').replace('|', '/')
                comparison = str(res.get('comparison', '-') or '-').replace('|', '/')
                md.append(f"| {assay} | {compounds} | {models} | {key_result} | {comparison} |\n")
            md.append("\n---\n\n")

        # Figures & Drawings (non-chemical)
        non_chem_figures = [f for f in figures if f.get('type') not in ('chemical_structure', 'reaction_scheme', 'generic_structure')]
        if non_chem_figures:
            md.append("## Figures & Drawings\n\n")
            for fig in non_chem_figures[:40]:
                label = str(fig.get('label', 'Figure') or 'Figure')
                fig_type = str(fig.get('type', 'other') or 'other')
                md.append(f"### {label}\n")
                md.append(f"**Type**: {fig_type.replace('_', ' ').title()}\n\n")
                if fig.get('title'):
                    md.append(f"**Title**: {str(fig['title'])}\n\n")
                if fig.get('x_axis'):
                    md.append(f"**X-axis**: {fig['x_axis']} | **Y-axis**: {fig.get('y_axis', 'N/A')}\n\n")
                if fig.get('data_series'):
                    series = fig['data_series'] if isinstance(fig['data_series'], list) else [fig['data_series']]
                    md.append(f"**Data Series**: {', '.join(str(s) for s in series)}\n\n")
                if fig.get('key_findings'):
                    md.append(f"**Key Findings**: {str(fig['key_findings'])}\n\n")
                elif fig.get('description'):
                    md.append(f"**Description**: {str(fig['description'])}\n\n")
                md.append(f"*Page {fig.get('page', 'N/A')}*\n\n")
            md.append("---\n\n")

        # Figures Index (machine-parseable summary table)
        if figures:
            md.append("## Figures Index\n\n")
            md.append("| Figure | Type | Page | Title | Key Finding |\n")
            md.append("|--------|------|-----:|-------|-------------|\n")
            for fig in figures[:50]:
                label = str(fig.get('label', '-') or '-')
                fig_type = str(fig.get('type', 'other') or 'other').replace('_', ' ')
                page = fig.get('page', '-')
                title = str(fig.get('title', '-') or '-').replace('|', '/')[:60]
                finding = str(fig.get('key_findings', '-') or '-').replace('|', '/').replace('\n', ' ')[:80]
                md.append(f"| {label} | {fig_type} | {page} | {title} | {finding} |\n")
            md.append("\n---\n\n")

        # Key Findings Summary
        if key_data.get('key_findings_summary'):
            md.append("## Key Findings Summary\n\n")
            md.append(f"{key_data['key_findings_summary']}\n\n---\n\n")

        # Quality Assessment
        md.append("## Quality Assessment\n\n")
        md.append(f"**Overall Quality**: {quality_scores['overall']}/10 - {quality_scores['assessment']}\n\n")
        md.append("**Component Scores**:\n")
        md.append(f"- Bibliographic Data: {quality_scores.get('bibliographic', 0)}/10\n")
        md.append(f"- Claims Extraction: {quality_scores.get('claims', 0)}/10\n")
        md.append(f"- Section Completeness: {quality_scores.get('sections', 0)}/10\n")
        md.append(f"- Figure Analysis: {quality_scores.get('figures', 0)}/10\n")
        md.append(f"- Chemical Data: {quality_scores.get('chemical_data', 0)}/10\n")
        md.append("\n---\n\n")

        # Processing Metadata
        md.append("## Processing Metadata\n\n")
        md.append(f"- **Extraction Method**: {method}\n")
        if is_scanned_pdf:
            md.append(f"- **Document Type**: Scanned/image-only PDF (OCR applied)\n")
        md.append(f"- **Vision Model**: {self.VISION_MODEL}\n")
        md.append(f"- **Total Pages**: {total_pages}\n")
        md.append(f"- **Figure Pages Analyzed**: {figure_pages_analyzed}\n")
        if ocr_page_count > 0:
            md.append(f"- **OCR Pages Recovered**: {ocr_page_count}\n")
        if text_enhanced_count > 0:
            md.append(f"- **Text Pages Vision-Enhanced**: {text_enhanced_count}\n")
        if reclassified_count > 0:
            md.append(f"- **Text Pages Reclassified as Figure**: {reclassified_count}\n")
        md.append(f"- **Total Figures Described**: {len(figures)}\n")
        md.append(f"- **Text Length**: {full_text_length:,} characters\n")
        if total_api_calls > 0:
            md.append(f"- **Total API Calls**: {total_api_calls}\n")
        md.append(f"- **Processing Date**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")

        return "".join(md)

    # ─── Main Processing Orchestration ─────────────────────────────────────

    def process_single_patent(self, pdf_path: Path, no_vision: bool = False,
                             claims_only: bool = False, max_figure_pages: Optional[int] = None,
                             skip_existing: bool = False) -> bool:
        logger.info(f"\n{'='*70}")
        logger.info(f"PATENT PIPELINE: {pdf_path.name}")
        logger.info(f"{'='*70}")

        self._api_call_count = 0

        # Stage 0: Characterize PDF
        pdf_info = self.characterize_pdf(pdf_path)
        page_data = pdf_info['page_data']
        total_pages = pdf_info['total_pages']

        # Stage 1: Bibliographic data (filename + Vision AI for cover page)
        cover_text = page_data[0]['text'] if page_data else ""
        bib = self.extract_bibliographic_data(cover_text, pdf_path)

        # Always use Vision AI for cover page — pdfplumber garbles multi-column WIPO/EPO layouts
        if not no_vision:
            vision_bib = self.extract_bibliographic_vision(pdf_path)
            if vision_bib:
                for key, val in vision_bib.items():
                    if val and (not bib.get(key) or key in ('applicants', 'inventors', 'title')):
                        bib[key] = val

        patent_id = bib.get('patent_number', pdf_path.stem)

        if skip_existing and any(patent_id in f for f in self._existing_files):
            logger.info(f"Patent {patent_id} already processed - skipping")
            return True

        # Stage 0.5: Detect image-only PDF and OCR text pages
        is_scanned_pdf = self._detect_image_only_pdf(pdf_info)
        ocr_page_count = 0
        ocr_engine_used = None
        if is_scanned_pdf and not no_vision:
            logger.info("  DETECTED: Image-only/scanned PDF — initiating OCR workflow")

            if self._tesseract_available and self._ocr_engine != 'vision':
                # Tesseract path: free, local, processes ALL pages
                logger.info("  Using Tesseract OCR (local, free) for bulk text extraction")
                ocr_engine_used = 'tesseract'
                page_type_map = self._classify_pages_heuristic(pdf_path, page_data)
                text_page_indices = page_type_map.get('text_pages', [])
                if text_page_indices:
                    ocr_page_count = self._ocr_text_pages_tesseract(
                        pdf_path, page_data, text_page_indices)
                    # Check if claims were captured
                    claims_pages = self._find_claims_pages(page_data)
                    if claims_pages:
                        logger.info(f"  Claims found on pages: "
                                   f"{', '.join(str(p + 1) for p in claims_pages[:5])}")
                    else:
                        logger.info("  Claims not found in Tesseract output — "
                                   "will attempt Vision AI on estimated claims region")
                        # Try Vision AI on estimated claims pages (60-80% through text)
                        if self._budget_remaining() > 10 and text_page_indices:
                            start = int(len(text_page_indices) * 0.55)
                            end = int(len(text_page_indices) * 0.85)
                            claims_region = text_page_indices[start:end][:30]
                            vision_ocr = self._ocr_text_pages(
                                pdf_path, page_data, claims_region)
                            ocr_page_count += vision_ocr
                            ocr_engine_used = 'tesseract_hybrid'
            else:
                # Vision AI path: existing behavior (budget-limited)
                logger.info("  Using Vision AI OCR (Tesseract unavailable)")
                ocr_engine_used = 'vision'
                page_type_map = self._sample_page_types_vision(pdf_path, page_data)
                text_page_indices = page_type_map.get('text_pages', [])
                if text_page_indices:
                    ocr_page_count = self._ocr_text_pages(
                        pdf_path, page_data, text_page_indices)

            if ocr_page_count > 0:
                # Reclassify remaining pages as figure pages
                figure_page_indices = page_type_map.get('figure_pages', [])
                for idx in figure_page_indices:
                    if idx < len(page_data) and page_data[idx]['classification'] != 'TEXT_HEAVY':
                        page_data[idx]['classification'] = 'FIGURE_PAGE'
                # Update classifications dict
                pdf_info['classifications'] = {}
                for pd_item in page_data:
                    cls = pd_item['classification']
                    pdf_info['classifications'][cls] = pdf_info['classifications'].get(cls, 0) + 1
                logger.info(f"  OCR complete: {ocr_page_count} text pages recovered "
                           f"(engine: {ocr_engine_used})")
                logger.info(f"  Updated classifications: {pdf_info['classifications']}")
            else:
                logger.info("  No text pages detected — treating as figures-only patent")

        # Stage 1.5: Text quality assessment and selective Vision AI re-extraction
        vision_enhanced_count = 0
        reclassified_count = 0
        if not no_vision:
            pages_needing_vision = self.assess_all_pages_quality(page_data)
            reclassified_count = sum(1 for p in page_data if p.get('remediation') == 'reclassify_figure')

            if pages_needing_vision:
                enhanced_texts = self.enhance_text_pages_vision(pdf_path, pages_needing_vision)
                for page_num, enhanced_text in enhanced_texts.items():
                    page_data[page_num]['text'] = enhanced_text
                    page_data[page_num]['text_length'] = len(enhanced_text)
                vision_enhanced_count = len(enhanced_texts)
            else:
                logger.info("  All text pages pass quality threshold — no Vision AI needed")
        else:
            logger.info("Stage 1.5: Skipped (--no-vision mode)")

        # Stage 2: Full text extraction & section splitting
        full_text = self.extract_full_text(page_data)
        sections = self.split_into_sections(full_text)

        # Title fallback: extract from page 2 if Vision AI didn't provide it
        if not bib.get('title') and len(page_data) > 1:
            lines = page_data[1]['text'][:500].split('\n')
            for line in lines[:5]:
                stripped = line.strip()
                if len(stripped) > 20 and stripped.isupper():
                    bib['title'] = stripped.title()
                    break

        # Stage 3: Claims parsing
        claims_data = self.parse_claims(sections.get('claims', ''))

        if claims_only:
            logger.info("Claims-only mode - generating claims output...")
            summaries = self.generate_summaries(sections, claims_data, bib) if not no_vision else {}
            quality_scores = self.calculate_quality_score(bib, claims_data, sections, [], {})
            markdown = self.generate_markdown(bib, sections, claims_data, [], {}, summaries,
                                            quality_scores, len(full_text), total_pages, 0)
            out_stem = self._generate_output_filename(patent_id, bib)
            output_path = self.output_dir / f"{out_stem}.md"
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(markdown)
            self._existing_files.add(output_path.stem)
            logger.info(f"Saved claims-only output: {output_path}")
            self.log_quality(patent_id, quality_scores, saved=True, filename=pdf_path.name)
            return True

        # Cost estimation before expensive stages
        self._estimate_and_log_cost(page_data, no_vision)

        # Stage 4: Vision AI figure analysis
        figures = []
        figure_pages_analyzed = 0
        if not no_vision:
            figures = self.analyze_figure_pages(pdf_path, page_data, sections, max_pages=max_figure_pages)
            figure_pages_analyzed = len([p for p in page_data if p['classification'] == 'FIGURE_PAGE'])
            if max_figure_pages:
                figure_pages_analyzed = min(figure_pages_analyzed, max_figure_pages)
        else:
            logger.info("Stage 4: Skipped (--no-vision mode)")

        # Stage 5: Key data extraction
        key_data = {}
        if not no_vision:
            key_data = self.extract_key_data(sections)
        else:
            logger.info("Stage 5: Skipped (--no-vision mode)")

        # Stage 5.5: SMILES refinement for chemical structures
        if not no_vision and figures:
            figures = self.refine_chemical_structures(pdf_path, figures, sections)

        # Stage 6: Summaries & quality
        summaries = {}
        if not no_vision:
            summaries = self.generate_summaries(sections, claims_data, bib)
        else:
            logger.info("Stage 6: Summaries skipped (--no-vision mode)")

        quality_scores = self.calculate_quality_score(bib, claims_data, sections, figures, key_data)
        logger.info(f"Quality: {quality_scores['overall']}/10 - {quality_scores['assessment']}")

        # Stage 7: Generate markdown
        logger.info("Stage 7: Generating markdown output...")
        markdown = self.generate_markdown(
            bib=bib, sections=sections, claims_data=claims_data,
            figures=figures, key_data=key_data, summaries=summaries,
            quality_scores=quality_scores, full_text_length=len(full_text),
            total_pages=total_pages, figure_pages_analyzed=figure_pages_analyzed,
            text_enhanced_count=vision_enhanced_count,
            reclassified_count=reclassified_count,
            is_scanned_pdf=is_scanned_pdf, ocr_page_count=ocr_page_count,
            total_api_calls=self._api_call_count,
            ocr_engine_used=ocr_engine_used
        )

        out_stem = self._generate_output_filename(patent_id, bib)
        output_path = self.output_dir / f"{out_stem}.md"
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(markdown)
        self._existing_files.add(output_path.stem)

        logger.info(f"\n{'='*70}")
        logger.info(f"SUCCESS: {output_path}")
        logger.info(f"  Pages: {total_pages} | Text: {len(full_text):,} chars")
        logger.info(f"  Claims: {claims_data.get('total_claims', 0)} | Figures: {len(figures)}")
        logger.info(f"  Quality: {quality_scores['overall']}/10 ({quality_scores['assessment']})")
        logger.info(f"{'='*70}")

        self.log_quality(patent_id, quality_scores, saved=True, filename=pdf_path.name)
        return True

    def process_all_patents(self, input_path: Path, no_vision: bool = False,
                           max_figure_pages: Optional[int] = None,
                           skip_existing: bool = True):
        logger.info("="*70)
        logger.info("PATENT PROCESSING PIPELINE")
        logger.info("="*70)
        logger.info(f"Input: {input_path}")
        logger.info(f"Output: {self.output_dir}")
        logger.info(f"Recursive: {self._recursive}")
        logger.info(f"Naming: {self._naming}")
        logger.info(f"Vision AI: {'DISABLED' if no_vision else f'ENABLED ({self.VISION_MODEL})'}")
        logger.info(f"Skip existing: {skip_existing}")
        logger.info("="*70)

        if not self.quality_log_path.exists() or self.quality_log_path.stat().st_size == 0:
            with open(self.quality_log_path, 'w', encoding='utf-8') as f:
                f.write("TIMESTAMP\tPATENT_ID\tQUALITY_SCORE\tASSESSMENT\tSTATUS\tFILENAME\n")

        if self._recursive:
            pdf_files = list(input_path.rglob("*.pdf"))
        else:
            pdf_files = list(input_path.glob("*.pdf"))

        if not pdf_files:
            logger.error("No PDF files found")
            return

        logger.info(f"Found {len(pdf_files)} PDF files")

        success_count = 0
        fail_count = 0

        for i, pdf_path in enumerate(pdf_files):
            logger.info(f"\nPatent {i+1}/{len(pdf_files)}")
            try:
                success = self.process_single_patent(
                    pdf_path, no_vision=no_vision,
                    max_figure_pages=max_figure_pages, skip_existing=skip_existing
                )
                if success:
                    success_count += 1
                else:
                    fail_count += 1
            except Exception as e:
                logger.error(f"Error processing {pdf_path.name}: {e}")
                import traceback
                logger.error(traceback.format_exc())
                fail_count += 1

        logger.info(f"\n{'='*70}")
        logger.info(f"COMPLETE: {success_count} succeeded, {fail_count} failed out of {len(pdf_files)} patents")
        logger.info(f"{'='*70}")


def main():
    import argparse

    parser = argparse.ArgumentParser(
        description='Process patent PDFs to structured markdown with selective Vision AI',
        epilog='Vision AI is used for cover page + figure/drawing pages. Text pages use native extraction.'
    )
    parser.add_argument('--input', type=str,
                       help='Path to folder containing patent PDFs')
    parser.add_argument('--output', type=str, default='output_patents',
                       help='Output directory for markdown files (default: output_patents)')
    parser.add_argument('--single', type=str,
                       help='Process single patent PDF file (provide full path)')
    parser.add_argument('--no-vision', action='store_true',
                       help='Skip Vision AI analysis (text-only extraction)')
    parser.add_argument('--max-figure-pages', type=int, default=None,
                       help='Limit figure pages for Vision AI (default: 50, 0=unlimited)')
    parser.add_argument('--render-dpi', type=int, default=200,
                       help='DPI for rendering pages to images (default: 200)')
    parser.add_argument('--no-skip', action='store_true',
                       help='Re-process patents even if output files exist (default: skip existing)')
    parser.add_argument('--claims-only', action='store_true',
                       help='Extract only claims section (fast, no Vision AI for figures)')
    parser.add_argument('--budget', type=int, default=200,
                       help='Max API calls per patent (default: 200, 0=unlimited)')
    parser.add_argument('--ocr-engine', choices=['auto', 'tesseract', 'vision'],
                       default='auto',
                       help='OCR engine for scanned PDFs: auto (Tesseract if available), '
                            'tesseract (local only), vision (API only) (default: auto)')
    parser.add_argument('--naming', choices=['default', 'detailed', 'dated'],
                       default='default',
                       help='Output filename scheme: default (patent_ID), '
                            'detailed (patent_ID_applicant_title), '
                            'dated (pubdate_ID_title) (default: default)')
    parser.add_argument('--recursive', action='store_true',
                       help='Recursively search input folder for PDF files')
    parser.add_argument('--verbose', action='store_true',
                       help='Enable verbose/debug logging')

    args = parser.parse_args()

    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    if not args.single and not args.input:
        parser.error("Either --single or --input must be provided")

    try:
        pipeline = PatentPipeline(output_dir=args.output, budget=args.budget,
                                   ocr_engine=args.ocr_engine,
                                   naming=args.naming,
                                   recursive=args.recursive)
        pipeline.RENDER_DPI = args.render_dpi

        if args.single:
            pdf_path = Path(args.single)
            if not pdf_path.exists():
                logger.error(f"File not found: {pdf_path}")
                return
            pipeline.process_single_patent(
                pdf_path,
                no_vision=args.no_vision or args.claims_only,
                claims_only=args.claims_only,
                max_figure_pages=args.max_figure_pages,
                skip_existing=not args.no_skip
            )
        else:
            input_path = Path(args.input)
            if not input_path.exists():
                logger.error(f"Directory not found: {input_path}")
                return
            pipeline.process_all_patents(
                input_path, no_vision=args.no_vision,
                max_figure_pages=args.max_figure_pages,
                skip_existing=not args.no_skip
            )

    except Exception as e:
        logger.error(f"Pipeline failed: {e}")
        import traceback
        logger.error(traceback.format_exc())
        raise


if __name__ == '__main__':
    main()
