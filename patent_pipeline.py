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

    # Chemical structure refinement
    CHEMICAL_STRUCTURE_DPI = 250
    MAX_IMAGE_DIMENSION_CHEMICAL = 2048
    SMILES_VALID_CHARS = re.compile(r'^[A-Za-z0-9@+\-\[\]()=#$/\\.%:]+$')

    # Text quality thresholds for Vision AI triggering
    TEXT_QUALITY_GOOD_THRESHOLD = 0.65
    TEXT_QUALITY_GARBLED_THRESHOLD = 0.30
    MAX_VISION_TEXT_PAGES = 30
    VISION_DPI_TEXT_ENHANCEMENT = 300
    MAX_IMAGE_DIMENSION_TEXT = 2048
    PAGES_PER_VISION_TEXT_BATCH = 2

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
        'seventy': 'severity',
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
        'claims': r'(?:Claims|CLAIMS|What\s+[Ii]s\s+[Cc]laimed)',
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

    def __init__(self, output_dir: str = "output_patents"):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)
        self.quality_log_path = self.output_dir / "quality_log.txt"
        self._client = None
        self._existing_files = set(f.stem for f in self.output_dir.glob("*.md"))

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

    def encode_image_to_base64(self, image, format="JPEG", max_dimension=None) -> Tuple[Optional[str], Optional[str]]:
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
            image.save(buffered, format=format, quality=85 if format == "JPEG" else None)
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

    # ─── Stage 1: Bibliographic Data Extraction ────────────────────────────

    def extract_bibliographic_data(self, cover_text: str, pdf_path: Path) -> Dict:
        logger.info("Stage 1: Extracting bibliographic data...")
        bib = {}

        # Patent number from filename (most reliable)
        stem = pdf_path.stem.lower()
        wo_file_match = re.match(r'wo(\d{6,})', stem)
        ep_file_match = re.match(r'ep(\d{6,})', stem)
        us_file_match = re.match(r'us(\d{6,})', stem)

        if wo_file_match:
            num = wo_file_match.group(1)
            bib['patent_number'] = f"WO{num}A1"
            bib['patent_number_formatted'] = f"WO {num[:4]}/{num[4:]} A1"
        elif ep_file_match:
            num = ep_file_match.group(1)
            bib['patent_number'] = f"EP{num}"
            bib['patent_number_formatted'] = f"EP {num}"
        elif us_file_match:
            num = us_file_match.group(1)
            bib['patent_number'] = f"US{num}"
            bib['patent_number_formatted'] = f"US {num}"
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
            if matches:
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
        if re.match(r'(?:a|an)\s+(compound|composition|conjugate|antibody|polypeptide|nucleic acid|vector|cell|adc)', first_100):
            return 'composition'
        elif re.match(r'(?:a|an)\s+(method|process)\s', first_100):
            return 'method'
        elif re.match(r'(use\s+of|.*for\s+use\s+)', first_100):
            return 'use'
        elif re.match(r'(?:a|an)\s+pharmaceutical', first_100):
            return 'pharmaceutical'
        elif re.match(r'(?:the|an?)\s+(?:compound|composition|conjugate|antibody|adc)', first_100):
            return 'composition'
        elif re.match(r'(?:the|an?)\s+(?:method|process)', first_100):
            return 'method'
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

        if max_pages and len(figure_pages) > max_pages:
            logger.info(f"  Limiting from {len(figure_pages)} to {max_pages} figure pages")
            figure_pages = figure_pages[:max_pages]

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
                for idx in figure_page_indices:
                    fig_start = max(1, idx * 2)
                    for fn in range(fig_start, fig_start + 4):
                        fig_nums_to_search.add(fn)
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
        """Re-extract SMILES for chemical structures with low/medium confidence.

        Uses higher DPI + RAG context (figure descriptions, patent text, scaffold info)
        for a focused SMILES-only extraction pass.
        """
        chem_figs = [f for f in figures
                     if f.get('type') == 'chemical_structure'
                     and f.get('smiles_confidence') in ('low', 'medium', None)
                     and f.get('page')]

        if not chem_figs:
            logger.info("Stage 5.5: No chemical structures need SMILES refinement")
            return figures

        logger.info(f"Stage 5.5: Refining SMILES for {len(chem_figs)} chemical structures...")

        client = self.get_anthropic_client()
        if not client:
            return figures

        drawings_text = sections.get('drawings', '')
        claims_text = sections.get('claims', '')[:4000]
        scaffold_context = ""
        scaffold_match = re.search(r'formula\s*\(I+\)[^.]*\.', claims_text, re.IGNORECASE)
        if scaffold_match:
            scaffold_context = f"\nCore scaffold from claims: {scaffold_match.group(0)[:500]}\n"

        pages_needed = sorted(set(f['page'] - 1 for f in chem_figs))
        images = self.render_pages_to_images(pdf_path, pages_needed,
                                              dpi=self.CHEMICAL_STRUCTURE_DPI)

        refined_count = 0
        for fig in chem_figs:
            page_0 = fig['page'] - 1
            if page_0 not in images:
                continue

            img_base64, media_type = self.encode_image_to_base64(
                images[page_0], max_dimension=self.MAX_IMAGE_DIMENSION_CHEMICAL)
            if not img_base64:
                continue

            label = fig.get('label', 'Unknown')
            description = fig.get('description', '')[:500]

            legend = ""
            fig_match = re.search(rf'FIG[S.]?\s*{re.escape(str(fig["page"]))}\b[^\n]*',
                                  drawings_text, re.IGNORECASE)
            if fig_match:
                legend = f"\nFigure legend: {fig_match.group(0).strip()[:300]}\n"

            compound_mention = ""
            compound_id = re.search(r'(\d{3,4})', label)
            if compound_id:
                for section_name, section_text in sections.items():
                    idx = section_text.find(compound_id.group(1))
                    if idx >= 0:
                        start = max(0, idx - 200)
                        end = min(len(section_text), idx + 300)
                        compound_mention = f"\nPatent text about this compound: ...{section_text[start:end]}...\n"
                        break

            prompt = f"""Focus on the chemical structure labeled "{label}" on page {fig['page']}.

Previous description: {description}
{scaffold_context}{legend}{compound_mention}
Provide ONLY a JSON object with:
- "smiles": Canonical SMILES notation. Include stereochemistry (@, @@) where visible. For complex structures, provide the FULL structure, not just the core scaffold.
- "smiles_confidence": "high" if simple/clear structure, "medium" if moderate complexity, "low" if uncertain
- "inchi": InChI string if possible, null otherwise
- "molecular_formula": e.g. "C23H28BrN7O" if determinable
- "molecular_weight": approximate MW if formula is known, null otherwise

Return ONLY valid JSON. No other text."""

            content = [
                {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": img_base64}},
                {"type": "text", "text": prompt}
            ]

            try:
                response = client.messages.create(
                    model=self.VISION_MODEL,
                    max_tokens=1024,
                    messages=[{"role": "user", "content": content}]
                )
                result = self.parse_json_response(response.content[0].text)
                if result and result.get('smiles'):
                    new_smiles = result['smiles']
                    if self.validate_smiles(new_smiles):
                        old_conf = fig.get('smiles_confidence', 'none')
                        new_conf = result.get('smiles_confidence', 'medium')
                        fig['smiles'] = new_smiles
                        fig['smiles_confidence'] = new_conf
                        if result.get('inchi'):
                            fig['inchi'] = result['inchi']
                        if result.get('molecular_formula'):
                            fig['molecular_formula'] = result['molecular_formula']
                        if result.get('molecular_weight'):
                            fig['molecular_weight'] = result['molecular_weight']
                        refined_count += 1
                        logger.info(f"    {label}: SMILES refined ({old_conf} -> {new_conf})")
                    else:
                        logger.warning(f"    {label}: Refined SMILES failed validation")
            except Exception as e:
                logger.error(f"    {label}: SMILES refinement failed: {e}")

        for page_num in list(images.keys()):
            del images[page_num]

        logger.info(f"  SMILES refinement complete: {refined_count}/{len(chem_figs)} improved")

        for fig in figures:
            if fig.get('smiles') and not fig.get('smiles_validated'):
                fig['smiles_validated'] = self.validate_smiles(fig['smiles'])

        return figures

    # ─── Stage 6: Executive Summary & Protection Scope ─────────────────────

    def generate_summaries(self, sections: Dict[str, str], claims_data: Dict, bib: Dict) -> Dict[str, str]:
        logger.info("Stage 6: Generating executive summary and protection scope...")
        client = self.get_anthropic_client()
        if not client:
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
                         text_enhanced_count: int = 0, reclassified_count: int = 0) -> str:
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
        method = "native_vision_enhanced" if text_enhanced_count > 0 else "native_vision_selective"
        md.append(f"extraction_method: {method}\n")
        md.append(f"vision_model: {self.VISION_MODEL}\n")
        md.append(f"processing_date: {datetime.now().strftime('%Y-%m-%d')}\n")
        md.append(f"figure_pages_analyzed: {figure_pages_analyzed}\n")
        if text_enhanced_count > 0:
            md.append(f"text_pages_enhanced: {text_enhanced_count}\n")
        if reclassified_count > 0:
            md.append(f"figure_pages_reclassified: {reclassified_count}\n")
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
                label = fig.get('label', 'Structure')
                smiles = fig.get('smiles', '-') if str(fig.get('smiles', '')).lower() not in ('null', 'none', '') else '-'
                conf = fig.get('smiles_confidence', '-') or '-'
                formula = fig.get('molecular_formula', '-') or '-'
                mw = fig.get('molecular_weight', '-') or '-'
                page = fig.get('page', '-')
                md.append(f"| {label} | {label} | `{smiles}` | {conf} | {formula} | {mw} | {page} |\n")
            md.append("\n")
            # Detailed descriptions below the table
            for fig in chem_figures[:20]:
                label = fig.get('label', 'Structure')
                md.append(f"### {label}\n")
                if fig.get('description'):
                    md.append(f"**Description**: {fig['description']}\n\n")
                if fig.get('smiles') and str(fig['smiles']).lower() not in ('null', 'none', ''):
                    confidence = fig.get('smiles_confidence', 'unknown')
                    md.append(f"**SMILES** (confidence: {confidence}): `{fig['smiles']}`\n\n")
                if fig.get('inchi'):
                    md.append(f"**InChI**: `{fig['inchi']}`\n\n")
                if fig.get('molecular_formula'):
                    md.append(f"**Formula**: {fig['molecular_formula']}")
                    if fig.get('molecular_weight'):
                        md.append(f" | **MW**: {fig['molecular_weight']}")
                    md.append("\n\n")
                if fig.get('key_findings'):
                    md.append(f"**Key Data**: {fig['key_findings']}\n\n")
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
                assay = res.get('assay_type', '-') or '-'
                compounds = ', '.join(res.get('compounds_tested', ['-'])) if res.get('compounds_tested') else '-'
                models = ', '.join(res.get('cell_lines_or_models', res.get('cell_lines', ['-']))) if res.get('cell_lines_or_models') or res.get('cell_lines') else '-'
                key_result = (res.get('key_result', '-') or '-').replace('|', '/')
                comparison = (res.get('comparison', '-') or '-').replace('|', '/')
                md.append(f"| {assay} | {compounds} | {models} | {key_result} | {comparison} |\n")
            md.append("\n---\n\n")

        # Figures & Drawings (non-chemical)
        non_chem_figures = [f for f in figures if f.get('type') not in ('chemical_structure', 'reaction_scheme', 'generic_structure')]
        if non_chem_figures:
            md.append("## Figures & Drawings\n\n")
            for fig in non_chem_figures[:40]:
                label = fig.get('label', 'Figure')
                fig_type = fig.get('type', 'other')
                md.append(f"### {label}\n")
                md.append(f"**Type**: {fig_type.replace('_', ' ').title()}\n\n")
                if fig.get('title'):
                    md.append(f"**Title**: {fig['title']}\n\n")
                if fig.get('x_axis'):
                    md.append(f"**X-axis**: {fig['x_axis']} | **Y-axis**: {fig.get('y_axis', 'N/A')}\n\n")
                if fig.get('data_series'):
                    series = fig['data_series'] if isinstance(fig['data_series'], list) else [fig['data_series']]
                    md.append(f"**Data Series**: {', '.join(str(s) for s in series)}\n\n")
                if fig.get('key_findings'):
                    md.append(f"**Key Findings**: {fig['key_findings']}\n\n")
                elif fig.get('description'):
                    md.append(f"**Description**: {fig['description']}\n\n")
                md.append(f"*Page {fig.get('page', 'N/A')}*\n\n")
            md.append("---\n\n")

        # Figures Index (machine-parseable summary table)
        if figures:
            md.append("## Figures Index\n\n")
            md.append("| Figure | Type | Page | Title | Key Finding |\n")
            md.append("|--------|------|-----:|-------|-------------|\n")
            for fig in figures[:50]:
                label = fig.get('label', '-')
                fig_type = fig.get('type', 'other').replace('_', ' ')
                page = fig.get('page', '-')
                title = (fig.get('title', '-') or '-').replace('|', '/')[:60]
                finding = (fig.get('key_findings', '-') or '-').replace('|', '/').replace('\n', ' ')[:80]
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
        md.append(f"- **Vision Model**: {self.VISION_MODEL}\n")
        md.append(f"- **Total Pages**: {total_pages}\n")
        md.append(f"- **Figure Pages Analyzed**: {figure_pages_analyzed}\n")
        if text_enhanced_count > 0:
            md.append(f"- **Text Pages Vision-Enhanced**: {text_enhanced_count}\n")
        if reclassified_count > 0:
            md.append(f"- **Text Pages Reclassified as Figure**: {reclassified_count}\n")
        md.append(f"- **Total Figures Described**: {len(figures)}\n")
        md.append(f"- **Text Length**: {full_text_length:,} characters\n")
        md.append(f"- **Processing Date**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")

        return "".join(md)

    # ─── Main Processing Orchestration ─────────────────────────────────────

    def process_single_patent(self, pdf_path: Path, no_vision: bool = False,
                             claims_only: bool = False, max_figure_pages: Optional[int] = None,
                             skip_existing: bool = False) -> bool:
        logger.info(f"\n{'='*70}")
        logger.info(f"PATENT PIPELINE: {pdf_path.name}")
        logger.info(f"{'='*70}")

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

        if skip_existing and f"patent_{patent_id}" in self._existing_files:
            logger.info(f"Patent {patent_id} already processed - skipping")
            return True

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
            output_path = self.output_dir / f"patent_{patent_id}.md"
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(markdown)
            logger.info(f"Saved claims-only output: {output_path}")
            self.log_quality(patent_id, quality_scores, saved=True, filename=pdf_path.name)
            return True

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
            reclassified_count=reclassified_count
        )

        output_path = self.output_dir / f"patent_{patent_id}.md"
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
                           skip_existing: bool = False):
        logger.info("="*70)
        logger.info("PATENT PROCESSING PIPELINE")
        logger.info("="*70)
        logger.info(f"Input: {input_path}")
        logger.info(f"Output: {self.output_dir}")
        logger.info(f"Vision AI: {'DISABLED' if no_vision else f'ENABLED ({self.VISION_MODEL})'}")
        logger.info("="*70)

        if not self.quality_log_path.exists() or self.quality_log_path.stat().st_size == 0:
            with open(self.quality_log_path, 'w', encoding='utf-8') as f:
                f.write("TIMESTAMP\tPATENT_ID\tQUALITY_SCORE\tASSESSMENT\tSTATUS\tFILENAME\n")

        pdf_files = list(input_path.glob("*.pdf"))
        if not pdf_files:
            pdf_files = list(input_path.rglob("*.pdf"))

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
                       help='Limit number of figure pages analyzed with Vision AI (default: all)')
    parser.add_argument('--render-dpi', type=int, default=200,
                       help='DPI for rendering pages to images (default: 200)')
    parser.add_argument('--skip-existing', action='store_true',
                       help='Skip patents that already have output files')
    parser.add_argument('--claims-only', action='store_true',
                       help='Extract only claims section (fast, no Vision AI for figures)')
    parser.add_argument('--verbose', action='store_true',
                       help='Enable verbose/debug logging')

    args = parser.parse_args()

    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    if not args.single and not args.input:
        parser.error("Either --single or --input must be provided")

    try:
        pipeline = PatentPipeline(output_dir=args.output)
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
                skip_existing=False
            )
        else:
            input_path = Path(args.input)
            if not input_path.exists():
                logger.error(f"Directory not found: {input_path}")
                return
            pipeline.process_all_patents(
                input_path, no_vision=args.no_vision,
                max_figure_pages=args.max_figure_pages,
                skip_existing=args.skip_existing
            )

    except Exception as e:
        logger.error(f"Pipeline failed: {e}")
        import traceback
        logger.error(traceback.format_exc())
        raise


if __name__ == '__main__':
    main()
