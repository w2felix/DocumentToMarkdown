"""
Scientific Talk Processing Pipeline with Vision Analysis
Converts multi-slide presentation PDFs (screenshot-based) to structured markdown using AI-powered vision analysis.

Designed for AACR 2026 conference talks where PDFs contain slide screenshots (no extractable text).
Uses batch Vision AI processing to extract slide content and generate narrative summaries.
"""

import os
import re
import html
import base64
import argparse
from pathlib import Path
from typing import List, Dict, Tuple, Optional, Any
from io import BytesIO
from concurrent.futures import ThreadPoolExecutor, as_completed
import logging
from datetime import datetime

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Load credentials and shared auth
from pipeline_security import validate_path, validate_output_path, sanitize_filename
from auth import get_anthropic_client as _get_shared_client
from anthropic_helpers import first_text


class TalkPipeline:
    """Pipeline for processing multi-slide scientific talk PDFs into structured markdown.

    Each talk PDF contains slide screenshots (pure images, no extractable text).
    Uses batch Vision AI to extract content from slides and generate summaries.
    """

    VISION_MODEL = "claude-sonnet-4-6"
    SLIDES_PER_BATCH = 5
    MAX_CONCURRENT_BATCHES = 5
    RENDER_DPI = 150
    MAX_IMAGE_DIMENSION = 1568
    MAX_TOKENS_SLIDE_EXTRACTION = 8192
    MAX_TOKENS_SUMMARY = 2048
    MAX_TOKENS_TAKEAWAYS = 1024

    NAMING_SCHEMES = {
        'default': 'talk_{talk_number}_{session_code}_{speaker}_{title_slug}',
        'detailed': 'talk_{session_code}_{speaker}_{title_slug}',
        'dated': '{date}_{session_code}_{speaker}_{title_slug}',
    }

    def __init__(self, talks_folder: str, metadata_excel: str, output_dir: str = "output_talks",
                 recursive: bool = False, naming: str = "default"):
        self.talks_folder = Path(talks_folder)
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)
        self._recursive = recursive
        self._naming = naming

        self.processing_log_path = self.output_dir / "processing_log.txt"
        self._init_processing_log()

        self._existing_files = set(f.stem for f in self.output_dir.glob("*.md"))
        # Build O(1) lookup for processed talk keys (talk_number + session_code)
        self._processed_keys = set()
        for f in self._existing_files:
            match = re.match(r'^talk_(\d+)_([A-Z0-9]+)_', f, re.IGNORECASE)
            if match:
                self._processed_keys.add(f"{match.group(1)}_{match.group(2)}")

        self.metadata_excel = Path(metadata_excel)
        self.metadata_df = None
        self._load_metadata()

        self._client = None

        conda_prefix = os.environ.get('CONDA_PREFIX')
        if conda_prefix and not os.environ.get('TESSDATA_PREFIX'):
            tessdata_dir = os.path.join(conda_prefix, 'Library', 'share', 'tessdata')
            if os.path.exists(tessdata_dir):
                os.environ['TESSDATA_PREFIX'] = tessdata_dir

    def _load_metadata(self):
        """Load abstracts metadata from Excel."""
        try:
            import pandas as pd
            self.metadata_df = pd.read_excel(self.metadata_excel, sheet_name='Sheet1')
            logger.info(f"Loaded metadata: {len(self.metadata_df)} rows")
        except Exception as e:
            logger.warning(f"Could not load metadata: {e}")
            self.metadata_df = None

    def _init_processing_log(self):
        """Initialize the processing log file with header if needed."""
        if not self.processing_log_path.exists() or self.processing_log_path.stat().st_size == 0:
            with open(self.processing_log_path, 'w', encoding='utf-8') as f:
                f.write("TIMESTAMP\tTALK_NUM\tSESSION\tSPEAKER\tSLIDES\tOCR_CHARS\t"
                        "CONTENT_CHARS\tSUMMARY_CHARS\tABSTRACT_MATCH\tQUALITY\tASSESSMENT\tSTATUS\n")

    def log_processing(self, file_info: Dict, num_slides: int, ocr_chars: int,
                       content_chars: int, summary_chars: int,
                       abstract_matched: bool, quality_scores: Optional[Dict],
                       status: str):
        """Log processing results for a single talk."""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        quality = quality_scores['overall'] if quality_scores else ''
        assessment = quality_scores['assessment'] if quality_scores else ''
        entry = (f"{timestamp}\t{file_info['talk_number']}\t{file_info['session_code']}\t"
                 f"{file_info['speaker_last_name']}\t{num_slides}\t{ocr_chars}\t"
                 f"{content_chars}\t{summary_chars}\t"
                 f"{'Yes' if abstract_matched else 'No'}\t{quality}\t{assessment}\t{status}\n")
        with open(self.processing_log_path, 'a', encoding='utf-8') as f:
            f.write(entry)

    def get_anthropic_client(self):
        """Get configured Anthropic client instance (cached via shared auth module)."""
        if self._client is not None:
            return self._client
        try:
            self._client = _get_shared_client()
        except RuntimeError:
            logger.error("API credentials not found")
            return None
        return self._client

    def list_talk_files(self) -> List[Path]:
        """List all PDF files in the talks folder."""
        if self._recursive:
            pdf_files = sorted(self.talks_folder.rglob("*.pdf"))
        else:
            pdf_files = sorted(self.talks_folder.glob("*.pdf"))
        logger.info(f"Found {len(pdf_files)} talk PDFs")
        return pdf_files

    def parse_filename(self, pdf_path: Path) -> Dict[str, str]:
        """Parse talk filename into components.

        Expected format: NN_SESSION_Speaker_Title.pdf
        e.g., 04_ED03_C.Bunne_Toward virtual patients with multimodal foundation models.pdf
        """
        stem = pdf_path.stem
        parts = stem.split('_', 3)

        result = {
            'talk_number': parts[0] if len(parts) >= 1 else '',
            'session_code': parts[1] if len(parts) >= 2 else '',
            'speaker': parts[2] if len(parts) >= 3 else '',
            'title': parts[3] if len(parts) >= 4 else stem,
        }

        speaker = result['speaker']
        name_parts = speaker.replace('-', ' ').split('.')
        result['speaker_last_name'] = name_parts[-1].strip() if name_parts else speaker
        result['speaker_formatted'] = speaker.replace('.', '. ').replace('  ', ' ').strip()

        return result

    def match_abstract(self, file_info: Dict) -> Optional[Dict]:
        """Match a talk to its abstract in the metadata Excel by speaker name and title keywords."""
        if self.metadata_df is None:
            return None

        import pandas as pd
        last_name = file_info['speaker_last_name']
        title_from_file = file_info.get('title', '')

        matches = self.metadata_df[
            self.metadata_df['Abstract Authors'].str.contains(last_name, case=False, na=False)
        ]

        if matches.empty:
            return None

        title_words = [w.lower() for w in title_from_file.split() if len(w) > 3]

        def _score_title(abstract_title):
            clean_title = re.sub(r'<[^>]+>', '', str(abstract_title)).lower()
            return sum(1 for w in title_words if w in clean_title)

        matches = matches.copy()
        matches['_match_score'] = matches['Abstract Title'].apply(_score_title)
        best_idx = matches['_match_score'].idxmax()
        best_score = matches.loc[best_idx, '_match_score']
        best_row = matches.loc[best_idx]

        if best_score >= 2:
            def clean_html(val):
                if pd.isna(val):
                    return None
                return html.unescape(re.sub(r'<[^>]+>', '', str(val)))

            return {
                'abstract_title': clean_html(best_row['Abstract Title']),
                'abstract_text': clean_html(best_row['Abstract Text']),
                'abstract_number': str(best_row['Abstract Number']) if pd.notna(best_row['Abstract Number']) else None,
                'authors': clean_html(best_row['Abstract Authors']),
                'companies': clean_html(best_row['Abstract Companies']),
                'session_title': str(best_row['Session Title']) if pd.notna(best_row['Session Title']) else None,
                'start_time': str(best_row['Start']) if pd.notna(best_row['Start']) else None,
                'abstract_url': str(best_row['AbstractUrl']) if pd.notna(best_row['AbstractUrl']) else None,
                'match_score': best_score,
            }

        return None

    def convert_pdf_to_slide_images(self, pdf_path: Path) -> List[Any]:
        """Convert all pages of a talk PDF to PIL images."""
        import pdf_utils

        num_pages = pdf_utils.page_count(pdf_path)
        logger.info(f"Converting {num_pages} slides at {self.RENDER_DPI} DPI...")

        rendered = pdf_utils.render_pages(pdf_path, dpi=self.RENDER_DPI)
        if not rendered:
            return []
        images = [rendered[i] for i in sorted(rendered.keys())]
        if images:
            logger.info(f"Converted {len(images)} slides ({images[0].size[0]}x{images[0].size[1]} px)")
        return images

    def ocr_slide_images(self, slide_images: List[Any]) -> List[str]:
        """Run OCR on each slide image to extract text as RAG context for Vision AI."""
        try:
            import pytesseract

            ocr_texts = []
            for img in slide_images:
                text = pytesseract.image_to_string(img)
                ocr_texts.append(text.strip())

            total_chars = sum(len(t) for t in ocr_texts)
            non_empty = sum(1 for t in ocr_texts if len(t) > 20)
            logger.info(f"OCR complete: {total_chars} chars from {non_empty}/{len(slide_images)} slides with content")
            return ocr_texts

        except ImportError:
            logger.warning("pytesseract not installed — skipping OCR pre-pass")
            return [""] * len(slide_images)
        except Exception as e:
            logger.warning(f"OCR failed: {e} — proceeding without OCR context")
            return [""] * len(slide_images)

    def encode_image_to_base64(self, image, format="JPEG") -> Tuple[Optional[str], Optional[str]]:
        """Convert PIL Image to base64 string for API, resizing if needed."""
        try:
            from PIL import Image

            width, height = image.size
            if max(width, height) > self.MAX_IMAGE_DIMENSION:
                scale = self.MAX_IMAGE_DIMENSION / max(width, height)
                image = image.resize((int(width * scale), int(height * scale)), Image.Resampling.LANCZOS)

            if image.mode == 'RGBA':
                image = image.convert('RGB')

            buffered = BytesIO()
            image.save(buffered, format=format, quality=85)
            img_base64 = base64.b64encode(buffered.getvalue()).decode('utf-8')
            media_type = f"image/{format.lower()}"
            return img_base64, media_type

        except Exception as e:
            logger.error(f"Error encoding image: {e}")
            return None, None

    def extract_slides_batch(self, encoded_images: List[Tuple[str, str]], batch_start: int,
                            ocr_texts: List[str] = None) -> str:
        """Extract content from a batch of slides using Vision AI with OCR as RAG context.

        Args:
            encoded_images: List of (base64_data, media_type) tuples, pre-encoded.
            batch_start: Index of the first slide in this batch (0-based).
            ocr_texts: OCR reference text per slide.
        """
        client = self.get_anthropic_client()
        if not client:
            return ""

        content_blocks = []
        for i, (img_b64, media_type) in enumerate(encoded_images):
            content_blocks.append({
                "type": "image",
                "source": {"type": "base64", "media_type": media_type, "data": img_b64}
            })
            ocr_context = ""
            if ocr_texts and i < len(ocr_texts) and ocr_texts[i]:
                ocr_context = f"\n[OCR reference: {ocr_texts[i][:500]}]"
            content_blocks.append({
                "type": "text",
                "text": f"[Slide {batch_start + i + 1}]{ocr_context}"
            })

        content_blocks.append({
            "type": "text",
            "text": """For each slide above, extract ALL text content and describe any figures, charts, or diagrams.
Use the OCR reference text (if provided) as a guide — it may contain errors, so verify against the image and correct mistakes.

Format your response as:

## Slide N: [Slide Title or Topic]
**Text content**: [All text on the slide, corrected and complete]
**Visual elements**: [Description of any figures, charts, plots, diagrams — what they show, axes, legends, key data points]
**Key point**: [One-sentence summary of what this slide communicates]

Be thorough — include all bullet points, labels, legends, axis titles, and footnotes.
For title slides, section dividers, or acknowledgments/disclosures, note them briefly.
Do NOT skip any data-containing slides."""
        })

        try:
            message = client.messages.create(
                model=self.VISION_MODEL,
                max_tokens=self.MAX_TOKENS_SLIDE_EXTRACTION,
                messages=[{"role": "user", "content": content_blocks}]
            )
            result = first_text(message)
            logger.info(f"  Batch {batch_start+1}-{batch_start+len(encoded_images)}: "
                       f"{len(result)} chars (in:{message.usage.input_tokens}, out:{message.usage.output_tokens})")
            return result
        except Exception as e:
            logger.error(f"Error in batch slide extraction: {e}")
            return ""

    def extract_all_slides(self, slide_images: List[Any]) -> Tuple[str, int]:
        """Run OCR on all slides, then process in concurrent batches with Vision AI.

        Returns:
            Tuple of (slide_content, total_ocr_chars)
        """
        logger.info("Running OCR pre-pass on all slides...")
        ocr_texts = self.ocr_slide_images(slide_images)
        total_ocr_chars = sum(len(t) for t in ocr_texts)

        logger.info("Encoding slide images to base64 per batch...")
        total_slides = len(slide_images)
        batches = []
        for batch_start in range(0, total_slides, self.SLIDES_PER_BATCH):
            batch_end = min(batch_start + self.SLIDES_PER_BATCH, total_slides)
            # Encode only this batch's images, then free them immediately
            batch_encoded = []
            for idx in range(batch_start, batch_end):
                img_b64, media_type = self.encode_image_to_base64(slide_images[idx])
                if img_b64 is not None:
                    batch_encoded.append((img_b64, media_type))
                else:
                    batch_encoded.append(("", "image/jpeg"))
                # Free the PIL image immediately after encoding
                try:
                    slide_images[idx].close()
                except Exception:
                    pass
                slide_images[idx] = None
            batches.append((batch_start, batch_encoded,
                           ocr_texts[batch_start:batch_end]))
        del slide_images

        workers = min(self.MAX_CONCURRENT_BATCHES, len(batches))
        logger.info(f"Sending {len(batches)} batches with concurrency={workers}...")
        results = [None] * len(batches)

        with ThreadPoolExecutor(max_workers=workers) as executor:
            future_to_idx = {}
            for idx, (batch_start, batch_encoded, batch_ocr) in enumerate(batches):
                future = executor.submit(self.extract_slides_batch, batch_encoded, batch_start, batch_ocr)
                future_to_idx[future] = idx

            for future in as_completed(future_to_idx):
                idx = future_to_idx[future]
                batch_start = batches[idx][0]
                batch_end = batch_start + len(batches[idx][1])
                try:
                    results[idx] = future.result()
                except Exception as e:
                    logger.error(f"Batch {batch_start+1}-{batch_end} failed: {e}")
                    results[idx] = ""

        all_content = [r for r in results if r]
        return "\n\n".join(all_content), total_ocr_chars

    def generate_summary_and_takeaways(self, slide_content: str, file_info: Dict,
                                       abstract: Optional[Dict]) -> Tuple[str, str]:
        """Generate executive summary and key takeaways in a single API call.

        Returns:
            Tuple of (executive_summary, key_takeaways)
        """
        client = self.get_anthropic_client()
        if not client:
            return "", ""

        abstract_context = ""
        if abstract and abstract.get('abstract_text'):
            abstract_context = f"""
The published abstract for this talk is:
---
{abstract['abstract_text'][:3000]}
---
Use this as additional context to validate and enrich your summary.
"""

        MAX_SUMMARY_CONTENT = 40000
        if len(slide_content) > MAX_SUMMARY_CONTENT:
            slides = re.split(r'(?=## Slide \d+)', slide_content)
            kept = []
            total = 0
            for s in slides:
                if total + len(s) > MAX_SUMMARY_CONTENT:
                    break
                kept.append(s)
                total += len(s)
            content_for_summary = ''.join(kept)
        else:
            content_for_summary = slide_content

        prompt = f"""You are summarizing a scientific conference talk by {file_info['speaker_formatted']}.
Talk title: "{file_info['title']}"
Session: {file_info['session_code']}
{abstract_context}
Below is the extracted content from all slides in this talk:
---
{content_for_summary}
---

Produce TWO sections separated by the exact delimiter "---TAKEAWAYS---":

SECTION 1: Executive Summary (3-5 paragraphs)
- State the main research question or topic
- Describe the key methods or approaches presented
- Highlight the most important results and findings
- Note any clinical implications, novel insights, or future directions
- Capture the overall narrative arc of the presentation
- Write in scientific prose, not bullet points. Be specific about data and findings.
- Do NOT include a heading — start directly with the summary text.

---TAKEAWAYS---

SECTION 2: Key Takeaways (4-7 bullet points)
- Return ONLY markdown bullet points (using "- " prefix), one per takeaway.
- Each takeaway should be a single concise sentence capturing one key finding or message.
- Focus on novel results, clinical implications, and methodological advances.
- Do NOT include a heading — just the bullet points."""

        try:
            message = client.messages.create(
                model=self.VISION_MODEL,
                max_tokens=self.MAX_TOKENS_SUMMARY + self.MAX_TOKENS_TAKEAWAYS,
                messages=[{"role": "user", "content": prompt}]
            )
            response = first_text(message)
            logger.info(f"Summary + takeaways: {len(response)} chars "
                       f"(in:{message.usage.input_tokens}, out:{message.usage.output_tokens})")

            if "---TAKEAWAYS---" in response:
                parts = response.split("---TAKEAWAYS---", 1)
                summary = parts[0].strip()
                takeaways = parts[1].strip()
            else:
                summary = response
                takeaways = ""

            return summary, takeaways
        except Exception as e:
            logger.error(f"Error generating summary and takeaways: {e}")
            return "", ""

    def calculate_quality_score(self, slide_content: str, executive_summary: str,
                               key_takeaways: str, num_slides: int,
                               ocr_chars: int, abstract_matched: bool) -> Dict[str, Any]:
        """Calculate quality scores for the extracted talk content.

        Returns dictionary with component scores, overall score (0-10), and assessment label.
        """
        scores = {}

        # 1. Content coverage (0-10): chars extracted per slide
        content_per_slide = len(slide_content) / max(num_slides, 1)
        scores['content_coverage'] = round(min(10, content_per_slide / 80), 1)

        # 2. OCR support (0-10): ratio of OCR chars to content
        if len(slide_content) > 0 and ocr_chars > 0:
            ocr_ratio = ocr_chars / len(slide_content)
            scores['ocr_support'] = round(min(10, ocr_ratio * 15), 1)
        else:
            scores['ocr_support'] = 0.0

        # 3. Summary quality (0-10): executive summary length + takeaway count
        summary_score = 0.0
        if executive_summary:
            summary_score += min(6, len(executive_summary) / 500)
        if key_takeaways:
            takeaway_count = key_takeaways.count('- ') + key_takeaways.count('* ')
            summary_score += min(4, takeaway_count * 0.8)
        scores['summary_quality'] = round(min(10, summary_score), 1)

        # 4. Slide coverage (0-10): % of slides that produced content
        slide_headers = len(re.findall(r'## Slide \d+', slide_content))
        coverage_ratio = slide_headers / max(num_slides, 1)
        scores['slide_coverage'] = round(min(10, coverage_ratio * 10), 1)

        # 5. Abstract enrichment (0-10)
        scores['abstract_enrichment'] = 10.0 if abstract_matched else 5.0

        # Weighted overall
        weights = {
            'content_coverage': 0.30,
            'ocr_support': 0.15,
            'summary_quality': 0.25,
            'slide_coverage': 0.20,
            'abstract_enrichment': 0.10,
        }
        overall = sum(scores[k] * weights[k] for k in weights)
        overall = round(max(0, min(10, overall)), 1)
        scores['overall'] = overall

        if overall >= 8.0:
            scores['assessment'] = 'Excellent'
        elif overall >= 5.5:
            scores['assessment'] = 'Good'
        elif overall >= 4.0:
            scores['assessment'] = 'Fair'
        else:
            scores['assessment'] = 'Poor - Consider manual review'

        logger.info(f"  Quality breakdown:")
        logger.info(f"    Content coverage: {scores['content_coverage']:.1f}/10 (weight 0.30)")
        logger.info(f"    OCR support: {scores['ocr_support']:.1f}/10 (weight 0.15)")
        logger.info(f"    Summary quality: {scores['summary_quality']:.1f}/10 (weight 0.25)")
        logger.info(f"    Slide coverage: {scores['slide_coverage']:.1f}/10 (weight 0.20)")
        logger.info(f"    Abstract enrichment: {scores['abstract_enrichment']:.1f}/10 (weight 0.10)")
        logger.info(f"    OVERALL: {overall:.1f}/10 — {scores['assessment']}")

        return scores

    @staticmethod
    def _yaml_escape(value: str) -> str:
        """Escape a string for safe YAML double-quoted output."""
        return value.replace('\\', '\\\\').replace('"', '\\"')

    def generate_markdown(self, file_info: Dict, abstract: Optional[Dict],
                         slide_content: str, executive_summary: str,
                         key_takeaways: str, num_slides: int,
                         quality_scores: Dict = None) -> str:
        """Generate the final structured markdown document."""
        md = []

        # YAML frontmatter
        md.append("---")
        md.append(f"talk_number: {file_info['talk_number']}")
        md.append(f"session_code: {file_info['session_code']}")
        if abstract and abstract.get('session_title'):
            md.append(f"session_title: \"{self._yaml_escape(abstract['session_title'])}\"")
        md.append(f"speaker: \"{self._yaml_escape(file_info['speaker_formatted'])}\"")
        md.append(f"title: \"{self._yaml_escape(file_info['title'])}\"")
        md.append(f"num_slides: {num_slides}")
        if abstract:
            if abstract.get('abstract_number'):
                md.append(f"abstract_number: {abstract['abstract_number']}")
            if abstract.get('abstract_url'):
                md.append(f"abstract_url: \"{self._yaml_escape(abstract['abstract_url'])}\"")
            if abstract.get('start_time'):
                md.append(f"start_time: \"{self._yaml_escape(abstract['start_time'])}\"")
        md.append(f"processing_date: {datetime.now().strftime('%Y-%m-%d')}")
        md.append(f"vision_model: {self.VISION_MODEL}")
        if quality_scores:
            md.append(f"quality_overall: {quality_scores['overall']}/10")
            md.append(f"quality_assessment: {quality_scores['assessment']}")
        md.append("---\n")

        # Title and header
        md.append(f"# {file_info['title']}\n")
        md.append(f"**Speaker**: {file_info['speaker_formatted']}")
        session_title = abstract['session_title'] if abstract and abstract.get('session_title') else ''
        if session_title:
            md.append(f"**Session**: {file_info['session_code']} — {session_title}")
        else:
            md.append(f"**Session**: {file_info['session_code']}")
        if abstract and abstract.get('authors'):
            md.append(f"\n**Authors**: {abstract['authors']}")
        if abstract and abstract.get('companies'):
            md.append(f"\n**Affiliations**: {abstract['companies']}")
        md.append("\n---\n")

        # Executive Summary
        if executive_summary:
            md.append("## Executive Summary\n")
            md.append(f"{executive_summary}\n")
            md.append("---\n")

        # Abstract (if available from metadata)
        if abstract and abstract.get('abstract_text'):
            md.append("## Abstract\n")
            md.append(f"{abstract['abstract_text']}\n")
            md.append("---\n")

        # Slide Content
        if slide_content:
            md.append("## Slide Content\n")
            md.append(slide_content)
            md.append("\n---\n")

        # Key Takeaways
        if key_takeaways:
            md.append("## Key Takeaways\n")
            md.append(f"{key_takeaways}\n")

        return "\n".join(md)

    def check_if_processed(self, file_info: Dict) -> bool:
        """Check if a talk has already been processed (uses cached filenames)."""
        key = f"{file_info['talk_number']}_{file_info['session_code']}"
        return key in self._processed_keys

    def generate_output_filename(self, file_info: Dict, abstract: Optional[Dict] = None) -> str:
        """Generate output filename based on naming scheme."""
        speaker_clean = re.sub(r'[^a-zA-Z0-9]', '_', file_info['speaker_last_name'])
        title_words = re.sub(r'[^a-zA-Z0-9\s]', '', file_info['title']).split()[:5]
        title_slug = '_'.join(w.lower() for w in title_words)

        if self._naming == 'default':
            return f"talk_{file_info['talk_number']}_{file_info['session_code']}_{speaker_clean}_{title_slug}.md"
        elif self._naming == 'detailed':
            return f"talk_{file_info['session_code']}_{speaker_clean}_{title_slug}.md"
        elif self._naming == 'dated':
            date = 'undated'
            if abstract and abstract.get('day'):
                date = str(abstract['day']).replace(' ', '-').replace('/', '-')
            return f"{date}_{file_info['session_code']}_{speaker_clean}_{title_slug}.md"
        return f"talk_{file_info['talk_number']}_{file_info['session_code']}_{speaker_clean}_{title_slug}.md"

    def process_single_talk(self, pdf_path: Path, skip_existing: bool = True) -> bool:
        """Process a single talk PDF through the full pipeline."""
        logger.info(f"\n{'='*60}")
        logger.info(f"Processing: {pdf_path.name}")
        logger.info(f"{'='*60}")

        # File size check
        file_size_mb = pdf_path.stat().st_size / (1024 * 1024)
        if file_size_mb > 500:
            logger.error(f"  File too large: {file_size_mb:.1f}MB (max 500MB) — skipping")
            return False

        # Step 1: Parse filename
        file_info = self.parse_filename(pdf_path)
        logger.info(f"  Speaker: {file_info['speaker_formatted']}")
        logger.info(f"  Session: {file_info['session_code']}")
        logger.info(f"  Title: {file_info['title']}")

        if skip_existing and self.check_if_processed(file_info):
            logger.info(f"  Already processed — skipping")
            return True

        # Step 2: Match abstract
        abstract = self.match_abstract(file_info)
        if abstract:
            logger.info(f"  Abstract matched (score={abstract['match_score']}): "
                       f"{abstract['abstract_title'][:60]}...")
        else:
            logger.info(f"  No abstract match — will rely on slide content only")

        # Step 3: Convert slides to images
        logger.info("\n--- Converting slides to images ---")
        slide_images = self.convert_pdf_to_slide_images(pdf_path)
        if not slide_images:
            logger.error("Failed to convert PDF to images")
            self.log_processing(file_info, 0, 0, 0, 0, abstract is not None, None, "FAILED_CONVERSION")
            return False
        num_slides = len(slide_images)

        # Step 4: Extract slide content via Vision AI (with OCR pre-pass)
        logger.info(f"\n--- Extracting content from {num_slides} slides ---")
        slide_content, ocr_chars = self.extract_all_slides(slide_images)
        if not slide_content:
            logger.error("Failed to extract slide content")
            self.log_processing(file_info, num_slides, ocr_chars, 0, 0,
                              abstract is not None, None, "FAILED_EXTRACTION")
            return False
        logger.info(f"Total extracted content: {len(slide_content)} chars")

        # Step 5: Generate executive summary + key takeaways (single API call)
        logger.info("\n--- Generating executive summary and key takeaways ---")
        executive_summary, key_takeaways = self.generate_summary_and_takeaways(
            slide_content, file_info, abstract
        )

        # Step 6: Quality assessment
        logger.info("\n--- Quality Assessment ---")
        quality_scores = self.calculate_quality_score(
            slide_content=slide_content,
            executive_summary=executive_summary,
            key_takeaways=key_takeaways,
            num_slides=num_slides,
            ocr_chars=ocr_chars,
            abstract_matched=abstract is not None
        )

        # Step 7: Generate markdown
        logger.info("\n--- Generating markdown ---")
        markdown = self.generate_markdown(
            file_info=file_info,
            abstract=abstract,
            slide_content=slide_content,
            executive_summary=executive_summary,
            key_takeaways=key_takeaways,
            num_slides=num_slides,
            quality_scores=quality_scores
        )

        # Save output
        output_filename = self.generate_output_filename(file_info, abstract)
        output_path = self.output_dir / output_filename
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(markdown)
        self._existing_files.add(output_path.stem)

        # Log
        self.log_processing(file_info, num_slides, ocr_chars, len(slide_content),
                           len(executive_summary), abstract is not None, quality_scores, "SAVED")

        logger.info(f"\n{'='*60}")
        logger.info(f"SAVED: {output_path}")
        logger.info(f"  Slides: {num_slides}")
        logger.info(f"  Content: {len(slide_content)} chars")
        logger.info(f"  Summary: {len(executive_summary)} chars")
        logger.info(f"  Quality: {quality_scores['overall']}/10 — {quality_scores['assessment']}")
        logger.info(f"  Abstract: {'Yes' if abstract else 'No'}")
        logger.info(f"{'='*60}")

        return True

    def process_all_talks(self, skip_existing: bool = True):
        """Process all talks in the folder."""
        logger.info("=" * 70)
        logger.info("TALK PROCESSING PIPELINE")
        logger.info("=" * 70)
        logger.info(f"Talks folder: {self.talks_folder}")
        logger.info(f"Output directory: {self.output_dir}")
        logger.info(f"Recursive: {self._recursive}")
        logger.info(f"Naming: {self._naming}")
        logger.info(f"Skip existing: {skip_existing}")
        logger.info(f"Vision model: {self.VISION_MODEL}")
        logger.info(f"Batch size: {self.SLIDES_PER_BATCH} slides/call, max {self.MAX_CONCURRENT_BATCHES} concurrent")
        logger.info("=" * 70)

        pdf_files = self.list_talk_files()
        if not pdf_files:
            logger.error("No talk PDFs found")
            return

        success_count = 0
        fail_count = 0

        for i, pdf_path in enumerate(pdf_files):
            logger.info(f"\n[{i+1}/{len(pdf_files)}] {pdf_path.name}")
            try:
                if self.process_single_talk(pdf_path, skip_existing=skip_existing):
                    success_count += 1
                else:
                    fail_count += 1
            except KeyboardInterrupt:
                logger.info("\nProcessing interrupted by user")
                raise
            except Exception as e:
                logger.error(f"Unexpected error processing {pdf_path.name}: {e}")
                fail_count += 1

        logger.info(f"\n{'='*70}")
        logger.info(f"PIPELINE COMPLETE: {success_count} succeeded, {fail_count} failed "
                   f"out of {len(pdf_files)}")
        logger.info(f"Processing log: {self.processing_log_path}")
        logger.info(f"{'='*70}")


def main():
    parser = argparse.ArgumentParser(
        description="Process scientific talk PDFs into structured markdown using Vision AI"
    )
    parser.add_argument("--input", "--talks", dest="talks",
                       type=str, required=True,
                       help="Folder containing talk PDFs")
    parser.add_argument("--metadata", default=r"test_poster\AACR2026_Abstracts.xlsx",
                       help="Excel file with abstract metadata")
    parser.add_argument("--output", default="output_talks",
                       help="Output directory for markdown files")
    parser.add_argument("--single", default=None,
                       help="Process a single PDF file only")
    parser.add_argument("--recursive", action="store_true",
                       help="Recursively search subfolders for PDF files")
    parser.add_argument("--no-skip", action="store_true",
                       help="Reprocess files that already exist")
    parser.add_argument("--naming", choices=["default", "detailed", "dated"],
                       default="default",
                       help="Output filename scheme: default (talk_NUM_SESSION_Speaker_Title), "
                            "detailed (talk_SESSION_Speaker_Title), "
                            "dated (date_SESSION_Speaker_Title)")
    parser.add_argument("--verbose", action="store_true",
                       help="Enable verbose/debug logging")

    args = parser.parse_args()

    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    # Validate paths
    try:
        validate_path(args.talks, must_exist=True, allow_dir=True, allow_file=False)
        validate_output_path(args.output)
        if args.metadata:
            validate_path(args.metadata, must_exist=True, allow_file=True, allow_dir=False)
        if args.single:
            validate_path(args.single, must_exist=True, allow_file=True, allow_dir=False)
    except (ValueError, FileNotFoundError) as e:
        logger.error(f"Path validation failed: {e}")
        return

    pipeline = TalkPipeline(
        talks_folder=args.talks,
        metadata_excel=args.metadata,
        output_dir=args.output,
        recursive=args.recursive,
        naming=args.naming
    )

    if args.single:
        single_path = Path(args.single)
        pipeline.process_single_talk(single_path, skip_existing=not args.no_skip)
    else:
        pipeline.process_all_talks(skip_existing=not args.no_skip)


if __name__ == "__main__":
    main()
