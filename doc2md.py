"""DocumentToMarkdown public API.

Import this module for programmatic access to document extraction.
Each function takes a file path and returns structured data — no disk writes.

Usage:
    import sys
    sys.path.insert(0, "/path/to/DocumentToMarkdown")
    from doc2md import classify_document, extract_text, extract_figures
"""
from __future__ import annotations

import shutil
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional


@dataclass
class ExtractionResult:
    """Result of document text extraction."""
    text: str
    sections: dict
    metadata: dict
    extraction_method: str
    page_count: int
    markdown: str
    quality_score: dict


@dataclass
class FigureResult:
    """A single extracted figure from a PDF."""
    page_num: int
    figure_number: str
    figure_type: str
    description: str
    key_findings: list
    relevance: str
    statistical_notes: str = ""
    title: str = ""
    image_bytes: Optional[bytes] = None


def classify_document(pdf_path) -> str:
    """Classify a PDF document type.

    Returns one of: "paper", "poster", "talk", "patent", "presentation"
    """
    import fitz

    pdf_path = Path(pdf_path)
    doc = fitz.open(str(pdf_path))
    num_pages = len(doc)

    if num_pages == 0:
        doc.close()
        return 'paper'

    # Single page with large dimensions = poster
    if num_pages <= 3:
        page = doc[0]
        width, height = page.rect.width, page.rect.height
        if max(width, height) > 2000:
            doc.close()
            return 'poster'

    # Patent indicators in first page
    first_text = doc[0].get_text()[:2000].upper()
    patent_markers = ['INTERNATIONAL APPLICATION', 'WORLD INTELLECTUAL PROPERTY',
                      'PCT/', 'WO 20', 'EP 20', 'US 20', 'PATENT APPLICATION']
    if any(m in first_text for m in patent_markers):
        doc.close()
        return 'patent'

    # Talk detection: many pages with no extractable text (image-only slides)
    text_pages = 0
    for i in range(min(5, num_pages)):
        if len(doc[i].get_text().strip()) > 100:
            text_pages += 1

    doc.close()

    if text_pages == 0 and num_pages > 5:
        return 'talk'

    return 'paper'


def extract_text(pdf_path, *, max_vision_pages: int = 8,
                 budget: int = 50,
                 poster_options: dict | None = None) -> ExtractionResult:
    """Extract full text and structure from a PDF or PPTX.

    Routes to appropriate pipeline (paper/poster/talk/presentation) based on
    file type and classification. Returns structured data without writing to disk.

    poster_options: optional dict forwarded to the poster pipeline. Recognized
        keys:
          - enable_detailed_analysis (bool, default True): when False, skips
            Stage 2 per-figure detailed analysis. ~60% faster; figure blocks
            still get Stage 1 identification with Caption/Description/Key
            Findings but no Detailed Analysis.
        Ignored for non-poster inputs.
    """
    pdf_path = Path(pdf_path)

    if pdf_path.suffix.lower() == '.pptx':
        no_vision = (max_vision_pages == 0)
        return _extract_with_presentation_pipeline(pdf_path, no_vision)

    doc_type = classify_document(pdf_path)

    if doc_type == 'paper' or doc_type == 'patent':
        return _extract_with_paper_pipeline(pdf_path, max_vision_pages, budget)
    elif doc_type == 'poster':
        return _extract_with_poster_pipeline(pdf_path, budget, poster_options or {})
    elif doc_type == 'talk':
        return _extract_with_talk_pipeline(pdf_path, budget)
    else:
        return _extract_with_paper_pipeline(pdf_path, max_vision_pages, budget)


_VALID_RELEVANCE = {'HIGH', 'MEDIUM', 'LOW'}


def extract_figures(pdf_path, *, max_pages: int = 10,
                    include_images: bool = False,
                    dpi: int = 150) -> list:
    """Extract and analyze figure pages from a PDF.

    Identifies pages with little text, renders them, analyzes with Vision AI.
    Returns list of FigureResult with optional JPEG bytes.
    """
    from paper_pipeline import PaperPipeline

    pdf_path = Path(pdf_path)
    pipeline = PaperPipeline.from_file(pdf_path, max_vision_pages=max_pages)
    try:
        page_data, _ = pipeline.characterize_pdf(pdf_path)
        if not page_data:
            return []

        figure_pages = pipeline.identify_figure_pages(page_data)
        if not figure_pages:
            return []

        full_text = pipeline.assemble_full_text(page_data)
        raw_figures = pipeline.analyze_figures_batch(
            pdf_path, figure_pages, full_text,
            include_images=include_images, dpi=dpi,
        )
    finally:
        tmpdir = getattr(pipeline, '_tmpdir', None)
        if tmpdir:
            shutil.rmtree(tmpdir, ignore_errors=True)

    results = []
    for fig in raw_figures:
        rel = fig.get('relevance', 'MEDIUM').upper()
        results.append(FigureResult(
            page_num=fig.get('page_num', 0),
            figure_number=str(fig.get('figure_number', 'Unknown')),
            figure_type=fig.get('figure_type', 'unknown'),
            description=fig.get('description', ''),
            key_findings=fig.get('key_findings', []),
            relevance=rel if rel in _VALID_RELEVANCE else 'MEDIUM',
            statistical_notes=fig.get('statistical_notes', ''),
            title=fig.get('title', ''),
            image_bytes=fig.get('image_bytes'),
        ))

    return results


# ── Internal helpers ──────────────────────────────────────────────────────────

def _extract_with_paper_pipeline(pdf_path: Path, max_vision_pages: int,
                                  budget: int) -> ExtractionResult:
    from paper_pipeline import PaperPipeline

    pipeline = PaperPipeline.from_file(
        pdf_path, max_vision_pages=max_vision_pages, budget=budget)
    pipeline._api_call_count = 0

    try:
        page_data, extraction_method = pipeline.characterize_pdf(pdf_path)
        if not page_data:
            return ExtractionResult(
                text="", sections={}, metadata={}, extraction_method="error",
                page_count=0, markdown="", quality_score={})

        full_text = pipeline.assemble_full_text(page_data)
        page_count = len(page_data)

        # Metadata
        filename_meta = pipeline.extract_metadata_from_filename(pdf_path)
        text_meta = pipeline.extract_metadata_from_text(full_text, page_data)
        no_vision = (max_vision_pages == 0)
        vision_meta = {} if no_vision else pipeline.extract_metadata_vision(pdf_path, page_data)
        metadata = pipeline.merge_metadata(vision_meta, text_meta, filename_meta)

        # Sections
        sections = pipeline.split_into_sections(full_text)
        if metadata.get('abstract') and 'abstract' not in sections:
            sections['abstract'] = metadata['abstract']
        if 'abstract' not in sections:
            abstract = pipeline._extract_abstract_from_first_page(page_data)
            if abstract:
                sections['abstract'] = abstract
        if len(sections) < 3 and not no_vision:
            ai_sections = pipeline.split_sections_with_ai(full_text, metadata)
            if ai_sections:
                for k, v in ai_sections.items():
                    if k not in sections or len(v) > len(sections.get(k, '')):
                        sections[k] = v

        # Figures + tables
        figures = []
        if not no_vision:
            figure_pages = pipeline.identify_figure_pages(page_data)
            if figure_pages:
                figures = pipeline.analyze_figures_batch(pdf_path, figure_pages, full_text)
        tables = pipeline.extract_tables(pdf_path, page_data)

        # Summary + quality
        summary = "" if no_vision else pipeline.generate_executive_summary(sections, metadata)
        quality = pipeline.calculate_quality_score(full_text, metadata, sections, figures, tables)

        method_str = f"{extraction_method}_with_vision" if not no_vision else f"{extraction_method}_text_only"

        markdown = pipeline.generate_markdown(
            metadata, sections, figures, tables, summary,
            quality, method_str, page_count, full_text)

        return ExtractionResult(
            text=full_text,
            sections=sections,
            metadata=metadata,
            extraction_method=method_str,
            page_count=page_count,
            markdown=markdown,
            quality_score=quality,
        )
    finally:
        tmpdir = getattr(pipeline, '_tmpdir', None)
        if tmpdir:
            shutil.rmtree(tmpdir, ignore_errors=True)


def _extract_with_poster_pipeline(pdf_path: Path, budget: int,
                                  options: dict | None = None) -> ExtractionResult:
    from poster_pipeline import PosterPipeline

    options = options or {}
    enable_detailed = bool(options.get("enable_detailed_analysis", True))

    with tempfile.TemporaryDirectory(prefix="doc2md_poster_") as tmpdir:
        pipeline = PosterPipeline(
            sharepoint_folder=str(pdf_path.parent),
            output_dir=tmpdir,
            naming='default',
        )
        success = pipeline.process_single_poster(
            pdf_path,
            skip_existing=False,
            enable_detailed_analysis=enable_detailed,
        )
        if success:
            output_files = list(Path(tmpdir).glob("*.md"))
            if output_files:
                markdown = output_files[0].read_text(encoding='utf-8')
                return ExtractionResult(
                    text=markdown,
                    sections={},
                    metadata={},
                    extraction_method="poster_pipeline",
                    page_count=1,
                    markdown=markdown,
                    quality_score={},
                )

    return ExtractionResult(
        text="", sections={}, metadata={}, extraction_method="error",
        page_count=0, markdown="", quality_score={})


def _extract_with_talk_pipeline(pdf_path: Path, budget: int) -> ExtractionResult:
    """Extract talk content. Falls back to paper pipeline if TalkPipeline
    can't be instantiated (requires metadata_excel)."""
    try:
        from talk_pipeline import TalkPipeline

        with tempfile.TemporaryDirectory(prefix="doc2md_talk_") as tmpdir:
            # TalkPipeline needs a metadata_excel; create an empty one as stub
            stub_excel = Path(tmpdir) / "stub.xlsx"
            try:
                import openpyxl
                wb = openpyxl.Workbook()
                wb.save(str(stub_excel))
            except Exception:
                return _extract_with_paper_pipeline(pdf_path, 0, budget)

            pipeline = TalkPipeline(
                talks_folder=str(pdf_path.parent),
                metadata_excel=str(stub_excel),
                output_dir=tmpdir,
                naming='default',
            )
            success = pipeline.process_single_talk(pdf_path, skip_existing=False)
            if success:
                output_files = list(Path(tmpdir).glob("*.md"))
                if output_files:
                    markdown = output_files[0].read_text(encoding='utf-8')
                    return ExtractionResult(
                        text=markdown,
                        sections={},
                        metadata={},
                        extraction_method="talk_pipeline",
                        page_count=0,
                        markdown=markdown,
                        quality_score={},
                    )
    except Exception:
        pass

    return _extract_with_paper_pipeline(pdf_path, 0, budget)


def _extract_with_presentation_pipeline(pptx_path: Path, no_vision: bool) -> ExtractionResult:
    from presentation_pipeline import PresentationPipeline

    with tempfile.TemporaryDirectory(prefix="doc2md_pptx_") as tmpdir:
        pipeline = PresentationPipeline(
            input_folder=str(pptx_path.parent),
            output_dir=tmpdir,
            naming='default',
            no_vision=no_vision,
        )
        success = pipeline.process_single_presentation(pptx_path, skip_existing=False)
        if success:
            output_files = list(Path(tmpdir).glob("*.md"))
            if output_files:
                markdown = output_files[0].read_text(encoding='utf-8')
                return ExtractionResult(
                    text=markdown,
                    sections={},
                    metadata={},
                    extraction_method="presentation_pipeline",
                    page_count=0,
                    markdown=markdown,
                    quality_score={},
                )

    return ExtractionResult(
        text="", sections={}, metadata={}, extraction_method="error",
        page_count=0, markdown="", quality_score={})
