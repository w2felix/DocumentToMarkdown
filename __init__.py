"""DocumentToMarkdown: AI-powered document processing for scientific/pharma PDFs.

Public API (programmatic, no file I/O):
    from doc2md import classify_document, extract_text, extract_figures

Pipeline classes (CLI / batch processing):
    from paper_pipeline import PaperPipeline
    from ci_pipeline import CIPipeline
    from poster_pipeline import PosterPipeline
    from talk_pipeline import TalkPipeline
    from presentation_pipeline import PresentationPipeline
"""
from doc2md import (
    classify_document,
    extract_text,
    extract_figures,
    ExtractionResult,
    FigureResult,
)
