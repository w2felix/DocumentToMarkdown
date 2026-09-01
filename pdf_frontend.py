"""pdf-inspector frontend adapter.

Optional replacement for the pdfplumber-based ``characterize_pdf`` step used
by ``paper_pipeline.PaperPipeline``. pdf-inspector is a Rust library that
parses layout (columns, tables), returns per-page markdown, and flags which
pages actually need OCR, letting us OCR selectively instead of the current
all-or-nothing fallback.

Contract: :func:`extract_pages` returns a :class:`FrontendResult` with a
``page_data`` list shaped exactly like the existing pipeline's page dicts
(minus ``classification``, which the caller applies) plus a bag of
document-level flags. If pdf-inspector is missing or throws, callers should
catch and fall through to the legacy pdfplumber path.

Used by default in ``paper_pipeline.characterize_pdf``. Set
``DOC2MD_USE_PDF_INSPECTOR=0`` to force the legacy pdfplumber path.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass
class FrontendResult:
    page_data: list[dict]
    doc_flags: dict[str, Any]
    method: str
    pages_needing_ocr: list[int] = field(default_factory=list)


def is_available() -> bool:
    try:
        import pdf_inspector  # noqa: F401
        return True
    except Exception:
        return False


def extract_pages(pdf_path: str | Path) -> FrontendResult:
    """Run pdf-inspector on ``pdf_path`` and shape the result.

    Raises :class:`ImportError` if pdf-inspector is not installed, or
    propagates any exception the library itself raises. Callers should
    catch and fall back to their legacy extractor.
    """
    import pdf_inspector as pi

    result = pi.extract_pages_markdown(str(pdf_path))

    page_data: list[dict] = []
    ocr_pages: list[int] = []
    for pm in result.pages:
        text = pm.markdown or ""
        page_data.append({
            "page_num": int(pm.page),
            "text": text,
            "char_count": len(text.strip()),
        })
        if bool(pm.needs_ocr):
            ocr_pages.append(int(pm.page))

    page_data.sort(key=lambda p: p["page_num"])
    ocr_pages = sorted(set(ocr_pages))

    doc_flags = {
        "is_complex_layout": bool(getattr(result, "is_complex", False)),
        "pages_with_columns": list(getattr(result, "pages_with_columns", []) or []),
        "pages_with_tables": list(getattr(result, "pages_with_tables", []) or []),
        "pages_needing_ocr": list(getattr(result, "pages_needing_ocr", []) or []),
        "ocr_reasons_by_page": {
            int(r.page): list(r.reasons)
            for r in (getattr(result, "ocr_reasons_by_page", []) or [])
        },
        "page_count": len(page_data),
    }

    method = "pdf_inspector+ocr" if ocr_pages else "pdf_inspector"
    return FrontendResult(
        page_data=page_data,
        doc_flags=doc_flags,
        method=method,
        pages_needing_ocr=ocr_pages,
    )
