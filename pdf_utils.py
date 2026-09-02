"""Shared PDF primitives for the DocumentToMarkdown pipelines.

Bytes-first, backend-selected:
- `pdf_inspector` for text + classification + per-page OCR-need flags.
- `fitz` (PyMuPDF) for rendering, page counts, page truncation.
- `pytesseract` for image OCR (post-render).
- `pdfplumber` remains the table extractor (called directly by pipelines
  that need it; not wrapped here).

Every function accepts either ``bytes`` or a filesystem path. Pipelines
that used to open their own ``fitz.open(...)`` loop should call the
primitive here instead so backend swaps happen in one place and error
handling stays consistent. Pipeline-specific policy (which DPI, which
pages, when to fall back to OCR) stays in the pipeline.
"""
from __future__ import annotations

import logging
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path
from typing import Any, Iterable, Optional, Union

logger = logging.getLogger(__name__)

PdfSource = Union[bytes, str, Path]


# ── dataclasses ───────────────────────────────────────────────────────────────

@dataclass
class PageText:
    page_num: int          # 0-indexed
    text: str
    char_count: int
    needs_ocr: bool = False


@dataclass
class Classification:
    pdf_type: str          # "text_based" | "scanned" | "image_based" | "mixed"
    page_count: int
    pages_needing_ocr: list[int] = field(default_factory=list)
    is_complex_layout: bool = False
    pages_with_columns: list[int] = field(default_factory=list)
    pages_with_tables: list[int] = field(default_factory=list)
    has_encoding_issues: bool = False


# ── input coercion ────────────────────────────────────────────────────────────

def _as_bytes(pdf: PdfSource) -> bytes:
    if isinstance(pdf, (bytes, bytearray)):
        return bytes(pdf)
    return Path(pdf).read_bytes()


def _as_path(pdf: PdfSource, tmp_dir: Optional[Path] = None) -> Path:
    """Ensure a filesystem path. Writes bytes to a temp file if needed.

    Callers that write a temp file are responsible for cleanup; this helper
    returns the path unchanged when one is provided so the common case is free.
    """
    if isinstance(pdf, (bytes, bytearray)):
        import tempfile
        d = tmp_dir or Path(tempfile.gettempdir())
        d.mkdir(parents=True, exist_ok=True)
        fd, name = tempfile.mkstemp(suffix=".pdf", dir=str(d))
        import os
        with os.fdopen(fd, "wb") as f:
            f.write(pdf)
        return Path(name)
    return Path(pdf)


# ── page count / classification ───────────────────────────────────────────────

def page_count(pdf: PdfSource) -> int:
    """Return the number of pages, or 0 on any failure."""
    try:
        import fitz
        data = _as_bytes(pdf)
        with fitz.open(stream=data, filetype="pdf") as doc:
            return len(doc)
    except Exception as e:
        logger.debug(f"page_count failed: {e}")
        return 0


def classify(pdf: PdfSource) -> Optional[Classification]:
    """Classify the PDF via pdf-inspector. Returns ``None`` if unavailable."""
    try:
        import pdf_inspector as _pi
    except ImportError:
        return None
    try:
        data = _as_bytes(pdf)
        r = _pi.classify_pdf_bytes(data)
        return Classification(
            pdf_type=str(getattr(r, "pdf_type", "unknown")),
            page_count=int(getattr(r, "page_count", 0)),
            pages_needing_ocr=list(getattr(r, "pages_needing_ocr", []) or []),
        )
    except Exception as e:
        logger.debug(f"classify failed: {e}")
        return None


# ── text extraction ───────────────────────────────────────────────────────────

def extract_text_pages(pdf: PdfSource,
                       max_pages: Optional[int] = None) -> list[PageText]:
    """Extract per-page text via pdf-inspector when available.

    Falls back silently to pdfplumber. Returns [] on both failing. This
    is a primitive: callers decide whether to run OCR on flagged pages.
    """
    data = _as_bytes(pdf)

    try:
        import pdf_inspector as _pi
        result = _pi.extract_pages_markdown_bytes(data)
        pages = list(result.pages)
        if max_pages is not None:
            pages = pages[:max_pages]
        return [
            PageText(
                page_num=int(p.page),
                text=p.markdown or "",
                char_count=len((p.markdown or "").strip()),
                needs_ocr=bool(p.needs_ocr),
            )
            for p in pages
        ]
    except ImportError:
        pass
    except Exception as e:
        logger.debug(f"pdf-inspector extract failed, falling back: {e}")

    try:
        import pdfplumber
        out: list[PageText] = []
        with pdfplumber.open(BytesIO(data)) as pdf_obj:
            pages = pdf_obj.pages
            if max_pages is not None:
                pages = pages[:max_pages]
            for i, p in enumerate(pages):
                text = p.extract_text() or ""
                out.append(PageText(i, text, len(text.strip())))
        return out
    except Exception as e:
        logger.debug(f"pdfplumber extract failed: {e}")
        return []


def extract_text_joined(pdf: PdfSource, max_pages: Optional[int] = None,
                        separator: str = "\n\n") -> str:
    """Convenience: extract per-page text and join into a single string."""
    return separator.join(p.text for p in extract_text_pages(pdf, max_pages) if p.text.strip())


# ── rendering ─────────────────────────────────────────────────────────────────

def render_pages(pdf: PdfSource,
                 page_indices: Optional[Iterable[int]] = None,
                 dpi: int = 200,
                 max_dimension: Optional[int] = None) -> dict[int, Any]:
    """Render the requested pages to PIL Images.

    Args:
        pdf: bytes or path.
        page_indices: 0-indexed pages to render. ``None`` means every page.
        dpi: rasterization DPI; 200 for OCR, 300+ for vision-AI-quality.
        max_dimension: if set, downscale so max(width, height) <= this.
            Vision APIs prefer <=1568 px; pass that here to enforce it.

    Returns ``{page_num: PIL.Image}``. Empty dict on failure. Callers
    close the images.
    """
    try:
        import fitz
        from PIL import Image
        Image.MAX_IMAGE_PIXELS = 300_000_000
    except ImportError as e:
        logger.error(f"render_pages requires fitz + PIL: {e}")
        return {}

    data = _as_bytes(pdf)
    out: dict[int, Any] = {}
    try:
        with fitz.open(stream=data, filetype="pdf") as doc:
            total = len(doc)
            if page_indices is None:
                indices = range(total)
            else:
                indices = [i for i in page_indices if 0 <= i < total]
            zoom = dpi / 72
            mat = fitz.Matrix(zoom, zoom)
            for i in indices:
                pix = doc[i].get_pixmap(matrix=mat)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                if max_dimension and max(img.size) > max_dimension:
                    scale = max_dimension / max(img.size)
                    new_size = (int(img.width * scale), int(img.height * scale))
                    img = img.resize(new_size, Image.LANCZOS)
                out[i] = img
    except Exception as e:
        logger.error(f"render_pages failed: {e}")
        for img in out.values():
            try:
                img.close()
            except Exception:
                pass
        return {}
    return out


# ── page truncation ───────────────────────────────────────────────────────────

def truncate(pdf: PdfSource, max_pages: int, output: Path) -> Optional[Path]:
    """Write a copy of ``pdf`` containing only the first ``max_pages`` pages.

    Returns the output path on success, ``None`` on failure. Uses fitz;
    caller is responsible for cleaning up the produced file.
    """
    try:
        import fitz
        data = _as_bytes(pdf)
        src = fitz.open(stream=data, filetype="pdf")
        dst = fitz.open()
        try:
            last = min(max_pages - 1, len(src) - 1)
            if last < 0:
                return None
            dst.insert_pdf(src, from_page=0, to_page=last)
            output.parent.mkdir(parents=True, exist_ok=True)
            dst.save(str(output))
            return output
        finally:
            dst.close()
            src.close()
    except Exception as e:
        logger.error(f"truncate failed: {e}")
        return None


# ── OCR ───────────────────────────────────────────────────────────────────────

def ocr_pages(pdf: PdfSource,
              page_indices: Optional[Iterable[int]] = None,
              dpi: int = 200,
              config: Optional[str] = None) -> list[PageText]:
    """OCR the requested pages via Tesseract, after rendering with fitz.

    Args:
        config: Tesseract CLI config string (e.g. ``"--psm 6 --oem 1"``).
            ``None`` uses Tesseract defaults.

    Returns a list of :class:`PageText` (marked ``needs_ocr=True`` for
    provenance). Empty list on any failure.
    """
    images = render_pages(pdf, page_indices=page_indices, dpi=dpi)
    if not images:
        return []
    try:
        import pytesseract
    except ImportError:
        logger.debug("pytesseract not installed, skipping OCR")
        for img in images.values():
            try:
                img.close()
            except Exception:
                pass
        return []

    results: list[PageText] = []
    for page_num, img in sorted(images.items()):
        try:
            kwargs = {"config": config} if config else {}
            text = pytesseract.image_to_string(img, **kwargs) or ""
        except Exception as e:
            logger.debug(f"ocr failed on page {page_num}: {e}")
            text = ""
        finally:
            try:
                img.close()
            except Exception:
                pass
        results.append(PageText(page_num, text, len(text.strip()), needs_ocr=True))
    return results
