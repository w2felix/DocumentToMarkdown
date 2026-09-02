"""Unit tests for pdf_utils.

Uses a real, small PDF from the repo (`test_poster/160.pdf`) as a fixture
and monkeypatches the backends for failure-mode coverage.

Run: python -m pytest tests/test_pdf_utils.py -v
"""
from __future__ import annotations

import sys
from pathlib import Path
from unittest.mock import patch

import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))

import pdf_utils as U  # noqa: E402


@pytest.fixture(scope="module")
def sample_pdf() -> Path:
    """Small real PDF that ships with the repo (a poster)."""
    p = REPO_ROOT / "test_poster" / "160.pdf"
    if not p.exists():
        pytest.skip(f"fixture PDF missing: {p}")
    return p


@pytest.fixture(scope="module")
def sample_bytes(sample_pdf: Path) -> bytes:
    return sample_pdf.read_bytes()


# ── _as_bytes / _as_path ──────────────────────────────────────────────────────

class TestInputCoercion:
    def test_as_bytes_passthrough(self):
        assert U._as_bytes(b"raw") == b"raw"

    def test_as_bytes_from_path(self, sample_pdf, sample_bytes):
        assert U._as_bytes(sample_pdf) == sample_bytes

    def test_as_path_passthrough(self, sample_pdf):
        assert U._as_path(sample_pdf) == sample_pdf

    def test_as_path_writes_temp_for_bytes(self, sample_bytes, tmp_path):
        p = U._as_path(sample_bytes, tmp_dir=tmp_path)
        try:
            assert p.exists()
            assert p.read_bytes() == sample_bytes
        finally:
            p.unlink(missing_ok=True)


# ── page_count ────────────────────────────────────────────────────────────────

class TestPageCount:
    def test_from_path(self, sample_pdf):
        assert U.page_count(sample_pdf) >= 1

    def test_from_bytes(self, sample_bytes):
        assert U.page_count(sample_bytes) >= 1

    def test_bytes_and_path_agree(self, sample_pdf, sample_bytes):
        assert U.page_count(sample_pdf) == U.page_count(sample_bytes)

    def test_invalid_input_returns_zero(self):
        assert U.page_count(b"not a pdf") == 0
        assert U.page_count(b"") == 0

    def test_missing_file_returns_zero(self, tmp_path):
        assert U.page_count(tmp_path / "nope.pdf") == 0


# ── classify ──────────────────────────────────────────────────────────────────

class TestClassify:
    def test_returns_classification(self, sample_pdf):
        c = U.classify(sample_pdf)
        assert c is not None
        assert c.pdf_type in {"text_based", "scanned", "image_based", "mixed", "unknown"}
        assert c.page_count >= 1
        assert isinstance(c.pages_needing_ocr, list)

    def test_pdf_inspector_missing_returns_none(self, sample_pdf):
        # Simulate pdf_inspector import failure
        import builtins
        real_import = builtins.__import__
        def fake_import(name, *a, **kw):
            if name == "pdf_inspector":
                raise ImportError("mocked")
            return real_import(name, *a, **kw)
        with patch("builtins.__import__", side_effect=fake_import):
            assert U.classify(sample_pdf) is None


# ── extract_text_pages / extract_text_joined ─────────────────────────────────

class TestExtractText:
    def test_returns_pages(self, sample_pdf):
        pages = U.extract_text_pages(sample_pdf)
        assert isinstance(pages, list)
        assert all(hasattr(p, "page_num") and hasattr(p, "text") for p in pages)

    def test_max_pages_honored(self, sample_pdf):
        pages = U.extract_text_pages(sample_pdf, max_pages=1)
        assert len(pages) <= 1

    def test_joined_matches_pages(self, sample_pdf):
        pages = U.extract_text_pages(sample_pdf, max_pages=2)
        joined = U.extract_text_joined(sample_pdf, max_pages=2)
        # joined is the non-empty page texts joined
        expected_parts = [p.text for p in pages if p.text.strip()]
        assert joined == "\n\n".join(expected_parts)

    def test_invalid_input_returns_empty(self):
        assert U.extract_text_pages(b"not a pdf") == []
        assert U.extract_text_joined(b"not a pdf") == ""


# ── render_pages ──────────────────────────────────────────────────────────────

class TestRenderPages:
    def test_render_single_page(self, sample_pdf):
        images = U.render_pages(sample_pdf, page_indices=[0], dpi=100)
        assert 0 in images
        img = images[0]
        try:
            assert img.width > 0 and img.height > 0
        finally:
            img.close()

    def test_render_all_pages(self, sample_pdf):
        images = U.render_pages(sample_pdf, dpi=72)
        try:
            assert len(images) == U.page_count(sample_pdf)
        finally:
            for img in images.values():
                img.close()

    def test_max_dimension_downscales(self, sample_pdf):
        images = U.render_pages(sample_pdf, page_indices=[0], dpi=300, max_dimension=800)
        try:
            img = images[0]
            assert max(img.size) <= 800
        finally:
            for img in images.values():
                img.close()

    def test_out_of_range_indices_dropped(self, sample_pdf):
        n = U.page_count(sample_pdf)
        images = U.render_pages(sample_pdf, page_indices=[0, n + 100], dpi=72)
        try:
            assert set(images.keys()) == {0}
        finally:
            for img in images.values():
                img.close()

    def test_invalid_input_returns_empty(self):
        assert U.render_pages(b"not a pdf") == {}


# ── truncate ──────────────────────────────────────────────────────────────────

class TestTruncate:
    def test_produces_smaller_file(self, sample_pdf, tmp_path):
        # Only meaningful if the sample has 2+ pages; skip if not.
        if U.page_count(sample_pdf) < 2:
            pytest.skip("sample has <2 pages")
        out = tmp_path / "cut.pdf"
        result = U.truncate(sample_pdf, max_pages=1, output=out)
        assert result == out
        assert out.exists()
        assert U.page_count(out) == 1

    def test_bad_input_returns_none(self, tmp_path):
        assert U.truncate(b"not a pdf", 1, tmp_path / "out.pdf") is None


# ── ocr_pages ─────────────────────────────────────────────────────────────────

class TestOcrPages:
    def test_pytesseract_missing_returns_empty(self, sample_pdf):
        import builtins
        real_import = builtins.__import__
        def fake_import(name, *a, **kw):
            if name == "pytesseract":
                raise ImportError("mocked")
            return real_import(name, *a, **kw)
        with patch("builtins.__import__", side_effect=fake_import):
            assert U.ocr_pages(sample_pdf, page_indices=[0]) == []

    def test_config_forwarded(self, sample_pdf):
        # Mock pytesseract.image_to_string to observe the config kwarg.
        seen_kwargs: dict = {}
        def fake_ocr(image, **kwargs):
            seen_kwargs.update(kwargs)
            return "sentinel-text"
        with patch("pytesseract.image_to_string", side_effect=fake_ocr):
            results = U.ocr_pages(
                sample_pdf, page_indices=[0], dpi=100,
                config="--psm 6 --oem 1",
            )
        assert seen_kwargs.get("config") == "--psm 6 --oem 1"
        assert results and results[0].text == "sentinel-text"
        assert results[0].needs_ocr is True

    def test_default_config_omits_kwarg(self, sample_pdf):
        seen_kwargs: dict = {}
        def fake_ocr(image, **kwargs):
            seen_kwargs.update(kwargs)
            return ""
        with patch("pytesseract.image_to_string", side_effect=fake_ocr):
            U.ocr_pages(sample_pdf, page_indices=[0], dpi=100)
        assert "config" not in seen_kwargs
