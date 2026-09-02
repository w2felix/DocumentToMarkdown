"""Unit tests for crossref_meta.

Run: python -m pytest tests/ -v
"""
from __future__ import annotations

import json
import sys
import time
from pathlib import Path
from unittest.mock import patch

import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))

import crossref_meta as cr  # noqa: E402


# ── _clean_doi ────────────────────────────────────────────────────────────────

class TestCleanDoi:
    def test_plain(self):
        assert cr._clean_doi("10.1038/s41568-019-0133-9") == "10.1038/s41568-019-0133-9"

    def test_trailing_period(self):
        assert cr._clean_doi("10.1038/nature12373.") == "10.1038/nature12373"

    def test_trailing_paren(self):
        assert cr._clean_doi("10.1038/nature12373)") == "10.1038/nature12373"

    def test_markdown_link_junk(self):
        """pdf-inspector emits markdown-link syntax around DOIs."""
        raw = "10.1038/s41568-019-0133-9](https://doi.org/10.1038/s41568-019-0133-9)"
        assert cr._clean_doi(raw) == "10.1038/s41568-019-0133-9"

    def test_unicode_replacement_suffix(self):
        assert cr._clean_doi("10.1186/s12935-022-02660-5�") == "10.1186/s12935-022-02660-5"

    def test_bold_marker_suffix(self):
        assert cr._clean_doi("10.1038/s41591-026-04521-4**") == "10.1038/s41591-026-04521-4"

    def test_pipe_junk_rejected(self):
        """Not a real DOI: no alphanumeric after the slash."""
        assert cr._clean_doi("10.1158/|") is None

    def test_empty(self):
        assert cr._clean_doi("") is None
        assert cr._clean_doi(None) is None

    def test_no_doi_in_string(self):
        assert cr._clean_doi("just some text") is None


# ── sanitize_title ────────────────────────────────────────────────────────────

class TestSanitizeTitle:
    def test_leading_hash(self):
        assert cr.sanitize_title("# The Great Paper") == "The Great Paper"

    def test_bold_wrapping(self):
        assert cr.sanitize_title("**Nature Medicine**") == "Nature Medicine"

    def test_replacement_char(self):
        assert cr.sanitize_title("antibody�drug conjugates") == "antibody drug conjugates"

    def test_collapses_whitespace(self):
        assert cr.sanitize_title("A   long\ttitle\n\nhere") == "A long title here"

    def test_empty(self):
        assert cr.sanitize_title("") == ""
        assert cr.sanitize_title(None) == ""


# ── _tokens (Jaccard base) ────────────────────────────────────────────────────

class TestTokens:
    def test_lowercases(self):
        assert cr._tokens("The Great Paper") == {"the", "great", "paper"}

    def test_drops_short(self):
        assert "of" not in cr._tokens("Analysis of RNA")
        assert "rna" in cr._tokens("Analysis of RNA")

    def test_alphanumeric_only(self):
        toks = cr._tokens("SARS-CoV-2 spike")
        assert "sars" in toks and "cov" in toks


# ── _normalize_authors / _shape_work ─────────────────────────────────────────

class TestShapeWork:
    def test_basic(self):
        msg = {
            "DOI": "10.1038/Test",
            "title": ["Some Paper"],
            "author": [
                {"given": "Alice", "family": "Smith", "sequence": "first"},
                {"given": "Bob", "family": "Jones", "sequence": "additional"},
            ],
            "issued": {"date-parts": [[2024, 3, 15]]},
            "container-title": ["Journal of X"],
            "volume": "12",
            "page": "1-10",
        }
        out = cr._shape_work(msg)
        assert out["doi"] == "10.1038/test"
        assert out["title"] == "Some Paper"
        assert out["year"] == "2024"
        assert out["journal"] == "Journal of X"
        assert len(out["authors"]) == 2
        assert out["authors"][0]["family"] == "Smith"
        assert out["authors"][0]["sequence"] == "first"

    def test_missing_author_names_skipped(self):
        msg = {"DOI": "10.1234/x", "author": [{"given": "", "family": ""},
                                            {"family": "Ok"}]}
        out = cr._shape_work(msg)
        assert len(out["authors"]) == 1
        assert out["authors"][0]["family"] == "Ok"

    def test_no_authors(self):
        out = cr._shape_work({"DOI": "10.1234/x"})
        assert out["authors"] == []
        assert out["title"] == ""


# ── reconcile ─────────────────────────────────────────────────────────────────

class TestReconcile:
    def test_none_returns_empty(self):
        assert cr.reconcile({}, {}, None) == {}

    def test_empty_dict_returns_empty(self):
        assert cr.reconcile({}, {}, {}) == {}

    def test_authors_ordered_by_sequence(self):
        record = {
            "title": "T", "year": "2020",
            "authors": [
                {"name": "B Later", "family": "Later", "sequence": "additional"},
                {"name": "A First", "family": "First", "sequence": "first"},
            ],
        }
        out = cr.reconcile({}, {}, record)
        assert out["authors"] == ["A First", "B Later"]
        assert out["first_author"] == "First"

    def test_vision_wins_for_non_biblio(self):
        record = {"title": "T", "authors": []}
        vision = {"abstract": "V abstract", "keywords": ["k"]}
        text = {"abstract": "T abstract"}
        out = cr.reconcile(text, vision, record)
        assert out["abstract"] == "V abstract"
        assert out["keywords"] == ["k"]

    def test_text_used_when_vision_missing(self):
        record = {"title": "T"}
        out = cr.reconcile({"abstract": "from text"}, {}, record)
        assert out["abstract"] == "from text"


# ── fetch_by_doi (with mocked HTTP) ──────────────────────────────────────────

class _FakeResp:
    def __init__(self, status_code: int, payload: dict):
        self.status_code = status_code
        self._payload = payload

    def json(self):
        return self._payload


class TestFetchByDoi:
    def test_success(self, tmp_path, monkeypatch):
        monkeypatch.setattr(cr, "_CACHE_DIR", tmp_path / "cache")
        payload = {"message": {
            "DOI": "10.1038/nature12373",
            "title": ["Paper"],
            "author": [{"given": "A", "family": "B", "sequence": "first"}],
            "issued": {"date-parts": [[2020]]},
            "container-title": ["Nature"],
        }}
        with patch("requests.get", return_value=_FakeResp(200, payload)):
            out = cr.fetch_by_doi("10.1038/nature12373")
        assert out and out["title"] == "Paper"
        assert out["year"] == "2020"
        # cache write happened — one file, no subdirectories
        cached = list((tmp_path / "cache").glob("*.json"))
        assert len(cached) == 1, f"expected 1 cache file, got {cached}"
        assert "/" not in cached[0].name and "\\" not in cached[0].name
        # second call should hit cache — no HTTP
        with patch("requests.get", side_effect=AssertionError("should not fire")):
            again = cr.fetch_by_doi("10.1038/nature12373")
        assert again == out

    def test_404(self, tmp_path, monkeypatch):
        monkeypatch.setattr(cr, "_CACHE_DIR", tmp_path / "cache")
        with patch("requests.get", return_value=_FakeResp(404, {})):
            assert cr.fetch_by_doi("10.9999/nonexistent") is None
        # 404 responses are NOT cached
        assert not (tmp_path / "cache").exists() or \
               not list((tmp_path / "cache").glob("*.json"))

    def test_bad_doi_short_circuits(self):
        with patch("requests.get", side_effect=AssertionError("should not fire")):
            assert cr.fetch_by_doi("not-a-doi") is None
            assert cr.fetch_by_doi("") is None

    def test_network_error(self, tmp_path, monkeypatch):
        monkeypatch.setattr(cr, "_CACHE_DIR", tmp_path / "cache")
        import requests
        with patch("requests.get", side_effect=requests.ConnectionError("no net")):
            assert cr.fetch_by_doi("10.1234/x") is None

    def test_cache_ttl_expiry(self, tmp_path, monkeypatch):
        monkeypatch.setattr(cr, "_CACHE_DIR", tmp_path / "cache")
        monkeypatch.setattr(cr, "_CACHE_TTL_S", 1)  # expire in 1s
        payload = {"message": {"DOI": "10.1234/x", "title": ["Old"]}}
        with patch("requests.get", return_value=_FakeResp(200, payload)):
            cr.fetch_by_doi("10.1234/x")
        # backdate cache mtime by 10s
        p = next((tmp_path / "cache").glob("*.json"))
        old = time.time() - 10
        import os as _os
        _os.utime(p, (old, old))
        # next call should refetch
        payload2 = {"message": {"DOI": "10.1234/x", "title": ["Fresh"]}}
        with patch("requests.get", return_value=_FakeResp(200, payload2)) as mock:
            r = cr.fetch_by_doi("10.1234/x")
            assert mock.called
        assert r["title"] == "Fresh"


# ── resolve_doi_by_title (with mocked HTTP) ──────────────────────────────────

class TestResolveDoiByTitle:
    def test_min_length_skip(self):
        with patch("requests.get", side_effect=AssertionError("should not fire")):
            assert cr.resolve_doi_by_title("short", None) is None
            assert cr.resolve_doi_by_title("", None) is None

    def test_close_match_returns_doi(self):
        query = "Molecular subtypes small cell lung cancer synthesis"
        payload = {"message": {"items": [{
            "DOI": "10.1038/nrc-2019-042",
            "title": ["Molecular subtypes small cell lung cancer synthesis mouse"],
        }]}}
        with patch("requests.get", return_value=_FakeResp(200, payload)):
            doi = cr.resolve_doi_by_title(query, None)
        assert doi == "10.1038/nrc-2019-042"

    def test_low_jaccard_returns_none(self):
        query = "Foo bar baz quux something interesting"
        payload = {"message": {"items": [{
            "DOI": "10.1234/wrong",
            "title": ["Something totally unrelated about proteins"],
        }]}}
        with patch("requests.get", return_value=_FakeResp(200, payload)):
            assert cr.resolve_doi_by_title(query, None) is None

    def test_sanitizes_before_search(self):
        """`# Markdown # title` should be cleaned before the Jaccard check."""
        payload = {"message": {"items": [{
            "DOI": "10.1234/ok",
            "title": ["Clean title here about proteins in cancer research"],
        }]}}
        query = "# Clean title here about proteins in cancer research # #"
        with patch("requests.get", return_value=_FakeResp(200, payload)):
            assert cr.resolve_doi_by_title(query, None) == "10.1234/ok"
