"""Crossref-based metadata reconciliation.

When a DOI is available (from text-regex extraction or pdf-inspector's
title-based search), Crossref returns canonical title, authors, year,
journal, and volume. This is deterministic and defeats vision-based
author hallucination (correct surname, invented given-name).

Public API:
- :func:`fetch_by_doi(doi)` -> canonical dict or ``None``
- :func:`resolve_doi_by_title(title, first_author=None)` -> DOI or ``None``
- :func:`reconcile(text_meta, vision_meta, crossref_meta)` -> dict

Responses are cached under ``~/.cache/doc2md/crossref/{doi_slug}.json``
with a 7-day TTL. All network calls are best-effort: on any failure
the caller should silently fall back to the existing merge order.
"""
from __future__ import annotations

import json
import logging
import re
import time
from pathlib import Path
from typing import Any, Optional

logger = logging.getLogger(__name__)

_CROSSREF_WORKS = "https://api.crossref.org/works"
_CACHE_DIR = Path.home() / ".cache" / "doc2md" / "crossref"
_CACHE_TTL_S = 7 * 24 * 3600
_HTTP_TIMEOUT_S = 8.0
_USER_AGENT = "DocumentToMarkdown/1.0 (mailto:felix.geist@merckgroup.com)"

_DOI_CLEAN = re.compile(r"[^A-Za-z0-9._/\-]")


def _cache_path(doi: str) -> Path:
    slug = _DOI_CLEAN.sub("_", doi.lower()).strip("_")
    return _CACHE_DIR / f"{slug}.json"


def _read_cache(path: Path) -> Optional[dict]:
    try:
        if not path.is_file():
            return None
        if time.time() - path.stat().st_mtime > _CACHE_TTL_S:
            return None
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None


def _write_cache(path: Path, data: dict) -> None:
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(data), encoding="utf-8")
    except Exception as e:
        logger.debug(f"Crossref cache write failed: {e}")


def _normalize_authors(msg: dict) -> list[dict]:
    """Extract [{given, family, sequence}] from a Crossref work record."""
    out = []
    for a in msg.get("author") or []:
        given = (a.get("given") or "").strip()
        family = (a.get("family") or "").strip()
        if not family and not given:
            continue
        out.append({
            "given": given,
            "family": family,
            "sequence": a.get("sequence") or "additional",
            "name": (f"{given} {family}").strip(),
        })
    return out


def _shape_work(msg: dict) -> dict:
    title = ""
    tlist = msg.get("title") or []
    if tlist:
        title = tlist[0]
    year = None
    for k in ("published-print", "published-online", "issued", "created"):
        v = (msg.get(k) or {}).get("date-parts")
        if v and v[0]:
            year = str(v[0][0])
            break
    journal = ""
    cnt = msg.get("container-title") or []
    if cnt:
        journal = cnt[0]
    return {
        "doi": (msg.get("DOI") or "").lower(),
        "title": title.strip(),
        "authors": _normalize_authors(msg),
        "year": year,
        "journal": journal,
        "volume": msg.get("volume"),
        "issue": msg.get("issue"),
        "pages": msg.get("page"),
        "publisher": msg.get("publisher"),
        "type": msg.get("type"),
        "url": msg.get("URL"),
    }


_DOI_STRICT = re.compile(r"10\.\d{4,}/[A-Za-z0-9][A-Za-z0-9\-._/:]*")


def _clean_doi(doi: str) -> Optional[str]:
    """Extract a strict DOI substring, dropping markdown/encoding cruft."""
    if not doi:
        return None
    doi = doi.strip()
    m = _DOI_STRICT.search(doi)
    return m.group(0).rstrip(").,;") if m else None


def fetch_by_doi(doi: str) -> Optional[dict]:
    """Fetch and cache a Crossref work record. Returns a shaped dict or ``None``."""
    doi_clean = _clean_doi(doi)
    if not doi_clean:
        return None
    cache = _cache_path(doi_clean)
    hit = _read_cache(cache)
    if hit is not None:
        return hit
    try:
        import requests
    except ImportError:
        logger.debug("requests not installed, skipping Crossref lookup")
        return None
    try:
        resp = requests.get(
            f"{_CROSSREF_WORKS}/{doi_clean}",
            headers={"User-Agent": _USER_AGENT, "Accept": "application/json"},
            timeout=_HTTP_TIMEOUT_S,
        )
        if resp.status_code != 200:
            logger.debug(f"Crossref {resp.status_code} for DOI {doi_clean}")
            return None
        msg = (resp.json() or {}).get("message") or {}
        if not msg:
            return None
        shaped = _shape_work(msg)
        _write_cache(cache, shaped)
        return shaped
    except Exception as e:
        logger.debug(f"Crossref lookup failed for {doi_clean}: {e}")
        return None


def resolve_doi_by_title(title: str,
                         first_author: Optional[str] = None) -> Optional[str]:
    """Best-effort title -> DOI resolution when no DOI was extracted.

    Uses Crossref's bibliographic search. Requires the top result to be a
    close string match to the query title (>=0.75 Jaccard on word sets)
    to avoid returning a wrong paper.
    """
    if not title or len(title) < 15:
        return None
    try:
        import requests
    except ImportError:
        return None

    try:
        params: dict[str, Any] = {"query.bibliographic": title[:300], "rows": 3}
        if first_author:
            params["query.author"] = first_author
        resp = requests.get(
            _CROSSREF_WORKS,
            params=params,
            headers={"User-Agent": _USER_AGENT, "Accept": "application/json"},
            timeout=_HTTP_TIMEOUT_S,
        )
        if resp.status_code != 200:
            return None
        items = ((resp.json() or {}).get("message") or {}).get("items") or []
    except Exception as e:
        logger.debug(f"Crossref title search failed: {e}")
        return None

    query_tokens = _tokens(title)
    if not query_tokens:
        return None
    for item in items:
        cand_titles = item.get("title") or []
        if not cand_titles:
            continue
        cand_tokens = _tokens(cand_titles[0])
        if not cand_tokens:
            continue
        jaccard = len(query_tokens & cand_tokens) / len(query_tokens | cand_tokens)
        if jaccard >= 0.75:
            doi = _clean_doi(item.get("DOI") or "")
            if doi:
                return doi
    return None


def _tokens(text: str) -> set[str]:
    return set(re.findall(r"[a-z0-9]{3,}", (text or "").lower()))


def reconcile(text_meta: dict, vision_meta: dict,
              crossref_meta: Optional[dict]) -> dict:
    """Merge metadata with Crossref as the authority when present.

    Priority when Crossref record exists:
      title, authors, year, journal, volume, pages, doi: Crossref
      abstract, keywords, correspondence_author, affiliations: vision > text
    When no Crossref record: preserves the existing vision > text > filename
    behavior (caller is expected to fall through to the legacy merger).
    """
    if not crossref_meta:
        return {}

    out: dict[str, Any] = {}
    if crossref_meta.get("title"):
        out["title"] = crossref_meta["title"]
    authors = crossref_meta.get("authors") or []
    if authors:
        # sort by Crossref-declared sequence: "first" before "additional"
        seq_rank = {"first": 0, "additional": 1}
        ordered = sorted(authors,
                         key=lambda a: seq_rank.get(a.get("sequence"), 2))
        out["authors"] = [a["name"] for a in ordered if a.get("name")]
        first = next((a for a in ordered if a.get("sequence") == "first"), ordered[0])
        if first.get("family"):
            out["first_author"] = first["family"]
    if crossref_meta.get("year"):
        out["year"] = crossref_meta["year"]
    if crossref_meta.get("journal"):
        out["journal"] = crossref_meta["journal"]
    if crossref_meta.get("volume"):
        out["volume"] = crossref_meta["volume"]
    if crossref_meta.get("pages"):
        out["pages"] = crossref_meta["pages"]
    if crossref_meta.get("doi"):
        out["doi"] = crossref_meta["doi"]

    # Non-Crossref fields: keep vision/text preferences
    for k in ("abstract", "keywords", "correspondence_author", "affiliations"):
        v = vision_meta.get(k) or text_meta.get(k)
        if v:
            out[k] = v

    out["_crossref_verified"] = True
    return out
