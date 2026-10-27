"""Shared Anthropic credential loading for DocumentToMarkdown pipelines."""

import os
import re
import sys
import threading

from pipeline_security import validate_base_url


def _load_credentials():
    """Load Anthropic credentials into environment if not already set."""
    needed = ("ANTHROPIC_AUTH_TOKEN", "ANTHROPIC_BASE_URL", "ANTHROPIC_API_KEY")
    if all(os.environ.get(v) for v in needed[:2]) or os.environ.get(needed[2]):
        return
    try:
        from credentials_loader import load_credentials as _load_from_toml
        _load_from_toml()
        if all(os.environ.get(v) for v in needed[:2]) or os.environ.get(needed[2]):
            return
    except Exception:
        pass
    if sys.platform != "win32":
        return
    try:
        import winreg
    except ImportError:
        return
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Environment", 0, winreg.KEY_READ)
        try:
            for var in needed:
                if not os.environ.get(var):
                    try:
                        val, _ = winreg.QueryValueEx(key, var)
                        if val and len(str(val)) > 20:
                            if var == "ANTHROPIC_BASE_URL":
                                val = validate_base_url(str(val))
                            os.environ[var] = str(val)
                    except FileNotFoundError:
                        pass
                    except ValueError as e:
                        print(f"  [!] Rejecting {var} from registry: {e}")
        finally:
            winreg.CloseKey(key)
    except OSError as e:
        print(f"  [!] Warning: could not read credentials from registry: {e}")


_anthropic_client = None


def get_anthropic_client():
    """Return a cached Anthropic client instance. Creates one on first call."""
    global _anthropic_client
    if _anthropic_client is not None:
        return _anthropic_client

    import anthropic

    _load_credentials()
    auth_token = os.environ.get("ANTHROPIC_AUTH_TOKEN")
    base_url = os.environ.get("ANTHROPIC_BASE_URL")
    api_key = os.environ.get("ANTHROPIC_API_KEY")

    if auth_token and base_url:
        _anthropic_client = anthropic.Anthropic(
            api_key="foundry-bearer-auth-not-used",
            base_url=base_url,
            default_headers={"Authorization": f"Bearer {auth_token}"},
            timeout=120,
        )
    elif api_key:
        _anthropic_client = anthropic.Anthropic(api_key=api_key)
    else:
        raise RuntimeError("No Anthropic credentials found.")
    return _anthropic_client


# ── Model selection ───────────────────────────────────────────────────────────
# Newest Sonnet / Haiku the endpoint exposes, resolved once per process on first
# use. An env var (VISION_MODEL / SUMMARY_MODEL) overrides; the literals below
# apply only when the endpoint's model list can't be read.

_FALLBACK_MODELS = {"sonnet": "claude-sonnet-4-6", "haiku": "claude-haiku-4-5-20251001"}

# claude-sonnet-5-5, claude-sonnet-5, claude-haiku-4-5-20251001, ...
_MODEL_ID_RE = re.compile(r"^claude-(sonnet|haiku)-(\d+)(?:-(\d{1,2}))?(?:-(\d{8}))?$")

_resolved_models = {}
_resolve_lock = threading.Lock()


def latest_models(model_ids):
    """Pick the highest-versioned model ID per family ('sonnet', 'haiku')."""
    best = {}
    for model_id in model_ids:
        m = _MODEL_ID_RE.match(model_id)
        if not m:
            continue
        family, major, minor, date = m.groups()
        key = (int(major), int(minor or 0), date or "")
        if family not in best or key > best[family][0]:
            best[family] = (key, model_id)
    return {family: model_id for family, (_, model_id) in best.items()}


def _resolve_models():
    with _resolve_lock:
        if _resolved_models:
            return _resolved_models
        found = {}
        try:
            client = get_anthropic_client().with_options(timeout=10.0, max_retries=1)
            found = latest_models(m.id for m in client.models.list(limit=1000))
        except Exception as e:
            print(f"  [!] Could not list models from endpoint ({e}); using fallback model IDs")
        for family, fallback in _FALLBACK_MODELS.items():
            _resolved_models[family] = found.get(family) or fallback
        return _resolved_models


def get_model(family, env_var=None):
    """Return the model ID for 'sonnet' or 'haiku' (env_var override first)."""
    return (env_var and os.environ.get(env_var)) or _resolve_models()[family]


class LazyModel:
    """Class attribute that resolves to a model ID on first access, so the
    network lookup never runs at import time. Behaves like a plain string
    attribute for `self.VISION_MODEL` / `Pipeline.VISION_MODEL` readers."""

    def __init__(self, family, env_var):
        self.family = family
        self.env_var = env_var

    def __get__(self, obj, objtype=None):
        return get_model(self.family, self.env_var)
