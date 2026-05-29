"""Shared Anthropic credential loading for DocumentToMarkdown pipelines."""

import os
import sys

from pipeline_security import validate_base_url


def _load_credentials():
    """Load Anthropic credentials from Windows registry into environment if not already set."""
    needed = ("ANTHROPIC_AUTH_TOKEN", "ANTHROPIC_BASE_URL", "ANTHROPIC_API_KEY")
    if all(os.environ.get(v) for v in needed[:2]) or os.environ.get(needed[2]):
        return
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
        )
    elif api_key:
        _anthropic_client = anthropic.Anthropic(api_key=api_key)
    else:
        raise RuntimeError("No Anthropic credentials found.")
    return _anthropic_client
