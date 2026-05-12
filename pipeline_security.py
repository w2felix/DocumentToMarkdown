"""
Shared security utilities for document processing pipelines.

Provides input validation, path sanitization, URL allowlisting,
and file safety checks used across all pipelines.
"""

import os
import re
import logging
from pathlib import Path
from typing import Optional, List
from urllib.parse import urlparse

logger = logging.getLogger(__name__)

# --- URL Validation ---

ALLOWED_BASE_URL_DOMAINS = [
    'api.anthropic.com',
    'foundry.merckgroup.com',
    'api.merck-foundry.com',
    'localhost',
    '127.0.0.1',
]


def validate_base_url(url: str) -> str:
    """Validate ANTHROPIC_BASE_URL is from an allowed domain.

    Raises ValueError if the URL is unsafe.
    """
    if not url or not url.strip():
        raise ValueError("Base URL is empty")

    parsed = urlparse(url.strip())

    # Require a scheme
    if not parsed.scheme:
        raise ValueError(f"Base URL missing scheme (http/https): {url}")

    # Require HTTPS except for localhost
    if parsed.scheme != 'https' and parsed.hostname not in ('localhost', '127.0.0.1'):
        raise ValueError(f"Base URL must use HTTPS: {url}")

    # Reject credentials in URL
    if parsed.username or parsed.password or '@' in (parsed.netloc or ''):
        raise ValueError(f"Base URL must not contain credentials: {url}")

    # Check domain allowlist
    hostname = parsed.hostname or ''
    if not any(hostname == domain or hostname.endswith('.' + domain)
               for domain in ALLOWED_BASE_URL_DOMAINS):
        raise ValueError(
            f"Base URL domain '{hostname}' not in allowlist. "
            f"Allowed: {ALLOWED_BASE_URL_DOMAINS}"
        )

    return url.strip()


# --- Path Validation ---

def validate_path(path_str: str, must_exist: bool = True, allow_file: bool = True,
                  allow_dir: bool = True) -> Path:
    """Validate a user-provided path is safe.

    Checks for:
    - Path traversal (.. components after resolution)
    - Null bytes
    - Excessively long paths
    - Existence (if must_exist=True)
    """
    if not path_str or not path_str.strip():
        raise ValueError("Path is empty")

    # Null byte injection
    if '\x00' in path_str:
        raise ValueError("Path contains null bytes")

    # Length check (Windows MAX_PATH=260, but long path support exists)
    if len(path_str) > 32767:
        raise ValueError(f"Path exceeds maximum length: {len(path_str)} chars")

    path = Path(path_str).resolve()

    if must_exist and not path.exists():
        raise FileNotFoundError(f"Path does not exist: {path_str}")

    if must_exist:
        if path.is_file() and not allow_file:
            raise ValueError(f"Expected directory, got file: {path_str}")
        if path.is_dir() and not allow_dir:
            raise ValueError(f"Expected file, got directory: {path_str}")

    return path


def validate_output_path(path_str: str) -> Path:
    """Validate an output directory path (created if needed)."""
    if not path_str or not path_str.strip():
        raise ValueError("Output path is empty")

    if '\x00' in path_str:
        raise ValueError("Output path contains null bytes")

    path = Path(path_str).resolve()

    # Don't allow writing to system directories
    system_dirs = [
        Path(os.environ.get('SYSTEMROOT', r'C:\Windows')).resolve(),
        Path(os.environ.get('PROGRAMFILES', r'C:\Program Files')).resolve(),
        Path(os.environ.get('PROGRAMFILES(X86)', r'C:\Program Files (x86)')).resolve(),
    ]
    for sys_dir in system_dirs:
        try:
            is_system = (path == sys_dir or path.is_relative_to(sys_dir))
        except (TypeError, AttributeError):
            is_system = False
        if is_system:
            raise ValueError(f"Cannot write to system directory: {path}")

    return path


# --- PPTX Safety ---

def check_pptx_safe(pptx_path: Path) -> bool:
    """Check if a PPTX file is safe to open (no macros/ActiveX).

    Returns True if safe, False if potentially malicious.
    """
    try:
        from zipfile import ZipFile, BadZipFile
    except ImportError:
        return True  # Can't check, allow

    try:
        with ZipFile(pptx_path, 'r') as zf:
            names = zf.namelist()

            # Check for VBA macros (PPTM files have these)
            vba_indicators = [n for n in names if 'vbaProject' in n or 'vba' in n.lower()]
            if vba_indicators:
                logger.warning(f"PPTX contains VBA macros — rejecting: {pptx_path.name}")
                logger.warning(f"  Macro files: {vba_indicators[:5]}")
                return False

            # Check for ActiveX controls
            activex_indicators = [n for n in names if 'activeX' in n]
            if activex_indicators:
                logger.warning(f"PPTX contains ActiveX controls — rejecting: {pptx_path.name}")
                return False

            # Check for external OLE objects that could execute code
            ole_indicators = [n for n in names if 'oleObject' in n or 'embeddings' in n.lower()]
            if len(ole_indicators) > 20:
                logger.warning(f"PPTX contains excessive OLE objects ({len(ole_indicators)}) — suspicious")
                # Don't reject, just warn — some legitimate presentations have embedded objects

    except BadZipFile:
        logger.warning(f"PPTX is not a valid ZIP file — rejecting: {pptx_path.name}")
        return False
    except Exception as e:
        logger.warning(f"Could not validate PPTX safety: {e} — rejecting as precaution")
        return False

    return True


# --- Excel Safety ---

MAX_EXCEL_SIZE_MB = 100  # 100MB limit


def check_excel_safe(excel_path: Path) -> bool:
    """Validate an Excel file for size and basic safety.

    Returns True if safe to process.
    """
    # File size check
    file_size_mb = excel_path.stat().st_size / (1024 * 1024)
    if file_size_mb > MAX_EXCEL_SIZE_MB:
        logger.error(f"Excel file too large: {file_size_mb:.1f}MB (max: {MAX_EXCEL_SIZE_MB}MB)")
        return False

    return True


def sanitize_excel_cell(value) -> str:
    """Sanitize an Excel cell value to prevent formula injection.

    Prefixes dangerous characters that could be interpreted as formulas
    by spreadsheet applications.
    """
    if value is None:
        return ''
    s = str(value)
    # Formula injection: cells starting with = + - @ could execute formulas
    if s and s[0] in ('=', '+', '-', '@'):
        return "'" + s  # Prefix with single quote to force text interpretation
    return s


# --- Filename Sanitization ---

# Windows reserved filenames
_RESERVED_NAMES = frozenset({
    'CON', 'PRN', 'AUX', 'NUL',
    'COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9',
    'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9',
})

_ILLEGAL_CHARS_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def sanitize_filename(filename: str, max_length: int = 200) -> str:
    """Sanitize a filename for cross-platform filesystem safety.

    Handles: illegal characters, reserved names, length limits,
    leading/trailing dots and spaces, Unicode normalization.
    """
    import unicodedata

    if not filename:
        return 'unnamed'

    # Normalize Unicode
    filename = unicodedata.normalize('NFKC', filename)

    # Remove control characters
    filename = ''.join(ch for ch in filename if unicodedata.category(ch)[0] != 'C')

    # Replace illegal filesystem characters
    filename = _ILLEGAL_CHARS_RE.sub('_', filename)

    # Remove leading/trailing dots and spaces
    filename = filename.strip('. ')

    # Handle reserved names
    name_part = filename.split('.')[0].upper()
    if name_part in _RESERVED_NAMES:
        filename = f"file_{filename}"

    # Enforce length limit
    if len(filename) > max_length:
        name, ext = os.path.splitext(filename)
        filename = name[:max_length - len(ext)] + ext

    # Ensure non-empty
    if not filename:
        filename = 'unnamed'

    return filename


# --- Credential Loading (shared) ---

def load_credentials_from_registry() -> None:
    """Load ANTHROPIC credentials from Windows User registry with validation."""
    if os.environ.get('ANTHROPIC_AUTH_TOKEN') and os.environ.get('ANTHROPIC_BASE_URL'):
        return  # Already set

    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, 'Environment', 0, winreg.KEY_READ)
        try:
            if not os.environ.get('ANTHROPIC_AUTH_TOKEN'):
                try:
                    token, _ = winreg.QueryValueEx(key, 'ANTHROPIC_AUTH_TOKEN')
                    if token and len(token) > 10:
                        os.environ['ANTHROPIC_AUTH_TOKEN'] = token
                        logger.info("Loaded ANTHROPIC_AUTH_TOKEN from Windows User environment")
                    else:
                        logger.debug("ANTHROPIC_AUTH_TOKEN in registry is empty or too short")
                except FileNotFoundError:
                    pass
                except OSError as e:
                    logger.debug(f"Could not read ANTHROPIC_AUTH_TOKEN from registry: {e}")

            if not os.environ.get('ANTHROPIC_BASE_URL'):
                try:
                    url, _ = winreg.QueryValueEx(key, 'ANTHROPIC_BASE_URL')
                    validated_url = validate_base_url(url)
                    os.environ['ANTHROPIC_BASE_URL'] = validated_url
                    logger.info("Loaded ANTHROPIC_BASE_URL from Windows User environment")
                except FileNotFoundError:
                    pass
                except ValueError as e:
                    logger.warning(f"Rejecting ANTHROPIC_BASE_URL from registry: {e}")
                except OSError as e:
                    logger.debug(f"Could not read ANTHROPIC_BASE_URL from registry: {e}")
        finally:
            winreg.CloseKey(key)
    except (ImportError, OSError) as e:
        logger.debug(f"Could not load credentials from Windows registry: {e}")
