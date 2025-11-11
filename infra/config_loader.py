# infra/config_loader.py
"""
Tiny configuration loader.

Design:
- Single Responsibility: provide app-wide config values.
- Keeps sane defaults; allows simple environment overrides for quick tweaks.
"""

from __future__ import annotations

import os
from typing import Any, Dict, List


_DEFAULT: Dict[str, Any] = {
    "max_filename_length": 120,
    "target_extensions": [".docx", ".pptx", ".pdf",".xlsx"],
    "ignore_dirs": [".git", "__pycache__", "venv"],
    "log_level": "INFO",  # DEBUG/INFO/WARNING/ERROR
}


def load_config() -> Dict[str, Any]:
    """
    Return a config dict. Environment variables can override some keys:

    - XRAY_MAX_FILENAME_LENGTH
    - XRAY_LOG_LEVEL
    - XRAY_IGNORE_DIRS (comma-separated)
    - XRAY_TARGET_EXTS (comma-separated, e.g. ".docx,.pptx,.pdf")
    """
    cfg = dict(_DEFAULT)

    # Numeric override
    val = os.getenv("XRAY_MAX_FILENAME_LENGTH")
    if val and val.isdigit():
        cfg["max_filename_length"] = int(val)

    # String override
    lvl = os.getenv("XRAY_LOG_LEVEL")
    if lvl:
        cfg["log_level"] = lvl.upper()

    # Lists
    ig = os.getenv("XRAY_IGNORE_DIRS")
    if ig:
        cfg["ignore_dirs"] = _split_list(ig)

    ex = os.getenv("XRAY_TARGET_EXTS")
    if ex:
        cfg["target_extensions"] = [e if e.startswith(".") else f".{e}" for e in _split_list(ex)]

    return cfg


def _split_list(s: str) -> List[str]:
    return [part.strip() for part in s.split(",") if part.strip()]
