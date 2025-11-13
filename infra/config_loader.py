# infra/config_loader.py
"""
Central configuration loader for the app.

Responsibilities (SRP):
- Provide a single place to define default configuration values.
- Allow simple environment variable overrides for quick tweaks (no code changes).

Why here (OCP/DIP):
- Adding new config flags doesn't require modifying orchestrator, processors, or checks.
- High-level logic depends on config values rather than hard-coded constants.

Environment variables (examples):
- XRAY_MAX_FILENAME_LENGTH
- XRAY_LOG_LEVEL
- XRAY_IGNORE_DIRS                (comma-separated)
- XRAY_TARGET_EXTS                (comma-separated, e.g. ".docx,.pptx,.pdf,.xlsx")

Spelling/Grammar (new):
- XRAY_ENABLE_SPELLING            ("1"/"true"/"yes" -> True)
- XRAY_ENABLE_GRAMMAR             ("1"/"true"/"yes" -> True)
- XRAY_GRAMMAR_ENGINE             ("basic" | "language_tool" | "word_com")
- XRAY_LANGUAGE_CODE              (e.g., "en")
- XRAY_MAX_TEXT_CHARS             (int; cap for text extraction, default 5_000_000)
- XRAY_SPELLING_FAIL_THRESHOLD    (int; default 10)
- XRAY_MAX_MISSPELLINGS_REPORTED  (int; default 100)
- XRAY_GRAMMAR_FAIL_THRESHOLD     (int; default 5)
"""

from __future__ import annotations

import os
from typing import Any, Dict, List


_DEFAULT: Dict[str, Any] = {
    # Core
    "max_filename_length": 120,
    "target_extensions": [".docx", ".pptx", ".pdf", ".xlsx"],
    "ignore_dirs": [".git", "__pycache__", "venv"],
    "log_level": "INFO",  # DEBUG/INFO/WARNING/ERROR

    # Spelling (basic phase 1 checks)
    "enable_spelling": True,
    "language_code": "en",
    "max_text_chars": 5_000_000,          # extraction cap to protect memory/CPU
    "spelling_fail_threshold": 10,        # >10 misspellings => ERROR (policy)
    "max_misspellings_reported": 100,     # cap list in CheckResult.extra

    # Grammar - disabled by default
    "enable_grammar": False,
    "grammar_engine": "basic",            # "basic" heuristics; can switch later
    "grammar_fail_threshold": 5,          # default policy for later
}


def load_config() -> Dict[str, Any]:
    """
    Return a config dict. Environment variables can override some keys.

    Lists:
      - XRAY_IGNORE_DIRS (comma-separated)
      - XRAY_TARGET_EXTS (comma-separated, e.g. ".docx,.pptx,.pdf,.xlsx")

    Booleans:
      - XRAY_ENABLE_SPELLING ("1"/"true"/"yes" -> True)
      - XRAY_ENABLE_GRAMMAR ("1"/"true"/"yes" -> True)

    Strings:
      - XRAY_LOG_LEVEL
      - XRAY_GRAMMAR_ENGINE
      - XRAY_LANGUAGE_CODE

    Integers:
      - XRAY_MAX_FILENAME_LENGTH
      - XRAY_MAX_TEXT_CHARS
      - XRAY_SPELLING_FAIL_THRESHOLD
      - XRAY_MAX_MISSPELLINGS_REPORTED
      - XRAY_GRAMMAR_FAIL_THRESHOLD
    """
    cfg = dict(_DEFAULT)

    # Numeric overrides
    _int_env(cfg, "max_filename_length", "XRAY_MAX_FILENAME_LENGTH")
    _int_env(cfg, "max_text_chars", "XRAY_MAX_TEXT_CHARS")
    _int_env(cfg, "spelling_fail_threshold", "XRAY_SPELLING_FAIL_THRESHOLD")
    _int_env(cfg, "max_misspellings_reported", "XRAY_MAX_MISSPELLINGS_REPORTED")
    _int_env(cfg, "grammar_fail_threshold", "XRAY_GRAMMAR_FAIL_THRESHOLD")

    # Boolean overrides
    _bool_env(cfg, "enable_spelling", "XRAY_ENABLE_SPELLING")
    _bool_env(cfg, "enable_grammar", "XRAY_ENABLE_GRAMMAR")

    # String overrides
    _str_upper_env(cfg, "log_level", "XRAY_LOG_LEVEL")
    _str_env(cfg, "grammar_engine", "XRAY_GRAMMAR_ENGINE")
    _str_env(cfg, "language_code", "XRAY_LANGUAGE_CODE")

    # List overrides
    ig = os.getenv("XRAY_IGNORE_DIRS")
    if ig:
        cfg["ignore_dirs"] = _split_list(ig)

    ex = os.getenv("XRAY_TARGET_EXTS")
    if ex:
        # ensure dot-prefixed, lowercase extensions
        cfg["target_extensions"] = [
            e if e.startswith(".") else f".{e}"
            for e in (s.lower() for s in _split_list(ex))
        ]

    return cfg


# ----------------- helpers -----------------

def _split_list(s: str) -> List[str]:
    return [part.strip() for part in s.split(",") if part.strip()]


def _bool_env(cfg: Dict[str, Any], key: str, env_key: str) -> None:
    val = os.getenv(env_key)
    if val is None:
        return
    s = val.strip().lower()
    cfg[key] = s in {"1", "true", "yes", "on"}


def _int_env(cfg: Dict[str, Any], key: str, env_key: str) -> None:
    val = os.getenv(env_key)
    if val and val.strip().isdigit():
        try:
            cfg[key] = int(val)
        except ValueError:
            pass


def _str_env(cfg: Dict[str, Any], key: str, env_key: str) -> None:
    val = os.getenv(env_key)
    if val is not None:
        cfg[key] = val.strip()


def _str_upper_env(cfg: Dict[str, Any], key: str, env_key: str) -> None:
    val = os.getenv(env_key)
    if val is not None:
        cfg[key] = val.strip().upper()