# utils/path_utils.py
"""
Path utilities: discover target files for scanning.

Goals:
- Single responsibility: file discovery only (no parsing, no UI).
- Windows-friendly, but cross-platform safe.
- Fast directory walking with pruning (skip .git, venv, etc.).
- Easy to extend: allow optional overrides for extensions and ignored dirs.

Default target extensions: .docx, .pptx, .pdf
Default ignored dirs: .git, __pycache__, venv
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Iterable, Iterator, Optional, Set

# Defaults are small and explicit; callers can override if needed.
_DEFAULT_EXTS: Set[str] = {".docx", ".pptx", ".pdf"}
_DEFAULT_IGNORES: Set[str] = {".git", "__pycache__", "venv"}


def iter_target_files(
    root: Path | str,
    exts: Optional[Iterable[str]] = None,
    ignore_dirs: Optional[Iterable[str]] = None,
) -> Iterator[Path]:
    """
    Yield files under 'root' whose extension is in 'exts', skipping folders in 'ignore_dirs'.

    Parameters
    ----------
    root : Path | str
        Folder to scan.
    exts : Optional[Iterable[str]]
        Allowed file extensions (case-insensitive). Defaults to .docx, .pptx, .pdf.
    ignore_dirs : Optional[Iterable[str]]
        Directory names to skip entirely (case-insensitive). Defaults to .git, __pycache__, venv.

    Yields
    ------
    Path
        Absolute Path objects for matching files.

    Notes
    -----
    - Uses os.walk(topdown=True) so we can prune directories in-place (fast).
    - Does not follow symlinks (safer; avoids infinite loops).
    - Normalizes extensions and ignore names to lowercase for consistent matching.
    """
    root_path = Path(root).resolve()
    allowed_exts = _normalize_exts(exts or _DEFAULT_EXTS)
    ignored = _normalize_names(ignore_dirs or _DEFAULT_IGNORES)

    # Prune as we walk for speed.
    for dirpath, dirnames, filenames in os.walk(
        root_path, topdown=True, followlinks=False
    ):
        # In-place pruning of ignored directories (case-insensitive match on name).
        # Example: if "venv" in ignored, any folder named "venv" is skipped (wherever it occurs).
        dirnames[:] = [d for d in dirnames if d.lower() not in ignored]

        # Emit only target files by extension, case-insensitive.
        for fname in filenames:
            ext = _suffix_lower(fname)
            if ext in allowed_exts:
                yield Path(dirpath) / fname


# ---------- helpers (module-internal) ----------

def _normalize_exts(exts: Iterable[str]) -> Set[str]:
    """
    Return a normalized set of extensions like {".pdf", ".docx"}.
    Accepts items with or without leading dot; coerces to lowercase.
    """
    norm: Set[str] = set()
    for e in exts:
        s = str(e).strip().lower()
        if not s:
            continue
        if not s.startswith("."):
            s = "." + s
        norm.add(s)
    return norm


def _normalize_names(names: Iterable[str]) -> Set[str]:
    """Lowercase each name for case-insensitive comparisons."""
    return {str(n).lower() for n in names if str(n).strip()}


def _suffix_lower(filename: str) -> str:
    """
    Fast, allocation-light way to get a lowercase suffix from a filename.
    We avoid constructing a Path for each file for speed inside tight loops.
    """
    # Find last dot; if none, return empty string
    idx = filename.rfind(".")
    if idx == -1:
        return ""
    return filename[idx:].lower()
