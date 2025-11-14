from pathlib import Path
import os
from typing import List
import streamlit as st

from core import registry

def _normalize_exts(exts: List[str]) -> List[str]:
    return sorted({(e if e.startswith(".") else f".{e}").lower() for e in exts})

def sidebar_extension_selector(root: Path) -> List[str]:
    """
    Render a sidebar multiselect for supported file extensions present in `root`.

    Returns the selected extensions (dot-prefixed, lowercase). If none are present,
    shows an informational message and returns [].

    If the user selects nothing, treat as 'no filters' upstream (scan all supported types).
    """
    # Compute supported extensions from registered processors
    supported: set[str] = set()
    for proc in registry.processors():
        try:
            for ext in proc.supports():
                if isinstance(ext, str) and ext:
                    supported.add((ext if ext.startswith(".") else f".{ext}").lower())
        except Exception:
            # Defensive: skip misbehaving processors
            continue

    # Walk once to compute present extensions
    present: set[str] = set()
    root_str = str(root.resolve())
    for dirpath, _dirnames, filenames in os.walk(root_str):
        for fname in filenames:
            ext = Path(fname).suffix.lower()
            if ext and ext in supported:
                present.add(ext)

    options = sorted(present)
    state_key = f"ext_selection::{root_str}"
    if not options:
        st.sidebar.info("No supported file types found in the selected folder.")
        st.session_state[state_key] = []
        return []

    # Initialize state with all selected by default
    if state_key not in st.session_state:
        st.session_state[state_key] = options

    selected = st.sidebar.multiselect(
        "File types to include in check:",
        options=options,
        default=st.session_state[state_key],
        help="Uncheck to limit the scan. If none selected, all supported types will be scanned.",
    )
    st.session_state[state_key] = selected
    return _normalize_exts(selected)

def sidebar_discovery_summary(count: int):
    """Optional helper to display discovery count."""
    st.sidebar.caption(f"Discovered files: {count}")
