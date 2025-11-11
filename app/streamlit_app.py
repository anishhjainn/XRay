# app/streamlit_app.py
"""
XRay Builder — Streamlit UI entry point.

Flow:
1) Configure logging + load config.
2) Let user pick a folder and (optionally) a cutoff date/time.
3) Import processors/checks (self-register into the registry).
4) Run orchestrator with a progress callback.
5) Show summary + table; allow CSV/JSON export.

Run:
    streamlit run app/streamlit_app.py
"""
from __future__ import annotations
# --- ensure project root is on sys.path ---
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]  # one level up from /app
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
# ------------------------------------------



import logging
from pathlib import Path

import pandas as pd
import streamlit as st

from infra.config_loader import load_config
from infra.logging_config import configure_logging

# Import processors/checks so they self-register with the registry on import.
# (Dependency Inversion: the app never references their internals.)
import processors.docx_processor  # noqa: F401
import processors.pptx_processor  # noqa: F401
import processors.pdf_processor   # noqa: F401
import processors.xlsx_processor  # noqa: F401

import checks.base_checks         # noqa: F401
import checks.docx_checks         # noqa: F401
import checks.pptx_checks         # noqa: F401
import checks.pdf_checks          # noqa: F401
import checks.xlsx_checks         # noqa: F401

from checks.settings import set_modified_cutoff, clear_modified_cutoff
from services.orchestrator import Orchestrator
from ui.components import folder_picker, cutoff_input, run_controls, summary_panel, results_table, downloads, progress_widgets
from utils.path_utils import iter_target_files


def main():
    st.set_page_config(page_title="XRay Builder — Document Checker", layout="wide")
    cfg = load_config()
    configure_logging(cfg.get("log_level", "INFO"))
    log = logging.getLogger("app")

    st.title("XRay Builder — Document Checker")
    st.caption("Pre-Archival checks for .docx, .pptx, .pdf")

    root = folder_picker()
    cutoff_dt = cutoff_input()

    colA, colB = st.columns([1, 3])
    run_clicked = colA.button("Scan folder", type="primary", use_container_width=True)
    reset_clicked = colB.button("Reset cutoff", use_container_width=True, help="Clears the configured cutoff date")

    if reset_clicked:
        clear_modified_cutoff()
        st.success("Cutoff cleared for this session.")

    if run_clicked:
        if not root:
            st.error("Please enter a folder path.")
            return

        root_path = Path(root)
        if not root_path.exists() or not root_path.is_dir():
            st.error("Folder does not exist or is not a directory.")
            return

        # Apply cutoff setting for this run
        set_modified_cutoff(cutoff_dt)

        # Progress widgets + discovery count
        set_total, on_progress = progress_widgets()
        discovered = list(iter_target_files(root_path, exts=cfg.get("target_extensions"), ignore_dirs=cfg.get("ignore_dirs")))
        set_total(len(discovered))

        # Run scan
        orchestrator = Orchestrator(on_progress=on_progress)
        report = orchestrator.run_scan_v2(root_path, config_snapshot=cfg)

        # Show results
        st.divider()
        summary_panel(report)
        filtered_df = results_table(report)
        downloads(filtered_df)

        log.info("Scan completed on %s. Results: %d rows.", root_path, len(report.files))

    st.sidebar.header("Config")
    st.sidebar.write("**Target extensions:**", ", ".join(cfg.get("target_extensions", [])))
    st.sidebar.write("**Ignored folders:**", ", ".join(cfg.get("ignore_dirs", [])))
    st.sidebar.write("**Log level:**", cfg.get("log_level", "INFO"))

    st.sidebar.markdown(
        "> Tip: Use the text filter to quickly find all errors for a specific file or rule."
    )


if __name__ == "__main__":
    main()
