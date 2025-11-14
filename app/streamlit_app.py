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
from ui.components import folder_picker, cutoff_input, run_controls, summary_panel, results_table, downloads, progress_widgets, yellow_cells_drilldown
from app.components import sidebar_extension_selector, sidebar_discovery_summary
from utils.path_utils import iter_target_files
from services.orchestrator import Orchestrator


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
import checks.spelling_checks     # noqa: F401  # Minimal wiring: enable SpellingCheck via self-registration

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
    st.caption("Pre-Archival checks for .docx, .pptx, .pdf")  # Keep minimal; caption can be updated later

    root = folder_picker()
    cutoff_dt = cutoff_input()

    # Initialize session state for two-stage scan
    if "extensions_populated" not in st.session_state:
        st.session_state.extensions_populated = False
    if "current_root" not in st.session_state:
        st.session_state.current_root = None
    
    # Reset extension population state if folder changes
    if root and st.session_state.current_root != root:
        st.session_state.extensions_populated = False
        st.session_state.current_root = root

    colA, colB = st.columns([1, 3])
    
    # Change button label based on state
    button_label = "Select file types" if not st.session_state.extensions_populated else "Scan folder"
    run_clicked = colA.button(button_label, type="primary", use_container_width=True)
    reset_clicked = colB.button("Reset cutoff", use_container_width=True, help="Clears the configured cutoff date")

    if reset_clicked:
        clear_modified_cutoff()
        st.success("Cutoff cleared for this session.")

    # Show extension selector in sidebar after first click
    enabled_exts = []
    if root and st.session_state.extensions_populated:
        root_path = Path(root)
        enabled_exts = sidebar_extension_selector(root_path)
        
        # Show discovery count for current selection
        ext_filter = enabled_exts if enabled_exts else None
        discovery_count = 0
        for _ in iter_target_files(root_path, exts=ext_filter):
            discovery_count += 1
        sidebar_discovery_summary(discovery_count)

    if run_clicked:
        if not root:
            st.error("Please enter a folder path.")
            return

        root_path = Path(root)
        if not root_path.exists() or not root_path.is_dir():
            st.error("Folder does not exist or is not a directory.")
            return

        # First click: populate extensions in sidebar
        if not st.session_state.extensions_populated:
            st.session_state.extensions_populated = True
            st.rerun()  # Rerun to show the extension selector
            return

        # Second click: perform actual scan
        # Apply cutoff setting for this run
        set_modified_cutoff(cutoff_dt)

        # Progress widgets
        set_total, on_progress = progress_widgets()

        # Use already computed enabled_exts from above
        ext_filter = enabled_exts if enabled_exts else None

        # Run scan
        orchestrator = Orchestrator(on_progress=on_progress)
        report = orchestrator.run_scan_v2(root_path, config_snapshot=cfg, exts=ext_filter)

        # Show results
        st.divider()
        summary_panel(report)
        filtered_df = results_table(report)
        from ui.components import yellow_cells_drilldown  # add at top with other UI imports, or inline here
        yellow_cells_drilldown(filtered_df)
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