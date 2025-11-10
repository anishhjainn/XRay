# ui/__init__.py
"""
UI package convenience exports.

This keeps import sites clean:
    from ui import folder_picker, summary_panel, results_table
instead of:
    from ui.components import folder_picker, summary_panel, results_table
"""

from __future__ import annotations

from .components import (
    folder_picker,
    cutoff_input,
    run_controls,
    summary_panel,
    results_table,
    downloads,
    progress_widgets,
)

__all__ = [
    "folder_picker",
    "cutoff_input",
    "run_controls",
    "summary_panel",
    "results_table",
    "downloads",
    "progress_widgets",
]

__version__ = "0.1.0"
