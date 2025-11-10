# ui/components.py
"""
Streamlit UI helpers (pure rendering/inputs; no business logic).

Functions:
- folder_picker() -> Optional[str]
- cutoff_input()  -> Optional[date | datetime]
- run_controls()  -> bool (returns True if user pressed "Scan")
- summary_panel(report)
- results_table(report) -> pandas.DataFrame
- downloads(df)  -> renders CSV/JSON download buttons
- progress_widgets() -> (set_total, update) closures for orchestrator callback
"""

from __future__ import annotations

from datetime import date, datetime, time
from pathlib import Path
from typing import Optional, Tuple, Callable, Iterable

import pandas as pd
import streamlit as st

from core.models import AggregateReport, Severity, CheckResult

# --- report shape detection helpers ---
def _is_v2_scan_report(report) -> bool:
    # v2 has header + files list
    return hasattr(report, "header") and hasattr(report, "files")

def _flatten_v2_checks(report):
    # Turn v2 (file-centric) into flat rows like v1 for the table
    rows = []
    for f in getattr(report, "files", []):
        for r in getattr(f, "results", []):
            rows.append({
                "File": str(f.file),
                "Verdict": getattr(f.verdict, "value", str(f.verdict)),
                "Check": r.check_name,
                "Severity": r.severity.value,
                "Passed": bool(r.passed),
                "Message": r.message,
                "Extra": r.extra,
            })
    return rows

# -------- Inputs --------

def folder_picker() -> Optional[str]:
    """
    Windows-friendly folder entry.
    Streamlit doesn't have a native folder dialog, so we use a text input.
    """
    return st.text_input(
        "Folder path to scan",
        placeholder=r"C:\path\to\your\folder",
        help="Paste or type the folder path. Subfolders will be scanned.",
    ).strip() or None


def cutoff_input() -> Optional[datetime]:
    """
    Two-part input: a date (required to set) and an optional time.
    If no date is chosen, returns None (no cutoff).
    """
    with st.expander("Optional: Set 'last allowed modification' cutoff", expanded=False):
        col1, col2 = st.columns([2, 1])
        d: Optional[date] = col1.date_input(
            "Cutoff date (inclusive)", value=None, format="YYYY-MM-DD"
        )
        t = col2.time_input("Time (optional)", value=time(23, 59), help="Defaults to end of day")
        if d is None:
            return None
        # Combine as naive local datetime; checks.settings will attach local tz
        return datetime.combine(d, t)


def run_controls() -> bool:
    """
    Render a primary action button and return True if clicked.
    """
    return st.button("Scan folder", type="primary")


# -------- Progress wiring --------

def progress_widgets() -> Tuple[Callable[[int], None], Callable[[int, int, Path], None]]:
    """
    Create progress placeholders and return two closures:
    - set_total(total_files)
    - on_progress(i, total, path)
    """
    bar = st.progress(0)
    status = st.empty()

    total = {"value": 0}

    def set_total(n: int) -> None:
        total["value"] = max(0, int(n))
        if n <= 0:
            bar.progress(0)
            status.info("No files to scan.")
        else:
            status.info(f"Discovered {n} file(s). Starting scan…")

    def on_progress(i: int, n: int, path: Path) -> None:
        if n > 0:
            bar.progress(min(int(i / n * 100), 100))
        status.write(f"[{i}/{n}] {path}")

    return set_total, on_progress


# -------- Results rendering --------

def _results_to_df(results: Iterable[CheckResult]) -> pd.DataFrame:
    rows = []
    for r in results:
        rows.append(
            {
                "File": str(r.file),
                "Check": r.check_name,
                "Severity": r.severity.value,
                "Passed": r.passed,
                "Message": r.message,
                # Flatten a couple of common extras; keep full extra as JSON string
                "Extra": r.extra,
            }
        )
    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.sort_values(["Severity", "Passed", "File", "Check"]).reset_index(drop=True)
    return df


def summary_panel(report) -> None:
    st.subheader("Summary")
    if _is_v2_scan_report(report):
        # Prefer header totals (fast and consistent)
        total = report.header.total_checks
        errors = report.header.total_errors
        warnings = report.header.total_warnings
        infos = report.header.total_infos
        files = report.header.total_files
    else:
        # v1 AggregateReport (flat)
        total = len(report.results)
        errors = sum(1 for r in report.results if not r.passed and r.severity.name == "ERROR")
        warnings = sum(1 for r in report.results if not r.passed and r.severity.name == "WARNING")
        infos = sum(1 for r in report.results if r.severity.name == "INFO")
        files = len({str(r.file) for r in report.results})

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Files", files)
    c2.metric("Checks run", total)
    c3.metric("Errors", errors)
    c4.metric("Warnings", warnings)
    c5.metric("Info", infos)



def results_table(report) -> pd.DataFrame:
    """
    Render the check-level table.
    Works with both v1 (AggregateReport) and v2 (ScanReport).
    Returns the filtered DataFrame (for export).
    """
    if _is_v2_scan_report(report):
        base_rows = _flatten_v2_checks(report)
    else:
        base_rows = [{
            "File": str(r.file),
            "Verdict": "",  # not available in v1
            "Check": r.check_name,
            "Severity": r.severity.value,
            "Passed": r.passed,
            "Message": r.message,
            "Extra": r.extra,
        } for r in report.results]

    df = pd.DataFrame(base_rows)
    if df.empty:
        st.info("No results to show.")
        return df

    with st.expander("Checks (filters)", expanded=False):
        cols = st.columns(5 if "Verdict" in df.columns else 4)
        sev = cols[0].multiselect("Severity", options=sorted(df["Severity"].unique()), default=list(sorted(df["Severity"].unique())))
        passed = cols[1].multiselect("Passed", options=[True, False], default=[True, False])
        if "Verdict" in df.columns and df["Verdict"].notna().any():
            verdicts = cols[2].multiselect("Verdict", options=sorted([v for v in df["Verdict"].unique() if v]), default=sorted([v for v in df["Verdict"].unique() if v]))
            col_idx = 3
        else:
            verdicts = None
            col_idx = 2
        substr = cols[col_idx].text_input("Text filter", value="")
        show_cols = cols[col_idx + 1].multiselect("Columns", options=list(df.columns), default=list(df.columns))

    mask = df["Severity"].isin(sev) & df["Passed"].isin(passed)
    if verdicts is not None:
        mask &= df["Verdict"].isin(verdicts)
    if substr:
        s = substr.lower()
        mask &= (
            df["File"].str.lower().str.contains(s, na=False)
            | df["Check"].str.lower().str.contains(s, na=False)
            | df["Message"].str.lower().str.contains(s, na=False)
        )
    fdf = df[mask]
    if show_cols:
        fdf = fdf[show_cols]

    st.dataframe(fdf, use_container_width=True)
    return fdf



def downloads(df) -> None:
    if df.empty:
        return
    csv = df.to_csv(index=False).encode("utf-8")
    json = df.to_json(orient="records", force_ascii=False, indent=2).encode("utf-8")
    c1, c2 = st.columns(2)
    c1.download_button("Download CSV", data=csv, file_name="xray_results.csv", mime="text/csv")
    c2.download_button("Download JSON", data=json, file_name="xray_results.json", mime="application/json")

def files_summary_table(report):
    """
    For v2: show one row per file with verdict and counts.
    Returns the DataFrame (or None for v1).
    """
    if not _is_v2_scan_report(report):
        return None

    rows = [{
        "File": str(f.file),
        "Extension": f.extension,
        "SizeBytes": f.size_bytes,
        "Verdict": getattr(f.verdict, "value", str(f.verdict)),
        "Errors": f.errors,
        "Warnings": f.warnings,
        "Infos": f.infos,
    } for f in report.files]

    df = pd.DataFrame(rows)
    if df.empty:
        st.info("No files found.")
        return df

    with st.expander("Files (summary)", expanded=True):
        # simple filters
        cols = st.columns(3)
        verdicts = cols[0].multiselect("Verdict", options=sorted(df["Verdict"].unique()), default=list(sorted(df["Verdict"].unique())))
        ext_sel = cols[1].multiselect("Extension", options=sorted(df["Extension"].unique()), default=list(sorted(df["Extension"].unique())))
        substr = cols[2].text_input("Text filter", value="")

        mask = df["Verdict"].isin(verdicts) & df["Extension"].isin(ext_sel)
        if substr:
            s = substr.lower()
            mask &= df["File"].str.lower().str.contains(s, na=False)
        fdf = df[mask].sort_values(["Verdict", "Errors", "Warnings", "File"]).reset_index(drop=True)

        st.dataframe(fdf, use_container_width=True)
        return fdf
