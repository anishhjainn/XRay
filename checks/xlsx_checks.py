# checks/xlsx_checks.py
"""
XLSX checks: small, single-purpose rules for .xlsx files.

Each class implements the Check interface and inspects fields in
artifact.metadata that were populated by the XLSX processor.

Design goals:
- SRP: one rule per class; easy to reason about and test.
- OCP: add/remove checks by adding/removing classes; no central branching.
- LSP/DIP: the orchestrator calls these via the Check interface (not concrete types).
- ISP: checks depend only on the Check interface & FileArtifact (no UI or IO).
"""

from __future__ import annotations

from typing import Dict, Any, List

# These imports assume the same locations/interfaces as your existing checks.
# If your project’s module paths differ, adjust the imports accordingly.
from core.interfaces import Check
from core.models import FileArtifact, CheckResult, Severity
from core.registry import register_check


def _unreadable_result(artifact: FileArtifact, name: str, message: str, meta: Dict[str, Any]) -> CheckResult:
    """
    Helper: if the artifact couldn't be read fully (e.g., encrypted or corrupted),
    return a WARNING with context so the user sees *why* the check couldn't run.
    """
    return CheckResult(
        file=artifact.path,
        check_name=name,
        severity=Severity.WARNING,
        passed=False,
        message=message,
        extra={
            "read_error": True,
            "read_error_detail": meta.get("read_error_detail"),
        },
    )


# ============== 1) Hidden & Very Hidden Sheets ================================

class XlsxHiddenSheetsCheck(Check):
    """
    Fail if any sheet is hidden or very hidden.
    """

    def name(self) -> str:
        return "xlsx_hidden_sheets"

    def description(self) -> str:
        return "Workbook must not contain hidden or very hidden sheets."

    def applies_to(self) -> List[str]:
        return [".xlsx"]

    def run(self, artifact: FileArtifact) -> CheckResult:
        meta = artifact.metadata or {}
        if meta.get("read_error"):
            return _unreadable_result(artifact, self.name(), "Unreadable XLSX (parse error)", meta)

        hidden = int(meta.get("hidden_sheet_count") or 0)
        very_hidden = int(meta.get("very_hidden_sheet_count") or 0)
        passed = (hidden == 0 and very_hidden == 0)

        return CheckResult(
            file=artifact.path,
            check_name=self.name(),
            severity=Severity.INFO if passed else Severity.ERROR,
            passed=passed,
            message=(
                "OK: no hidden/very hidden sheets"
                if passed
                else f"Found {hidden} hidden and {very_hidden} very hidden sheet(s)"
            ),
            extra={
                "hidden_sheet_count": hidden,
                "very_hidden_sheet_count": very_hidden,
            },
        )


# ============== 2) Formula Errors (formulas allowed, errors not) ==============

class XlsxFormulaErrorsCheck(Check):
    """
    Formulas are allowed, but any error cells or error tokens should fail.
    We consider:
      - error_cell_count: cells with type 'e' (<c t="e">)
      - formula_ref_error_count: any '#REF!' inside formula text
      - other_error_token_count: other well-known error tokens in formulas/values
    """

    def name(self) -> str:
        return "xlsx_formula_errors"

    def description(self) -> str:
        return "Formulas are allowed, but no error cells or error tokens are permitted."

    def applies_to(self) -> List[str]:
        return [".xlsx"]

    def run(self, artifact: FileArtifact) -> CheckResult:
        meta = artifact.metadata or {}
        if meta.get("read_error"):
            return _unreadable_result(artifact, self.name(), "Unreadable XLSX (parse error)", meta)

        error_cells = int(meta.get("error_cell_count") or 0)
        ref_err = int(meta.get("formula_ref_error_count") or 0)
        other_errs = int(meta.get("other_error_token_count") or 0)
        formula_count = int(meta.get("formula_count") or 0)

        passed = (error_cells == 0 and ref_err == 0 and other_errs == 0)

        return CheckResult(
            file=artifact.path,
            check_name=self.name(),
            severity=Severity.INFO if passed else Severity.ERROR,
            passed=passed,
            message=(
                "OK: no formula errors"
                if passed
                else (
                    f"Formula issues: error_cells={error_cells}, #REF!={ref_err}, other_errors={other_errs}"
                )
            ),
            extra={
                "formula_count": formula_count,
                "error_cell_count": error_cells,
                "formula_ref_error_count": ref_err,
                "other_error_token_count": other_errs,
            },
        )


# ============== 3) External Links (fail if any) ===============================

class XlsxExternalLinksCheck(Check):
    """
    Fail if the workbook contains external relationships (URLs/UNC/TargetMode=External)
    or externalLinks parts.
    """

    def name(self) -> str:
        return "xlsx_external_links"

    def description(self) -> str:
        return "Workbook must not contain external links."

    def applies_to(self) -> List[str]:
        return [".xlsx"]

    def run(self, artifact: FileArtifact) -> CheckResult:
        meta = artifact.metadata or {}
        if meta.get("read_error"):
            return _unreadable_result(artifact, self.name(), "Unreadable XLSX (parse error)", meta)

        count = int(meta.get("external_links_count") or 0)
        passed = (count == 0)

        return CheckResult(
            file=artifact.path,
            check_name=self.name(),
            severity=Severity.INFO if passed else Severity.ERROR,
            passed=passed,
            message=(
                "OK: no external links" if passed else f"Found {count} external link(s)"
            ),
            extra={"external_links_count": count},
        )


# ============== 4) Data Connections (fail if any) ============================

class XlsxDataConnectionsCheck(Check):
    """
    Fail if the workbook declares any data connections (Power Query / OLE DB / Web).
    """

    def name(self) -> str:
        return "xlsx_data_connections"

    def description(self) -> str:
        return "Workbook must not contain data connections."

    def applies_to(self) -> List[str]:
        return [".xlsx"]

    def run(self, artifact: FileArtifact) -> CheckResult:
        meta = artifact.metadata or {}
        if meta.get("read_error"):
            return _unreadable_result(artifact, self.name(), "Unreadable XLSX (parse error)", meta)

        count = int(meta.get("data_connections_count") or 0)
        passed = (count == 0)

        return CheckResult(
            file=artifact.path,
            check_name=self.name(),
            severity=Severity.INFO if passed else Severity.ERROR,
            passed=passed,
            message=(
                "OK: no data connections" if passed else f"Found {count} data connection(s)"
            ),
            extra={"data_connections_count": count},
        )


# ============== 5) Workbook Protection (fail if encrypted; warn if structure) =

class XlsxWorkbookProtectionCheck(Check):
    """
    - Fail if the workbook is password-encrypted (file-level).
    - Flag (warning) if workbook structure/windows protection is set.
    """

    def name(self) -> str:
        return "xlsx_workbook_protection"

    def description(self) -> str:
        return "Fail if encrypted; warn if workbook structure/windows are protected."

    def applies_to(self) -> List[str]:
        return [".xlsx"]

    def run(self, artifact: FileArtifact) -> CheckResult:
        meta = artifact.metadata or {}
        # Note: For encryption, we *want* to elevate to ERROR even if other reads failed.
        encrypted = bool(meta.get("password_encrypted_workbook", False))

        if encrypted:
            return CheckResult(
                file=artifact.path,
                check_name=self.name(),
                severity=Severity.ERROR,
                passed=False,
                message="Encrypted workbook (password protected)",
                extra={
                    "password_encrypted_workbook": True,
                    "workbook_structure_protected": bool(meta.get("workbook_structure_protected", False)),
                    "read_error": bool(meta.get("read_error", True)),
                    "read_error_detail": meta.get("read_error_detail"),
                },
            )

        # Not encrypted; if the artifact is otherwise unreadable, return a warning.
        if meta.get("read_error"):
            return _unreadable_result(artifact, self.name(), "Unreadable XLSX (parse error)", meta)

        structure_protected = bool(meta.get("workbook_structure_protected", False))
        passed = not structure_protected

        # This rule is intentionally a WARNING when structure protection exists (policy choice).
        severity = Severity.INFO if passed else Severity.WARNING

        return CheckResult(
            file=artifact.path,
            check_name=self.name(),
            severity=severity,
            passed=passed,
            message=(
                "OK: no workbook protection"
                if passed
                else "Workbook structure/windows protection enabled"
            ),
            extra={
                "password_encrypted_workbook": False,
                "workbook_structure_protected": structure_protected,
            },
        )


# ============== 6) Comments/Notes (fail if any) ==============================

class XlsxCommentsCheck(Check):
    """
    Treat legacy notes and threaded comments as comments; fail if any exist.
    """

    def name(self) -> str:
        return "xlsx_comments"

    def description(self) -> str:
        return "Workbook must not contain comments or notes (threaded or legacy)."

    def applies_to(self) -> List[str]:
        return [".xlsx"]

    def run(self, artifact: FileArtifact) -> CheckResult:
        meta = artifact.metadata or {}
        if meta.get("read_error"):
            return _unreadable_result(artifact, self.name(), "Unreadable XLSX (parse error)", meta)

        notes = int(meta.get("comments_count") or 0)
        threaded = int(meta.get("threaded_comments_count") or 0)
        total = notes + threaded
        passed = (total == 0)

        return CheckResult(
            file=artifact.path,
            check_name=self.name(),
            severity=Severity.INFO if passed else Severity.ERROR,
            passed=passed,
            message=(
                "OK: no comments/notes"
                if passed
                else f"Found {total} comment(s) (notes={notes}, threaded={threaded})"
            ),
            extra={
                "comments_count": notes,
                "threaded_comments_count": threaded,
                "total_comments": total,
            },
        )


# ============== 7) VBA project present in .xlsx (fail if found) ==============

class XlsxVbaInXlsxCheck(Check):
    """
    .xlsx should not contain macros; presence of xl/vbaProject.bin indicates macro content.
    """

    def name(self) -> str:
        return "xlsx_vba_in_xlsx"

    def description(self) -> str:
        return "Workbook must not embed a VBA project when using .xlsx format."

    def applies_to(self) -> List[str]:
        return [".xlsx"]

    def run(self, artifact: FileArtifact) -> CheckResult:
        meta = artifact.metadata or {}
        if meta.get("read_error"):
            return _unreadable_result(artifact, self.name(), "Unreadable XLSX (parse error)", meta)

        has_vba = bool(meta.get("has_vba_project", False))
        passed = (not has_vba)

        return CheckResult(
            file=artifact.path,
            check_name=self.name(),
            severity=Severity.INFO if passed else Severity.ERROR,
            passed=passed,
            message=(
                "OK: no VBA project" if passed else "VBA project embedded in workbook"
            ),
            extra={"has_vba_project": has_vba},
        )


# Register on import so the orchestrator discovers these via the registry.
register_check(XlsxHiddenSheetsCheck())
register_check(XlsxFormulaErrorsCheck())
register_check(XlsxExternalLinksCheck())
register_check(XlsxDataConnectionsCheck())
register_check(XlsxWorkbookProtectionCheck())
register_check(XlsxCommentsCheck())
register_check(XlsxVbaInXlsxCheck())

#Implementing Yellow cells check

# -------------------------------------------------------------------
# 8) Yellow-highlighted Cells (fail if any yellow cell found)
# 9) Yellow-highlighted Sheet Tabs (fail if any sheet tab is yellow)
# -------------------------------------------------------------------

class XlsxYellowCellsCheck(Check):
    """
    Fail if any cells are highlighted with standard yellow fill.
    Uses metadata['yellow_cell_count'] populated by the XLSX processor.
    """

    def name(self) -> str:
        return "xlsx_yellow_cells"

    def description(self) -> str:
        return "Workbook must not contain yellow-highlighted cells."

    def applies_to(self) -> list[str]:
        return [".xlsx"]

    def run(self, artifact: FileArtifact) -> CheckResult:
        meta = artifact.metadata or {}

        if meta.get("read_error"):
            return _unreadable_result(self.name(), self.description(), meta)

        yellow_cells = int(meta.get("yellow_cell_count") or 0)
        passed = yellow_cells == 0

        return CheckResult(
            file = artifact.path,
            check_name=self.name(),
            passed=passed,
            severity=Severity.INFO if passed else Severity.ERROR,
            message=(
                "OK: no yellow cells" if passed else "Yellow cells are present in the workbook"
            ),
            extra={"yellow_cell_count": yellow_cells},
        )


class XlsxYellowSheetTabsCheck(Check):
    """
    Fail if any sheet tab is highlighted with standard yellow color.
    Uses metadata['yellow_tab_sheets'] and ['yellow_tab_sheet_count'].
    """

    def name(self) -> str:
        return "xlsx_yellow_sheet_tabs"

    def description(self) -> str:
        return "Workbook must not contain sheets with yellow tab color."

    def applies_to(self) -> list[str]:
        return [".xlsx"]

    def run(self, artifact: FileArtifact) -> CheckResult:
        meta = artifact.metadata or {}

        if meta.get("read_error"):
            return _unreadable_result(self.name(), self.description(), meta)

        count = int(meta.get("yellow_tab_sheet_count") or 0)
        sheets = meta.get("yellow_tab_sheets") or []
        passed = count == 0

        return CheckResult(
            file = artifact.path,
            check_name=self.name(),
            passed=passed,
            severity=Severity.INFO if passed else Severity.ERROR,
            message=(
                "OK: no yellow tabs" if passed else "Yellow tabs are present in the workbook"
            ),
            extra={"yellow_tab_sheet_count": count, "yellow_tab_sheets": sheets},
        )


# Register the new checks so they’re picked up by the orchestrator
register_check(XlsxYellowCellsCheck())
register_check(XlsxYellowSheetTabsCheck())
