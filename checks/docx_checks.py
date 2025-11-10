# checks/docx_checks.py
"""
DOCX checks:
1) No comments (treat ANY comment as a violation, resolved or not).
2) No tracked changes (Track Changes) anywhere in the document/headers/footers.
3) No highlights (explicit <w:highlight> or shading <w:shd> with a non-auto fill).

Each check:
- Reads only from FileArtifact.metadata (filled by DocxProcessor).
- Returns WARNING if the file couldn't be read (parse error).
- Registers itself at import time.
"""

#Imports
from __future__ import annotations

from core.interfaces import Check
from core.models import CheckResult, Severity, FileArtifact
from core.registry import register_check

#If Unreadable
def _unreadable(artifact: FileArtifact) -> CheckResult | None:
    """
    If the DOCX couldn't be parsed, return a WARNING result.
    Otherwise return None and let the caller continue.
    """
    if artifact.metadata.get("read_error"):
        return CheckResult(
            file=artifact.path,
            check_name="DOCX unreadable",
            severity=Severity.WARNING,
            passed=False,
            message="Unreadable DOCX (parse error)",
            extra={"detail": artifact.metadata.get("read_error_detail")},
        )
    return None

#Checks for no comments
class DocxNoCommentsCheck(Check):
    def name(self) -> str:
        return "DOCX: no comments"

    def applies_to(self):
        return [".docx"]

    def run(self, artifact: FileArtifact) -> CheckResult:
        # Guard: unreadable file
        unreadable = _unreadable(artifact)
        if unreadable:
            return unreadable

        count = int(artifact.metadata.get("comments_count", 0) or 0)
        passed = count == 0
        return CheckResult(
            file=artifact.path,
            check_name=self.name(),
            severity=Severity.ERROR if not passed else Severity.INFO,
            passed=passed,
            message="OK: no comments" if passed else f"Found {count} comment(s)",
            extra={"comments_count": count},
        )

#Checks for track changes
class DocxNoTrackedChangesCheck(Check):
    def name(self) -> str:
        return "DOCX: no tracked changes"

    def applies_to(self):
        return [".docx"]

    def run(self, artifact: FileArtifact) -> CheckResult:
        unreadable = _unreadable(artifact)
        if unreadable:
            return unreadable

        count = int(artifact.metadata.get("tracked_changes_count", 0) or 0)
        passed = count == 0
        return CheckResult(
            file=artifact.path,
            check_name=self.name(),
            severity=Severity.ERROR if not passed else Severity.INFO,
            passed=passed,
            message="OK: no tracked changes" if passed else f"Found {count} tracked change(s)",
            extra={"tracked_changes_count": count},
        )

#Checks for highlights
class DocxNoHighlightsCheck(Check):
    def name(self) -> str:
        return "DOCX: no highlights"

    def applies_to(self):
        return [".docx"]

    def run(self, artifact: FileArtifact) -> CheckResult:
        unreadable = _unreadable(artifact)
        if unreadable:
            return unreadable

        explicit_h = int(artifact.metadata.get("highlight_run_count", 0) or 0)
        shading_h = int(artifact.metadata.get("shading_highlight_count", 0) or 0)
        total = explicit_h + shading_h
        passed = total == 0

        return CheckResult(
            file=artifact.path,
            check_name=self.name(),
            severity=Severity.ERROR if not passed else Severity.INFO,
            passed=passed,
            message=(
                "OK: no highlights"
                if passed
                else f"Found {total} highlighted region(s) "
                     f"(explicit={explicit_h}, shading={shading_h})"
            ),
            extra={
                "highlight_run_count": explicit_h,
                "shading_highlight_count": shading_h,
                "total_highlight_like": total,
            },
        )


# Register all on import
register_check(DocxNoCommentsCheck())
register_check(DocxNoTrackedChangesCheck())
register_check(DocxNoHighlightsCheck())
