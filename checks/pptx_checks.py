# checks/pptx_checks.py
"""
PPTX checks:
1) No comments (counts <p:cm> across /ppt/comments/comment*.xml via metadata).

Notes:
- PowerPoint doesn't have tracked changes like Word; we skip that check.
- Returns WARNING if the file couldn't be read.
- Registers itself at import time.
"""

#Imports
from __future__ import annotations

from core.interfaces import Check
from core.models import CheckResult, Severity, FileArtifact
from core.registry import register_check

#Unreadable
def _unreadable(artifact: FileArtifact) -> CheckResult | None:
    if artifact.metadata.get("read_error"):
        return CheckResult(
            file=artifact.path,
            check_name="PPTX unreadable",
            severity=Severity.WARNING,
            passed=False,
            message="Unreadable PPTX (parse error)",
            extra={"detail": artifact.metadata.get("read_error_detail")},
        )
    return None

#Checks for any comments
class PptxNoCommentsCheck(Check):
    def name(self) -> str:
        return "PPTX: no comments"

    def applies_to(self):
        return [".pptx"]

    def run(self, artifact: FileArtifact) -> CheckResult:
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

# Register on import
register_check(PptxNoCommentsCheck())
