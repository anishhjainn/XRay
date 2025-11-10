# checks/pdf_checks.py
"""
PDF checks:
1) No comments      -> fail if Text/FreeText annotations exist.
2) No highlights    -> fail if ANY annotation other than Text/FreeText exists
                       (you asked to treat all kinds of annotations as disallowed
                        in this rule, not just text markups).

Edge cases:
- If the PDF is encrypted or unreadable, we cannot inspect annotations reliably:
  emit a WARNING with passed=False.

Relies on PdfProcessor to provide:
- metadata["encrypted"] (bool)
- metadata["annots_summary"] (dict of Subtype -> count, names without leading '/')
"""

from __future__ import annotations

from typing import Dict

from core.interfaces import Check
from core.models import CheckResult, Severity, FileArtifact
from core.registry import register_check

#Unreadable checks
def _unreadable_or_encrypted(artifact: FileArtifact) -> CheckResult | None:
    meta = artifact.metadata
    if meta.get("read_error"):
        return CheckResult(
            file=artifact.path,
            check_name="PDF unreadable",
            severity=Severity.WARNING,
            passed=False,
            message="Unreadable PDF (parse error)",
            extra={"detail": meta.get("read_error_detail")},
        )
    if meta.get("encrypted"):
        return CheckResult(
            file=artifact.path,
            check_name="PDF encrypted",
            severity=Severity.WARNING,
            passed=False,
            message="Encrypted PDF: annotations cannot be verified",
            extra={"encrypted": True},
        )
    return None

#checks for annotations
def _get_annots(meta: Dict) -> Dict[str, int]:
    ann = meta.get("annots_summary") or {}
    # Ensure keys are strings and values are ints
    return {str(k): int(v) for k, v in ann.items()}


#checks for any comments
class PdfNoCommentsCheck(Check):
    def name(self) -> str:
        return "PDF: no comments"

    def applies_to(self):
        return [".pdf"]

    def run(self, artifact: FileArtifact) -> CheckResult:
        early = _unreadable_or_encrypted(artifact)
        if early:
            return early

        ann = _get_annots(artifact.metadata)
        comment_count = ann.get("Text", 0) + ann.get("FreeText", 0)
        passed = comment_count == 0

        return CheckResult(
            file=artifact.path,
            check_name=self.name(),
            severity=Severity.ERROR if not passed else Severity.INFO,
            passed=passed,
            message="OK: no comments" if passed else f"Found {comment_count} comment annotation(s)",
            extra={"comments_Text": ann.get("Text", 0), "comments_FreeText": ann.get("FreeText", 0)},
        )

#Checks for any highlights
class PdfNoHighlightsCheck(Check):
    def name(self) -> str:
        return "PDF: no highlights/markups/annotations"

    def applies_to(self):
        return [".pdf"]

    def run(self, artifact: FileArtifact) -> CheckResult:
        early = _unreadable_or_encrypted(artifact)
        if early:
            return early

        ann = _get_annots(artifact.metadata)

        # Treat ALL annotation subtypes as disallowed for this rule,
        # EXCEPT Text/FreeText (which are handled by the comments check).
        disallowed = 0
        details = {}

        for subtype, count in ann.items():
            if subtype in {"Text", "FreeText"}:
                continue  # counted by PdfNoCommentsCheck
            if count:
                disallowed += count
                details[subtype] = count

        passed = disallowed == 0

        return CheckResult(
            file=artifact.path,
            check_name=self.name(),
            severity=Severity.ERROR if not passed else Severity.INFO,
            passed=passed,
            message=(
                "OK: no highlights/markups/annotations"
                if passed
                else f"Found {disallowed} non-comment annotation(s)"
            ),
            extra={"disallowed_annotations": details, "all_annotations": ann},
        )

# Register on import
register_check(PdfNoCommentsCheck())
register_check(PdfNoHighlightsCheck())
