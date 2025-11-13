# checks/spelling_checks.py
"""
SpellingCheck: flags misspelled words across .docx, .pptx, .xlsx, and .pdf files.

Design (SOLID):
- SRP: This check ONLY evaluates spelling over a pre-extracted text sample
  (no file I/O, no parsing libraries).
- OCP: Adding/removing checks requires no orchestrator changes (self-registers).
- LSP: Conforms to the Check interface; orchestrator treats it like any other check.
- ISP: Depends only on FileArtifact and config; not on processor internals.
- DIP: High-level flow depends on the Check abstraction, and this check depends
  on plain text provided by processors (not on concrete parsing libs).

How it works:
- Processors populate artifact.metadata["text_sample"] (normalized text) and related fields.
- This check tokenizes words and uses pyspellchecker to find unknown words.
- Severity policy (configurable):
    0 misspellings                     -> INFO,   passed=True
    1..spelling_fail_threshold (10)    -> WARNING, passed=False
    > spelling_fail_threshold          -> ERROR,  passed=False
"""

from __future__ import annotations

from typing import List, Dict, Any

from core.interfaces import Check
from core.models import FileArtifact, CheckResult, Severity
from core.registry import register_check

from infra.config_loader import load_config
from utils.text_extract import tokenize_words

try:
    # pyspellchecker is a small, pure-Python dependency
    from spellchecker import SpellChecker
except Exception:
    SpellChecker = None  # We will handle missing dependency gracefully


class SpellingCheck(Check):
    def name(self) -> str:
        return "spelling"

    def applies_to(self) -> List[str]:
        # Applies to all current text-bearing types
        return [".docx", ".pptx", ".xlsx", ".pdf"]

    def run(self, artifact: FileArtifact) -> CheckResult:
        cfg = load_config()

        # 1) Feature toggle: allow disabling via config without code changes.
        if not cfg.get("enable_spelling", True):
            return CheckResult(
                file=artifact.path,
                check_name=self.name(),
                severity=Severity.INFO,
                passed=True,
                message="Spelling check disabled by config",
                extra={"reason": "disabled"},
            )

        # 2) Guard: dependency presence
        if SpellChecker is None:
            return CheckResult(
                file=artifact.path,
                check_name=self.name(),
                severity=Severity.WARNING,
                passed=False,
                message="Spelling engine unavailable (pyspellchecker not installed)",
                extra={"reason": "missing_dependency"},
            )

        meta: Dict[str, Any] = artifact.metadata or {}
        text = meta.get("text_sample") or ""
        text_len = int(meta.get("text_length") or len(text) or 0)

        # If processors couldn't provide text (e.g., encrypted/unreadable), warn gracefully.
        if not text:
            return CheckResult(
                file=artifact.path,
                check_name=self.name(),
                severity=Severity.WARNING,
                passed=False,
                message="No text available for spelling check",
                extra={
                    "text_length": text_len,
                    "text_extraction_error": bool(meta.get("text_extraction_error", False)),
                    "text_extraction_error_detail": meta.get("text_extraction_error_detail"),
                },
            )

        # 3) Tokenize and find misspellings
        tokens = tokenize_words(text)  # lowercased word-ish tokens for English
        unique_tokens = set(tokens)

        language_code = (cfg.get("language_code") or "en").strip() or "en"
        try:
            sp = SpellChecker(language=language_code)
        except Exception:
            # If the requested language is not available, fall back to English
            sp = SpellChecker(language="en")
            language_code = "en"

        misspelled = sp.unknown(unique_tokens)  # set of unknown words among unique tokens
        misspelling_count = len(misspelled)

        # 4) Severity policy (configurable thresholds)
        threshold = int(cfg.get("spelling_fail_threshold", 10) or 10)
        max_list = int(cfg.get("max_misspellings_reported", 100) or 100)

        if misspelling_count == 0:
            severity = Severity.INFO
            passed = True
            msg = "OK: no misspellings"
        elif misspelling_count <= threshold:
            severity = Severity.WARNING
            passed = False
            msg = f"Found {misspelling_count} misspelling(s)"
        else:
            severity = Severity.ERROR
            passed = False
            msg = f"Found {misspelling_count} misspelling(s) (over threshold={threshold})"

        # Keep the payload small and stable
        sample_misspellings = sorted(list(misspelled))[:max_list]

        return CheckResult(
            file=artifact.path,
            check_name=self.name(),
            severity=severity,
            passed=passed,
            message=msg,
            extra={
                "language_code": language_code,
                "total_tokens": len(tokens),
                "unique_misspellings_count": misspelling_count,
                "sample_misspellings": sample_misspellings,
                "text_length": text_len,
                "capped_list_at": max_list,
            },
        )


# Register on import so the orchestrator discovers it via the registry.
register_check(SpellingCheck())