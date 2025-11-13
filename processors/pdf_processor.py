# processors/pdf_processor.py
"""
PDF technician (FileProcessor): builds a read-only FileArtifact for .pdf files.

Existing metadata produced (consumed by current checks):
- encrypted: bool
- pages: int | None
- annots_summary: {Subtype -> count}  e.g. {"Text": 3, "Highlight": 5, "Ink": 1, ...}
- mod_date: ISO 8601 string parsed from PDF /ModDate (fallback handled by checks)
- read_error / read_error_detail on parse failures

New (Phase 3: for spelling now, grammar later):
- text_sample: str                    (normalized text extracted from pages, up to max_text_chars)
- text_length: int
- token_count: int
- text_extraction_error: bool
- text_extraction_error_detail: str | None

Design (SOLID):
- SRP: Processor extracts read-only facts (including a text sample); checks apply policy using those facts.
- OCP: Adding checks (spelling/grammar) requires no orchestrator changes; they just read metadata.
- LSP: Still honors FileProcessor; orchestrator treats it the same.
- ISP: Checks depend only on FileArtifact, not on pypdf internals.
- DIP: Checks depend on plain text (string) rather than concrete parsing libraries.

Notes:
- We do not mark the entire file unreadable if text extraction fails; we set text_extraction_error instead.
- For encrypted PDFs, pypdf reports reader.is_encrypted == True. We record 'encrypted' and skip annotation/page parsing.
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict, Any, Optional
from datetime import datetime, timezone, timedelta

from pypdf import PdfReader
from pypdf.errors import PdfReadError

from core.interfaces import FileProcessor
from core.models import FileArtifact
from core.registry import register_processor

# NEW: spelling/grammar helpers and config
from utils.text_extract import extract_pdf_text, tokenize_words
from infra.config_loader import load_config


class PdfProcessor(FileProcessor):
    def supports(self):
        return [".pdf"]

    def build_artifact(self, path: Path) -> FileArtifact:
        size = path.stat().st_size

        metadata: Dict[str, Any] = {
            "kind": "pdf",
            "encrypted": False,
            "pages": None,
            "annots_summary": {},
            "mod_date": None,
            "read_error": False,
            # NEW: text extraction for spelling/grammar checks
            "text_sample": "",
            "text_length": 0,
            "token_count": 0,
            "text_extraction_error": False,
            # "text_extraction_error_detail": "...",
        }

        # --- Core PDF parsing (annotations, pages, mod date) ---
        try:
            reader = PdfReader(str(path))
            encrypted = bool(getattr(reader, "is_encrypted", False))
            metadata["encrypted"] = encrypted

            if not encrypted:
                # Pages
                try:
                    metadata["pages"] = len(reader.pages)
                except Exception:
                    metadata["pages"] = None

                # Annotations summary (all kinds)
                annots: Dict[str, int] = {}
                for page in reader.pages:
                    try:
                        raw_annots = page.get("/Annots", []) or []
                    except Exception:
                        raw_annots = []
                    for ref in raw_annots:
                        try:
                            obj = ref.get_object()
                            subtype = obj.get("/Subtype")
                            if subtype is None:
                                continue
                            # subtype like NameObject('/Highlight') -> 'Highlight'
                            name = str(subtype)
                            if name.startswith("/"):
                                name = name[1:]
                            annots[name] = annots.get(name, 0) + 1
                        except Exception:
                            # Ignore malformed annotations; continue
                            continue
                metadata["annots_summary"] = annots

                # Modified date from document info (if present)
                try:
                    info = reader.metadata  # DocumentInformation or dict-like
                    raw_mod = None
                    if info:
                        raw_mod = info.get("/ModDate") or info.get("/ModificationDate")
                    metadata["mod_date"] = _pdf_date_to_iso(raw_mod)
                except Exception:
                    metadata["mod_date"] = None

        except (PdfReadError, OSError, ValueError, Exception) as exc:
            metadata["read_error"] = True
            metadata["read_error_detail"] = str(exc) or exc.__class__.__name__

        # --- NEW: Plain-text sample extraction (for spelling/grammar) ---
        # Keep this separate from core metadata parsing. If it fails, do not set read_error.
        try:
            cfg = load_config()
            if cfg.get("enable_spelling", True) or cfg.get("enable_grammar", False):
                max_chars = int(cfg.get("max_text_chars", 5_000_000))
                sample = extract_pdf_text(path, max_chars)
                metadata["text_sample"] = sample
                metadata["text_length"] = len(sample)
                metadata["token_count"] = len(tokenize_words(sample))
        except Exception as exc:
            metadata["text_sample"] = ""
            metadata["text_length"] = 0
            metadata["token_count"] = 0
            metadata["text_extraction_error"] = True
            metadata["text_extraction_error_detail"] = str(exc) or exc.__class__.__name__

        return FileArtifact(
            path=path,
            extension=".pdf",
            size_bytes=size,
            metadata=metadata,
        )


# --- helpers ---

def _pdf_date_to_iso(raw: Optional[str]) -> Optional[str]:
    """
    Parse PDF date strings like 'D:YYYYMMDDHHmmSS+HH'mm'' to ISO 8601.
    Returns None if parsing fails or raw is falsy.
    """
    if not raw:
        return None
    s = raw.strip()
    if s.startswith("D:"):
        s = s[2:]

    # Minimal tolerant parse:
    # YYYY MM DD HH mm SS O HH ' mm
    # Lengths vary; we'll slice what we can find.
    def take(n, default="00"):
        nonlocal s
        if len(s) >= n:
            part, s = s[:n], s[n:]
            return part
        part, s = s, ""
        return (part + default)[:n]

    try:
        year = int(take(4, "0000"))
        month = int(take(2, "01"))
        day = int(take(2, "01"))
        hour = int(take(2, "00"))
        minute = int(take(2, "00"))
        second = int(take(2, "00"))

        # Timezone offset if present: e.g., +05'30', -08'00'
        tz = timezone.utc
        if s:
            sign = s[0]
            if sign in "+-":
                s = s[1:]
                hh = 0
                mm = 0
                if "'" in s:
                    # format HH'mm'
                    try:
                        hh_str, rest = s.split("'", 1)
                        mm_str = rest.split("'")[0]
                        hh = int(hh_str or "0")
                        mm = int(mm_str or "0")
                    except Exception:
                        hh = mm = 0
                else:
                    # format HHmm (rare)
                    hh = int((s[:2] or "0"))
                    mm = int((s[2:4] or "0"))
                delta = timedelta(hours=hh, minutes=mm)
                if sign == "-":
                    delta = -delta
                tz = timezone(delta)

        dt = datetime(year, month, day, hour, minute, second, tzinfo=tz)
        return dt.astimezone(timezone.utc).isoformat()
    except Exception:
        return None


# Register on import
register_processor(PdfProcessor())