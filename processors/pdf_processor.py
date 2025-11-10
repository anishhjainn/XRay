# processors/pdf_processor.py
"""
PDF technician (FileProcessor): builds a read-only FileArtifact for .pdf files.

New metadata:
- annots_summary: {Subtype -> count}  e.g. {"Text": 3, "Highlight": 5, "Ink": 1, ...}
- mod_date: ISO 8601 string parsed from PDF /ModDate (fallback handled by checks)

Requires: pypdf
"""

from pathlib import Path
from typing import Dict, Any, Optional
from datetime import datetime, timezone, timedelta
from pypdf import PdfReader
from pypdf.errors import PdfReadError

from core.interfaces import FileProcessor
from core.models import FileArtifact
from core.registry import register_processor


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
        }

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
                            # Ignore a bad/indirect annotation; keep scanning
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
