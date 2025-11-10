# processors/docx_processor.py
"""
DOCX technician (FileProcessor): builds a read-only FileArtifact for .docx files.

New metadata:
- comments_present: bool
- comments_count: int
- tracked_changes_count: int
- highlight_run_count: int          (w:highlight)
- shading_highlight_count: int      (w:shd with non-empty/non-auto fill)
- core_modified: ISO 8601 (already present earlier)

Implementation notes:
- Use python-docx for core properties (simple & stable).
- Use zipfile + iterparse to stream-read XML parts (fast, low memory).
"""

from __future__ import annotations
from zipfile import ZipFile, BadZipFile
from pathlib import Path
from typing import Any, Dict, Optional
from zipfile import ZipFile
from xml.etree.ElementTree import iterparse

from docx import Document
from docx.opc.exceptions import PackageNotFoundError

from core.interfaces import FileProcessor
from core.models import FileArtifact
from core.registry import register_processor


DOC_PARTS_TO_SCAN = (
    "word/document.xml",
    # headers / footers may contain highlights or tracked changes
    # We'll scan any present files matching these prefixes.
)


class DocxProcessor(FileProcessor):
    def supports(self):
        return [".docx"]

    def build_artifact(self, path: Path) -> FileArtifact:
        size = path.stat().st_size
        metadata: Dict[str, Any] = {
            "kind": "docx",
            "paragraph_count": None,
            "table_count": None,
            "core_author": None,
            "core_created": None,
            "core_modified": None,
            "comments_present": False,
            "comments_count": 0,
            "tracked_changes_count": 0,
            "highlight_run_count": 0,
            "shading_highlight_count": 0,
            "read_error": False,
        }

        # 1) Core props via python-docx (graceful if it fails)
        try:
            doc = Document(str(path))
            metadata["paragraph_count"] = len(doc.paragraphs)
            metadata["table_count"] = len(doc.tables)
            core = doc.core_properties
            metadata["core_author"] = getattr(core, "author", None)
            metadata["core_created"] = _safe_iso(getattr(core, "created", None))
            metadata["core_modified"] = _safe_iso(getattr(core, "modified", None))
        except Exception:
            # We'll still try XML scanning even if python-docx stumbles.
            metadata["paragraph_count"] = metadata.get("paragraph_count")
            metadata["table_count"] = metadata.get("table_count")

        # 2) XML scan inside the .docx (ZIP)
        try:
            with ZipFile(str(path)) as zf:
                # Comments (legacy + modern)
                c_count = 0
                # legacy comments: /word/comments.xml with <w:comment>
                if "word/comments.xml" in zf.namelist():
                    c_count += _count_tags_in_zip(zf, "word/comments.xml", {"comment"})
                # modern/extended comments: /word/commentsExtended.xml with w15:commentEx
                if "word/commentsExtended.xml" in zf.namelist():
                    c_count += _count_tags_in_zip(zf, "word/commentsExtended.xml", {"commentEx"})
                metadata["comments_count"] = c_count
                metadata["comments_present"] = c_count > 0

                # Tracked changes across main doc + headers/footers
                tracked_tags = {
                    "ins",
                    "del",
                    "moveFrom",
                    "moveTo",
                    "tblPrChange",
                    "trPrChange",
                    "tcPrChange",
                    "pPrChange",
                    "rPrChange",
                }
                tracked_total = 0

                # Main document
                if "word/document.xml" in zf.namelist():
                    tracked_total += _count_tags_in_zip(zf, "word/document.xml", tracked_tags)

                # Headers/footers (names vary: header1.xml, footer2.xml, etc.)
                for name in zf.namelist():
                    if name.startswith("word/header") and name.endswith(".xml"):
                        tracked_total += _count_tags_in_zip(zf, name, tracked_tags)
                    if name.startswith("word/footer") and name.endswith(".xml"):
                        tracked_total += _count_tags_in_zip(zf, name, tracked_tags)

                metadata["tracked_changes_count"] = tracked_total

                # Highlights: explicit w:highlight, plus shading w:shd fill
                highlight_count = 0
                shading_count = 0

                scan_targets = ["word/document.xml"]
                scan_targets += [n for n in zf.namelist() if n.startswith("word/header") and n.endswith(".xml")]
                scan_targets += [n for n in zf.namelist() if n.startswith("word/footer") and n.endswith(".xml")]

                for part in scan_targets:
                    h, s = _count_highlights_in_part(zf, part)
                    highlight_count += h
                    shading_count += s

                metadata["highlight_run_count"] = highlight_count
                metadata["shading_highlight_count"] = shading_count

        except (PackageNotFoundError, KeyError, ValueError, OSError, Exception) as exc:
            metadata["read_error"] = True
            metadata["read_error_detail"] = str(exc) or exc.__class__.__name__

        return FileArtifact(
            path=path,
            extension=".docx",
            size_bytes=size,
            metadata=metadata,
        )


# --- helpers ---

def _safe_iso(dt) -> Optional[str]:
    try:
        return dt.isoformat() if dt is not None else None
    except Exception:
        return None


def _count_tags_in_zip(zf: ZipFile, member: str, localname_set: set[str]) -> int:
    """
    Count elements whose tag's local-name is in localname_set.
    local-name = tag without namespace, e.g., '{ns}comment' -> 'comment'.
    """
    count = 0
    with zf.open(member) as fp:
        # 'events' = ("end",) is cheaper; clear elements as we go to save memory
        for _event, elem in iterparse(fp, events=("end",)):
            if _local_name(elem.tag) in localname_set:
                count += 1
            elem.clear()
    return count


def _count_highlights_in_part(zf: ZipFile, member: str) -> tuple[int, int]:
    """
    Return (highlight_count, shading_count) for a single XML part.
    - highlight_count: number of <w:highlight> occurrences
    - shading_count: number of <w:shd> with a non-empty/non-'auto' fill
    """
    h_count = 0
    s_count = 0
    with zf.open(member) as fp:
        for _event, elem in iterparse(fp, events=("end",)):
            lname = _local_name(elem.tag)
            if lname == "highlight":
                h_count += 1
            elif lname == "shd":
                # Attributes may be namespaced; accept any attr ending with 'fill'
                has_fill = False
                for k, v in elem.attrib.items():
                    if k.endswith("fill"):
                        val = (v or "").strip().lower()
                        if val and val not in {"auto", "none"}:
                            has_fill = True
                            break
                if has_fill:
                    s_count += 1
            elem.clear()
    return h_count, s_count


def _local_name(tag: str) -> str:
    """
    Strip the XML namespace from a tag: '{namespace}name' -> 'name'.
    """
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


# Register on import
register_processor(DocxProcessor())
