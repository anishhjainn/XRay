# processors/pptx_processor.py
"""
PPTX technician (FileProcessor): builds a read-only FileArtifact for .pptx files.

New metadata:
- comments_count: int (counts <p:cm> across /ppt/comments/comment*.xml)

Keeps existing fields like slide_count, total_shapes, has_any_notes, and core props.

Requires: python-pptx
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, Optional
from zipfile import ZipFile
from xml.etree.ElementTree import iterparse

from pptx import Presentation
from pptx.exc import PackageNotFoundError as PptxPackageNotFoundError

from core.interfaces import FileProcessor
from core.models import FileArtifact
from core.registry import register_processor


class PptxProcessor(FileProcessor):
    def supports(self):
        return [".pptx"]

    def build_artifact(self, path: Path) -> FileArtifact:
        size = path.stat().st_size
        metadata: Dict[str, Any] = {
            "kind": "pptx",
            "slide_count": None,
            "total_shapes": None,
            "has_any_notes": None,
            "core_author": None,
            "core_created": None,
            "core_modified": None,
            "comments_count": 0,
            "read_error": False,
        }

        # 1) High-level facts via python-pptx (slides, shapes, core props)
        try:
            prs = Presentation(str(path))
            slides = list(prs.slides)
            metadata["slide_count"] = len(slides)
            metadata["total_shapes"] = sum(len(s.shapes) for s in slides)
            try:
                metadata["has_any_notes"] = any(
                    getattr(s, "has_notes_slide", False) for s in slides
                )
            except Exception:
                def _has_notes(slide):
                    try:
                        return slide.notes_slide is not None
                    except Exception:
                        return False
                metadata["has_any_notes"] = any(_has_notes(s) for s in slides)

            core = prs.core_properties
            metadata["core_author"] = getattr(core, "author", None)
            metadata["core_created"] = _safe_iso(getattr(core, "created", None))
            metadata["core_modified"] = _safe_iso(getattr(core, "modified", None))
        except Exception:
            # Keep going; we can still count comments via ZIP scan
            pass

        # 2) Count comments via ZIP scan of /ppt/comments/comment*.xml (<p:cm>)
        try:
            with ZipFile(str(path)) as zf:
                total = 0
                for name in zf.namelist():
                    if name.startswith("ppt/comments/comment") and name.endswith(".xml"):
                        total += _count_tag_local(zf, name, "cm")
                metadata["comments_count"] = total
        except (PptxPackageNotFoundError, OSError, ValueError, KeyError, Exception) as exc:
            metadata["read_error"] = True
            metadata["read_error_detail"] = str(exc) or exc.__class__.__name__

        return FileArtifact(
            path=path,
            extension=".pptx",
            size_bytes=size,
            metadata=metadata,
        )


# --- helpers ---

def _safe_iso(dt) -> Optional[str]:
    try:
        return dt.isoformat() if dt is not None else None
    except Exception:
        return None


def _count_tag_local(zf: ZipFile, member: str, local_name: str) -> int:
    """
    Count how many elements with the given local-name appear in an XML part.
    For PPTX comments we look for local-name 'cm' inside comment parts.
    """
    count = 0
    with zf.open(member) as fp:
        for _event, elem in iterparse(fp, events=("end",)):
            tag = elem.tag
            if "}" in tag:
                tag = tag.split("}", 1)[1]
            if tag == local_name:
                count += 1
            elem.clear()
    return count


# Register on import
register_processor(PptxProcessor())
