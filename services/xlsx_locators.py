# services/xlsx_locators.py
"""
On-demand XLSX locators (XML streaming).

Public API:
    list_yellow_cells(path: Path) -> list[tuple[str, str]]

Returns (sheet_name, cell_ref) rows for cells with *classic* Excel yellow fill:
- Solid pattern fill with RGB 'FFFF00' (normalized from ARGB too) OR
- Indexed fill '6' (Excel built-in yellow)

Notes:
- This detects *statically applied* fills (cellXfs -> fillId).
- Conditional formats (DXF) are ignored. Theme+tint colors are resolved when
  they evaluate to the classic yellow RGB.
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Tuple
from zipfile import ZipFile, BadZipFile
from xml.etree.ElementTree import iterparse

from services.xlsx_theme import resolve_theme_color, theme_rgb_map_from_zip

# Normalized 6-hex yellow and indexed yellow
_YELLOW_6_HEX = {"FFFF00"}  # compare after normalization to last 6 hex chars
_YELLOW_INDEXED = {"6"}     # standard indexed yellow (5 is typically blue)


# ------------------------------- Public API ----------------------------------

def list_yellow_cells(path: Path) -> List[Tuple[str, str]]:
    """
    Return a list of (sheet_name, cell_ref) tuples for cells with classic yellow fill.
    Raises:
        FileNotFoundError, BadZipFile, KeyError (if critical parts missing)
    """
    p = Path(path)
    if not p.exists() or not p.is_file():
        raise FileNotFoundError(f"No such file: {p}")

    results: List[Tuple[str, str]] = []

    with ZipFile(str(p)) as zf:
        names = set(zf.namelist())

        # Build 'xl/worksheets/sheetN.xml' -> 'Display Name'
        theme_rgb_map = theme_rgb_map_from_zip(zf, names)
        part_to_name = _map_worksheet_parts_to_names(zf, names)

        # Build mapping: style index -> fillId, and which fillIds are yellow
        style_to_fill, yellow_fill_ids = _parse_styles_for_yellow(zf, names, theme_rgb_map)
        if not style_to_fill or not yellow_fill_ids:
            # No styles or no yellow fills present => nothing to report
            return results

        # Walk each worksheet and collect yellow cells
        for member in sorted(n for n in names if n.startswith("xl/worksheets/") and n.endswith(".xml")):
            sheet_name = part_to_name.get(member, _basename_no_ext(member))
            results.extend(
                _yield_sheet_cells_with_yellow(zf, member, sheet_name, style_to_fill, yellow_fill_ids)
            )

    return results


# ----------------------------- Internal helpers ------------------------------

def _map_worksheet_parts_to_names(zf: ZipFile, names: set[str]) -> Dict[str, str]:
    """
    Map each worksheet part path to its display name:
        "xl/worksheets/sheet1.xml" -> "Sheet1"
    We parse:
      - xl/_rels/workbook.xml.rels : rId -> Target (worksheets/sheet*.xml)
      - xl/workbook.xml            : <sheet name="..." r:id="rIdN">
    """
    mapping: Dict[str, str] = {}
    if "xl/workbook.xml" not in names:
        return mapping

    # Resolve relationships (rId -> xl/worksheets/sheet*.xml)
    rels: Dict[str, str] = {}
    if "xl/_rels/workbook.xml.rels" in names:
        with zf.open("xl/_rels/workbook.xml.rels") as fp:
            for _event, elem in iterparse(fp, events=("end",)):
                if _local(elem.tag) == "Relationship":
                    rid = elem.attrib.get("Id")
                    tgt = elem.attrib.get("Target", "")
                    if rid and tgt:
                        if not tgt.startswith("xl/"):
                            tgt = f"xl/{tgt.lstrip('./')}"
                        rels[rid] = tgt
                elem.clear()

    with zf.open("xl/workbook.xml") as fp:
        for _event, elem in iterparse(fp, events=("end",)):
            if _local(elem.tag) == "sheet":
                name = elem.attrib.get("name")
                rid = (
                    elem.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                    or elem.attrib.get("r:id")
                )
                if name and rid:
                    tgt = rels.get(rid)
                    if tgt and tgt.startswith("xl/worksheets/"):
                        mapping[tgt] = name
            elem.clear()

    return mapping


def _parse_styles_for_yellow(
    zf: ZipFile,
    names: set[str],
    theme_rgb_map: Dict[int, str],
) -> tuple[Dict[int, int], set[int]]:
    """
    Parse xl/styles.xml to build:
      - style_to_fill: cellXfs index (cell style 's' on <c>) -> fillId (int; defaults to 0 when absent)
      - yellow_fill_ids: set of fillIds considered classic yellow
        (solid pattern fill, fg/bg RGB FFFF00 after normalization OR indexed=6)
    """
    style_to_fill: Dict[int, int] = {}
    yellow_fill_ids: set[int] = set()

    if "xl/styles.xml" not in names:
        return style_to_fill, yellow_fill_ids

    in_fills = False
    in_cellxfs = False
    fill_id = -1
    xf_idx = -1

    with zf.open("xl/styles.xml") as fp:
        for _event, elem in iterparse(fp, events=("end",)):
            tag = _local(elem.tag)

            if tag == "fills":
                in_fills = True
            elif tag == "cellXfs":
                in_cellxfs = True

            # Collect fillId -> is_yellow?
            if in_fills and tag == "fill":
                fill_id += 1
                # Look for <patternFill patternType="solid"><fgColor/bgColor .../></patternFill>
                for child in list(elem):
                    if _local(child.tag) == "patternFill" and child.attrib.get("patternType") == "solid":
                        for sub in list(child):
                            if _local(sub.tag) in ("fgColor", "bgColor"):
                                if _color_attrs_are_classic_yellow(sub.attrib, theme_rgb_map):
                                    yellow_fill_ids.add(fill_id)
                elem.clear()
                continue

            # Collect cellXfs index -> fillId (default to 0 if missing)
            if in_cellxfs and tag == "xf":
                xf_idx += 1
                fid = elem.attrib.get("fillId")
                try:
                    style_to_fill[xf_idx] = int(fid) if fid is not None else 0
                except ValueError:
                    style_to_fill[xf_idx] = 0
                elem.clear()
                continue

            elem.clear()

    return style_to_fill, yellow_fill_ids


def _yield_sheet_cells_with_yellow(
    zf: ZipFile,
    member: str,
    sheet_name: str,
    style_to_fill: Dict[int, int],
    yellow_fill_ids: set[int],
) -> List[tuple[str, str]]:
    """
    Stream a worksheet XML and collect cells whose style's fillId resolves to yellow.
    We rely only on cell attributes:
      - r: cell reference (e.g., 'B12')
      - s: style index (maps via cellXfs to a fillId)
    """
    out: List[tuple[str, str]] = []
    if not style_to_fill or not yellow_fill_ids:
        return out

    with zf.open(member) as fp:
        for _event, elem in iterparse(fp, events=("end",)):
            if _local(elem.tag) == "c":  # <c r="B12" s="17" ...>
                ref = elem.attrib.get("r")
                s = elem.attrib.get("s")
                if s is not None and ref:
                    try:
                        style_idx = int(s)
                    except ValueError:
                        style_idx = None

                    if style_idx is not None:
                        # Default fillId to 0 if style not found (Excel default)
                        fill_id = style_to_fill.get(style_idx, 0)
                        if fill_id in yellow_fill_ids:
                            out.append((sheet_name, ref))
                elem.clear()
            else:
                elem.clear()
    return out


def _local(tag: str) -> str:
    """Strip XML namespace to get the local tag name."""
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def _normalize_rgb(rgb: str) -> str:
    """
    Normalize ARGB/RGB to last 6 hex digits so 'FFFFFFFF' -> 'FFFFFF', '00FFFF00' -> 'FFFF00'.
    """
    s = (rgb or "").upper()
    return s[-6:] if len(s) >= 6 else s


def _color_attrs_are_classic_yellow(attrs: Dict[str, str], theme_rgb_map: Dict[int, str]) -> bool:
    """
    Decide whether a <fgColor>/<bgColor> element describes classic Excel yellow.
    Considers rgb, indexed, and theme+tint representations.
    """
    rgb_norm = _normalize_rgb(attrs.get("rgb", ""))
    if rgb_norm in _YELLOW_6_HEX:
        return True

    idx = attrs.get("indexed")
    if idx in _YELLOW_INDEXED:
        return True

    theme_idx = attrs.get("theme")
    if theme_idx is not None:
        resolved = resolve_theme_color(theme_rgb_map, theme_idx, attrs.get("tint"))
        if resolved and _normalize_rgb(resolved) in _YELLOW_6_HEX:
            return True

    return False


def _basename_no_ext(member: str) -> str:
    """Fallback sheet name if mapping is unavailable."""
    base = Path(member).name
    if base.lower().endswith(".xml"):
        return base[:-4]
    return base
