"""
Helpers for resolving Excel theme colors (theme + tint) to concrete RGB hex.

Provides:
    - theme_rgb_map_from_zip / theme_rgb_map_from_path: parse theme parts.
    - resolve_theme_color: turn a theme index + tint into an RGB string.
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict, Iterable, Optional
from zipfile import ZipFile, BadZipFile
from xml.etree.ElementTree import ParseError, fromstring

_THEME_ORDER = [
    "lt1",
    "dk1",
    "lt2",
    "dk2",
    "accent1",
    "accent2",
    "accent3",
    "accent4",
    "accent5",
    "accent6",
    "hlink",
    "folHlink",
]
_THEME_TAG_TO_INDEX = {name: idx for idx, name in enumerate(_THEME_ORDER)}


def _local_name(tag: str) -> str:
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def _normalize_rgb(value: str | None) -> str:
    s = (value or "").upper()
    return s[-6:] if len(s) >= 6 else s


def _apply_tint(rgb: str, tint: float) -> str:
    """
    Apply Excel's tint math to an RGB hex string (without alpha).
    tint in [-1, 1]; negative darkens, positive lightens.
    """
    base = _normalize_rgb(rgb)
    if len(base) != 6:
        return base

    try:
        tint_f = float(tint)
    except (TypeError, ValueError):
        tint_f = 0.0

    def _adjust(component: int) -> int:
        if tint_f < 0:
            value = component * (1.0 + tint_f)
        else:
            value = component + (255 - component) * tint_f
        return max(0, min(255, int(round(value))))

    r = _adjust(int(base[0:2], 16))
    g = _adjust(int(base[2:4], 16))
    b = _adjust(int(base[4:6], 16))
    return f"{r:02X}{g:02X}{b:02X}"


def _extract_rgb_from_theme_elem(elem) -> Optional[str]:
    """
    Theme color nodes contain either <a:srgbClr val="..."> or
    <a:sysClr lastClr="...">. Return whichever is available.
    """
    for child in list(elem):
        lname = _local_name(child.tag)
        if lname == "srgbClr":
            val = child.attrib.get("val")
            if val:
                return val
        elif lname == "sysClr":
            val = child.attrib.get("lastClr")
            if val:
                return val
    return None


def theme_rgb_map_from_zip(zf: ZipFile, names: Iterable[str]) -> Dict[int, str]:
    """
    Build a mapping theme_index -> RGB (string without alpha) from a ZipFile.
    """
    names = list(names)
    theme_part = None
    for candidate in names:
        if candidate.startswith("xl/theme/") and candidate.endswith(".xml"):
            theme_part = candidate
            break

    if not theme_part:
        return {}

    mapping: Dict[int, str] = {}
    try:
        with zf.open(theme_part) as fp:
            data = fp.read()
    except KeyError:
        return {}

    try:
        root = fromstring(data)
    except ParseError:
        return {}

    clr_scheme = None
    for elem in root.iter():
        if _local_name(elem.tag) == "clrScheme":
            clr_scheme = elem
            break

    if clr_scheme is None:
        return {}

    for child in clr_scheme:
        lname = _local_name(child.tag)
        if lname in _THEME_TAG_TO_INDEX:
            rgb = _extract_rgb_from_theme_elem(child)
            if rgb:
                mapping[_THEME_TAG_TO_INDEX[lname]] = _normalize_rgb(rgb)

    return mapping


def theme_rgb_map_from_path(path: Path) -> Dict[int, str]:
    """
    Convenience wrapper: open the .xlsx as ZIP and return the theme map.
    """
    try:
        with ZipFile(str(path)) as zf:
            names = zf.namelist()
            return theme_rgb_map_from_zip(zf, names)
    except (BadZipFile, OSError, FileNotFoundError):
        return {}


def resolve_theme_color(theme_map: Dict[int, str], theme_idx, tint=None) -> Optional[str]:
    """
    Given a theme index (int/str) and optional tint, return the resolved RGB string.
    """
    if not theme_map:
        return None

    try:
        idx = int(theme_idx)
    except (TypeError, ValueError):
        return None

    base = theme_map.get(idx)
    if not base:
        return None

    tint_val = 0.0
    if tint not in (None, "", 0, 0.0):
        try:
            tint_val = float(tint)
        except (TypeError, ValueError):
            tint_val = 0.0

    rgb = _normalize_rgb(base)
    if tint_val == 0.0:
        return rgb
    return _apply_tint(rgb, tint_val)
