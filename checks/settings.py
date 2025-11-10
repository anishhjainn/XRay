# checks/settings.py
"""
Run-time settings for checks (kept tiny and UI-agnostic).

Currently stores a single setting:
- modified_cutoff (the latest allowed modification datetime)

Design:
- Single responsibility: hold values checks need (not UI logic).
- Dependency inversion: Streamlit (or any UI) sets values here; checks read them.
- Values are stored as timezone-aware UTC datetimes for consistent comparisons.

API:
- set_modified_cutoff(value)  -> None
- get_modified_cutoff()       -> Optional[datetime]
- clear_modified_cutoff()     -> None
- is_cutoff_set()             -> bool
"""

from __future__ import annotations

from datetime import date, datetime, time, timezone
from typing import Optional, Union

# Internal module-level store (kept private).
# Always an aware UTC datetime or None.
__modified_cutoff_utc: Optional[datetime] = None


def set_modified_cutoff(value: Optional[Union[datetime, date, str]]) -> None:
    """
    Set the global 'modified cutoff' used by checks.

    Args:
        value:
            - datetime (naive or aware): naive is assumed to be in local time.
            - date: interpreted as "end of that day" (23:59:59.999999) in local time.
            - str: ISO-8601 string parsable by datetime.fromisoformat()
                   (e.g., '2025-01-31' or '2025-01-31T18:30:00+05:30').
            - None: clears the cutoff (checks will treat as 'no cutoff configured').

    Behavior:
        Value is normalized to a timezone-aware UTC datetime for stable comparisons.
    """
    global __modified_cutoff_utc

    if value is None:
        __modified_cutoff_utc = None
        return

    dt = _coerce_to_datetime(value)
    dt_aware = _ensure_aware_in_local_tz(dt)
    __modified_cutoff_utc = dt_aware.astimezone(timezone.utc)


def get_modified_cutoff() -> Optional[datetime]:
    """
    Return the stored cutoff as a timezone-aware UTC datetime, or None if not set.
    """
    return __modified_cutoff_utc


def clear_modified_cutoff() -> None:
    """Clear any stored cutoff."""
    set_modified_cutoff(None)


def is_cutoff_set() -> bool:
    """Convenience helper: True if a cutoff is configured."""
    return __modified_cutoff_utc is not None


# ---------- helpers (internal) ----------

def _coerce_to_datetime(value: Union[datetime, date, str]) -> datetime:
    """
    Convert supported input types to a datetime (may be naive).
    - date -> end of day (23:59:59.999999)
    - str  -> datetime.fromisoformat(...)
    """
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        # End of the given day
        return datetime.combine(value, time(23, 59, 59, 999_999))
    if isinstance(value, str):
        try:
            # Supports 'YYYY-MM-DD' and 'YYYY-MM-DDTHH:MM:SS[.ffff][±HH:MM]'
            return datetime.fromisoformat(value)
        except ValueError as exc:
            raise ValueError(
                "Could not parse cutoff string. Use ISO formats like "
                "'2025-01-31' or '2025-01-31T18:30:00+05:30'."
            ) from exc
    raise TypeError("Unsupported cutoff type. Use datetime, date, str, or None.")


def _ensure_aware_in_local_tz(dt: datetime) -> datetime:
    """
    If dt is naive, assume it is in the machine's local timezone.
    Return a timezone-aware datetime in that local timezone.
    """
    if dt.tzinfo is not None:
        return dt
    # Derive local tzinfo by asking 'now' for its zone.
    local_tz = datetime.now().astimezone().tzinfo
    return dt.replace(tzinfo=local_tz)
