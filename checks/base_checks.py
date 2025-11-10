# checks/base_checks.py
"""
Shared checks that apply across file types.

Includes:
- ModifiedBeforeCutoffCheck: ensure file was not modified after a user-provided cutoff.

Behavior:
- If no cutoff is configured -> passes (INFO) with a friendly message.
- Prefers document metadata timestamps (Office/PDF) and falls back to filesystem mtime.
- Produces ERROR if modified time is after the cutoff; INFO otherwise.
- Produces WARNING if we cannot determine any modified time.

Timezones:
- The cutoff stored in checks.settings is timezone-aware UTC.
- Document/FS times are parsed to timezone-aware datetimes (naive -> local tz).
"""

#all imports sit her
from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from typing import Optional, Tuple

from core.interfaces import Check
from core.models import CheckResult, Severity, FileArtifact
from core.registry import register_check

from .settings import get_modified_cutoff

#Modified Before cut off check.. takes the modified cut off from the settings.py module and
"""
Modified Before cut off check.. takes the modified cut off from the settings.py module and
first checks if there is any cut off provided. If not, it throws INFO saying that there is
no cutoff configured.

If the code is not able to determine the last modified date and time of the file then it throws
an error saying "Unable to determine last modified time" 

and lastly the pass variable is given a bool value and the true condition is if Observed_DT is before
cutoff_DT.  
"""

class ModifiedBeforeCutoffCheck(Check):
    """All files must be modified on/before the configured cutoff."""

    def name(self) -> str:
        return "Modified on or before cutoff"

    def applies_to(self):
        # Applies to all supported file types in this app.
        return [".docx", ".pptx", ".pdf"]

    def run(self, artifact: FileArtifact) -> CheckResult:
        cutoff_utc = get_modified_cutoff()
        if cutoff_utc is None:
            # No policy configured -> pass with INFO
            return CheckResult(
                file=artifact.path,
                check_name=self.name(),
                severity=Severity.INFO,
                passed=True,
                message="No cutoff configured",
                extra={"reason": "no_cutoff"},
            )

        observed_dt, source = _choose_modified_datetime(artifact)

        if observed_dt is None:
            # Could not determine any modified time at all
            return CheckResult(
                file=artifact.path,
                check_name=self.name(),
                severity=Severity.WARNING,
                passed=False,
                message="Unable to determine last modified time",
                extra={"source": None, "cutoff_utc": cutoff_utc.isoformat()},
            )

        observed_utc = observed_dt.astimezone(timezone.utc)
        passed = observed_utc <= cutoff_utc

        return CheckResult(
            file=artifact.path,
            check_name=self.name(),
            severity=Severity.INFO if passed else Severity.ERROR,
            passed=passed,
            message=(
                "OK: modified <= cutoff"
                if passed
                else "Modified after cutoff"
            ),
            extra={
                "observed_source": source,
                "observed_utc": observed_utc.isoformat(),
                "cutoff_utc": cutoff_utc.isoformat(),
            },
        )


# -------- helpers (module-internal) --------

def _choose_modified_datetime(artifact: FileArtifact) -> Tuple[Optional[datetime], Optional[str]]:
    """
    Return (datetime, source) where source ∈ {"doc_property", "pdf_mod_date", "fs_mtime"} or None.
    Priority:
        - DOCX/PPTX: core_modified (ISO) -> FS mtime
        - PDF: mod_date (ISO) -> FS mtime
    """
    ext = (artifact.extension or "").lower()

    # 1) Prefer document metadata
    if ext in {".docx", ".pptx"}:
        iso = _safe_str(artifact.metadata.get("core_modified"))
        dt = _parse_iso_to_aware(iso)
        if dt:
            return dt, "doc_property"
    elif ext == ".pdf":
        iso = _safe_str(artifact.metadata.get("mod_date"))
        dt = _parse_iso_to_aware(iso)
        if dt:
            return dt, "pdf_mod_date"

    # 2) Fallback to filesystem modified time
    dt_fs = _fs_mtime_to_aware(artifact.path)
    if dt_fs:
        return dt_fs, "fs_mtime"

    return None, None


def _parse_iso_to_aware(s: Optional[str]) -> Optional[datetime]:
    """
    Parse an ISO-like string to a timezone-aware datetime.
    - Accepts 'YYYY-MM-DD', 'YYYY-MM-DDTHH:MM:SS', with/without offset.
    - Treats trailing 'Z' as UTC.
    - If naive, assume local timezone.
    """
    if not s:
        return None
    s = s.strip()
    if not s:
        return None
    # Handle trailing 'Z' (UTC) for compatibility.
    if s.endswith("Z"):
        s = s[:-1] + "+00:00"
    try:
        dt = datetime.fromisoformat(s)
    except ValueError:
        return None
    if dt.tzinfo is None:
        dt = _attach_local_tz(dt)
    return dt


def _fs_mtime_to_aware(path: Path) -> Optional[datetime]:
    """Filesystem mtime as a timezone-aware datetime in local tz."""
    try:
        mtime = path.stat().st_mtime
    except OSError:
        return None
    local = datetime.fromtimestamp(mtime).astimezone()  # attaches local tz
    return local


def _attach_local_tz(dt: datetime) -> datetime:
    """Attach the machine's local tz to a naive datetime (no conversion)."""
    return dt.replace(tzinfo=datetime.now().astimezone().tzinfo)


def _safe_str(value) -> Optional[str]:
    return str(value) if value is not None else None


# Register this check on import
register_check(ModifiedBeforeCutoffCheck())
