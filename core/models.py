# core/models.py
"""
Plain data shapes for the app (no UI, no parsing, no I/O).
These are the “contracts” that processors and checks speak.

Design goals:
- Minimal and framework-agnostic (easy to test and reason about).
- Immutable where it matters (to avoid accidental mutation).
- Extensible via generic metadata fields (no model changes for new checks).
"""

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any, Dict, List, Optional
from datetime import datetime

class Severity(Enum):
    """
    How serious a finding is. Enums give us a small, typo-proof vocabulary
    that the UI can style consistently (e.g., ERROR in red, WARNING in amber).
    """
    INFO = "INFO"
    WARNING = "WARNING"
    ERROR = "ERROR"


@dataclass(frozen=True)
class FileArtifact:
    """
    Lightweight, read-only description of a file produced by a FileProcessor.

    Why frozen=True?
    - Processors build this once. Checks must not mutate it.
    - Immutability avoids bugs where one check silently changes data for another.

    Fields:
    - path: absolute path to the file (Path is cross-platform friendly).
    - extension: normalized extension like ".pdf" (lowercase).
    - size_bytes: file size for generic checks (e.g., size thresholds).
    - metadata: a flexible bag for processor-specific facts, e.g.:
        PDF: {"encrypted": True, "pages": 12}
        DOCX: {"has_tracked_changes": False}
        PPTX: {"slide_count": 20}
    """
    path: Path
    extension: str
    size_bytes: int
    metadata: Dict[str, Any] = field(default_factory=dict)



@dataclass(frozen=True)
class CheckResult:
    """
    Outcome of running a single Check on a single FileArtifact.

    Why frozen=True?
    - A result is historical truth; keep it immutable after creation.

    Fields:
    - file: the file that was checked (Path for convenience).
    - check_name: stable, human-readable name (good for UI & exports).
    - severity: INFO/WARNING/ERROR to guide attention in the UI.
    - passed: True/False outcome for simple filtering.
    - message: short human explanation (single line preferred).
    - extra: structured details for power users/exports (counts, flags, etc.).
    """
    file: Path
    check_name: str
    severity: Severity
    passed: bool
    message: str
    extra: Dict[str, Any] = field(default_factory=dict)


@dataclass
class AggregateReport:
    """
    Container for all results from a scan. Mutable is fine here because
    the orchestrator will append results as it processes files.
    """
    results: List[CheckResult] = field(default_factory=list)


class FileVerdict(Enum):
    PASS_ = "PASS"   # trailing underscore to avoid keyword conflicts if used as attr
    WARN = "WARN"
    FAIL = "FAIL"


@dataclass(frozen=True)
class FileReport: 
    """
    All results related to a single file, plus a verdict and quick counts.
    """
    file: Path
    extension: str
    size_bytes: int
    verdict: FileVerdict
    errors: int
    warnings: int
    infos: int
    results: List["CheckResult"] = field(default_factory=list)


@dataclass(frozen=True)
class RunHeader:
    """
    Metadata about a scan run to make exports self-describing.
    """
    schema_version: str                 # e.g., "1.0-file-centric"
    run_id: str                         # uuid4 string
    root: Path                          # scanned folder
    started_at_utc: datetime            # timezone-aware
    finished_at_utc: datetime           # timezone-aware
    config_snapshot: Dict[str, Any]     # whatever config the UI used

    total_files: int
    total_checks: int
    total_errors: int
    total_warnings: int
    total_infos: int


@dataclass(frozen=True)
class ScanReport:
    """
    The new top-level report: header + file-centric details.
    """
    header: RunHeader
    files: List[FileReport] = field(default_factory=list)