# core/interfaces.py
"""
Stable abstractions the rest of the app depends on.
High-level code (like the Orchestrator and UI) imports only these interfaces,
not any concrete PDF/DOCX/PPTX libraries.

Design notes (SOLID):
- SRP: Each interface has one clear purpose.
- OCP: New processors/checks plug in by implementing these contracts.
- LSP: Any subtype of these interfaces can be used interchangeably.
- ISP: Small, focused method sets (no "god interfaces").
- DIP: High-level policy depends on these abstractions, not on concretes.
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path
from typing import Iterable, Protocol
from .models import FileArtifact, CheckResult, ScanReport

__all__ = ["FileProcessor", "Check", "ResultWriter", "ScanReportWriter"]


class FileProcessor(ABC):
    """
    Extracts a lightweight, read-only FileArtifact from a given file path.

    Implementations should be fast and avoid heavy parsing unless the checks
    truly need it. Never mutate files — processors are read-only by design.
    """

    @abstractmethod
    def supports(self) -> Iterable[str]:
        """
        Return the file extensions this processor can handle.
        Example: [".pdf"] or [".docx", ".dotx"]

        The Orchestrator will build a lookup from extension -> processor.
        """
        raise NotImplementedError

    @abstractmethod
    def build_artifact(self, path: Path) -> FileArtifact:
        """
        Produce a FileArtifact with essential metadata for checks.

        Must not mutate the file. Raise a clear exception if the file
        cannot be read; callers may catch and convert that into a CheckResult.
        """
        raise NotImplementedError


class Check(ABC):
    """
    A single read-only validation rule that inspects a FileArtifact
    and returns a CheckResult. Keep implementations side-effect free.
    """

    @abstractmethod
    def name(self) -> str:
        """
        Stable, human-friendly name (used by UI and exports).
        Keep it short and descriptive, e.g., "PDF not encrypted".
        """
        raise NotImplementedError

    @abstractmethod
    def applies_to(self) -> Iterable[str]:
        """
        Return the file extensions this check supports.
        Use ["*"] to indicate it applies to all supported file types.
        """
        raise NotImplementedError

    @abstractmethod
    def run(self, artifact: FileArtifact) -> CheckResult:
        """
        Execute the rule on the given artifact and return a result.
        Should not raise for ordinary "failed check" cases—encode that
        as passed=False in the returned CheckResult with a clear message.
        """
        raise NotImplementedError


class ResultWriter(Protocol):
    """
    Optional sink for results (e.g., CSV/JSON exporters).
    Using a Protocol enables structural typing: any object that provides
    a compatible 'write' method can serve as a ResultWriter.
    """

    def write(self, results: Iterable[CheckResult]) -> None:
        ...

class ScanReportWriter(Protocol):
    def write(self, report: ScanReport) -> None: ...