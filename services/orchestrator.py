# services/orchestrator.py
"""
High-level coordinator: discovers files, builds artifacts via processors,
runs applicable checks, and aggregates results.

This module deliberately depends only on:
- core.interfaces + core.models (abstractions)
- core.registry (to discover plugins)
- utils.path_utils (file discovery)

It does NOT import concrete processors or checks directly (DIP).
"""

from __future__ import annotations
import uuid
from collections import defaultdict
from datetime import datetime, timezone
from typing import Dict, List, Tuple
from pathlib import Path
from typing import Callable, Iterable, Optional, Dict
from core.registry import processors, checks
from utils.path_utils import iter_target_files
from core.models import (
    AggregateReport,
    CheckResult,
    FileArtifact,
    FileVerdict,
    FileReport,
    RunHeader,
    ScanReport,
    Severity,
)

# A tiny type alias for a UI-friendly progress callback:
# on_progress(current_index, total_files, current_path)
ProgressFn = Callable[[int, int, Path], None]


class Orchestrator:
    """
    Coordinates the end-to-end scan. Stateless aside from an optional
    progress callback retained for the duration of a run.

    Usage:
        orchestrator = Orchestrator(on_progress=my_progress_fn)
        report = orchestrator.run_scan("C:/docs")
    """

    def __init__(self, on_progress: Optional[ProgressFn] = None) -> None:
        self._on_progress = on_progress

    def run_scan(self, folder: str | Path) -> AggregateReport:
        
        """
        Walk the folder, route files to processors, run checks, and return an AggregateReport.

        - Converts unexpected processor/check exceptions into ERROR CheckResult entries.
        - Skips files with no matching processor (shouldn't happen if discovery filters are set).
        """
        root = Path(folder)
        file_list = list(iter_target_files(root))
        total = len(file_list)

        proc_index = self._index_processors()
        report = AggregateReport()

        for i, fpath in enumerate(file_list, start=1):
            # Inform the UI about progress if a callback was provided
            if self._on_progress:
                self._on_progress(i, total, fpath)

            ext = fpath.suffix.lower()
            processor = proc_index.get(ext)

            if processor is None:
                # Defensive: discovery normally filters to known extensions.
                report.results.append(
                    CheckResult(
                        file=fpath,
                        check_name="No processor found",
                        severity=Severity.WARNING,
                        passed=False,
                        message=f"No processor registered for extension: {ext}",
                        extra={"extension": ext},
                    )
                )
                continue

            artifact = self._safe_build_artifact(processor, fpath, report)
            if artifact is None:
                # Error already recorded as a CheckResult
                continue

            # Run only checks that claim to apply to this extension
            for chk in checks():
                if self._check_applies(chk.applies_to(), ext):
                    self._safe_run_check(chk, artifact, report)

        return report

    def run_scan_v2(self, folder: str | Path, config_snapshot: Dict = None) -> ScanReport:
        """
        File-centric scan that returns a ScanReport (header + files with verdicts).
        """
        start = datetime.now(timezone.utc)
        config_snapshot = dict(config_snapshot or {})

        # Reuse your v1 logic to get flat results
        flat = self.run_scan(folder)  # AggregateReport

        # Group by file
        by_file: Dict[Path, List[CheckResult]] = defaultdict(list)
        for r in flat.results:
            by_file[r.file].append(r)

        files_out: List[FileReport] = []
        total_errors = total_warnings = total_infos = 0

        # We need size/extension per file; derive from the first result via Path
        for file_path, results in by_file.items():
            # Counters
            errs = sum(1 for r in results if r.severity == Severity.ERROR and not r.passed)
            warns = sum(1 for r in results if r.severity == Severity.WARNING and not r.passed)
            infos = sum(1 for r in results if r.severity == Severity.INFO)

            # Verdict: any ERROR -> FAIL; else if any WARNING (failed) -> WARN; else PASS
            if errs > 0:
                verdict = FileVerdict.FAIL
            elif warns > 0:
                verdict = FileVerdict.WARN
            else:
                verdict = FileVerdict.PASS_

            # Try to get size quickly (filesystem) and extension
            try:
                size = file_path.stat().st_size
            except OSError:
                size = 0
            ext = file_path.suffix.lower()

            files_out.append(
                FileReport(
                    file=file_path,
                    extension=ext,
                    size_bytes=size,
                    verdict=verdict,
                    errors=errs,
                    warnings=warns,
                    infos=infos,
                    results=results,
                )
            )

            total_errors += errs
            total_warnings += warns
            total_infos += infos

        end = datetime.now(timezone.utc)
        header = RunHeader(
            schema_version="1.0-file-centric",
            run_id=str(uuid.uuid4()),
            root=Path(folder).resolve(),
            started_at_utc=start,
            finished_at_utc=end,
            config_snapshot=config_snapshot,
            total_files=len(files_out),
            total_checks=len(flat.results),
            total_errors=total_errors,
            total_warnings=total_warnings,
            total_infos=total_infos,
        )
        return ScanReport(header=header, files=sorted(files_out, key=lambda f: str(f.file).lower()))
    
    # ---------- helpers ----------

    def _index_processors(self) -> Dict[str, object]:
        """
        Build a mapping of extension -> processor instance for O(1) routing.
        If multiple processors claim the same extension, the last one wins.
        """
        index: Dict[str, object] = {}
        for p in processors():
            for ext in p.supports():
                index[ext.lower()] = p
        return index

    def _check_applies(self, targets: Iterable[str], ext: str) -> bool:
        """
        Return True if a check targets this extension.
        '*' means "applies to all".
        """
        lowered = [t.lower() for t in targets]
        return "*" in lowered or ext in lowered

    def _safe_build_artifact(
        self, processor, fpath: Path, report: AggregateReport
    ) -> Optional[FileArtifact]:
        """
        Build a FileArtifact and capture exceptions as structured ERROR results.
        """
        try:
            return processor.build_artifact(fpath)
        except Exception as exc:
            report.results.append(
                CheckResult(
                    file=fpath,
                    check_name="Artifact build failed",
                    severity=Severity.ERROR,
                    passed=False,
                    message=str(exc) or exc.__class__.__name__,
                    extra={"exception": exc.__class__.__name__},
                )
            )
            return None

    def _safe_run_check(self, chk, artifact: FileArtifact, report: AggregateReport) -> None:
        """
        Run a check and capture exceptions as structured ERROR results.
        """
        try:
            result = chk.run(artifact)
        except Exception as exc:
            # If the check crashes, we record that as an ERROR with the check's name.
            result = CheckResult(
                file=artifact.path,
                check_name=f"{chk.name()}",
                severity=Severity.ERROR,
                passed=False,
                message=f"Check raised exception: {exc}",
                extra={"exception": exc.__class__.__name__},
            )
        report.results.append(result)
