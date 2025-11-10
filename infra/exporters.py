# infra/exporters.py
"""
Writers for the file-centric ScanReport.

Exports:
- JSON (schema_version "1.0-file-centric"): header + files[] + embedded results[]
- CSV files table: one row per file (verdict + counts)
- CSV checks table: one row per (file × check)

Usage:
    from infra.exporters import JsonScanReportWriter, FilesCsvWriter, ChecksCsvWriter
    JsonScanReportWriter("scan.json").write(report)
    FilesCsvWriter("files.csv").write(report)
    ChecksCsvWriter("checks.csv").write(report)
"""

from __future__ import annotations

import json
from dataclasses import asdict, is_dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict

import csv

from core.interfaces import ScanReportWriter
from core.models import ScanReport, CheckResult


def _to_iso(obj):
    if isinstance(obj, datetime):
        return obj.isoformat()
    return obj


def _result_to_plain(r: CheckResult) -> Dict[str, Any]:
    return {
        "file": str(r.file),
        "check": r.check_name,
        "severity": r.severity.value,
        "passed": bool(r.passed),
        "message": r.message,
        "extra": r.extra,
    }


class JsonScanReportWriter(ScanReportWriter):
    def __init__(self, path: str | Path) -> None:
        self.path = Path(path)

    def write(self, report: ScanReport) -> None:
        payload = {
            "schema_version": report.header.schema_version,
            "header": {
                "run_id": report.header.run_id,
                "root": str(report.header.root),
                "started_at_utc": report.header.started_at_utc.isoformat(),
                "finished_at_utc": report.header.finished_at_utc.isoformat(),
                "config_snapshot": report.header.config_snapshot,
                "totals": {
                    "files": report.header.total_files,
                    "checks": report.header.total_checks,
                    "errors": report.header.total_errors,
                    "warnings": report.header.total_warnings,
                    "infos": report.header.total_infos,
                },
            },
            "files": [
                {
                    "file": str(f.file),
                    "extension": f.extension,
                    "size_bytes": f.size_bytes,
                    "verdict": f.verdict.value,
                    "counts": {
                        "errors": f.errors,
                        "warnings": f.warnings,
                        "infos": f.infos,
                    },
                    "results": [_result_to_plain(r) for r in f.results],
                }
                for f in report.files
            ],
        }
        self.path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


class FilesCsvWriter(ScanReportWriter):
    """
    Writes a file-level table:
    File,Extension,SizeBytes,Verdict,Errors,Warnings,Infos
    """
    def __init__(self, path: str | Path) -> None:
        self.path = Path(path)

    def write(self, report: ScanReport) -> None:
        with self.path.open("w", newline="", encoding="utf-8") as fp:
            w = csv.writer(fp)
            w.writerow(["File", "Extension", "SizeBytes", "Verdict", "Errors", "Warnings", "Infos"])
            for f in report.files:
                w.writerow([str(f.file), f.extension, f.size_bytes, f.verdict.value, f.errors, f.warnings, f.infos])


class ChecksCsvWriter(ScanReportWriter):
    """
    Writes a check-level table:
    File,Check,Severity,Passed,Message,Extra(JSON)
    """
    def __init__(self, path: str | Path) -> None:
        self.path = Path(path)

    def write(self, report: ScanReport) -> None:
        with self.path.open("w", newline="", encoding="utf-8") as fp:
            w = csv.writer(fp)
            w.writerow(["File", "Check", "Severity", "Passed", "Message", "Extra"])
            for f in report.files:
                for r in f.results:
                    w.writerow([
                        str(r.file),
                        r.check_name,
                        r.severity.value,
                        "TRUE" if r.passed else "FALSE",
                        r.message,
                        json.dumps(r.extra, ensure_ascii=False),
                    ])
