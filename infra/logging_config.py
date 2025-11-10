# infra/logging_config.py
"""
Centralized logging setup with rotating file handler.
"""

from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Optional


def configure_logging(level_name: str = "INFO", log_dir: Optional[Path] = None) -> None:
    """
    Configure console + rotating file logging.

    Args:
        level_name: "DEBUG" | "INFO" | "WARNING" | "ERROR"
        log_dir: optional custom log directory (defaults to ./logs)
    """
    level = getattr(logging, level_name.upper(), logging.INFO)
    log_dir = log_dir or (Path.cwd() / "logs")
    log_dir.mkdir(parents=True, exist_ok=True)

    fmt = logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s")

    # Console
    console = logging.StreamHandler()
    console.setLevel(level)
    console.setFormatter(fmt)

    # Rotating file (5 MB x 3 files)
    file_handler = RotatingFileHandler(log_dir / "app.log", maxBytes=5_000_000, backupCount=3, encoding="utf-8")
    file_handler.setLevel(level)
    file_handler.setFormatter(fmt)

    root = logging.getLogger()
    # Avoid duplicate handlers if user reruns the Streamlit script
    if not root.handlers:
        root.setLevel(level)
        root.addHandler(console)
        root.addHandler(file_handler)
    else:
        # Update existing handler levels on rerun
        for h in root.handlers:
            h.setLevel(level)
