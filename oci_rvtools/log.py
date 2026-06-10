# Copyright (c) 2026 Kim Tholstorf
# https://github.com/KimTholstorf/oci-rvtools-cost-estimator
# MIT License — see LICENSE file for details

"""Lightweight logging built on the stdlib ``logging`` module.

Output is intentionally identical to the previous ``print``-based helpers
(``[INFO] message`` / ``[WARN] message`` on stdout) so the CLI and the
browser (Pyodide) front-end keep showing the same console text. Using the
stdlib logger underneath means callers can adjust levels or attach their own
handlers when embedding the package.
"""

from __future__ import annotations

import logging
import sys

LOGGER_NAME = "oci_rvtools"

# Render WARNING as "WARN" to preserve the historical "[WARN]" prefix.
logging.addLevelName(logging.WARNING, "WARN")

logger = logging.getLogger(LOGGER_NAME)


def _ensure_configured() -> None:
    """Attach a stdout handler once, formatted as ``[LEVEL] message``."""
    if logger.handlers:
        return
    handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(logging.Formatter("[%(levelname)s] %(message)s"))
    logger.addHandler(handler)
    logger.setLevel(logging.INFO)
    logger.propagate = False


def info(msg: str) -> None:
    _ensure_configured()
    logger.info(msg)


def warn(msg: str) -> None:
    _ensure_configured()
    logger.warning(msg)


def error(msg: str) -> None:
    """Emit an error to stderr with the historical ``[ERROR]`` prefix."""
    print(f"[ERROR] {msg}", file=sys.stderr)
