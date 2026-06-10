# Copyright (c) 2026 Kim Tholstorf
# https://github.com/KimTholstorf/oci-rvtools-cost-estimator
# MIT License — see LICENSE file for details

"""Compute-shape registry — the single source of truth for OCI part numbers.

Loaded from the packaged ``data/shapes.json``. The web UI reads an identical
mirror at ``docs/shapes.json`` and the weekly prices workflow derives its part
list from the same file, so part numbers live in exactly one place.
"""

from __future__ import annotations

import json
from importlib import resources
from typing import Dict, List, Optional

_DATA: Dict = json.loads(
    resources.files("oci_rvtools").joinpath("data", "shapes.json").read_text(encoding="utf-8")
)

SHAPES: List[Dict[str, str]] = _DATA["shapes"]
STORAGE_PART: str = _DATA["storage_part"]
VPU_PART: str = _DATA["vpu_part"]


def _default_shape() -> Dict[str, str]:
    for shape in SHAPES:
        if shape.get("default"):
            return shape
    return SHAPES[0]


DEFAULT_SHAPE = _default_shape()
DEFAULT_OCPU_PART: str = DEFAULT_SHAPE["ocpu_part"]
DEFAULT_MEMORY_PART: str = DEFAULT_SHAPE["memory_part"]
DEFAULT_STORAGE_PART: str = STORAGE_PART
DEFAULT_VPU_PART: str = VPU_PART


def get_shape(shape_id: str) -> Optional[Dict[str, str]]:
    """Return the shape dict for an id (e.g. "e6ax"), or None."""
    for shape in SHAPES:
        if shape["id"] == shape_id:
            return shape
    return None


def all_part_numbers() -> List[str]:
    """Every distinct part number referenced by any shape plus storage and VPU."""
    parts: List[str] = []
    for shape in SHAPES:
        parts.append(shape["ocpu_part"])
        parts.append(shape["memory_part"])
    parts.append(STORAGE_PART)
    parts.append(VPU_PART)
    seen = set()
    ordered: List[str] = []
    for part in parts:
        if part not in seen:
            seen.add(part)
            ordered.append(part)
    return ordered
