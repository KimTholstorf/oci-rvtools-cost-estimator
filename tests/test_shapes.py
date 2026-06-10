"""Shape registry: single-source-of-truth guarantees."""

import json
from pathlib import Path

from importlib import resources

from oci_rvtools import shapes

REPO_ROOT = Path(__file__).resolve().parent.parent


def test_docs_mirror_matches_package_copy():
    """docs/shapes.json (web) must stay byte-identical to the packaged copy."""
    packaged = resources.files("oci_rvtools").joinpath("data", "shapes.json").read_text(encoding="utf-8")
    docs = (REPO_ROOT / "docs" / "shapes.json").read_text(encoding="utf-8")
    assert packaged == docs, "docs/shapes.json is out of sync with oci_rvtools/data/shapes.json"


def test_default_shape_parts():
    assert shapes.DEFAULT_OCPU_PART
    assert shapes.DEFAULT_MEMORY_PART
    assert shapes.DEFAULT_STORAGE_PART
    assert shapes.DEFAULT_VPU_PART


def test_all_part_numbers_unique_and_complete():
    parts = shapes.all_part_numbers()
    assert len(parts) == len(set(parts)), "duplicate part numbers"
    # every shape's parts plus storage and vpu are present
    for shape in shapes.SHAPES:
        assert shape["ocpu_part"] in parts
        assert shape["memory_part"] in parts
    assert shapes.STORAGE_PART in parts
    assert shapes.VPU_PART in parts


def test_get_shape_lookup():
    assert shapes.get_shape("e6ax")["label"] == "VM.Standard.E6.Ax.Flex"
    assert shapes.get_shape("does-not-exist") is None


def test_one_default_shape():
    defaults = [s for s in shapes.SHAPES if s.get("default")]
    assert len(defaults) == 1
