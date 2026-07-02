# Copyright (c) 2026 Kim Tholstorf
# https://github.com/KimTholstorf/oci-rvtools-cost-estimator
# MIT License — see LICENSE file for details

"""Backward-compatibility shim.

The implementation was split into focused modules (ingest, compute, pricing,
osmatch, report, cli). This module re-exports the public API at its historical
``oci_rvtools.core`` location so existing imports keep working.
"""

from __future__ import annotations

from .version import VERSION

# Configuration / defaults
from . import shapes
from .cli import (
    DEFAULT_CURRENCY,
    DEFAULT_HOURS,
    DEFAULT_MEMORY_PART,
    DEFAULT_OCPU_PART,
    DEFAULT_OUTPUT,
    DEFAULT_STORAGE_PART,
    DEFAULT_VPU,
    DEFAULT_VPU_PART,
    main,
    parse_args,
)

# Logging
from .log import info, warn

# OS classification
from .osmatch import classify_os, detect_os

# Ingestion
from .ingest import (
    ALIASES_VINFO,
    CANON_COLS_VINFO,
    INVALID_CLUSTER_VALUES,
    NUMERIC_PREFERRED_VINFO,
    TOKEN_MAP_VINFO,
    apply_vm_filter,
    canonicalize_vinfo,
    collapse_duplicate_columns,
    collect_rvtools_files,
    list_datacenters_and_clusters,
    load_vinfo_dataframe,
    prepare_vinfo_df as _prepare_vinfo_df,
    valid_cluster as _valid_cluster,
)

# Domain model
from .model import AggregatedUsage, LineItem, PriceRecord

# Compute
from .compute import (
    MIB_TO_GB,
    aggregate_from_rvtools,
    aggregate_vinfo,
    build_line_items,
)

# Pricing
from .pricing import API_BASE, PricingClient

# Anonymisation
from .anonymize import anonymize_file, build_anonymized

# Reporting
from .report import write_output
from .report.cost_summary import EXCEL_HEADERS, build_metadata_rows
from .report.os_summary import write_os_summary_sheet
from .report.styles import accounting_number_format as _accounting_number_format
from .report.vm_details import write_vm_detail_sheet

__all__ = [
    "VERSION",
    "main",
    "parse_args",
    "detect_os",
    "classify_os",
    "aggregate_from_rvtools",
    "aggregate_vinfo",
    "build_line_items",
    "write_output",
    "write_vm_detail_sheet",
    "write_os_summary_sheet",
    "PricingClient",
    "PriceRecord",
    "LineItem",
    "AggregatedUsage",
    "load_vinfo_dataframe",
    "canonicalize_vinfo",
    "collect_rvtools_files",
    "apply_vm_filter",
    "list_datacenters_and_clusters",
    "build_metadata_rows",
    "anonymize_file",
    "build_anonymized",
    "MIB_TO_GB",
    "API_BASE",
]
