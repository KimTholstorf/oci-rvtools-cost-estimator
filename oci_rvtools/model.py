# Copyright (c) 2026 Kim Tholstorf
# https://github.com/KimTholstorf/oci-rvtools-cost-estimator
# MIT License — see LICENSE file for details

"""Domain dataclasses shared across the calculation and reporting layers."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import List, Optional

import pandas as pd


@dataclass
class AggregatedUsage:
    """Totals aggregated across one or more RVTools exports."""
    total_vcpu: float = 0.0
    ram_gb: float = 0.0
    disk_total_gb: float = 0.0
    disk_used_gb: float = 0.0
    powered_on_vms: int = 0
    powered_off_vms: int = 0
    source_files: List[str] = field(default_factory=list)
    vm_dataframe: Optional[pd.DataFrame] = field(default=None)


@dataclass
class LineItem:
    """A single priced row in a cost section."""
    description: str
    part_number: str
    category: str
    raw_base_quantity: float
    usage_quantity: float
    unit_price: float


@dataclass
class PriceRecord:
    """A resolved price for an OCI part number."""
    part_number: str
    display_name: str
    unit_price: float
