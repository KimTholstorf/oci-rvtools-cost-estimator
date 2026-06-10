# Copyright (c) 2026 Kim Tholstorf
# https://github.com/KimTholstorf/oci-rvtools-cost-estimator
# MIT License — see LICENSE file for details

"""Top-level workbook assembly: Cost Summary, VM Details, OS Summary."""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
from openpyxl import Workbook

from ..log import info
from ..model import LineItem
from .cost_summary import COST_SUMMARY_SHEET, write_cost_summary
from .os_summary import write_os_summary_sheet
from .vm_details import write_vm_detail_sheet


def write_output(
    output_path: Path,
    metadata: Dict[str, object],
    sheets: Dict[str, List[LineItem]],
    currency: str,
    vm_df: Optional[pd.DataFrame] = None,
) -> None:
    info(f"Writing output to {output_path}")
    output_path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = COST_SUMMARY_SHEET

    write_cost_summary(ws, metadata, sheets, currency, vm_df=vm_df)

    if vm_df is not None and not vm_df.empty:
        write_vm_detail_sheet(
            wb, vm_df, metadata, currency,
            include_vms_off=bool(metadata.get("powered_off_included")),
            include_disks_off=bool(metadata.get("powered_off_disks_included")),
        )
        write_os_summary_sheet(wb, vm_df)

    wb.save(output_path)
