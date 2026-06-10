# Copyright (c) 2026 Kim Tholstorf
# https://github.com/KimTholstorf/oci-rvtools-cost-estimator
# MIT License — see LICENSE file for details

"""Usage aggregation and cost line-item construction."""

from __future__ import annotations

import math
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

import pandas as pd

from .ingest import (
    apply_vm_filter,
    load_vinfo_dataframe,
    prepare_vinfo_df,
)
from .log import info, warn
from .model import AggregatedUsage, LineItem
from .pricing import PricingClient

# RVTools reports disk in MiB; convert MiB -> TiB -> GB.
MIB_TO_GB = 1024.0 / 953_674.0


def aggregate_vinfo(
    df: pd.DataFrame, include_vms_off: bool, include_disks_off: bool
) -> Tuple[float, float, float, float, int, int]:
    if df is None or df.empty:
        return 0.0, 0.0, 0.0, 0.0, 0, 0

    df = prepare_vinfo_df(df)
    if df.empty:
        return 0.0, 0.0, 0.0, 0.0, 0, 0

    powered_mask = df["Powerstate"].astype(str).str.lower() == "poweredon"
    powered_on_count = int(powered_mask.sum())
    total_count = int(len(df))
    powered_off_count = total_count - powered_on_count

    vcpu_frame = df if include_vms_off else df[powered_mask]
    total_vcpu = float(pd.to_numeric(vcpu_frame["CPUs"], errors="coerce").fillna(0).sum())
    ram_gb = float(pd.to_numeric(vcpu_frame["Memory"], errors="coerce").fillna(0).sum()) / 1024.0

    disk_frame = df if include_disks_off else df[powered_mask]
    total_mib = pd.to_numeric(disk_frame["Total disk capacity MiB"], errors="coerce").fillna(0)
    prov_mib = pd.to_numeric(disk_frame["Provisioned MiB"], errors="coerce").fillna(0)
    effective_prov = total_mib.where(total_mib > 0, prov_mib)
    used_mib = pd.to_numeric(disk_frame["In Use MiB"], errors="coerce").fillna(0)

    disk_total_gb = float(effective_prov.sum() * MIB_TO_GB)
    disk_used_gb = float(used_mib.sum() * MIB_TO_GB)

    return total_vcpu, ram_gb, disk_total_gb, disk_used_gb, powered_on_count, powered_off_count


def aggregate_from_rvtools(
    files: Sequence[Path],
    include_vms_off: bool,
    include_disks_off: bool,
    datacenters: Optional[List[str]] = None,
    clusters: Optional[List[str]] = None,
) -> AggregatedUsage:
    usage = AggregatedUsage()
    vm_frames: List[pd.DataFrame] = []
    for file in files:
        info(f"Processing RVTools export: {file}")
        df = load_vinfo_dataframe(file)
        if df is None:
            continue
        df = apply_vm_filter(df, datacenters, clusters)
        if df.empty:
            warn(f"{file.name}: no VMs remain after applying datacenter/cluster filter, skipping")
            continue
        detail_df = prepare_vinfo_df(df, verbose=False)
        if not detail_df.empty:
            vm_frames.append(detail_df)
        total_vcpu, ram_gb, disk_total_gb, disk_used_gb, powered_on, powered_off = aggregate_vinfo(
            df, include_vms_off, include_disks_off
        )
        usage.total_vcpu += total_vcpu
        usage.ram_gb += ram_gb
        usage.disk_total_gb += disk_total_gb
        usage.disk_used_gb += disk_used_gb
        usage.powered_on_vms += powered_on
        usage.powered_off_vms += powered_off
        usage.source_files.append(str(file))

    if vm_frames:
        usage.vm_dataframe = pd.concat(vm_frames, ignore_index=True)

    return usage


def build_line_items(
    usage: AggregatedUsage,
    hours: float,
    vpu_value: float,
    pricing: PricingClient,
    parts: Dict[str, str],
) -> Dict[str, List[LineItem]]:
    raw_total_ocpus = usage.total_vcpu / 2.0
    raw_memory_gb = usage.ram_gb
    raw_total_disk_gb = usage.disk_total_gb
    raw_used_disk_gb = usage.disk_used_gb

    total_ocpus = math.ceil(raw_total_ocpus)
    memory_gb = math.ceil(raw_memory_gb)
    total_disk_gb = math.ceil(raw_total_disk_gb)
    used_disk_gb = math.ceil(raw_used_disk_gb)

    ocpu_hours = total_ocpus * hours
    memory_hours = memory_gb * hours

    info(f"Total OCPUs: {total_ocpus:.2f}, OCPU hours: {ocpu_hours:.2f}")
    info(f"Memory GB: {memory_gb:.2f}, GB hours: {memory_hours:.2f}")
    info(f"Total disk GB: {total_disk_gb:.2f}")
    info(f"Used disk GB: {used_disk_gb:.2f}")
    info(f"Hours per month: {hours}")
    info(f"VPU per GB: {vpu_value}")

    price_records = {
        key: pricing.get_price(parts[key])
        for key in ("ocpu", "memory", "storage", "vpu")
    }

    def make_lines(disk_gb: float, raw_disk_gb: float) -> List[LineItem]:
        raw_vpu_quantity = raw_disk_gb * vpu_value
        return [
            LineItem(
                description=price_records["ocpu"].display_name,
                part_number=price_records["ocpu"].part_number,
                category="ocpu",
                raw_base_quantity=raw_total_ocpus,
                usage_quantity=hours,
                unit_price=price_records["ocpu"].unit_price,
            ),
            LineItem(
                description=price_records["memory"].display_name,
                part_number=price_records["memory"].part_number,
                category="memory",
                raw_base_quantity=raw_memory_gb,
                usage_quantity=hours,
                unit_price=price_records["memory"].unit_price,
            ),
            LineItem(
                description=price_records["storage"].display_name,
                part_number=price_records["storage"].part_number,
                category="storage",
                raw_base_quantity=raw_disk_gb,
                usage_quantity=1.0,
                unit_price=price_records["storage"].unit_price,
            ),
            LineItem(
                description=price_records["vpu"].display_name,
                part_number=price_records["vpu"].part_number,
                category="vpu",
                raw_base_quantity=raw_vpu_quantity,
                usage_quantity=1.0,
                unit_price=price_records["vpu"].unit_price,
            ),
        ]

    return {
        "total_disk": make_lines(total_disk_gb, raw_total_disk_gb),
        "used_disk": make_lines(used_disk_gb, raw_used_disk_gb),
    }
