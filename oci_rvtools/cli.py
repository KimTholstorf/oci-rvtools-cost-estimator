# Copyright (c) 2026 Kim Tholstorf
# https://github.com/KimTholstorf/oci-rvtools-cost-estimator
# MIT License — see LICENSE file for details

"""Command-line interface."""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Iterable, Optional

from . import shapes
from .compute import aggregate_from_rvtools, build_line_items
from .ingest import collect_rvtools_files, list_datacenters_and_clusters
from .log import error, info
from .pricing import PricingClient
from .report import write_output
from .version import VERSION

DEFAULT_OUTPUT = "oci_cost_summary.xlsx"
DEFAULT_CURRENCY = "USD"
DEFAULT_HOURS = 730
DEFAULT_VPU = 10.0
DEFAULT_OCPU_PART = shapes.DEFAULT_OCPU_PART
DEFAULT_MEMORY_PART = shapes.DEFAULT_MEMORY_PART
DEFAULT_STORAGE_PART = shapes.DEFAULT_STORAGE_PART
DEFAULT_VPU_PART = shapes.DEFAULT_VPU_PART


def parse_args(argv: Optional[Iterable[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Calculate OCI monthly costs directly from RVTools exports."
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {VERSION}",
        help="Show program's version number and exit.",
    )
    parser.add_argument(
        "--rvtools",
        nargs="+",
        required=True,
        help="Path(s) to RVTools Excel exports (.xlsx) or directories containing them.",
    )
    parser.add_argument("--output", default=DEFAULT_OUTPUT, help="Output Excel file path.")
    parser.add_argument("--hours", type=float, default=DEFAULT_HOURS, help="Billing hours in month (default 730).")
    parser.add_argument("--currency", default=DEFAULT_CURRENCY, help="Currency code for pricing (default USD).")
    parser.add_argument("--ocpu-part", default=DEFAULT_OCPU_PART, help="OCI part number for OCPU per hour.")
    parser.add_argument("--memory-part", default=DEFAULT_MEMORY_PART, help="OCI part number for Memory GB per hour.")
    parser.add_argument("--storage-part", default=DEFAULT_STORAGE_PART, help="OCI part number for block storage GB per month.")
    parser.add_argument("--vpu-part", default=DEFAULT_VPU_PART, help="OCI part number for block volume performance units.")
    parser.add_argument("--vpu", type=float, default=DEFAULT_VPU, help="Block volume performance units per GB (default 10).")
    parser.add_argument(
        "--include-poweredoff-vms",
        action="store_true",
        default=False,
        help="Include powered-off VMs in CPU/RAM totals (default: powered-on only).",
    )
    parser.add_argument(
        "--include-poweredoff-disks",
        dest="include_poweredoff_disks",
        action="store_true",
        default=True,
        help="Include powered-off VMs when summing disk usage (default).",
    )
    parser.add_argument(
        "--exclude-poweredoff-disks",
        dest="include_poweredoff_disks",
        action="store_false",
        help="Exclude powered-off VMs when summing disk usage.",
    )
    parser.add_argument(
        "--list",
        action="store_true",
        default=False,
        dest="list_topology",
        help="List all Datacenter and Cluster names found in the input file(s) and exit.",
    )
    parser.add_argument(
        "--datacenter",
        nargs="+",
        metavar="NAME",
        default=None,
        help='Only include VMs in the given Datacenter(s). Quote names with spaces e.g. "DC East".',
    )
    parser.add_argument(
        "--cluster",
        nargs="+",
        metavar="NAME",
        default=None,
        help='Only include VMs in the given Cluster(s). Quote names with spaces e.g. "Cluster East 01".',
    )
    return parser.parse_args(argv)


def main(argv: Optional[Iterable[str]] = None) -> int:
    args = parse_args(argv)

    rvtools_files = collect_rvtools_files(args.rvtools)
    if not rvtools_files:
        error("No RVTools Excel exports found for the given paths.")
        return 1

    if args.list_topology:
        list_datacenters_and_clusters(rvtools_files)
        return 0

    if args.datacenter:
        info(f"Filtering by Datacenter(s): {', '.join(args.datacenter)}")
    if args.cluster:
        info(f"Filtering by Cluster(s): {', '.join(args.cluster)}")

    info(f"Powered-off VMs {'included' if args.include_poweredoff_vms else 'excluded'} in CPU/RAM totals.")
    info(f"Powered-off disks {'included' if args.include_poweredoff_disks else 'excluded'} in disk totals.")

    usage = aggregate_from_rvtools(
        rvtools_files,
        args.include_poweredoff_vms,
        args.include_poweredoff_disks,
        datacenters=args.datacenter,
        clusters=args.cluster,
    )
    if not usage.source_files:
        error("No usable vInfo data found in the provided RVTools exports.")
        return 1

    info(
        f"Aggregated totals - vCPU: {usage.total_vcpu:.2f}, RAM GB: {usage.ram_gb:.2f}, "
        f"Disk Total GB: {usage.disk_total_gb:.2f}, Disk Used GB: {usage.disk_used_gb:.2f}"
    )
    info(f"Powered-on VMs counted: {usage.powered_on_vms}")
    info(f"Powered-off VMs counted: {usage.powered_off_vms}")

    parts = {
        "ocpu": args.ocpu_part,
        "memory": args.memory_part,
        "storage": args.storage_part,
        "vpu": args.vpu_part,
    }

    pricing = PricingClient(args.currency)
    clamped_vpu = max(1.0, min(float(args.vpu), 120.0))

    try:
        sheets = build_line_items(usage, args.hours, clamped_vpu, pricing, parts)
    except Exception as exc:
        error(f"Failed to compute costs: {exc}")
        return 1

    metadata = {
        "source_files": ", ".join(usage.source_files),
        "filter_datacenters": args.datacenter or [],
        "filter_clusters": args.cluster or [],
        "hours_per_month": str(args.hours),
        "currency": args.currency.upper(),
        "vpu": str(clamped_vpu),
        "powered_on_vms": usage.powered_on_vms,
        "powered_off_vms": usage.powered_off_vms,
        "powered_off_included": args.include_poweredoff_vms,
        "powered_off_disks_included": args.include_poweredoff_disks,
    }

    output_path = Path(args.output).resolve()
    try:
        write_output(output_path, metadata, sheets, args.currency.upper(), vm_df=usage.vm_dataframe)
    except Exception as exc:
        error(f"Failed to write output: {exc}")
        return 1

    print("[OK] Cost summary generated successfully.")
    return 0
