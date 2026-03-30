"""
OCI Monthly Cost Calculator (direct RVTools input)

Reads one or more raw RVTools Excel exports, reproduces the aggregation logic
used by rvtools_summarizer.py for powered-on VM CPU/RAM and provisioned/used
disk capacity, then retrieves current OCI list prices from the CETools API and
produces a cost breakdown workbook.
"""

from __future__ import annotations

import argparse
import json
import math
from datetime import datetime
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple
from urllib import error as urlerror
from urllib import parse as urlparse
from urllib import request as urlrequest

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill


# =========================
# Version
# =========================

VERSION = "1.0.7"


# =========================
# CLI Defaults
# =========================

DEFAULT_INPUT = ""  # unused placeholder
DEFAULT_OUTPUT = "oci_cost_summary.xlsx"
DEFAULT_CURRENCY = "USD"
DEFAULT_HOURS = 730
DEFAULT_OCPU_PART = "B97384"
DEFAULT_MEMORY_PART = "B97385"
DEFAULT_STORAGE_PART = "B91961"
DEFAULT_VPU_PART = "B91962"
DEFAULT_VPU = 10.0


# =========================
# API
# =========================

API_BASE = "https://apexapps.oracle.com/pls/apex/cetools/api/v1/products/"


# =========================
# vInfo Column Mapping
# =========================

CANON_COLS_VINFO = [
    "Powerstate",
    "VM",
    "Datacenter",
    "Cluster",
    "CPUs",
    "Memory",
    "Video Ram KiB",
    "Provisioned MiB",
    "In Use MiB",
    "Total disk capacity MiB",
]

ALIASES_VINFO = {
    "vInfoVMName": "VM",
    "vInfoPowerstate": "Powerstate",
    "vInfoDataCenter": "Datacenter",
    "vInfoCluster": "Cluster",
    "vInfoHost": "Host",
    "vInfoCPUs": "CPUs",
    "vInfoMemory": "Memory",
    "vInfoVideoRamKiB": "Video Ram KiB",
    "vInfoProvisioned": "Provisioned MiB",
    "vInfoInUse": "In Use MiB",
    "vInfoTotalDiskCapacityMiB": "Total disk capacity MiB",
    "vInfoPrimaryIPAddress": "Primary IP Address",
    "vInfoPrimaryIP": "Primary IP Address",
    "vInfoOS": "OS according to the VMware Tools",
    "vInfoOSTools": "OS according to the VMware Tools",
}

TOKEN_MAP_VINFO = {
    "powerstate": "Powerstate",
    "vm": "VM",
    "vmname": "VM",
    "datacenter": "Datacenter",
    "datacentre": "Datacenter",
    "cluster": "Cluster",
    "cpus": "CPUs",
    "memory": "Memory",
    "videorammikb": "Video Ram KiB",
    "videorammikib": "Video Ram KiB",
    "provisioned": "Provisioned MiB",
    "provisionedmib": "Provisioned MiB",
    "inuse": "In Use MiB",
    "inusemib": "In Use MiB",
    "totaldiskcapacitymib": "Total disk capacity MiB",
}

NUMERIC_PREFERRED_VINFO = {
    "CPUs",
    "Memory",
    "Video Ram KiB",
    "Provisioned MiB",
    "In Use MiB",
    "Total disk capacity MiB",
}

INVALID_CLUSTER_VALUES = {"", "none", "nan", "unknown"}


# =========================
# Unit Conversion
# =========================

MIB_TO_GB = 1024.0 / 953_674.0  # RVTools MiB -> TiB -> GB conversion factor


# =========================
# Excel Layout
# =========================

TABLE_COLUMN_COUNT = 7
EXCEL_HEADERS = [
    "Description",
    "Part Number",
    "Part Qty",
    "Instance Qty",
    "Usage Qty",
    "Unit Price ({currency})",
    "Monthly Cost ({currency})",
]
COLUMN_WIDTHS = {
    "A": 36,
    "B": 16,
    "C": 16,
    "D": 16,
    "E": 16,
    "F": 16,
    "G": 16,
}
TITLE_ROW_HEIGHT = 40
HEADER_ROW_HEIGHT = 40
DEFAULT_ROW_HEIGHT = 20
DISCLAIMER_ROW_HEIGHT = 80


# =========================
# Excel Styles
# =========================

TITLE_FONT = Font(bold=True, size=14)
HEADER_FILL = PatternFill(start_color="FFDCE6F1", end_color="FFDCE6F1", fill_type="solid")
DISCLAIMER_ALIGNMENT = Alignment(wrap_text=True, vertical="top")


# =========================
# Logging helpers
# =========================

def info(msg: str) -> None:
    print(f"[INFO] {msg}")


def warn(msg: str) -> None:
    print(f"[WARN] {msg}")


# =========================
# RVTools helpers
# =========================

def _sheet_token(name: str) -> str:
    return "".join(ch for ch in (name or "").strip().lower() if ch.isalnum())


def _to_token(value: str) -> str:
    return "".join(ch for ch in value.strip().lower() if ch.isalnum() or ch == "#")


def collapse_duplicate_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Coalesce duplicate column names (sum numeric, else first non-empty)."""
    if df is None or df.empty:
        return df

    cols = list(df.columns)
    duplicates = {c for c in cols if cols.count(c) > 1}
    for col in duplicates:
        numeric = col in NUMERIC_PREFERRED_VINFO
        parts = [df.iloc[:, idx] for idx, name in enumerate(cols) if name == col]
        stacked = pd.concat(parts, axis=1)
        if numeric:
            coalesced = pd.to_numeric(stacked, errors="coerce").fillna(0).sum(axis=1)
        else:
            coalesced = stacked.apply(
                lambda row: next((v for v in row if pd.notna(v) and str(v).strip()), ""),
                axis=1,
            )
        first_idx = next(i for i, name in enumerate(cols) if name == col)
        mask = [True] * len(cols)
        seen = 0
        for i, name in enumerate(cols):
            if name == col:
                seen += 1
                if seen > 1:
                    mask[i] = False
        df = df.loc[:, mask]
        cols = list(df.columns)
        df.iloc[:, first_idx] = coalesced
        info(f"vInfo: coalesced duplicate column '{col}' ({len(parts)} -> 1)")
    return df


def canonicalize_vinfo(df: pd.DataFrame) -> pd.DataFrame:
    if not isinstance(df, pd.DataFrame):
        return df

    # Aliases
    aliases = {old: new for (old, new) in ALIASES_VINFO.items() if old in df.columns}
    if aliases:
        df = df.rename(columns=aliases)
        for old, new in aliases.items():
            info(f"vInfo: renaming column '{old}' -> '{new}' (alias)")

    # Token-based normalization
    renames: Dict[str, str] = {}
    for col in df.columns:
        token = _to_token(str(col))
        target = TOKEN_MAP_VINFO.get(token)
        if target and target != col:
            renames[col] = target
    if renames:
        df = df.rename(columns=renames)
        for old, new in renames.items():
            info(f"vInfo: renaming column '{old}' -> '{new}' (token)")

    df = collapse_duplicate_columns(df)

    # Backfill required columns
    for column in CANON_COLS_VINFO:
        if column not in df.columns:
            warn(f"vInfo: missing column '{column}', backfilling default")
            df[column] = 0 if column in NUMERIC_PREFERRED_VINFO else ""

    return df


def load_vinfo_dataframe(filepath: Path) -> Optional[pd.DataFrame]:
    try:
        xl = pd.read_excel(filepath, sheet_name=None, engine="openpyxl")
    except Exception as exc:
        warn(f"{filepath.name}: failed to load workbook ({exc})")
        return None

    vinfo_df = None
    for sheet_name, df in xl.items():
        token = _sheet_token(sheet_name)
        if token.endswith("vinfo") or "vinfo" in token:
            vinfo_df = df
            break

    if vinfo_df is None:
        warn(f"{filepath.name}: missing vInfo sheet, skipping")
        return None

    vinfo_df = canonicalize_vinfo(vinfo_df)
    return vinfo_df


# =========================
# Aggregation from vInfo
# =========================

def _valid_cluster(value: object) -> bool:
    if pd.isna(value):
        return False
    s = str(value).strip()
    if not s:
        return False
    return s.lower() not in INVALID_CLUSTER_VALUES


@dataclass
class AggregatedUsage:
    total_vcpu: float = 0.0
    ram_gb: float = 0.0
    disk_total_gb: float = 0.0
    disk_used_gb: float = 0.0
    powered_on_vms: int = 0
    powered_off_vms: int = 0
    source_files: List[str] = field(default_factory=list)


def aggregate_vinfo(df: pd.DataFrame, include_vms_off: bool, include_disks_off: bool) -> Tuple[float, float, float, float, int, int]:
    if df is None or df.empty:
        return 0.0, 0.0, 0.0, 0.0, 0, 0

    df = df.copy()

    # Remove RVTools housekeeping VMs
    df = df[~df["VM"].astype(str).str.startswith("vCLS", na=False)]

    # Keep only rows with valid cluster labels
    df = df[df["Cluster"].apply(_valid_cluster)]
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


def collect_rvtools_files(paths: Sequence[str]) -> List[Path]:
    files: List[Path] = []
    seen: set[Path] = set()
    skipped_temp: set[str] = set()

    def add_candidate(candidate: Path) -> None:
        resolved = candidate.resolve()
        if resolved in seen:
            return
        seen.add(resolved)
        files.append(candidate)

    for entry in paths:
        p = Path(entry).expanduser()
        if p.is_dir():
            for candidate in sorted(p.glob("*.xlsx")):
                if candidate.name.startswith("~$"):
                    skipped_temp.add(candidate.name)
                    continue
                add_candidate(candidate)
        elif p.is_file():
            if p.name.startswith("~$"):
                skipped_temp.add(p.name)
                continue
            add_candidate(p)
        else:
            warn(f"Input path not found: {p}")

    if skipped_temp:
        info(f"Skipped temporary workbook(s): {', '.join(sorted(skipped_temp))}")
    return files


def aggregate_from_rvtools(
    files: Sequence[Path],
    include_vms_off: bool,
    include_disks_off: bool,
) -> AggregatedUsage:
    usage = AggregatedUsage()
    for file in files:
        info(f"Processing RVTools export: {file}")
        df = load_vinfo_dataframe(file)
        if df is None:
            continue
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

    return usage


# =========================
# Pricing client
# =========================

@dataclass
class LineItem:
    description: str
    part_number: str
    category: str
    raw_base_quantity: float
    usage_quantity: float
    unit_price: float


@dataclass
class PriceRecord:
    part_number: str
    display_name: str
    unit_price: float


class PricingClient:
    def __init__(self, currency: str) -> None:
        self.currency = currency.upper()
        self._cache: Dict[str, PriceRecord] = {}

    def get_price(self, part_number: str) -> PriceRecord:
        part_number = part_number.strip()
        if part_number in self._cache:
            return self._cache[part_number]

        params = {"partNumber": part_number, "currencyCode": self.currency}
        url = f"{API_BASE}?{urlparse.urlencode(params)}"
        try:
            with urlrequest.urlopen(url) as response:
                payload = response.read().decode("utf-8")
        except urlerror.URLError as exc:
            raise RuntimeError(f"Failed to reach OCI price API ({url}): {exc}") from exc

        try:
            data = json.loads(payload)
        except json.JSONDecodeError as exc:
            raise RuntimeError(f"Invalid JSON from OCI price API for part {part_number}") from exc

        items = data.get("items")
        if not isinstance(items, list) or not items:
            raise RuntimeError(f"No pricing data for part {part_number} ({self.currency})")

        price: Optional[float] = None
        display_name: Optional[str] = None
        for item in items:
            if not isinstance(item, dict):
                continue
            if item.get("partNumber") and item["partNumber"] != part_number:
                continue
            price = self._extract_price(item, self.currency)
            display_name = item.get("displayName") or display_name
            if price is not None:
                break

        if price is None:
            for item in items:
                if isinstance(item, dict):
                    price = PricingClient._extract_price(item, self.currency)
                    display_name = item.get("displayName") or display_name
                    if price is not None:
                        break

        if price is None:
            raise RuntimeError(f"Could not determine unit price for part {part_number}: {payload[:500]}...")

        if not display_name:
            display_name = part_number

        record = PriceRecord(part_number=part_number, display_name=str(display_name), unit_price=float(price))
        self._cache[part_number] = record
        return record

    @staticmethod
    def _extract_price(item: Dict[str, object], currency: str) -> Optional[float]:
        currency = currency.upper()
        candidate_keys = [
            "price",
            "unitPrice",
            "unit_price",
            "unit_price_value",
            "unit_price_included",
            "netUnitPrice",
            "list_price",
            "usdPrice",
            "amount",
        ]
        for key in candidate_keys:
            if key in item:
                try:
                    return float(item[key])  # type: ignore[arg-type]
                except (TypeError, ValueError):
                    continue

        if "prices" in item and isinstance(item["prices"], list):
            for entry in item["prices"]:
                if isinstance(entry, dict):
                    maybe = PricingClient._extract_price(entry, currency)
                    if maybe is not None:
                        return maybe

        locs = item.get("currencyCodeLocalizations")
        if isinstance(locs, list):
            for loc in locs:
                if not isinstance(loc, dict):
                    continue
                loc_currency = loc.get("currencyCode")
                if loc_currency and loc_currency.upper() != currency:
                    continue
                prices = loc.get("prices")
                if isinstance(prices, list):
                    for entry in prices:
                        if not isinstance(entry, dict):
                            continue
                        if "value" in entry:
                            try:
                                return float(entry["value"])
                            except (TypeError, ValueError):
                                pass
                        maybe = PricingClient._extract_price(entry, currency)
                        if maybe is not None:
                            return maybe
        return None


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


# =========================
# Excel output
# =========================

def build_metadata_rows(metadata: Dict[str, object]) -> List[Tuple[str, str, str]]:
    return [
        ("Source Files", metadata["source_files"], ""),
        ("Hours per Month", metadata["hours_per_month"], ""),
        ("Currency", metadata["currency"], ""),
        ("VPU", metadata["vpu"], ""),
        ("Powered On VMs", str(metadata["powered_on_vms"]), "(included)"),
        (
            "Powered Off VMs",
            str(metadata["powered_off_vms"]),
            "(included)" if metadata["powered_off_included"] else "(excluded)",
        ),
        (
            "Powered Off Disks",
            "",
            "(included)" if metadata["powered_off_disks_included"] else "(excluded)",
        ),
    ]


def apply_column_widths(ws) -> None:
    for column, width in COLUMN_WIDTHS.items():
        ws.column_dimensions[column].width = width


def write_metadata(ws, rows: List[Tuple[str, str, str]], start_row: int = 2) -> int:
    row = start_row
    for title, value, status in rows:
        ws.cell(row=row, column=1, value=title)
        ws.cell(row=row, column=2, value=value)
        if status:
            ws.cell(row=row, column=3, value=status)
        row += 1
    return row - 1


def format_row(ws, row: int, *, bold: bool = False, vertical: str = "bottom") -> None:
    for col_idx in range(1, TABLE_COLUMN_COUNT + 1):
        cell = ws.cell(row=row, column=col_idx)
        if bold:
            cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="left", vertical=vertical)


def format_metadata_block(ws, start_row: int, end_row: int) -> None:
    for row in range(start_row, end_row + 1):
        format_row(ws, row)
    format_row(ws, start_row, bold=True)


def format_table_header(ws, header_row: int, headers: List[str]) -> None:
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col_idx, value=header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="left", vertical="center")
        cell.fill = HEADER_FILL


def format_data_rows(ws, start_row: int, end_row: int) -> None:
    for row in range(start_row, end_row):
        format_row(ws, row)


def format_total_row(ws, total_row: int) -> None:
    format_row(ws, total_row, bold=True)


def apply_row_heights(
    ws,
    metadata_end_row: int,
    blank_row_index: int,
    header_row: int,
    data_start: int,
    total_row: int,
    disclaimer_row: int,
) -> None:
    ws.row_dimensions[1].height = TITLE_ROW_HEIGHT
    for row in range(2, metadata_end_row + 1):
        ws.row_dimensions[row].height = DEFAULT_ROW_HEIGHT
    ws.row_dimensions[blank_row_index].height = DEFAULT_ROW_HEIGHT
    ws.row_dimensions[header_row].height = HEADER_ROW_HEIGHT
    for row in range(data_start, total_row + 1):
        ws.row_dimensions[row].height = DEFAULT_ROW_HEIGHT
    quote_row = disclaimer_row - 2
    if quote_row >= 2:
        ws.row_dimensions[quote_row].height = DEFAULT_ROW_HEIGHT
    ws.row_dimensions[disclaimer_row].height = DISCLAIMER_ROW_HEIGHT


def append_advisory(ws, total_row: int) -> Tuple[int, int]:
    quote_row = total_row + 3
    quote_cell = ws.cell(row=quote_row, column=1, value="Quote is for investment proposal only.")
    quote_cell.font = Font(bold=True, color="00FF0000")
    ws.merge_cells(start_row=quote_row, start_column=1, end_row=quote_row, end_column=TABLE_COLUMN_COUNT)

    disclaimer_row = quote_row + 2
    disclaimer_text = (
        "Disclaimer:  This sample quote is provided solely for evaluation purposes and is intended to further "
        "discussions between you and Oracle.  This sample quote is not eligible for acceptance by you and is "
        "not a binding contract between you and Oracle for the services specified.  If you would like to purchase "
        "the services specified in this sample quote, please request that Oracle issue you a formal quote (which "
        "may include an OMA or a CSA if you do not already have an appropriate agreement in place with Oracle) "
        "for your acceptance and execution.  Your formal quote will be effective only upon Oracle's acceptance "
        "of the formal quote (and the OMA or CSA, if required)."
    )
    disclaimer_cell = ws.cell(row=disclaimer_row, column=1, value=disclaimer_text)
    ws.merge_cells(start_row=disclaimer_row, start_column=1, end_row=disclaimer_row, end_column=TABLE_COLUMN_COUNT)
    disclaimer_cell.alignment = DISCLAIMER_ALIGNMENT
    ws.row_dimensions[disclaimer_row].height = DISCLAIMER_ROW_HEIGHT
    return quote_row, disclaimer_row


def write_output(
    output_path: Path,
    metadata: Dict[str, object],
    sheets: Dict[str, List[LineItem]],
    currency: str,
) -> None:
    info(f"Writing output to {output_path}")
    output_path.parent.mkdir(parents=True, exist_ok=True)

    headers = [header.format(currency=currency) for header in EXCEL_HEADERS]

    def roundup_formula(value: float) -> str:
        raw_str = f"{value:.10f}".rstrip("0").rstrip(".")
        if not raw_str:
            raw_str = "0"
        return f"=ROUNDUP({raw_str},0)"

    wb = Workbook()
    first_sheet = True
    for sheet_name, items in sheets.items():
        if first_sheet:
            ws = wb.active
            ws.title = sheet_name
            first_sheet = False
        else:
            ws = wb.create_sheet(title=sheet_name)

        ws.insert_rows(1)
        title_text = f"Oracle Investment Proposal (as of {datetime.now().strftime('%m/%d/%Y')})"
        title_cell = ws.cell(row=1, column=1, value=title_text)
        title_cell.font = TITLE_FONT
        title_cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=TABLE_COLUMN_COUNT)

        apply_column_widths(ws)

        metadata_end_row = write_metadata(ws, build_metadata_rows(metadata), start_row=2)
        blank_row_index = metadata_end_row + 1
        header_row = blank_row_index + 1
        format_table_header(ws, header_row, headers)

        data_start = header_row + 1
        current_row = data_start
        storage_row_index: Optional[int] = None

        for item in items:
            ws.cell(row=current_row, column=1, value=item.description)
            ws.cell(row=current_row, column=2, value=item.part_number)

            part_cell = ws.cell(row=current_row, column=3)
            instance_cell = ws.cell(row=current_row, column=4)
            usage_cell = ws.cell(row=current_row, column=5)

            if item.category == "vpu":
                if storage_row_index is None:
                    raise ValueError("Storage row must be written before VPU row.")
                part_cell.value = f"=C{storage_row_index}*B$5"
                instance_cell.value = 1.0
                usage_cell.value = 1.0
            else:
                part_cell.value = roundup_formula(item.raw_base_quantity)
                instance_cell.value = 1.0
                if item.category in {"ocpu", "memory"}:
                    usage_cell.value = "=B$3"
                elif item.category == "storage":
                    usage_cell.value = 1.0
                    storage_row_index = current_row
                else:
                    usage_cell.value = item.usage_quantity

            ws.cell(row=current_row, column=6, value=item.unit_price)
            ws.cell(row=current_row, column=7, value=f"=C{current_row}*D{current_row}*E{current_row}*F{current_row}")
            current_row += 1

        total_row = current_row
        ws.cell(row=total_row, column=1, value="Monthly Total")
        for col in range(2, 7):
            ws.cell(row=total_row, column=col, value="")
        ws.cell(row=total_row, column=7, value=f"=SUM(G{data_start}:G{current_row - 1})")

        _, disclaimer_row = append_advisory(ws, total_row)

        apply_row_heights(
            ws,
            metadata_end_row=metadata_end_row,
            blank_row_index=blank_row_index,
            header_row=header_row,
            data_start=data_start,
            total_row=total_row,
            disclaimer_row=disclaimer_row,
        )
        format_metadata_block(ws, 2, metadata_end_row)
        format_data_rows(ws, data_start, total_row)
        format_total_row(ws, total_row)

    wb.save(output_path)


# =========================
# CLI / Main
# =========================

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
    return parser.parse_args(argv)


def main(argv: Optional[Iterable[str]] = None) -> int:
    args = parse_args(argv)

    rvtools_files = collect_rvtools_files(args.rvtools)
    if not rvtools_files:
        print("[ERROR] No RVTools Excel exports found for the given paths.", file=sys.stderr)
        return 1

    info(f"Powered-off VMs {'included' if args.include_poweredoff_vms else 'excluded'} in CPU/RAM totals.")
    info(f"Powered-off disks {'included' if args.include_poweredoff_disks else 'excluded'} in disk totals.")

    usage = aggregate_from_rvtools(rvtools_files, args.include_poweredoff_vms, args.include_poweredoff_disks)
    if not usage.source_files:
        print("[ERROR] No usable vInfo data found in the provided RVTools exports.", file=sys.stderr)
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
        print(f"[ERROR] Failed to compute costs: {exc}", file=sys.stderr)
        return 1

    metadata = {
        "source_files": ", ".join(usage.source_files),
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
        write_output(output_path, metadata, sheets, args.currency.upper())
    except Exception as exc:
        print(f"[ERROR] Failed to write output: {exc}", file=sys.stderr)
        return 1

    print("[OK] Cost summary generated successfully.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
