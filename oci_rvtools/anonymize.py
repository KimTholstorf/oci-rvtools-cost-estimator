# Copyright (c) 2026 Kim Tholstorf
# https://github.com/KimTholstorf/oci-rvtools-cost-estimator
# MIT License — see LICENSE file for details

"""Anonymise an RVTools export for privacy-conscious sharing.

Produces a stripped copy of a full RVTools workbook that still contains
everything this tool needs to calculate cost and detect the guest OS, but
nothing else:

* only the ``vInfo`` and ``vCPU`` sheets are kept;
* ``vInfo`` is reduced to the sizing + OS-detection columns (plus VM, Cluster,
  Datacenter and a domain-stripped DNS Name);
* ``vCPU`` is reduced to VM + the two OS columns (VM is the join key);
* DNS names are always reduced to their short hostname
  (``server.corp.local`` -> ``server``);
* with ``anonymize_names`` the VM, hostname, Cluster and Datacenter values are
  replaced by deterministic tokens (``VM0001`` …) and a mapping key is emitted
  so the anonymisation can be translated back.

This module is deliberately self-contained: it does not reuse the cost/ingest
pipeline, only the generic string tokenisers used for robust sheet/column
matching. Sizing columns are never modified, so a cost run on the anonymised
file yields totals identical to the original.
"""

from __future__ import annotations

import csv
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

from .ingest import _sheet_token, _to_token
from .log import info, warn

# Canonical columns to keep, matched case/format-insensitively by token.
_VINFO_KEEP = [
    "Powerstate", "VM", "Datacenter", "Cluster", "CPUs", "Memory",
    "Total disk capacity MiB", "Provisioned MiB", "In Use MiB",
    "OS according to the VMware Tools", "OS according to the configuration file",
    "DNS Name",
]
_VCPU_KEEP = [
    "VM", "OS according to the VMware Tools", "OS according to the configuration file",
]

_VINFO_KEEP_TOKENS = {_to_token(c) for c in _VINFO_KEEP}
_VCPU_KEEP_TOKENS = {_to_token(c) for c in _VCPU_KEEP}

# Role tokens used to locate specific columns for transformation.
_TOK_VM = _to_token("VM")
_TOK_DNS = _to_token("DNS Name")
_TOK_CLUSTER = _to_token("Cluster")
_TOK_DC = _to_token("Datacenter")


@dataclass
class AnonymizeResult:
    vinfo: pd.DataFrame
    vcpu: Optional[pd.DataFrame]
    key_rows: List[Tuple[str, str, str]] = field(default_factory=list)  # (category, anonymized, original)


# ── Column helpers ───────────────────────────────────────────────────────────

def _select_columns(df: pd.DataFrame, keep_tokens: set) -> pd.DataFrame:
    """Keep the first column per matching token, preserving original headers."""
    keep: List = []
    seen: set = set()
    for col in df.columns:
        tok = _to_token(str(col))
        if tok in keep_tokens and tok not in seen:
            keep.append(col)
            seen.add(tok)
    return df[keep].copy()


def _find_col(df: Optional[pd.DataFrame], role_token: str) -> Optional[str]:
    if df is None:
        return None
    for col in df.columns:
        if _to_token(str(col)) == role_token:
            return col
    return None


def _strip_domain(value) -> object:
    """``server.corp.local`` -> ``server``; blanks/NaN pass through unchanged."""
    if pd.isna(value):
        return value
    s = str(value).strip()
    if not s:
        return s
    return s.split(".", 1)[0]


# ── Token mapping ────────────────────────────────────────────────────────────

def _build_map(
    values_in_order,
    prefix: str,
    min_width: int = 2,
    vcls_aware: bool = False,
) -> Dict[str, str]:
    """Map distinct real values to deterministic tokens, in first-appearance order.

    When ``vcls_aware`` is set, values starting with "vCLS" get a ``vCLS-`` token
    so the cost pipeline still recognises and excludes them.
    """
    ordered: List[str] = []
    seen: set = set()
    for v in values_in_order:
        if pd.isna(v):
            continue
        s = str(v)
        if not s.strip():
            continue
        if s not in seen:
            seen.add(s)
            ordered.append(s)

    mapping: Dict[str, str] = {}
    if vcls_aware:
        normal = [v for v in ordered if not v.lower().startswith("vcls")]
        vcls = [v for v in ordered if v.lower().startswith("vcls")]
        w = max(min_width, len(str(len(normal))))
        for i, v in enumerate(normal, 1):
            mapping[v] = f"{prefix}{i:0{w}d}"
        wv = max(min_width, len(str(len(vcls))))
        for i, v in enumerate(vcls, 1):
            mapping[v] = f"vCLS-{i:0{wv}d}"
    else:
        w = max(min_width, len(str(len(ordered))))
        for i, v in enumerate(ordered, 1):
            mapping[v] = f"{prefix}{i:0{w}d}"
    return mapping


def _apply_map(df: pd.DataFrame, col: Optional[str], mapping: Dict[str, str]) -> None:
    if col is None:
        return
    df[col] = df[col].map(lambda v: mapping.get(str(v), v) if pd.notna(v) else v)


# ── Core transform ───────────────────────────────────────────────────────────

def build_anonymized(
    vinfo_raw: pd.DataFrame,
    vcpu_raw: Optional[pd.DataFrame],
    anonymize_names: bool,
) -> AnonymizeResult:
    vinfo = _select_columns(vinfo_raw, _VINFO_KEEP_TOKENS)
    vcpu = _select_columns(vcpu_raw, _VCPU_KEEP_TOKENS) if vcpu_raw is not None else None

    # DNS domain stripping is always applied in anonymise mode.
    dns_col = _find_col(vinfo, _TOK_DNS)
    if dns_col is not None:
        vinfo[dns_col] = vinfo[dns_col].map(_strip_domain)

    key_rows: List[Tuple[str, str, str]] = []
    if not anonymize_names:
        # Names are kept real, but strip any embedded FQDN domain from VM names so
        # a domain isn't leaked via the VM column (host.corp.local -> host). VM is
        # the vInfo<->vCPU join key, so strip it identically in both sheets.
        vm_col_vi = _find_col(vinfo, _TOK_VM)
        vm_col_vc = _find_col(vcpu, _TOK_VM)
        if vm_col_vi is not None:
            vinfo[vm_col_vi] = vinfo[vm_col_vi].map(_strip_domain)
        if vcpu is not None and vm_col_vc is not None:
            vcpu[vm_col_vc] = vcpu[vm_col_vc].map(_strip_domain)
        return AnonymizeResult(vinfo=vinfo, vcpu=vcpu, key_rows=key_rows)

    vm_col_vi = _find_col(vinfo, _TOK_VM)
    vm_col_vc = _find_col(vcpu, _TOK_VM)
    cluster_col = _find_col(vinfo, _TOK_CLUSTER)
    dc_col = _find_col(vinfo, _TOK_DC)

    # VM map spans both sheets (vInfo order first) so the join key stays consistent.
    vm_values: List = []
    if vm_col_vi is not None:
        vm_values += list(vinfo[vm_col_vi])
    if vm_col_vc is not None and vcpu is not None:
        vm_values += list(vcpu[vm_col_vc])
    vm_map = _build_map(vm_values, "VM", min_width=4, vcls_aware=True)
    _apply_map(vinfo, vm_col_vi, vm_map)
    if vcpu is not None:
        _apply_map(vcpu, vm_col_vc, vm_map)
    key_rows += [("VM", tok, real) for real, tok in vm_map.items()]

    # Hostname (post domain-strip)
    if dns_col is not None:
        host_map = _build_map(list(vinfo[dns_col]), "host", min_width=4)
        _apply_map(vinfo, dns_col, host_map)
        key_rows += [("Hostname", tok, real) for real, tok in host_map.items()]

    # Cluster
    if cluster_col is not None:
        cl_map = _build_map(list(vinfo[cluster_col]), "cluster", min_width=2)
        _apply_map(vinfo, cluster_col, cl_map)
        key_rows += [("Cluster", tok, real) for real, tok in cl_map.items()]

    # Datacenter
    if dc_col is not None:
        dc_map = _build_map(list(vinfo[dc_col]), "dc", min_width=2)
        _apply_map(vinfo, dc_col, dc_map)
        key_rows += [("Datacenter", tok, real) for real, tok in dc_map.items()]

    return AnonymizeResult(vinfo=vinfo, vcpu=vcpu, key_rows=key_rows)


# ── File I/O wrapper (CLI) ───────────────────────────────────────────────────

def anonymize_file(
    input_path: Path,
    anonymize_names: bool,
    out_dir: Optional[Path] = None,
) -> Optional[Tuple[Path, Optional[Path]]]:
    """Anonymise a single RVTools workbook on disk.

    Returns ``(xlsx_path, key_path_or_None)`` or ``None`` if the file has no
    vInfo sheet.
    """
    input_path = Path(input_path)
    try:
        xl = pd.read_excel(input_path, sheet_name=None, engine="openpyxl")
    except Exception as exc:
        warn(f"{input_path.name}: failed to load workbook ({exc})")
        return None

    vinfo_raw = None
    vcpu_raw = None
    for name, df in xl.items():
        tok = _sheet_token(name)
        if tok.endswith("vinfo") or "vinfo" in tok:
            vinfo_raw = df
        elif tok == "vcpu":
            vcpu_raw = df

    if vinfo_raw is None:
        warn(f"{input_path.name}: no vInfo sheet, cannot anonymize, skipping")
        return None

    result = build_anonymized(vinfo_raw, vcpu_raw, anonymize_names)

    out_dir = out_dir or input_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    stem = input_path.stem

    xlsx_path = out_dir / f"{stem}_anonymized.xlsx"
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        result.vinfo.to_excel(writer, sheet_name="vInfo", index=False)
        if result.vcpu is not None:
            result.vcpu.to_excel(writer, sheet_name="vCPU", index=False)
    info(f"Wrote {xlsx_path}")

    key_path: Optional[Path] = None
    if anonymize_names and result.key_rows:
        key_path = out_dir / f"{stem}_anonymized_key.csv"
        with open(key_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Category", "Anonymized", "Original"])
            writer.writerows(result.key_rows)
        info(f"Wrote {key_path}")

    return xlsx_path, key_path
