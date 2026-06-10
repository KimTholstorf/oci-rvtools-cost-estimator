# Copyright (c) 2026 Kim Tholstorf
# https://github.com/KimTholstorf/oci-rvtools-cost-estimator
# MIT License — see LICENSE file for details

"""RVTools ingestion: locate files, load the vInfo sheet, normalise columns."""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional, Sequence

import pandas as pd

from .log import info, warn

# =========================
# vInfo column mapping
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
    "OS according to the VMware Tools",
    "OS according to the configuration file",
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
    "vInfoOSConfig": "OS according to the configuration file",
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
    "osaccordingtovmwaretools": "OS according to the VMware Tools",
    "osaccordingtotheconfigurationfile": "OS according to the configuration file",
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

_OS_COLS = ["OS according to the VMware Tools", "OS according to the configuration file"]


# =========================
# Helpers
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
    vcpu_df = None
    for sheet_name, df in xl.items():
        token = _sheet_token(sheet_name)
        if token.endswith("vinfo") or "vinfo" in token:
            vinfo_df = df
        elif token == "vcpu":
            vcpu_df = df

    if vinfo_df is None:
        warn(f"{filepath.name}: missing vInfo sheet, skipping")
        return None

    vinfo_df = canonicalize_vinfo(vinfo_df)

    # If vInfo is missing OS columns, try to pull them from vCPU (common in trimmed exports)
    missing_os = [
        c for c in _OS_COLS
        if c not in vinfo_df.columns or vinfo_df[c].astype(str).str.strip().eq("").all()
    ]
    if missing_os and vcpu_df is not None and "VM" in vcpu_df.columns:
        available = [c for c in _OS_COLS if c in vcpu_df.columns]
        if available:
            os_lookup = (
                vcpu_df[["VM"] + available]
                .drop_duplicates(subset="VM")
            )
            vinfo_df = vinfo_df.merge(os_lookup, on="VM", how="left", suffixes=("", "_vcpu"))
            # Prefer the freshly merged column; drop the _vcpu suffixed duplicate if any
            for col in available:
                vcpu_col = f"{col}_vcpu"
                if vcpu_col in vinfo_df.columns:
                    vinfo_df[col] = vinfo_df[col].where(
                        vinfo_df[col].astype(str).str.strip().ne(""), vinfo_df[vcpu_col]
                    )
                    vinfo_df.drop(columns=[vcpu_col], inplace=True)
            info(f"vInfo: pulled OS columns from vCPU sheet ({', '.join(available)})")

    return vinfo_df


# =========================
# Cluster / filter helpers
# =========================

def valid_cluster(value: object) -> bool:
    if pd.isna(value):
        return False
    s = str(value).strip()
    if not s:
        return False
    return s.lower() not in INVALID_CLUSTER_VALUES


def prepare_vinfo_df(df: pd.DataFrame, verbose: bool = True) -> pd.DataFrame:
    """Remove vCLS housekeeping VMs and apply cluster filter when cluster data exists."""
    df = df.copy()
    df = df[~df["VM"].astype(str).str.startswith("vCLS", na=False)]
    if df["Cluster"].apply(valid_cluster).any():
        df = df[df["Cluster"].apply(valid_cluster)]
    elif verbose:
        info("No valid cluster data detected — cluster filter skipped, all VMs included")
    return df


def collect_rvtools_files(paths: Sequence[str]) -> List[Path]:
    files: List[Path] = []
    seen: set = set()
    skipped_temp: set = set()

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


def apply_vm_filter(
    df: pd.DataFrame,
    datacenters: Optional[List[str]],
    clusters: Optional[List[str]],
) -> pd.DataFrame:
    """Filter vInfo rows by Datacenter and/or Cluster (case-insensitive, AND between flags)."""
    if not datacenters and not clusters:
        return df

    mask = pd.Series(True, index=df.index)

    if datacenters:
        dc_lower = [d.lower() for d in datacenters]
        df_dc = df["Datacenter"].astype(str).str.strip().str.lower()
        matched = set(df_dc[df_dc.isin(dc_lower)].unique())
        for d in dc_lower:
            if d not in matched:
                warn(f"No VMs found for Datacenter '{d}' — check spelling and case")
        mask &= df_dc.isin(dc_lower)

    if clusters:
        cl_lower = [c.lower() for c in clusters]
        df_cl = df["Cluster"].astype(str).str.strip().str.lower()
        matched = set(df_cl[df_cl.isin(cl_lower)].unique())
        for c in cl_lower:
            if c not in matched:
                warn(f"No VMs found for Cluster '{c}' — check spelling and case")
        mask &= df_cl.isin(cl_lower)

    return df[mask]


def list_datacenters_and_clusters(files: Sequence[Path]) -> None:
    """Print all unique Datacenter and Cluster names found across the given files."""
    topology: Dict[str, set] = {}
    for file in files:
        df = load_vinfo_dataframe(file)
        if df is None:
            continue
        df = df[df["Cluster"].apply(valid_cluster)]
        for _, row in df[["Datacenter", "Cluster"]].drop_duplicates().iterrows():
            dc = str(row["Datacenter"]).strip() or "(unknown)"
            cl = str(row["Cluster"]).strip()
            topology.setdefault(dc, set()).add(cl)
    if not topology:
        warn("No Datacenter/Cluster data found in the provided files.")
        return
    for dc in sorted(topology):
        print(f"Datacenter: {dc}")
        for cl in sorted(topology[dc]):
            print(f"  Cluster: {cl}")
