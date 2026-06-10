"""Ingestion: column canonicalisation and the vCPU OS-column fallback."""

import pandas as pd

from oci_rvtools.ingest import canonicalize_vinfo, load_vinfo_dataframe


def test_canonicalize_token_rename_and_backfill():
    df = pd.DataFrame({
        "VM": ["a"],
        "Powerstate": ["poweredOn"],
        "CPUs": [4],
        "Memory": [8192],
    })
    out = canonicalize_vinfo(df)
    # required columns backfilled
    for col in ("Cluster", "Datacenter", "Provisioned MiB",
                "OS according to the VMware Tools",
                "OS according to the configuration file"):
        assert col in out.columns


def test_alias_rename():
    df = pd.DataFrame({"vInfoVMName": ["a"], "vInfoCPUs": [2]})
    out = canonicalize_vinfo(df)
    assert "VM" in out.columns and "CPUs" in out.columns


def test_os_fallback_from_vcpu_sheet(tmp_path):
    vinfo = pd.DataFrame({
        "VM": ["app1", "app2"],
        "Powerstate": ["poweredOn", "poweredOn"],
        "CPUs": [4, 8],
        "Memory": [8192, 16384],
        "Total disk capacity MiB": [102400, 204800],
        "Provisioned MiB": [102400, 204800],
        "In Use MiB": [51200, 102400],
    })
    vcpu = pd.DataFrame({
        "VM": ["app1", "app2"],
        "CPUs": [4, 8],
        "OS according to the VMware Tools": ["Ubuntu Linux (64-bit)", "CentOS 4/5 (64-bit)"],
        "OS according to the configuration file": ["Ubuntu Linux (64-bit)", "CentOS 4/5 (64-bit)"],
    })
    path = tmp_path / "trimmed.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as xl:
        vinfo.to_excel(xl, sheet_name="vInfo", index=False)
        vcpu.to_excel(xl, sheet_name="vCPU", index=False)

    out = load_vinfo_dataframe(path)
    assert out is not None
    osvals = out.set_index("VM")["OS according to the VMware Tools"].to_dict()
    assert osvals["app1"] == "Ubuntu Linux (64-bit)"
    assert osvals["app2"] == "CentOS 4/5 (64-bit)"


def test_missing_vinfo_returns_none(tmp_path):
    path = tmp_path / "nope.xlsx"
    pd.DataFrame({"x": [1]}).to_excel(path, sheet_name="Other", index=False)
    assert load_vinfo_dataframe(path) is None
