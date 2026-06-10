"""Aggregation and line-item math."""

import math

import pandas as pd

from oci_rvtools.compute import MIB_TO_GB, aggregate_vinfo, build_line_items
from oci_rvtools.model import AggregatedUsage, PriceRecord


def _vinfo(rows):
    cols = [
        "VM", "Powerstate", "Cluster", "CPUs", "Memory",
        "Total disk capacity MiB", "Provisioned MiB", "In Use MiB",
    ]
    return pd.DataFrame(rows, columns=cols)


def test_aggregate_powered_on_only_by_default():
    df = _vinfo([
        ["web1", "poweredOn", "cl1", 4, 8192, 102400, 102400, 51200],
        ["web2", "poweredOff", "cl1", 8, 16384, 204800, 204800, 102400],
    ])
    vcpu, ram, disk_total, disk_used, on, off = aggregate_vinfo(df, include_vms_off=False, include_disks_off=True)
    assert vcpu == 4.0          # only powered-on VM counted for CPU
    assert ram == 8192 / 1024.0
    assert on == 1 and off == 1
    # disks include powered-off by default
    assert disk_total > 0 and disk_used > 0


def test_aggregate_include_poweredoff_vms():
    df = _vinfo([
        ["web1", "poweredOn", "cl1", 4, 8192, 0, 102400, 51200],
        ["web2", "poweredOff", "cl1", 8, 16384, 0, 204800, 102400],
    ])
    vcpu, ram, *_ = aggregate_vinfo(df, include_vms_off=True, include_disks_off=True)
    assert vcpu == 12.0
    assert ram == (8192 + 16384) / 1024.0


def test_aggregate_exclude_poweredoff_disks():
    df = _vinfo([
        ["web1", "poweredOn", "cl1", 4, 8192, 102400, 102400, 51200],
        ["web2", "poweredOff", "cl1", 8, 16384, 999999, 999999, 999999],
    ])
    _, _, disk_total_incl, _, _, _ = aggregate_vinfo(df, include_vms_off=False, include_disks_off=True)
    _, _, disk_total_excl, _, _, _ = aggregate_vinfo(df, include_vms_off=False, include_disks_off=False)
    assert disk_total_excl < disk_total_incl


def test_provisioned_falls_back_when_total_zero():
    df = _vinfo([["web1", "poweredOn", "cl1", 2, 4096, 0, 51200, 25600]])
    _, _, disk_total, _, _, _ = aggregate_vinfo(df, include_vms_off=False, include_disks_off=True)
    assert disk_total == 51200 * MIB_TO_GB


def test_vcls_rows_are_dropped():
    df = _vinfo([
        ["vCLS-abc", "poweredOn", "cl1", 2, 2048, 1024, 1024, 512],
        ["app1", "poweredOn", "cl1", 4, 8192, 102400, 102400, 51200],
    ])
    vcpu, _, _, _, on, _ = aggregate_vinfo(df, include_vms_off=False, include_disks_off=True)
    assert vcpu == 4.0 and on == 1


class _StubPricing:
    PRICES = {
        "B97384": ("OCPU", 0.025),
        "B97385": ("Memory", 0.0015),
        "B91961": ("Storage", 0.0255),
        "B91962": ("VPU", 0.0017),
    }

    def get_price(self, part):
        name, price = self.PRICES[part]
        return PriceRecord(part_number=part, display_name=name, unit_price=price)


def test_build_line_items_structure_and_rounding():
    usage = AggregatedUsage(
        total_vcpu=7.0,        # -> 3.5 OCPU raw, ceil 4
        ram_gb=10.4,
        disk_total_gb=100.7,
        disk_used_gb=40.2,
    )
    parts = {"ocpu": "B97384", "memory": "B97385", "storage": "B91961", "vpu": "B91962"}
    sheets = build_line_items(usage, hours=730, vpu_value=10.0, pricing=_StubPricing(), parts=parts)

    assert set(sheets) == {"total_disk", "used_disk"}
    cats = [li.category for li in sheets["total_disk"]]
    assert cats == ["ocpu", "memory", "storage", "vpu"]

    ocpu_item = sheets["total_disk"][0]
    assert ocpu_item.raw_base_quantity == 3.5  # raw (rounding happens in Excel formula layer)
    # used_disk storage uses the used GB figure
    assert sheets["used_disk"][2].raw_base_quantity == 40.2
    # vpu raw quantity scales disk by vpu_value
    assert math.isclose(sheets["total_disk"][3].raw_base_quantity, 100.7 * 10.0)
