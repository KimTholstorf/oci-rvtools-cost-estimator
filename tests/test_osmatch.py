"""Locks the OS → OCI compatibility classification table."""

import pytest

from oci_rvtools.osmatch import classify_os, detect_os

# (os_string, expected_verdict, expected_note_substring_or_empty)
CASES = [
    # Windows Server — supported
    ("Microsoft Windows Server 2025 (64-bit)", "yes", ""),
    ("Microsoft Windows Server 2022 (64-bit)", "yes", ""),
    ("Microsoft Windows Server 2019 (64-bit)", "yes", ""),
    ("Microsoft Windows Server 2016 (64-bit)", "yes", ""),
    ("Microsoft Windows Server 2016 or later (64-bit)", "yes", ""),
    # Windows Server — unsupported / EOL
    ("Microsoft Windows Server 2012 R2 (64-bit)", "no", ""),
    ("Microsoft Windows Server 2012 (64-bit)", "no", ""),
    ("Microsoft Windows Server 2008 R2 (64-bit)", "no", ""),
    ("Microsoft Windows Server 2003 (64-bit)", "no", ""),
    ("Microsoft Windows Server 2000", "no", ""),
    ("Microsoft Windows NT 4.0", "no", ""),
    # Windows Server generic
    ("Microsoft Windows Server (64-bit)", "maybe", ""),
    # Windows Desktop
    ("Microsoft Windows 10 (64-bit)", "maybe", "Secure Desktop"),
    ("Microsoft Windows 11 (64-bit)", "maybe", "Secure Desktop"),
    ("Microsoft Windows 7 (64-bit)", "no", ""),
    ("Microsoft Windows 8.1 (64-bit)", "no", ""),
    ("Microsoft Windows XP Professional (32-bit)", "no", ""),
    ("Microsoft Windows Vista (32-bit)", "no", ""),
    # 32-bit catch
    ("Other (32-bit)", "no", "32-bit"),
    # Oracle Linux
    ("Oracle Linux 8 (64-bit)", "yes", ""),
    ("Oracle Autonomous Linux 8 (64-bit)", "yes", ""),
    # Ubuntu
    ("Ubuntu Linux 22.04 LTS (64-bit)", "yes", ""),
    ("Ubuntu Linux 20.04 (64-bit)", "yes", ""),
    ("Ubuntu Linux (64-bit)", "maybe", ""),
    # RHEL
    ("Red Hat Enterprise Linux 9 (64-bit)", "yes", ""),
    ("Red Hat Enterprise Linux 6 (64-bit)", "yes", ""),
    ("Red Hat Enterprise Linux 5 (64-bit)", "maybe", ""),
    ("Red Hat Enterprise Linux 3 (64-bit)", "no", ""),
    ("Red Hat Enterprise Linux 2.1 (64-bit)", "no", ""),
    # CentOS
    ("CentOS 7 (64-bit)", "yes", ""),
    ("CentOS Stream 9 (64-bit)", "yes", ""),
    ("CentOS 4/5 (64-bit)", "maybe", ""),
    ("CentOS 4/5/6/7 (64-bit)", "maybe", ""),
    # SUSE
    ("SUSE Linux Enterprise 11 (64-bit)", "yes", ""),
    ("SUSE Linux Enterprise 15 (64-bit)", "yes", ""),
    ("SUSE Linux Enterprise Server (64-bit)", "maybe", ""),
    ("openSUSE Leap 15 (64-bit)", "yes", ""),
    # Debian / FreeBSD
    ("Debian GNU/Linux 11 (64-bit)", "yes", ""),
    ("Debian GNU/Linux (64-bit)", "maybe", ""),
    ("FreeBSD (64-bit)", "maybe", ""),
    # Not supported platforms
    ("VMware ESXi 7.0", "no", ""),
    ("VMware ESX Server", "no", ""),
    ("Apple Mac OS X (64-bit)", "no", ""),
    ("Oracle Solaris 11 (64-bit)", "no", ""),
    # Generic catch-alls
    ("Other Linux (64-bit)", "maybe", ""),
    ("Other (64-bit)", "maybe", ""),
    # Distro names with "Linux" fused to the word must still match the catch-all
    ("AlmaLinux (64-bit)", "maybe", ""),
    ("RockyLinux (64-bit)", "maybe", ""),
]


@pytest.mark.parametrize("os_str,verdict,note_sub", CASES)
def test_classify(os_str, verdict, note_sub):
    got_verdict, got_note = classify_os(os_str)
    assert got_verdict == verdict, f"{os_str!r} -> {got_verdict} (want {verdict})"
    if note_sub:
        assert note_sub.lower() in got_note.lower()


def test_blank_is_unknown():
    assert classify_os("") == ("unknown", "")
    assert classify_os("   ") == ("unknown", "")
    assert classify_os("Plan 9 from Bell Labs") == ("unknown", "")


def test_detect_os_prefers_vmware_tools():
    row = {
        "OS according to the VMware Tools": "Oracle Linux 9 (64-bit)",
        "OS according to the configuration file": "CentOS 4/5 (64-bit)",
    }
    assert detect_os(row) == ("Oracle Linux 9 (64-bit)", "yes", "")


def test_detect_os_falls_back_to_config_file():
    row = {
        "OS according to the VMware Tools": "",
        "OS according to the configuration file": "Red Hat Enterprise Linux 8 (64-bit)",
    }
    os_str, verdict, _ = detect_os(row)
    assert os_str == "Red Hat Enterprise Linux 8 (64-bit)"
    assert verdict == "yes"


def test_detect_os_empty_row():
    row = {"OS according to the VMware Tools": "", "OS according to the configuration file": ""}
    assert detect_os(row) == ("", "unknown", "")
