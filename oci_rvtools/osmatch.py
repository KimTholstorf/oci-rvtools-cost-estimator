# Copyright (c) 2026 Kim Tholstorf
# https://github.com/KimTholstorf/oci-rvtools-cost-estimator
# MIT License — see LICENSE file for details

"""Operating-system → OCI compatibility classification.

RVTools reports an OS string per VM (from VMware Tools, falling back to the
configuration file). We classify each string as one of:

    yes      – an OCI platform image or BYOI-supported OS
    maybe    – plausibly works but the RVTools string is too generic, or BYOI
               with caveats (e.g. Windows Desktop on OCI Secure Desktop)
    no       – not supported on OCI (32-bit, ESXi, macOS, EOL Windows Server …)
    unknown  – no OS string available

Rules are an *ordered* list of compiled regular expressions; the first match
wins. Word boundaries (\\b) keep short tokens like "windows 7" from matching
inside longer strings, replacing the brittle trailing-space substring hacks
the table used previously.

Last reviewed against OCI documentation: 2026-06
  - Platform images:  https://docs.oracle.com/en-us/iaas/Content/Compute/References/images.htm
  - BYOI (Linux):     https://docs.oracle.com/en-us/iaas/Content/Compute/Tasks/importingcustomimagelinux.htm
  - BYOI (Windows):   https://docs.oracle.com/en-us/iaas/Content/Compute/Tasks/importingcustomimagewindows.htm
"""

from __future__ import annotations

import re
from typing import List, Tuple

# Verdict constants
YES = "yes"
MAYBE = "maybe"
NO = "no"
UNKNOWN = "unknown"

_SECURE_DESKTOP_NOTE = "Only supported on OCI Secure Desktop"
_BIT32_NOTE = "32-bit OS not supported on OCI"

# Ordered (regex, verdict, note). First match wins. Patterns assume a
# lower-cased input string. Keep specific rules before generic catch-alls.
_RULES_SOURCE: List[Tuple[str, str, str]] = [
    # ── Definite "no" ──────────────────────────────────────────────────────
    (r"vmware esx",              NO,    ""),
    (r"\besxi\b",               NO,    ""),
    (r"\bmac ?os\b",            NO,    ""),
    (r"\bsolaris\b",            NO,    ""),
    (r"\bnetware\b",            NO,    ""),
    (r"\bms-dos\b",             NO,    ""),
    (r"\bwindows xp\b",         NO,    ""),
    (r"\bwindows vista\b",      NO,    ""),
    (r"\bwindows 7\b",          NO,    ""),
    (r"\bwindows 8\b",          NO,    ""),
    (r"\(32-bit\)",             NO,    _BIT32_NOTE),
    # ── Windows Desktop → "maybe" (OCI Secure Desktop only) ─────────────────
    (r"\bwindows 10\b",         MAYBE, _SECURE_DESKTOP_NOTE),
    (r"\bwindows 11\b",         MAYBE, _SECURE_DESKTOP_NOTE),
    # ── Windows Server supported versions → "yes" ───────────────────────────
    (r"\bwindows server 2025\b", YES,  ""),
    (r"\bwindows server 2022\b", YES,  ""),
    (r"\bwindows server 2019\b", YES,  ""),
    (r"\bwindows server 2016\b", YES,  ""),
    # ── Windows Server old/unsupported versions → "no" ──────────────────────
    (r"\bwindows server 2012\b", NO,   ""),
    (r"\bwindows server 2008\b", NO,   ""),
    (r"\bwindows server 2003\b", NO,   ""),
    (r"\bwindows server 2000\b", NO,   ""),
    (r"\bwindows nt\b",          NO,   ""),
    # ── Windows Server generic (unknown/future version) → "maybe" ───────────
    (r"\bwindows server\b",      MAYBE, ""),
    # ── Oracle Linux / Autonomous → "yes" ───────────────────────────────────
    (r"\boracle autonomous linux\b", YES, ""),
    (r"\boracle linux\b",        YES,   ""),
    # ── Ubuntu: specific supported LTS → "yes"; generic → "maybe" ───────────
    (r"\bubuntu( linux)? 20\b",  YES,   ""),
    (r"\bubuntu( linux)? 22\b",  YES,   ""),
    (r"\bubuntu( linux)? 24\b",  YES,   ""),
    (r"\bubuntu\b",              MAYBE, ""),
    # ── RHEL: supported versions → "yes"; ancient → "no"; rest → "maybe" ────
    (r"\bred hat enterprise linux 6\b", YES, ""),
    (r"\bred hat enterprise linux 7\b", YES, ""),
    (r"\bred hat enterprise linux 8\b", YES, ""),
    (r"\bred hat enterprise linux 9\b", YES, ""),
    (r"\bred hat enterprise linux 2\b", NO,  ""),
    (r"\bred hat enterprise linux 3\b", NO,  ""),
    (r"\bred hat enterprise linux\b",   MAYBE, ""),
    # ── CentOS: modern specific → "yes"; old/generic → "maybe" ──────────────
    (r"\bcentos stream\b",       YES,   ""),
    (r"\bcentos 7\b",            YES,   ""),
    (r"\bcentos 6\b",            YES,   ""),
    (r"\bcentos\b",              MAYBE, ""),
    # ── SUSE: specific versions → "yes"; generic → "maybe" ──────────────────
    (r"\bsuse linux enterprise 11\b", YES, ""),
    (r"\bsuse linux enterprise 12\b", YES, ""),
    (r"\bsuse linux enterprise 15\b", YES, ""),
    (r"\bsuse linux enterprise\b",    MAYBE, ""),
    (r"\bopensuse\b",            YES,   ""),
    # ── Debian: versions 8+ → "yes"; generic → "maybe" ──────────────────────
    (r"\bdebian gnu/linux 8\b",  YES,   ""),
    (r"\bdebian gnu/linux 9\b",  YES,   ""),
    (r"\bdebian gnu/linux 10\b", YES,   ""),
    (r"\bdebian gnu/linux 11\b", YES,   ""),
    (r"\bdebian gnu/linux 12\b", YES,   ""),
    (r"\bdebian\b",              MAYBE, ""),
    # ── FreeBSD ──────────────────────────────────────────────────────────────
    (r"\bfreebsd\b",             MAYBE, ""),
    # ── Catch-all Linux / Other ─────────────────────────────────────────────
    # "linux" is a plain substring (no \b) so it still catches distro names where
    # "Linux" is fused to the word, e.g. "AlmaLinux", "RockyLinux".
    (r"\bother linux\b",         MAYBE, ""),
    (r"linux",                   MAYBE, ""),
    (r"\bother\b",               MAYBE, ""),
]

# Compiled (pattern, verdict, note) — built once at import.
_RULES: List[Tuple[re.Pattern, str, str]] = [
    (re.compile(pattern), verdict, note) for pattern, verdict, note in _RULES_SOURCE
]


def classify_os(os_str: str) -> Tuple[str, str]:
    """Classify a raw OS string. Returns ``(verdict, note)``.

    ``verdict`` is one of yes/maybe/no/unknown. ``note`` is a short advisory
    (may be empty). An empty/blank string yields ``(unknown, "")``.
    """
    if not os_str or not os_str.strip():
        return UNKNOWN, ""
    lower = os_str.lower()
    for pattern, verdict, note in _RULES:
        if pattern.search(lower):
            return verdict, note
    return UNKNOWN, ""


def detect_os(row_data) -> Tuple[str, str, str]:
    """Return ``(os_detected, verdict, note)`` for an RVTools vInfo row.

    Prefers "OS according to the VMware Tools" and falls back to
    "OS according to the configuration file" when the first is blank.
    """
    os_str = ""
    for col in ("OS according to the VMware Tools", "OS according to the configuration file"):
        raw = row_data.get(col, "")
        # Treat NaN/None/blank as missing without importing pandas here.
        if raw is None:
            continue
        text = str(raw).strip()
        if text and text.lower() != "nan":
            os_str = text
            break

    if not os_str:
        return "", UNKNOWN, ""

    verdict, note = classify_os(os_str)
    return os_str, verdict, note
