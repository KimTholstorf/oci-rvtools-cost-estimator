# OCI RVTools Cost Estimator — Implementation Notes

## Purpose
- Converts one or more RVTools Excel exports into an aggregated Oracle Cloud monthly cost estimate.
- Re-implements the cluster aggregation logic that previously lived in `rvtools_summarizer.py`, so the script no longer depends on intermediate summary files.
- Produces a heavily formatted Excel workbook with formulas so results stay editable/auditable.

## High-level flow
1. **CLI parsing** (argparse)  
   Flags cover input paths (`--rvtools`), pricing parts, hours, currency, VPU value, and powered-off inclusion toggles. Version is exposed via `--version`.
2. **File collection**  
   - Accepts a mixture of files & directories.  
   - Skips `~$` temp workbooks.  
   - Dedupes resolved paths, logging any skipped temp files once.
3. **Workbook loading**  
   - Uses `pandas.read_excel(..., sheet_name=None, engine="openpyxl")`.  
   - Searches sheets whose tokenised name contains `vinfo`. If none, the export is skipped with a warning.
4. **Column canonicalisation**  
   - Renames known aliases (`vInfoVMName` → `VM`, etc.).  
   - Tokenises headers to match canonical targets (e.g., `totaldiskcapacitymib` → `Total disk capacity MiB`).  
   - Collapses duplicate columns: numerics are summed, strings keep the first non-empty value.  
   - Backfills any missing canonical columns with zero/empty defaults.
5. **Aggregation** (`aggregate_vinfo`)  
   - Drops housekeeping VMs whose name starts with `vCLS`.  
   - Requires a non-empty cluster string (`unknown`, `nan`, etc. are ignored).  
   - Builds masks based on `Powerstate == poweredOn`.  
   - Controlled by two flags:  
     - `--include-poweredoff-vms` toggles whether powered-off machines contribute CPU/RAM.  
     - `--include-poweredoff-disks` / `--exclude-poweredoff-disks` control disk usage.  
   - Disk totals prefer `Total disk capacity MiB`; if zero, fall back to `Provisioned MiB`.  
   - Converts MiB to GiB via `effective_prov.sum() * (1024 / 953674)` (mirrors TiB→GiB as in the legacy script).  
   - Counts powered-on/off VMs separately for metadata reporting.
6. **Price lookup** (`PricingClient`)  
   - Caches results per part number.  
   - Calls CETools API with `partNumber` and `currencyCode`.  
   - JSON parsing is defensive: looks through common keys (`price`, `value`, nested `currencyCodeLocalizations`, etc.).  
   - Raises a descriptive error if no usable price is found.
7. **Line-item construction** (`build_line_items`)  
   - Normalises totals: rounds *up* (ceil) vCPUs/2 → OCPUs, RAM GB, and disk GB.  
   - OCPU and memory usage quantities multiply by `hours`.  
   - Disk usage quantity is monthly (1).  
   - VPU quantity derives from `disk_gb * vpu_value`, rounded up.  
   - Generates two sheet payloads: `total_disk` and `used_disk`.
8. **Excel generation** (`write_output`)  
   - Uses helper constants & functions for formatting (column widths, fonts, fills, row heights).  
   - Inserts a banner row with the current date.  
   - Metadata rows include source files, hours, currency, VPU, powered-on/off counts, and powered-off disk inclusion.  
   - Table header: “Description, Part Number, Part Qty, Instance Qty, Usage Qty, Unit Price, Monthly Cost”.  
   - Part Qty cells use `ROUNDUP(raw, 0)` formulas; VPU row references the storage part quantity (`=C{storage_row}*B$5`).  
   - Usage Qty for OCPU and memory references `B$3` (Hours per Month); storage/VPU rows use 1.  
   - Monthly cost column uses formula `=C*D*E*F`.  
   - Adds an advisory line (“Quote is for investment proposal only.”) in bold red, then a multi-line Oracle disclaimer with text wrapping.  
   - Sets explicit row heights (title 40, metadata 20, header 40, data 20, disclaimer 80) and left-aligns most content (table header vertically centred).

## Defaults & clamps
- Hours per month: `730`.
- Part numbers: `B97384` (OCPU), `B97385` (RAM GB-hr), `B91961` (Block storage GB-month), `B91962` (VPU).  
  All overridable via flags.
- VPU per GB: default `10`. Values `< 1` become `1`; values `> 120` clamp to `120`.
- Currency: `USD` (pass-through to the pricing API).

## Logging
- Logs inclusion/exclusion choices for powered-off VMs and disks.
- Emits aggregated totals (vCPU, RAM GB, disk GB) pre-rounding.
- Notes VPU per GB and hours used for pricing.
- On failure, errors are printed to stderr and the program exits with code `1`.

## Limitations / TODOs
- No built-in retry/backoff on pricing API failures (runs once per part).  
  Consider caching to disk or adding retry logic if this proves flaky.
- Currency support depends entirely on the API returning that currency. No conversion fallback implemented.
- Workbook formatting uses hard-coded column widths/row heights; if more rows are added, helper functions may need adjustment.
- Unit conversions assume RVTools MiB semantics; revisit if RVTools schema changes.
