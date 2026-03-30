<div align="center">
  <img src="logo_gh.png" width="345" height="93" alt="Logo"/>
  <h4 align="center">Turn RVTools exports into an Oracle Cloud monthly cost estimate</h4>
</div>

<br>
    
This utility ingests one or more RVTools `vInfo` sheets, pulls the latest Oracle Cloud Infrastructure list prices, and produces an Excel workbook with aggregate monthly costs for all included resources. 

Because OCI pricing scales linearly it It doesn’t price individual VMs. Instead it calculates the cost of a hypothetical single VM whose vCPU, RAM, and disk match the combined totals of the ingested workloads. That aggregated cost is identical to summing the per-VM prices, but a lot just easier to understand and calculate 🤓.

---

## 🚀 Features

- **Direct RVTools ingestion** – reads raw `RVTools_export_all.xlsx` files, normalises column names, and ignores housekeeping VMs (`vCLS-*`).
- **Configurable inclusion filters** – toggle powered-off VMs for CPU/RAM and powered-off disks for storage calculations independently.
- **Automatic unit handling** – converts MiB totals to GiB, rounds quantities up to whole units, and maps 2 vCPUs to 1 OCPU.
- **Live pricing lookup** – fetcheslist prices for configurable OCI part numbers via the [OCI pricing API](https://apexapps.oracle.com/pls/apex/cetools/api/v1/products/).
- **Polished Excel output** – writes `oci_cost_summary.xlsx` with two sheets (Total Disk vs. In Use Disk), metadata header, formulas, advisory text, and Oracle-styled formatting.
- **Console logging** – prints aggregation totals, pricing inputs, and powered-on/off inclusion choices to the console.

---

## ⚡ Quick start

```bash
# 1) Set up a virtual environment (optional but recommended)
python3 -m venv .venv
source .venv/bin/activate

# 2) Install dependencies
pip install pandas openpyxl

# 3) Run the estimator
python oci-rvtools-cost-estimator.py \
  --rvtools ./customer/RVTools_export_all.xlsx \
  --output oci_cost_summary.xlsx
```

The script contacts the OCI pricing API at runtime. Ensure the machine has outbound internet access.

---

## 🏗️ Installation options

### Local Python environment

Requirements:
- Python 3.9+
- `pandas`
- `openpyxl`

```bash
git clone https://github.com/KimTholstorf/oci-rvtools-cost-estimator.git
cd oci-rvtools-cost-estimator
python3 -m venv .venv
source .venv/bin/activate
pip install pandas openpyxl
```

You can also use [uv](https://github.com/astral-sh/uv) or `pipx` to keep dependencies isolated.

### One-off execution with uv

```bash
uv run --with pandas --with openpyxl oci-rvtools-cost-estimator.py \
  --rvtools ./customer/RVTools_export_all.xlsx
```

`uv` downloads dependencies into a cache and runs the script without a virtual environment.

---

## 📥 Input expectations

- RVTools workbook(s) in `.xlsx` format containing the `vInfo` sheet (default `RVTools_export_all.xlsx`).
- The script ignores temporary Excel lock files (`~$*.xlsx`) automatically.
- All calculations default to powered-on VMs, but powered-off VM CPU/RAM and disk capacity can be included via flags.

---

## 📤 Output workbook

The generated Excel file (`oci_cost_summary.xlsx` by default) contains:

1. **total_disk** – Monthly costs using the total provisioned disk capacity (TiB → GiB).
2. **used_disk** – Monthly costs using the reported “In Use” disk capacity.

Each sheet includes:

- Banner row stamped with the run date.
- Metadata block (source files, hours per month, currency, VPU value, powered-on/off inclusion flags).
- Pricing table with formulas for Part Qty, Instance Qty, Usage Qty, Unit Price, and Monthly Cost.
- Advisory text and Oracle disclaimer merged across all columns.

All quantities are rounded up to whole units before pricing. Block Volume Performance Units (VPU) scale with disk capacity (`VPU per GB` × GB).

---

## 🛠️ CLI reference

| Argument | Description |
| --- | --- |
| `--version` | Print the script version and exit. |
| `--rvtools PATH [PATH ...]` | One or more RVTools `.xlsx` files or directories to scan. Required. |
| `--output FILE` | Destination workbook path. Defaults to `oci_cost_summary.xlsx`. |
| `--hours HOURS` | Hours per month to bill. Defaults to `730`. |
| `--currency CODE` | Pricing currency (passed to OCI pricing API). Defaults to `USD`. |
| `--ocpu-part PART` | OCI part number for OCPU per hour (default `B97384`). |
| `--memory-part PART` | OCI part number for memory GB per hour (default `B97385`). |
| `--storage-part PART` | OCI part number for block storage capacity per month (default `B91961`). |
| `--vpu-part PART` | OCI part number for block volume performance units (default `B91962`). |
| `--vpu VALUE` | VPUs per GB (clamped 1–120, default `10`). |
| `--include-poweredoff-vms` | Include powered-off VMs when summing vCPU and RAM. |
| `--include-poweredoff-disks` | Include powered-off VMs when summing disk usage (default). |
| `--exclude-poweredoff-disks` | Ignore powered-off VMs when summing disk usage. |

Paths can point to folders; the script recursively picks up `.xlsx` files (skipping `~$` temp files). Duplicate files are de-duplicated.

---

## 📈 Examples

```bash
# Baseline run (powered-on VMs only, powered-off disks included)
python oci-rvtools-cost-estimator.py \
  --rvtools ./customer/RVTools_export_all.xlsx

# Aggregate multiple exports and change output name
python oci-rvtools-cost-estimator.py \
  --rvtools ./customer/site-a.xlsx ./customer/site-b.xlsx \
  --output reports/oci_cost_summary.xlsx

# Include powered-off VM CPU/RAM and exclude their disks
python oci-rvtools-cost-estimator.py \
  --rvtools ./customer/RVTools_export_all.xlsx \
  --include-poweredoff-vms \
  --exclude-poweredoff-disks

# Override pricing part numbers and hours per month
python oci-rvtools-cost-estimator.py \
  --rvtools ./customer/RVTools_export_all.xlsx \
  --hours 744 \
  --ocpu-part B12345 \
  --memory-part B67890 \
  --storage-part B54321 \
  --vpu-part B09876

# Cap VPU to 120 automatically; explicit value of 0 becomes 1
python oci-rvtools-cost-estimator.py \
  --rvtools ./customer/RVTools_export_all.xlsx \
  --vpu 0   # silently treated as 1
```

---

## 🧪 Testing hints

- Use `uv run --with pandas --with openpyxl` to execute without polluting a local environment.
- For offline validation, mock the Oracle pricing API response by monkeypatching `PricingClient.get_price`.
- Compare VM counts with RVTools (`Powerstate == poweredOn`) to confirm aggregation behaviour.

---

## ⚠️ Notes

- The script relies on real-time pricing data; expect run failures if the Oracle pricing API is unreachable.
- Pricing logic assumes USD list rates identical across regions. Adjust currency or part numbers as needed.
- Generated workbooks contain formulas and formatting; Excel recalculates automatically when opened.

---

Happy estimating! Contributions and pull requests are welcome.
