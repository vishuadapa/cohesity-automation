# storage_domain_report.py

Reports storage domain utilization across Cohesity clusters — physical/logical capacity, data reduction, dedup and compression configuration, encryption, cloud tiering, and view count.

Requires a **Helios API key**. The key is stored securely in the OS keychain after the first run.

Required packages:
```bash
pip install requests keyring openpyxl
```

**Version:** 1.0

---

## Usage

```bash
# Generate storage domain report for all clusters:
python3 storage_domain_report.py

# Filter to one cluster:
python3 storage_domain_report.py --cluster prod-cluster

# Custom output filename:
python3 storage_domain_report.py --output domains.xlsx

# Combine cluster filter and custom output:
python3 storage_domain_report.py --cluster prod-cluster --output prod-domains.xlsx

# First run — prompts for Helios API key:
python3 storage_domain_report.py --apikey <key>

# Remove stored credentials:
python3 storage_domain_report.py --clear-credentials
```

Output file is created automatically as `storage_domain_report_YYYYMMDD_HHMMSS.xlsx` in the directory where the script is run (unless `--output` is specified).

---

## CLI Options

| Option | Description |
|--------|-------------|
| `--apikey` | Helios API key (optional — keychain used if omitted, or you will be prompted) |
| `--clear-credentials` | Remove stored API key from the OS keychain and exit |
| `--cluster` | Filter to a specific cluster (partial match supported) |
| `--output` | Output Excel filename (default: `storage_domain_report_YYYYMMDD_HHMMSS.xlsx` in current directory) |

---

## Report Fields

| Column | Description |
|--------|-------------|
| Cluster | Cluster name |
| Domain Name | Storage domain name |
| Domain ID | Internal Cohesity storage domain identifier |
| Physical Capacity (TB) | Total capacity of the storage domain in TB |
| Logical Used (TB) | Logical bytes consumed (before dedup/compression) |
| Physical Used (TB) | Physical bytes on disk (after dedup/compression) |
| Data Reduction Ratio | Ratio of logical to physical (`Logical / Physical`) |
| Dedup Enabled | `True` if deduplication is active on this domain |
| Compression Policy | Compression setting: `Enabled`, `Disabled`, or policy name |
| Encryption Enabled | `True` if encryption is configured on this domain |
| Cloud Tiering Enabled | `True` if cloud tiering (archival to external targets) is configured |
| View Count | Number of NAS views hosted on this storage domain |
| Default Domain | `True` if this is the default domain (used for new views by default) |

---

## Behavior

- **Cluster filter:** Uses partial name matching. If omitted, all clusters are listed.
- **Data reduction metrics:**
  - `Data Reduction Ratio` = logical bytes ÷ physical bytes (e.g., 2.5 means domain achieves 2.5:1 reduction)
  - Represents actual dedup and compression effectiveness on this specific domain
- **Output:** Excel file with styled headers (bold white text on dark blue background), frozen header row, and auto-fit columns.

---

## Notes

- Each cluster typically has one default storage domain. Additional domains can be created for tiering, dedicated workloads, or multi-tenant isolation.
- Cloud tiering enabled means the domain is configured to send data to external targets (e.g., FortKnox), but does not indicate active archival of every backup to those targets — that depends on individual protection policies.
- Dedup and compression are cluster-wide features; enabling/disabling on a domain affects all new writes to that domain.
