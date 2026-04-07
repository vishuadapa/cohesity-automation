# storage_domain_report.py

Reports storage domain utilization across Cohesity clusters — physical/logical capacity, data reduction, dedup and compression configuration, encryption, cloud tiering, and view count.

Requires a **Helios API key**. The key is stored securely in the OS keychain after the first run.

Required packages:
```bash
pip install requests keyring openpyxl
```

**Version:** 1.1

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
| Logical Data (TB) | Logical bytes ingested (before dedup/compression) |
| Physical Used (TB) | Physical bytes on disk after all reduction |
| Data Reduction Ratio | Overall ratio of logical to physical (`Logical / Physical`) |
| Dedup Ratio | Reduction from deduplication only (`Logical / After-Dedup`) |
| Compression Ratio | Reduction from compression only (`After-Dedup / Physical`) |
| Dedup Enabled | `Yes` if inline deduplication is active on this domain |
| Compression Type | Compression algorithm: e.g., `Low`, `High` |
| Encryption Type | Encryption algorithm: e.g., `Strong` |
| Erasure Coding | `Yes` if erasure coding is enabled (instead of RF2/RF3) |
| Cloud Tiering | `Yes` if cloud down-waterfall tiering is configured |

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
- Cloud tiering means the domain is configured to automatically move cold data to external targets; it does not indicate every backup is archived — that depends on individual protection policies.
- Dedup and compression settings apply to all new writes to that domain.

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.1 | 2026-04-07 | Fixed empty columns — all field names corrected to match v2 API: `stats.dataInBytes` (logical), `stats.storageConsumedBytes` (physical), `storagePolicy.deduplicationParams.enabled`, `compressionParams.type`, `encryptionType`, `erasureCodingParams.enabled`, `cloudDownWaterFallParams.thresholdSecs`. Removed Physical Capacity, View Count, Default Domain (not in API). Added Dedup Ratio, Compression Ratio, Erasure Coding columns. |
| 1.0 | 2026-04-06 | Initial release |
