# capacity_report.py

Generates a cluster-level capacity report across all Cohesity clusters showing usable/used/free capacity, data reduction ratio, dedup and compression effectiveness, and total savings in TB.

Requires a **Helios API key**. The key is stored securely in the OS keychain after the first run.

Required packages:
```bash
pip install requests keyring openpyxl
```

**Version:** 1.1

---

## Usage

```bash
# Generate capacity report for all clusters:
python3 capacity_report.py

# Filter to one cluster:
python3 capacity_report.py --cluster prod-cluster

# Custom output filename:
python3 capacity_report.py --output capacity.xlsx

# Combine cluster filter and custom output:
python3 capacity_report.py --cluster prod-cluster --output prod-capacity.xlsx

# First run — prompts for Helios API key:
python3 capacity_report.py --apikey <key>

# Remove stored credentials:
python3 capacity_report.py --clear-credentials
```

Output file is created automatically as `capacity_report_YYYYMMDD_HHMMSS.xlsx` in the directory where the script is run (unless `--output` is specified).

---

## CLI Options

| Option | Description |
|--------|-------------|
| `--apikey` | Helios API key (optional — keychain used if omitted, or you will be prompted) |
| `--clear-credentials` | Remove stored API key from the OS keychain and exit |
| `--cluster` | Filter to a specific cluster (partial match supported) |
| `--output` | Output Excel filename (default: `capacity_report_YYYYMMDD_HHMMSS.xlsx` in current directory) |

---

## Report Fields

| Column | Description |
|--------|-------------|
| Cluster | Cluster name |
| Software Version | Cohesity cluster software version |
| Usable Capacity (TB) | Total capacity available for storing protected data |
| Used (TB) | Capacity currently in use by protected data and metadata |
| Free (TB) | Remaining capacity available (`Usable - Used`) |
| Used % | Percentage of usable capacity in use |
| Logical Data (TB) | Logical size of protected data (before dedup and compression) |
| Data Reduction Ratio | Ratio of logical data to physical data (`Logical / Physical`) |
| Dedup Ratio | Storage savings from deduplication as a ratio |
| Compression Ratio | Storage savings from compression as a ratio |
| Savings (TB) | Total storage saved via dedup and compression (`Logical - Physical`) |
| Node Count | Number of nodes in the cluster |

---

## Behavior

- **Cluster filter:** Uses partial name matching. If omitted, all clusters are listed.
- **Data reduction metrics:**
  - `Data Reduction Ratio` = logical bytes ÷ physical bytes (e.g., 3.2 means data is compressed to 1/3.2 of original size)
  - `Savings (TB)` = logical TB - physical TB (absolute space saved)
- **Output:** Excel file with styled headers (bold white text on dark blue background), frozen header row, and auto-fit columns.

---

## Notes

- These metrics reflect the current cluster state and include all storage used by snapshots, replication, archival metadata, and system data.
- Capacity is reported in decimal TB (thousands) as shown in the Cohesity UI.

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.1 | 2026-04-07 | Fixed empty columns: v1 `/public/cluster` returns `stats=null` by default — added `?fetchStats=true` parameter. Also fixed field paths: stats are nested under `usagePerfStats` / `dataUsageStats`, not at the top level |
| 1.0 | 2026-04-06 | Initial release |
