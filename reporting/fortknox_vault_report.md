# fortknox_vault_report.py

Reports FortKnox (RPaaS) vault data transfer by protection group using the
cluster-level **"Data Transferred to External Targets"** report API, routed
through Helios per-cluster. One row per protection group per vault.

Requires a **Helios API key** — FortKnox is a Helios-managed feature. The key is stored securely in the OS keychain after the first run; it is never written to disk in plaintext or visible in the process list.

Required packages:
```bash
pip install openpyxl keyring
```

Default date range is the last 7 days.

**Version:** 4.9

---

## API Endpoints Used

| Purpose | Method | Endpoint |
|---------|--------|----------|
| Cluster list | `GET` | `https://helios.cohesity.com/mcm/clusters/connectionStatus` |
| FortKnox vault IDs | `GET` | `https://helios.cohesity.com/irisservices/api/v1/public/vaults?includeFortKnoxVault=true` |
| Data transfer report | `GET` | `https://helios.cohesity.com/irisservices/api/v1/public/reports/dataTransferToVaults` |

The transfer report is fetched per cluster using the `accessClusterId` Helios
routing header. Parameters: `startTimeMsecs`, `endTimeMsecs`, `vaultIds` (one
per FortKnox vault ID on the cluster).

---

## Usage

```bash
# First run — prompts for API key and saves it to the OS keychain:
python3 fortknox_vault_report.py

# All subsequent runs — key retrieved from keychain automatically:
python3 fortknox_vault_report.py --days 30
python3 fortknox_vault_report.py --days 30 --mode trend
python3 fortknox_vault_report.py --start 2026-03-01 --end 2026-04-01 --mode trend
python3 fortknox_vault_report.py --cluster <cluster-name>

# One-time override with a different key (not saved):
python3 fortknox_vault_report.py --apikey <key>

# Remove stored credentials:
python3 fortknox_vault_report.py --clear-credentials
```

Output file is created automatically as `fortknox_vault_report_YYYYMMDD_HHMMSS.xlsx` in the directory where the script is run.

---

## CLI Options

| Option | Description |
|--------|-------------|
| `--apikey` | Helios API key (optional — if omitted, the key stored in the OS keychain is used, or you will be prompted) |
| `--clear-credentials` | Remove the stored API key from the OS keychain and exit |
| `--cluster` | Filter to a specific cluster (partial name match) |
| `--vault` | Filter to a specific vault (partial name match) |
| `--output` | Output Excel filename (default: `<script_name>_YYYYMMDD_HHMMSS.xlsx` in current directory) |
| `--days N` | Last N days (default: 7) |
| `--start YYYY-MM-DD` | Start date (inclusive) |
| `--end YYYY-MM-DD` | End date (inclusive, defaults to today) |
| `--start-msecs MS` | Exact start timestamp in milliseconds — paste directly from Chrome DevTools request URL to match UI exactly |
| `--end-msecs MS` | Exact end timestamp in milliseconds — paste directly from Chrome DevTools request URL |
| `--mode` | `summary` (default): one row per protection group for the full date range. `trend`: one row per protection group **per day** + a single **"Trend Charts"** sheet with all protection groups on one chart and column-filter dropdowns (▼) for navigation. |
| `--debug` | Print exact request URL and raw response sample for comparison with Chrome DevTools |

---

## Report Fields

Size columns report raw bytes as returned by the API. Apply your own unit conversion (e.g. ÷ 1,099,511,627,776 for TiB to match the Helios UI).

The script automatically queries the timezone from the first cluster via `GET /irisservices/api/v1/public/cluster` and uses it for all date boundary calculations, matching how the Helios UI interprets selected dates. The detected timezone and resolved timestamps are printed at startup for verification.

> **Trend Charts sheet (--mode trend):** Contains a single full-screen line chart (one series per protection group). The chart references the **Report** sheet data directly — no data is duplicated on the chart sheet. Report rows are sorted by (Protection Group, Cluster, Vault, Date) so each series maps to a contiguous row range. Protection groups where all Storage Consumed values are zero are excluded from the chart automatically.

> **Note — Primary field:** `Storage Consumed (Bytes)` is the key metric in this report. It reflects the total physical storage retained in the vault for a protection group across all snapshots, regardless of whether a transfer occurred in the queried period.

> **Caveat — Zero transfer values:** `Logical Transferred` and `Physical Transferred` reflect data moved to the vault **within the queried time window only**. The API always returns every protection group that has retained storage in the vault — even if no archival run occurred during the period. Those groups will show `0` for both transfer fields. This is expected and most visible in `--mode trend` where each day is queried separately: groups that archive every few days (per their policy schedule) will show `0` transfer on days with no run. `Storage Consumed` remains populated for all rows.

| Column | Source field | Description |
|--------|-------------|-------------|
| Period Start | _(script)_ | Start date of the report window (YYYY-MM-DD). Use with Period End to identify which run produced a row when stacking CSVs for trend analysis. |
| Period End | _(script)_ | End date of the report window (YYYY-MM-DD). |
| Cluster | `clusterName` | Cluster name as registered in Helios |
| Vault Name | `vaultName` | FortKnox vault target name |
| Vault Type | `vaultType` | Cloud platform: `kAzure`, `kAmazon`, etc. |
| Protection Group | `protectionJobName` | Protection group name |
| Logical Transferred (Bytes) | `numLogicalBytesTransferred` | Logical bytes sent to vault (pre-compression) |
| Physical Transferred (Bytes) | `numPhysicalBytesTransferred` | Bytes actually sent over the network (post-compression) |
| Storage Consumed (Bytes) | `storageConsumed` | Physical storage currently retained in this vault for this protection group across all retained snapshots |
| Storage Consumed (TB) | _(derived)_ | Same value converted to TB (÷ 1,000,000,000,000, 4 decimal places). Used as the Y axis in the trend chart. |

---

## Available Fields Not Currently Reported

Fields available in the `dataTransferToVaults` response that are not included by default.

| Field | API location | Notes |
|-------|-------------|-------|
| Vault-level logical total | `dataTransferSummary[].numLogicalBytesTransferred` | Aggregate across all protection groups for this vault |
| Vault-level physical total | `dataTransferSummary[].numPhysicalBytesTransferred` | Aggregate physical bytes for the vault |
| Vault-level storage total | `dataTransferSummary[].storageConsumedBytes` | Total retained storage across all groups for this vault |
| Number of protection jobs | `dataTransferSummary[].numProtectionJobs` | Count of protection groups writing to this vault |
| Logical transferred time-series | `dataTransferSummary[].logicalDataTransferredBytesDuringTimeRange[]` | Array of per-period logical transfer values |
| Physical transferred time-series | `dataTransferSummary[].physicalDataTransferredBytesDuringTimeRange[]` | Array of per-period physical transfer values |

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 4.9 | 2026-04-07 | Fixed Logical/Physical Transferred — switched from cumulative `dataTransferPerProtectionJob` fields to `GET /public/protectionRuns` per-run stats (time-windowed). Now matches UI values. |
| 4.8 | 2026-04-07 | Added Storage Consumed (TiB) column; fixed `--days` / default range to snap to calendar day start; fixed `end_of_day_msecs` to cover full last second |
| 4.7 | 2026-04-06 | Fixed x-axis dates not showing on trend charts |
| 4.6 | 2026-04-05 | Fixed axis labels; x-axis tick density auto-reduces |
| 4.5 | 2026-04-05 | One chart sheet per cluster |
| 4.3 | 2026-04-04 | Secure credential storage via OS keychain |
