# fortknox_vault_report.py

Reports FortKnox (RPaaS) vault data transfer by protection group using the
cluster-level **"Data Transferred to External Targets"** report API, routed
through Helios per-cluster. One row per protection group per vault.

Requires a **Helios API key** — FortKnox is a Helios-managed feature.

Default date range is the last 7 days.

**Version:** 3.4

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
# Last 7 days (default):
python3 fortknox_vault_report.py --apikey <key>

# Last 30 days:
python3 fortknox_vault_report.py --apikey <key> --days 30

# Specific date range:
python3 fortknox_vault_report.py --apikey <key> --start 2026-03-01 --end 2026-04-01

# Filter to one cluster:
python3 fortknox_vault_report.py --apikey <key> --cluster <cluster-name>

# Filter to one vault:
python3 fortknox_vault_report.py --apikey <key> --vault <vault-name>

# Debug (print raw API response):
python3 fortknox_vault_report.py --apikey <key> --debug
```

---

## CLI Options

| Option | Description |
|--------|-------------|
| `--apikey` | Helios API key (required) |
| `--cluster` | Filter to a specific cluster (partial name match) |
| `--vault` | Filter to a specific vault (partial name match) |
| `--output` | Output CSV filename (default: `fortknox_report.csv`) |
| `--days N` | Last N days (default: 7) |
| `--start YYYY-MM-DD` | Start date (inclusive) |
| `--end YYYY-MM-DD` | End date (inclusive, defaults to today) |
| `--start-msecs MS` | Exact start timestamp in milliseconds — paste directly from Chrome DevTools request URL to match UI exactly |
| `--end-msecs MS` | Exact end timestamp in milliseconds — paste directly from Chrome DevTools request URL |
| `--debug` | Print exact request URL and raw response sample for comparison with Chrome DevTools |

---

## Report Fields

Size columns are plain GB decimals — the `(GB)` label is in the header only so cells stay numeric for Excel sorting and pivot tables.

Size columns report raw bytes as returned by the API. Apply your own unit conversion (e.g. ÷ 1,099,511,627,776 for TiB to match the Helios UI).

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
