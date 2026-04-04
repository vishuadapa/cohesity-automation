# fortknox_vault_report.py

Reports FortKnox (RPaaS) vault archival activity per protection group using the
**Helios-native activity API** — a single POST call returns all FortKnox runs
across every connected cluster. Vault storage consumed is sourced from the
Helios Reporting API.

Requires a **Helios API key** — FortKnox is a Helios-managed feature; direct
cluster auth is not applicable.

Each row represents one FortKnox archival run. Default range is the last 7 days.

**Version:** 2.1

---

## API Endpoints Used

| Purpose | Method | Endpoint |
|---------|--------|----------|
| FortKnox activity | `POST` | `https://helios.cohesity.com/v2/mcm/data-protect/protection-group/activity` |
| Vault storage stats | `POST` | `https://helios.cohesity.com/heliosreporting/api/v1/public/components/1600/preview` |

Activity POST body:
```json
{
  "activityTypes": ["ArchivalRun"],
  "archivalRunParams": {"isRpaas": true},
  "startTimeUsecs": <epoch_usecs>,
  "endTimeUsecs":   <epoch_usecs>
}
```

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
```

---

## CLI Options

| Option | Description |
|--------|-------------|
| `--apikey` | Helios API key (required) |
| `--cluster` | Filter output to a specific cluster (partial name match) |
| `--vault` | Filter output to a specific vault (partial name match) |
| `--output` | Output CSV filename (default: `fortknox_report.csv`) |
| `--days N` | Last N days (default: 7) |
| `--start YYYY-MM-DD` | Start date (inclusive) |
| `--end YYYY-MM-DD` | End date (inclusive, defaults to today) |

---

## Report Fields

Size columns are plain GB decimals — the `(GB)` label is in the header only so cells stay numeric for Excel sorting and pivot tables.

| Column | Description |
|--------|-------------|
| Cluster | Cluster name as registered in Helios |
| Protection Group | Name of the protection group |
| Environment | Workload type: `kVMware`, `kPhysical`, `kOracle`, etc. |
| Policy | Name of the protection policy |
| Vault Name | FortKnox vault target name |
| Region | RPaaS region where the vault is hosted |
| Status | Run outcome: `Succeeded`, `Failed`, `Running`, `Canceled` |
| Start Time | Vault transfer start time (UTC) |
| End Time | Vault transfer end time (UTC) |
| Duration Mins | Transfer duration in minutes |
| Logical Size (GB) | Logical size of the data in this vault snapshot |
| Logical Transferred (GB) | Logical bytes sent to vault (pre-compression) |
| Physical Transferred (GB) | Bytes actually sent over the network (post-compression). Compare to Logical Transferred to see the effective compression ratio. |
| Vault Target Type | Cloud platform hosting the vault: `Azure`, `AWS`, etc. |
| Vault Storage Consumed (GB) | Total physical storage currently retained in this vault for this cluster, across all retained snapshots (`scRetainedBytes`). Vault-level — same value on every run row for the same cluster+vault. |
| Vault Storage Growth (GB) | Change in retained storage over the reporting period (`scRetainedBytesGrowth`). |
| Vault Cumul Transferred (GB) | Cumulative total bytes ever transferred to this vault from this cluster (`cumulativeDataTransferredBytes`). |
| Vault Period Transferred (GB) | Bytes transferred to the vault in the last 7 days (`dataTransferredBytes`). |
| Vault Daily Growth Pct | Average daily growth rate as a percentage of current vault storage (`scRetainedDailyGrowthPercent`). |

---

## Available Fields Not Currently Reported

| Field | API location | Notes |
|-------|-------------|-------|
| Object count | `archivalRunParams.objectsCount` | Number of objects in the run (if returned by API) |
| SLA violated | `archivalRunParams.isSlaViolated` | Whether the vault run breached SLA |
| Run type | `archivalRunParams.backupType` | Regular, Full, Log, etc. |
| Expiry time | `archivalRunParams.expiryTimeUsecs` | When this vault snapshot will be deleted |
| CAD (Cloud Archive Direct) | `archivalRunParams.isCadArchive` | Whether this was a direct cloud archive |
| WORM compliant | `archivalRunParams.wormProperties.isArchiveWormCompliant` | DataLock/WORM status |
| Objects succeeded/failed | `archivalRunParams.successfulObjectsCount` / `failedObjectsCount` | Per-run object-level counts |
