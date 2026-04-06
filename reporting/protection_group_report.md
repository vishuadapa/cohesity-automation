# protection_group_report.py

Generates an Excel report of protection group activity across Cohesity clusters. Supports two modes:

- **`--mode summary`** (default): last run per group, or all runs within a date range (one row per run). All timestamps in the cluster's local timezone.
- **`--mode trend`**: daily storage consumed per protection group. One row per group per day + a trend chart sheet per cluster. Mirrors the FortKnox report's trend capability for local cluster storage.

Requires a **Helios API key** (for Helios mode) or cluster credentials (for direct mode). Credentials are stored securely in the OS keychain after the first run.

> **Storage Consumed limitation:** Cohesity does not provide a historical per-day storage consumed value per protection group through any public API accessible via Helios. Both `--mode summary` and `--mode trend` source this value from `GET /stats/consumers`, which returns the **current state of the cluster at the time the script runs**. In trend mode the column is named **"Storage Consumed (As of End Date)"** — it will show the same value on every date row for a given protection group. For per-day backup activity trends, use the run detail columns (Data Written, Objects Succeeded, Run Status, etc.).

Required packages:
```bash
pip install openpyxl keyring
```

**Version:** 4.4

---

## Usage

```bash
# First run — prompts for Helios API key, saves to keychain:
python3 protection_group_report.py

# Summary mode (default) — last run per group:
python3 protection_group_report.py --days 30
python3 protection_group_report.py --start 2026-03-01 --end 2026-04-01

# Trend mode — daily storage consumed per group + chart per cluster:
python3 protection_group_report.py --mode trend --days 30
python3 protection_group_report.py --mode trend --start 2026-03-01 --end 2026-04-01
python3 protection_group_report.py --mode trend          # defaults to last 30 days

# Debug — print raw timeSeriesStats response to verify metric name:
python3 protection_group_report.py --mode trend --days 7 --debug

# Target one cluster via Helios:
python3 protection_group_report.py --mode trend --cluster <cluster-name> --days 30

# Direct cluster — password prompted and saved to keychain:
python3 protection_group_report.py --cluster <ip> --username admin --domain LOCAL

# One-time Helios key override (not saved):
python3 protection_group_report.py --apikey <key>

# Remove stored credentials:
python3 protection_group_report.py --clear-credentials
python3 protection_group_report.py --clear-credentials --cluster <ip>
```

Output file is created automatically as `protection_group_report_YYYYMMDD_HHMMSS.xlsx` in the directory where the script is run.

---

## CLI Options

| Option | Description |
|--------|-------------|
| `--apikey` | Helios API key (optional — keychain used if omitted, or you will be prompted) |
| `--clear-credentials` | Remove stored credentials from the OS keychain and exit. Without `--cluster`: clears Helios API key. With `--cluster`: clears the cluster password. |
| `--cluster` | Cluster name filter (Helios) or IP/hostname (direct auth) |
| `--username` | Username for direct cluster auth (default: `admin`) |
| `--domain` | Auth domain — `LOCAL` or AD domain name (default: `LOCAL`) |
| `--output` | Output Excel filename (default: `<script_name>_YYYYMMDD_HHMMSS.xlsx` in current directory) |
| `--mode` | `summary` (default): last run or all runs in date range. `trend`: daily storage consumed per group + chart. |
| `--debug` | Print raw `timeSeriesStats` API response — use to verify metric name if trend data is empty |
| `--days N` | Last N days (e.g. `--days 30`). Trend mode defaults to 30 days if no range given. |
| `--start YYYY-MM-DD` | Start date (inclusive) |
| `--end YYYY-MM-DD` | End date (inclusive, defaults to today) |

---

## Trend Mode Report Fields (`--mode trend`)

| Column | Description |
|--------|-------------|
| Date | Calendar date (YYYY-MM-DD) of the data point |
| Cluster | Cluster name |
| Protection Group | Name of the protection group |
| Environment | Workload type: `VMware`, `PhysicalFiles`, `Oracle`, etc. |
| Policy Name | Protection policy assigned to the group |
| Is Active | `True` if the group is active |
| Is Paused | `True` if the group is currently paused |
| Storage Consumed (As of End Date, Bytes) | Current physical storage this group occupies on the cluster, sourced from `GET /stats/consumers`. **Same value on every date row** — reflects cluster state at script run time, not a per-day historical value. See limitation note above. |
| Storage Consumed (As of End Date, TB) | Same value ÷ 1,000,000,000,000, 4 decimal places. Used as the Y axis in trend charts. |
| Runs On Day | Number of protection runs that started on this date (`0` = no run scheduled/completed) |
| Run Type | Type of the last run on this date: `Regular`, `Full`, `Log`, `System` |
| Sla Violated | `Yes` if any run on this date breached the policy SLA window |
| Run Status | Status of the last run: `Succeeded`, `Failed`, `SucceededWithWarning`, `Canceled`, `Running` |
| Run Start | Earliest run start timestamp on this date (cluster local time) |
| Run End | Latest run end timestamp on this date (cluster local time) |
| Duration Mins | Total wall-clock minutes across all runs on this date |
| Snapshot Expiry | Snapshot expiry of the last run on this date |
| Objects Succeeded | Sum of successfully backed-up objects across all runs on this date |
| Objects Failed | Sum of failed objects across all runs on this date |
| Objects Canceled | Sum of canceled objects across all runs on this date |
| Data Read (GB) | Total bytes read from source across all runs on this date |
| Data Written (GB) | Total bytes written to Cohesity storage across all runs on this date |
| Logical Size (GB) | Total logical size of data protected across all runs on this date |
| Replication Status | Replication status of the last run (pipe-separated if multiple targets) |
| Replication Targets | Replication target cluster names of the last run |
| Replication Expiry | Replication expiry dates of the last run |
| Archive Status | Archive status of the last run (pipe-separated if multiple targets) |
| Archive Targets | Archive target names of the last run |
| Archive Expiry | Archive expiry dates of the last run |

> **Multi-run days:** Object counts and data sizes are summed across all runs. Status, type, targets, and expiry reflect the last run. `Runs On Day = 0` means the group had no backup run on that date — storage consumed is still shown from the retained snapshot estimate.

---

## Summary / Historical Mode Report Fields (`--mode summary`)

| Column | Description |
|--------|-------------|
| Cluster | Cluster name as registered in Helios, or IP for direct auth |
| Protection Group | Name of the protection group |
| Environment | Workload type: `VMware`, `PhysicalFiles`, `Oracle`, etc. |
| Is Active | `True` if the group is active (not deleted or paused permanently) |
| Is Paused | `True` if the group is currently paused |
| Policy Name | Name of the protection policy assigned to the group (resolved from policy ID via a cached API lookup) |
| Storage Consumed (GB) | Total physical space this group currently occupies on the cluster across all retained snapshots, after dedup and compression. Group-level aggregate, not per-run. |
| Run # | Sequential run number within the group (historical mode only) |
| Run Type | Backup type: `Regular`, `Full`, `Log`, `System` |
| Sla Violated | `Yes` if the run breached the policy SLA window |
| Run Status | Outcome of the run: `Succeeded`, `Failed`, `SucceededWithWarning`, `Canceled`, `Running` |
| Run Start | Run start timestamp in the cluster's local timezone |
| Run End | Run end timestamp in the cluster's local timezone |
| Duration Mins | Elapsed wall-clock minutes from start to end |
| Snapshot Expiry | Date/time this snapshot will be automatically deleted per the policy retention |
| Objects Succeeded | Number of objects (VMs, hosts, databases) that backed up successfully |
| Objects Failed | Number of objects that failed |
| Objects Canceled | Number of objects that were canceled before completing |
| Data Read (GB) | Bytes read from the source during backup (logical source data scanned) |
| Data Written (GB) | Bytes written to Cohesity storage for this run (post-dedup/compression) |
| Logical Size (GB) | Logical size of the data protected in this run |
| Replication Status | Status of each replication target, pipe-separated if multiple |
| Replication Targets | Name of each remote cluster receiving replicated data |
| Replication Expiry | Retention expiry date on each remote cluster |
| Archive Status | Status of each archival target (e.g., S3, Azure, Tape) |
| Archive Targets | Name of each archive target |
| Archive Expiry | Retention expiry date on each archive target |

---

## Available Fields Not Currently Reported

Fields available in the Cohesity v2 API that are not included in the default output but can be added if needed.

| Field | API location | Notes |
|-------|-------------|-------|
| Per-object status | `runs[].objects[].status` | Detailed per-VM/host outcome; adds one row per object |
| Per-object name | `runs[].objects[].object.name` | Individual VM/host/DB names within the run |
| Per-object error message | `runs[].objects[].localSnapshotInfo.failedAttempts[].message` | Specific error text for failed objects |
| Bytes replicated (per target) | `runs[].replicationInfo.replicationTargetResults[].stats.physicalBytesTransferred` | Network bytes actually sent to each replication target |
| Bytes archived (per target) | `runs[].archivalInfo.archivalTargetResults[].stats.physicalBytesTransferred` | Bytes uploaded to each archive target |
| DataLock / WORM status | `runs[].localBackupInfo.dataLock` | Whether the snapshot is WORM-locked |
| Legal hold status | `runs[].localBackupInfo.onLegalHold` | Whether the snapshot is on legal hold |
| Warning messages | `runs[].localBackupInfo.messages` | Array of warning strings logged during the run |
| Policy ID | `protectionGroups[].policyId` | Raw policy UUID; Policy Name is reported instead but ID is available if needed |
| Storage domain (View Box) | `runs[].localBackupInfo.storageDomainId` | Which storage domain the snapshot lives on |
| Cloud archive tier | `runs[].archivalInfo.archivalTargetResults[].tierSettings` | Tiering info for cloud archives (S3 Standard, Glacier, etc.) |
| Cloud spin-up status | `runs[].cloudSpinInfo` | Status of any cloud spin-up (AWS/Azure DR) |
| Tenant ID | `protectionGroups[].tenantId` | Multi-tenant environment: which tenant owns this group |
