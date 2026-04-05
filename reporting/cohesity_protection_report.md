# cohesity_protection_report.py

Generates an Excel report of protection group run status across Cohesity clusters. Default mode reports the last run for each group. Historical mode reports every run within a date range (one row per run).

All timestamps are shown in the **cluster's configured local timezone** (queried at runtime). Size columns report in **GB** as plain decimal numbers — the column header carries the `(GB)` label so cells stay numeric for Excel sorting and pivot tables.

Requires a **Helios API key** (for Helios mode) or cluster credentials (for direct mode). Credentials are stored securely in the OS keychain after the first run — never written to disk in plaintext or visible in the process list.

Required packages:
```bash
pip install openpyxl keyring
```

**Version:** 3.3

---

## Usage

```bash
# First run — prompts for Helios API key, saves to keychain:
python3 cohesity_protection_report.py

# Subsequent runs — key retrieved automatically:
python3 cohesity_protection_report.py --days 30
python3 cohesity_protection_report.py --start 2026-03-01 --end 2026-04-01

# Target one cluster via Helios:
python3 cohesity_protection_report.py --cluster <cluster-name> --days 7

# Direct cluster — password prompted and saved to keychain:
python3 cohesity_protection_report.py --cluster <ip> --username admin --domain LOCAL

# One-time Helios key override (not saved):
python3 cohesity_protection_report.py --apikey <key>

# Remove stored Helios API key:
python3 cohesity_protection_report.py --clear-credentials

# Remove stored cluster password:
python3 cohesity_protection_report.py --clear-credentials --cluster <ip>
```

Output file is created automatically as `cohesity_protection_report_YYYYMMDD_HHMMSS.xlsx` in the directory where the script is run.

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
| `--days N` | Historical mode: last N days of runs |
| `--start YYYY-MM-DD` | Historical mode: start date (inclusive) |
| `--end YYYY-MM-DD` | Historical mode: end date (inclusive, defaults to today) |

---

## Report Fields

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
