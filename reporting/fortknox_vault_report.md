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

**Version:** 4.10

---

## API Endpoints Used

| Purpose | Method | Endpoint |
|---------|--------|----------|
| Cluster list | `GET` | `https://helios.cohesity.com/mcm/clusters/connectionStatus` |
| FortKnox vault IDs | `GET` | `https://helios.cohesity.com/irisservices/api/v1/public/vaults?includeFortKnoxVault=true` |
| Storage consumed (per protection group) | `GET` | `https://helios.cohesity.com/irisservices/api/v1/public/reports/dataTransferToVaults` |
| Logical / Physical transferred (per run) | `GET` | `https://helios.cohesity.com/irisservices/api/v1/public/protectionRuns` |
| Cluster timezone | `GET` | `https://helios.cohesity.com/irisservices/api/v1/public/cluster` |

Both report endpoints are fetched per cluster using the `accessClusterId` Helios routing header.

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

The script automatically queries the cluster timezone via `GET /public/cluster` and uses it for all date boundary calculations, matching how the Helios UI interprets selected dates.

---

### Why Logical and Physical Transferred do NOT come from the Helios UI

> **Important:** The Logical Transferred and Physical Transferred values in this report will differ from what you see in the Helios UI's "Data Transfer to Vaults" page when broken down by protection group. This is intentional — the UI figures are inaccurate at the per-protection-group level. Here is why:

The `dataTransferToVaults` API (which the Helios UI uses) exposes two field families:

| Field location | Field name | What it actually contains |
|---|---|---|
| Per protection group | `numLogicalBytesTransferred` | **Cumulative lifetime total** — all bytes ever transferred since the job was created, regardless of date range |
| Per protection group | `numPhysicalBytesTransferred` | Same — cumulative, not time-windowed |
| Vault level (aggregate) | `logicalDataTransferredBytesDuringTimeRange[]` | Correctly time-windowed daily array — but vault total only, not broken down per protection group |

The Cohesity API does not provide time-windowed logical/physical transfer data broken down per protection group via `dataTransferToVaults`. If you select "last 7 days" in the UI, the per-group logical/physical numbers shown still reflect the **entire lifetime** of that protection group's transfers to the vault — not just the last 7 days. This is a limitation of the Cohesity API at the per-group level.

### What this script does instead

To get **accurate, time-windowed** logical and physical transfer data per protection group, this script calls a second endpoint:

**`GET /public/protectionRuns`** — returns individual archival copy run records for the queried time window. Each run has:
- `stats.logicalBytesTransferred` — logical bytes transferred in **that specific run**
- `stats.physicalBytesTransferred` — physical bytes sent over the wire in that run

The script sums these across all successful (`kSuccess` / `kWarning`) FortKnox archival runs within your selected date range, per protection group per vault. This gives a true time-bounded total that reflects what was actually archived in the period you asked for.

`Storage Consumed` is sourced from `dataTransferToVaults` as before — it is a point-in-time snapshot and is not affected by the date range, so it matches the UI exactly.

---

### Column reference

| Column | Source | Description |
|--------|--------|-------------|
| Period Start | _(script)_ | Start of the report window (YYYY-MM-DD) |
| Period End | _(script)_ | End of the report window (YYYY-MM-DD) |
| Cluster | `dataTransferToVaults` | Cluster name as registered in Helios |
| Vault Name | `dataTransferToVaults → vaultName` | FortKnox vault target name |
| Vault Type | `dataTransferToVaults → vaultType` | Cloud platform: `kAzure`, `kAmazon`, etc. |
| Protection Group | `dataTransferToVaults → protectionJobName` | Protection group name |
| Logical Transferred (Bytes) | `protectionRuns → stats.logicalBytesTransferred` | Sum of logical bytes (pre-compression) across all archival runs in the queried window — **time-windowed, not cumulative** |
| Physical Transferred (Bytes) | `protectionRuns → stats.physicalBytesTransferred` | Sum of physical bytes (post-compression, over the wire) across all archival runs in the queried window — **time-windowed, not cumulative** |
| Storage Consumed (Bytes) | `dataTransferToVaults → storageConsumed` | Point-in-time snapshot of total storage retained in the vault for this protection group across all snapshots. Matches the Helios UI exactly. |
| Storage Consumed (TB) | _(derived)_ | Storage Consumed ÷ 1,000,000,000,000 |
| Storage Consumed (TiB) | _(derived)_ | Storage Consumed ÷ 1,099,511,627,776 — matches the binary TiB value shown in the Helios UI |

> **Zero transfer rows:** Groups with no archival run during the queried period will show `0` for Logical and Physical Transferred. This is expected — `Storage Consumed` remains populated because retained snapshots still exist in the vault from prior runs.

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
| 4.10 | 2026-04-07 | Fixed blank Logical/Physical Transferred: vault ID is at `target.archivalTarget.vaultId` in protectionRuns copyRun, not `target.vaultId` — vault match always failed leaving all rows at 0 |
| 4.9 | 2026-04-07 | Fixed Logical/Physical Transferred — switched from cumulative `dataTransferPerProtectionJob` fields to `GET /public/protectionRuns` per-run stats (time-windowed). Now matches UI values. |
| 4.8 | 2026-04-07 | Added Storage Consumed (TiB) column; fixed `--days` / default range to snap to calendar day start; fixed `end_of_day_msecs` to cover full last second |
| 4.7 | 2026-04-06 | Fixed x-axis dates not showing on trend charts |
| 4.6 | 2026-04-05 | Fixed axis labels; x-axis tick density auto-reduces |
| 4.5 | 2026-04-05 | One chart sheet per cluster |
| 4.3 | 2026-04-04 | Secure credential storage via OS keychain |
