# Changelog

All notable changes to this repository are documented here.
Format: `[YYYY-MM-DD] — description`

---

## [2026-04-04] — Fix: SSL verification disabled across all calls

### Fixed
- `reporting/cohesity_protection_report.py` — Set `verify=False` on all requests. Previously only the cluster discovery call had SSL disabled; the protection groups and runs calls still used `verify=True`, causing `SSLCertVerificationError` on corporate networks with SSL inspection proxies. Removed `verify_ssl` parameter entirely — it's always `False` in this environment.

---

## [2026-04-04] — fortknox_vault_report.py v2.2: Per-group vault storage

### Fixed
- `reporting/fortknox_vault_report.py` — `Vault Storage Consumed` now uses `cloudTotalPhysicalUsageBytes` from `GET /irisservices/api/v1/public/stats/consumers?consumerType=kProtectionRuns`, fetched once per unique cluster in the activity results and keyed by `(clusterName, groupName)`. Previously used the vault-level `scRetainedBytes` aggregate which was wildly inaccurate at the group level.

---

## [2026-04-04] — fortknox_vault_report.py v2.1: Fix storage stats + new vault columns

### Fixed
- `reporting/fortknox_vault_report.py` — Storage stats parsed from `resp['component']['data']` (was falling back to the `filters` echo list, giving 0 results and empty Storage Consumed column).

### Added
- New vault-level columns from Helios Reporting API: `Vault Target Type`, `Vault Storage Consumed (GB)`, `Vault Storage Growth (GB)`, `Vault Cumul Transferred (GB)`, `Vault Period Transferred (GB)`, `Vault Daily Growth Pct`.

---

## [2026-04-04] — fortknox_vault_report.py v2.0: Rewrite using Helios activity API

### Changed
- `reporting/fortknox_vault_report.py` — Full rewrite. Replaced per-cluster protection-group-runs approach with the Helios-native activity API (`POST /v2/mcm/data-protect/protection-group/activity` with `isRpaas=true`). One call returns all FortKnox runs across all clusters. Vault storage (`scRetainedBytes`) now sourced from Helios Reporting API (`POST /heliosreporting/api/v1/public/components/1600/preview` filtered to `managedBy=FortKnox`). Approach identified from ELF script (cohesityInv.ps1). Script is now Helios-only (`--apikey` required). Default date range is last 7 days.
- `reporting/fortknox_vault_report.md` — Updated with new API endpoints, field table, and CLI options.

---

## [2026-04-04] — Per-script READMEs

### Changed
- `reporting/README.md` — Slimmed to an index table linking to per-script docs.
- `reporting/cohesity_protection_report.md` — New dedicated README for the protection report: usage, CLI options, all 27 report fields, available-but-not-reported fields.
- `reporting/fortknox_vault_report.md` — New dedicated README for the FortKnox report: usage, CLI options, all 24 report fields, available-but-not-reported fields.

---

## [2026-04-04] — New script: fortknox_vault_report.py v1.0

### Added
- `reporting/fortknox_vault_report.py` — FortKnox (RPaaS) vault activity report. One row per FortKnox archival target per run per protection group. Filters on `targetType = kRpaas`. Reports: Vault Name, Run Type, SLA Violated, Vault Status, Run Start, Vault Start/End, Duration, Vault Expiry, Objects Succeeded/Failed/Canceled, Data Read (GB), Logical Size (GB), Logical Transferred (GB), Physical Transferred (GB), Vault Storage Consumed (GB). Supports `--vault` filter, `--days`/`--start`/`--end` date range, Helios and direct cluster auth. Cluster-local timestamps, GB-only size columns.
- `reporting/README.md` — Added `fortknox_vault_report.py` section: usage, CLI options, 24-column field table, available-but-not-reported fields table.

---

## [2026-04-04] — v3.2: Policy Name replaces Policy ID

### Changed
- `reporting/cohesity_protection_report.py` — `Policy ID` column replaced with `Policy Name`. Name is resolved via `GET /v2/data-protect/policies/{id}` with in-process caching so each unique policy is only fetched once per script run. Falls back to the raw ID if the lookup fails.
- `reporting/README.md` — Updated field table (Policy Name); moved Policy ID to the unreported fields table.

---

## [2026-04-04] — v3.1: GB-only size columns, updated README

### Changed
- `reporting/cohesity_protection_report.py` — Size values (Data Read, Data Written, Logical Size, Storage Consumed) now report as plain GB decimal numbers. The `(GB)` label is in the column header only, keeping cell values numeric for Excel pivot tables and sorting. Replaces the previous auto-scaling TB/GB/MB format.
- `reporting/README.md` — Full rewrite: added per-field descriptions table for all 25 columns, and a separate table listing available-but-not-reported API fields with their locations and notes.

---

## [2026-04-04] — v2.0: Historical run data, version numbering

### Added
- `reporting/cohesity_protection_report.py` — `__version__ = "2.0"`. Version history block in docstring. Version number shown in `--help`.
- `--days N` — pull last N days of runs (e.g. `--days 30`). Convenient shortcut.
- `--start YYYY-MM-DD` / `--end YYYY-MM-DD` — explicit date range. `--end` defaults to today if omitted. Each run in the range becomes its own CSV row.
- `Run #` column added in historical mode (sequential per protection group).
- `Logical Size Gb` column added (from `localSnapshotStats.logicalSizeBytes`).
- Column rename: `Last Run *` → `Run *` (accurate for both single and historical rows).

### Changed
- `get_last_run` replaced by `get_runs(start_usecs, end_usecs)` — returns a list. Default (no dates) returns last run as a single-element list. Date range returns all matching runs (capped at 1000 per group).
- `build_report` accepts `start_usecs`/`end_usecs` and emits one row per run.

---

## [2026-04-04] — Fix: Bytes read/written fields empty

### Fixed
- `reporting/cohesity_protection_report.py` — Bytes read/written (columns N, O) were always N/A. Stats are nested at `localBackupInfo.localSnapshotStats`, not `localBackupInfo.stats`. Fixed by changing `backup.get("stats", {})` to `backup.get("localSnapshotStats", {})`.

---

## [2026-04-04] — Fix: Correct Helios cluster routing header

### Fixed
- `reporting/cohesity_protection_report.py` — The Helios cluster routing header was wrong: `clusterIdentifier` does not exist. The correct header confirmed in `cohesity-api.ps1` (Brian Seltzer) is **`accessClusterId`**. This was causing Helios to proxy to no cluster, returning an empty result for every API call — hence "Found 0 active protection jobs".
- Reverted from v1 back to v2 API for both protection groups and runs, matching the `ClusterV2` pattern used in `cohesityInv.ps1`. Added `includeObjectDetails=true` to the runs call to get per-object success/failure/warning counts.
- Updated `extract_run_stats` to parse v2 response structure (`localBackupInfo`, plain-string status values).

---

## [2026-04-04] — Refactor: Switch to v1 API for job listing and run stats

### Changed
- `reporting/cohesity_protection_report.py` — Replaced v2 `data-protect/protection-groups` and `protection-groups/{id}/runs` endpoints with v1 `protectionJobs` and `protectionRuns` endpoints (same approach used in Brian Seltzer's reporting scripts).
  - v2 runs endpoint returned sparse/empty data through Helios; errors were silently swallowed returning "No runs yet" for all jobs
  - v1 `protectionRuns` returns complete stats in a single call: start/end times, bytes read from source, bytes written to Cohesity, and per-task success/failure/warning counts
  - Added explicit error logging on failed API calls (status code + truncated response body) instead of silent `return {}`
- Added `objects_warnings` column to CSV output

---

## [2026-04-04] — Fix: Helios cluster list endpoint

### Fixed
- `reporting/cohesity_protection_report.py` — Corrected Helios cluster discovery endpoint from `/mcm/clusters/info` (404) to `/mcm/clusters/connectionStatus`.

---

## [2026-04-04] — Fix: Helios SSL on corporate networks

### Fixed
- `reporting/cohesity_protection_report.py` — Changed `verify=True` to `verify=False` for Helios cluster discovery call. Corporate laptops with SSL inspection (proxy CA) caused a `ConnectionError` on the initial Helios connection even when the API key was valid.

---

## [2026-04-04]

### Added
- `reporting/cohesity_protection_report.py` — Initial script. Generates a CSV report of all protection group last-run status, timing, object counts, and data written/read. Supports Cohesity 7.1 and 7.3.

### Added — Auth
- Direct cluster auth: username/password → Bearer token via `/v2/users/sessions`
- Helios API key auth: `apiKey` header + `clusterIdentifier` routing via `helios.cohesity.com`. Auto-discovers all connected clusters or filters to a named cluster with `--cluster`.

### Added — Repo structure
- Organized repo into topic folders: `reporting/`, `protection/`, `recovery/`, `storage/`, `policies/`, `alerts/`, `infrastructure/`, `utils/`
- Added per-folder `README.md` with planned scripts for each area
- Added root `README.md` and this `CHANGELOG.md`
