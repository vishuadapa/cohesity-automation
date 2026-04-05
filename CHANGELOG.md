# Changelog

All notable changes to this repository are documented here.
Format: `[YYYY-MM-DD] — description`

---

## [2026-04-05] — Cohesity green branding across both reports

### Changed
- `reporting/fortknox_vault_report.py` v4.4 — Trend chart improvements: 16 pt bold chart title, legend repositioned to bottom (no overlap with chart area), x-axis labelled "Date" with `yyyy-mm-dd` tick format, y-axis labelled "Storage Consumed (TB)" with 4-decimal number format. Report header row color changed from dark blue to Cohesity green (`#00B388`).
- `reporting/cohesity_protection_report.py` v3.4 — Report header row color changed to Cohesity green (`#00B388`).

---

## [2026-04-04] — cohesity_protection_report.py v3.3: Secure credential storage + Excel output

### Changed
- `reporting/cohesity_protection_report.py` — `--apikey` is now optional. Helios API key is stored in the OS keychain (`keyring`) after the first run and retrieved automatically on subsequent runs. Direct cluster passwords are also stored in the keychain (keyed by `cluster:domain:username`). `--clear-credentials` removes the stored Helios key; `--clear-credentials --cluster <ip>` removes a cluster password. Output changed from CSV to Excel (`.xlsx`) with styled headers (bold white on dark blue), frozen header row, and auto-fit columns. Output filename is auto-generated as `<script_name>_YYYYMMDD_HHMMSS.xlsx` in the current working directory. Requires `pip install keyring openpyxl`.

---

## [2026-04-04] — fortknox_vault_report.py v4.3: Secure credential storage + auto output filename

### Changed
- `reporting/fortknox_vault_report.py` — `--apikey` is now optional. API key is stored in the OS keychain (`keyring`) after the first run and retrieved automatically on subsequent runs — never visible in the process list or written to disk in plaintext. `--clear-credentials` removes the stored key. Output filename is auto-generated as `<script_name>_YYYYMMDD_HHMMSS.xlsx` in the current working directory. Requires `pip install keyring`.

---

## [2026-04-04] — fortknox_vault_report.py v4.2: Storage Consumed (TB) column + TB trend chart

### Added
- `reporting/fortknox_vault_report.py` — `Storage Consumed (TB)` column added to every row (raw bytes ÷ 1,000,000,000,000, 4 decimal places). Trend chart Y axis now plots TB values instead of raw bytes, giving a human-readable scale.

---

## [2026-04-04] — fortknox_vault_report.py v4.1: Chart references Report sheet directly

### Changed
- `reporting/fortknox_vault_report.py` — Trend chart no longer duplicates data on the chart sheet. Report rows are sorted by (group, cluster, vault, date) so each series is a contiguous range that the chart references directly in the Report sheet. All-zero Storage Consumed series are excluded. Chart fixed to 18 × 28 cm to fit a single screen. "Trend Charts" sheet contains only the chart.

---

## [2026-04-04] — fortknox_vault_report.py v4.0: Single combined trend chart + column filters

### Changed
- `reporting/fortknox_vault_report.py` — Replaced per-group chart sheets with a single **"Trend Charts"** sheet. Contains a pivot table (Date × Protection Group) formatted as an Excel Table with ▼ column-filter dropdowns on every header for group/date navigation. Chart below shows all protection groups as series. Storage is summed across vaults per group per day. Native slicers require pivot tables (not supported by openpyxl); the column-filter dropdowns provide equivalent navigation.

---

## [2026-04-04] — fortknox_vault_report.py v3.9: Excel output + trend charts

### Changed
- `reporting/fortknox_vault_report.py` — Output changed from CSV to Excel (`.xlsx`) via `openpyxl`. Header row styled (bold white on dark blue), columns auto-fitted, header frozen. In `--mode trend`, one chart sheet is added per protection group showing `Storage Consumed (Bytes)` as a line chart over time. If a protection group archives to multiple vaults, each vault is a separate series on the same chart. Requires `pip install openpyxl`.

---

## [2026-04-04] — fortknox_vault_report.py v3.8: Zero-transfer caveat + primary field note

### Added
- `reporting/fortknox_vault_report.py` — Startup note explaining that `Storage Consumed` is the primary metric, and that zero `Logical/Physical Transferred` values are expected for groups with no archival run in the queried window.
- `reporting/fortknox_vault_report.md` — Same caveat and primary field note added to the Report Fields section.

---

## [2026-04-04] — fortknox_vault_report.py v3.7: Trend mode

### Added
- `reporting/fortknox_vault_report.py` — `--mode trend` flag. In trend mode the full date range is split into individual calendar days (using the auto-detected cluster timezone); each day is queried separately and produces its own set of rows. `Period Start` and `Period End` both show the specific date (YYYY-MM-DD) so rows from multiple runs can be stacked in Excel or a BI tool to plot a daily storage utilization trend per protection group. Vault IDs are fetched once per cluster and reused across all daily calls. Default behaviour (`--mode summary`) is unchanged.

---

## [2026-04-04] — fortknox_vault_report.py v3.6: Auto-detect cluster timezone

### Changed
- `reporting/fortknox_vault_report.py` — Timezone is now queried automatically from the cluster via `GET /irisservices/api/v1/public/cluster` before computing date boundaries. Removed the manual `--timezone` flag. When `--cluster` is specified the matching cluster is queried; otherwise the first available cluster is used. Startup banner shows the detected timezone and resolved start/end timestamps for verification.

---

## [2026-04-04] — fortknox_vault_report.py v3.5: Timezone-aware date boundaries

### Added
- `reporting/fortknox_vault_report.py` — `--timezone TZ` flag (default: `UTC`). When `--start`/`--end`/`--days` are used, start-of-day and end-of-day boundaries are now computed in the given timezone instead of UTC. The Helios UI uses the cluster's local time (e.g. `America/Chicago` for US Central / GMT-6) for its date range — setting `--timezone` to match eliminates the window mismatch. Startup banner now shows the resolved timestamps with timezone label for verification.

---

## [2026-04-04] — fortknox_vault_report.py v3.4: Exact timestamp flags + debug URL

### Added
- `reporting/fortknox_vault_report.py` — `--start-msecs` / `--end-msecs` flags accept exact millisecond timestamps pasted from the Chrome DevTools request URL, eliminating time boundary differences between script and UI. `--debug` now prints the exact request URL before the call so it can be compared directly against the Chrome network request.

---

## [2026-04-04] — fortknox_vault_report.py v3.3: Output raw bytes

### Changed
- `reporting/fortknox_vault_report.py` — Size columns now report raw bytes (no conversion). Column headers updated from `(TB)` to `(Bytes)`. API payload confirmed via Chrome DevTools to return values in bytes; all prior unit conversions produced values that did not match the UI.

---

## [2026-04-04] — fortknox_vault_report.py v3.2: Switch to TB units

### Changed
- `reporting/fortknox_vault_report.py` — Size columns changed from GB (÷ 1e9) to TB (÷ 1e12). Helios UI shows binary TiB; 1 TiB = 1.09951 TB so script TB values are ~9.95% higher than UI TiB figures — correct decimal equivalent. Column headers updated from `(GB)` to `(TB)`. Precision increased to 4 decimal places.

---

## [2026-04-04] — fortknox_vault_report.py v3.1: Correct GB conversion + period timestamp columns

### Changed
- `reporting/fortknox_vault_report.py` — GB conversion corrected from binary (÷ 1024³) to decimal SI (÷ 1,000,000,000). Added `Period Start` and `Period End` columns (YYYY-MM-DD) to every row so stacked CSV runs from periodic executions can be used to build a storage utilization trend in Excel or similar tools.

---

## [2026-04-04] — fortknox_vault_report.py v3.0: Rewrite using dataTransferToVaults only

### Changed
- `reporting/fortknox_vault_report.py` — Full rewrite. Removed Helios activity API and Reporting API component calls (data was inaccurate). Single source of truth is now `GET /irisservices/api/v1/public/reports/dataTransferToVaults` per cluster, the same report shown in the Helios single-cluster "Data Transferred to External Targets" UI. One row per protection group per vault. Columns: Cluster, Vault Name, Vault Type, Protection Group, Logical Transferred (GB), Physical Transferred (GB), Storage Consumed (GB).
- `reporting/fortknox_vault_report.md` — Updated for v3.0: new endpoint table, simplified field list, available-but-not-reported fields.

---

## [2026-04-04] — Fix: SSL verification disabled across all calls

### Fixed
- `reporting/cohesity_protection_report.py` — Set `verify=False` on all requests. Previously only the cluster discovery call had SSL disabled; the protection groups and runs calls still used `verify=True`, causing `SSLCertVerificationError` on corporate networks with SSL inspection proxies. Removed `verify_ssl` parameter entirely — it's always `False` in this environment.

---

## [2026-04-04] — fortknox_vault_report.py v2.4: Per-group retained storage via v1 report API

### Changed
- `reporting/fortknox_vault_report.py` — Vault Storage Consumed now sourced from `GET /irisservices/api/v1/public/reports/dataTransferToVaults` (Helios-routed per-cluster via `accessClusterId`), filtered by FortKnox vault IDs and the report date range. Vault IDs fetched from `GET /vaults?includeFortKnoxVault=true`. This is the same endpoint used by the Helios single-cluster "Data Transferred to External Targets" Protection Group tab.

---

## [2026-04-04] — fortknox_vault_report.py v2.3: Use component 1601 for retained storage

### Changed
- `reporting/fortknox_vault_report.py` — Vault Storage Consumed now sourced from Helios Reporting API component 1601 ("Data Transferred to External Targets"), filtered by `systemId` (clusterId:clusterIncarnationId) per cluster. This is the same data shown in the Helios UI report. Lookup keyed by (cluster, group, vault). `--debug` flag shows first component 1601 record to confirm field names.

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
