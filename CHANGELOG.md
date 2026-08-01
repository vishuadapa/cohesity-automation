# Changelog

All notable changes to this repository are documented here.
Format: `[YYYY-MM-DD] type(scope): description`

Commit types: `feat` (new feature), `fix` (bug fix), `refactor` (restructure), `docs` (docs only), `chore` (maintenance)

---

## [2026-08-01] docs(mcp): add v5 step-by-step testing guide (Word document)

### Added — `mcp/Cohesity_MCP_v5_Testing_Guide.docx`
- Step-by-step testing guide for `cohesity_mcp_v5.py` aimed at the average user: shared one-time setup (profiles, Helios, verification), then separate sections for Claude Desktop, Microsoft Copilot (GitHub Copilot in VS Code), and ChatGPT (HTTP transport + temporary ngrok tunnel, with security warnings), each with Windows and Mac instructions; 13-point test checklist with expected results and Pass/Fail column; troubleshooting table and post-test cleanup steps.

### Changed — `mcp/README.md`
- Linked the testing guide from the "Which Version Should I Use?" section.

---

## [2026-07-30] feat(mcp): cohesity_mcp_v5 v5.0 — multi-cluster profiles and Helios support

### Added — `mcp/cohesity_mcp_v5.py`
- **Named profiles**: store credentials for any number of clusters in the OS lockbox, namespaced per profile (`prod:api_key`, `dr:cluster`, …) under service `cohesity-mcp`, with a stored index (`_profiles`) and default (`_default`). The server selects its profile via the non-secret `COHESITY_PROFILE` env var → stored default → `default`, so one Claude Desktop entry per profile exposes several clusters side by side with zero secrets in the config.
- **Helios connection type**: a profile can target Helios (`helios.cohesity.com` by default, Helios API key required, SSL verification on by default). New MCP tools `list_clusters` (registered clusters with connection state, via `/public/mcm/clusters/connectionStatus` with `/v2/mcm/cluster-mgmt/info` fallback) and `select_cluster` (by name or id; refuses disconnected clusters). All other tools then route to the selected cluster via the `accessClusterId` passthrough header; cluster-scoped calls before selection return a clear instruction, and `get_cluster_health` degrades gracefully where the cluster-local `nexus` status endpoint is unavailable through passthrough.
- **Profile-aware CLI**: `setup` (add/update a profile, choose type, first profile auto-default), `list`, `use <name>`, `show [name]`, `test [name]` (Helios: prints every registered cluster and its connection state), `clear <name>` / `clear --all`.
- **Backward compatible**: a v4 lockbox (un-namespaced keys) is read automatically as profile `default` — no migration; v2-style env vars still override everything; `get_whoami` now also reports the active profile and connection route.

### Added — `mcp/cohesity_mcp_v5.md`
- Full why/what/how/where documentation: motivation for profiles and Helios, setup walkthrough, multi-entry Claude Desktop config, in-conversation Helios workflow (list → select → operate → switch), lockbox layout per profile, troubleshooting, and configuration-resolution reference.

### Changed — `mcp/README.md`
- "Which Version Should I Use?" now recommends v5 (multi-cluster + Helios), with v4 and v2 retained as alternatives; quick start updated to the profile workflow.

---

## [2026-07-30] feat(mcp): cohesity_mcp_v4 v4.0 — secure credential storage in OS lockbox

### Added — `mcp/cohesity_mcp_v4.py`
- **New v4 MCP server** based on v2 (all 14 tools, confirm pattern, RBAC/risk gating unchanged). Replaces plain-text env-var credentials in `claude_desktop_config.json` with secure storage in the OS credential lockbox via the cross-platform `keyring` library — macOS Keychain, Windows Credential Manager, Linux Secret Service. Stored under service name `cohesity-mcp`.
- **Built-in CLI**: `setup` (interactive one-time credential storage, secrets hidden while typing), `show` (config with masked secrets; warns when env vars override the lockbox), `test` (connect and verify auth with friendly error messages), `clear` (remove all stored credentials), `help`.
- **Config resolution order** per setting: environment variable → OS lockbox → default. Existing v2 env-var configs keep working unchanged.
- **MCP SDK 1.x/2.x compatibility**: imports `FastMCP` (SDK 1.x) or falls back to `MCPServer` (SDK 2.x rename); HTTP-transport port handling works on both.
- Passwords shown as `(set, hidden)` in `show`; API keys masked to last 4 characters.

### Added — `mcp/cohesity_mcp_v4.md`
- Full why/what/how/where documentation: the plain-text-key problem v4 solves, why the v3 TOTP approach was abandoned, 3-step setup, exact lockbox locations per OS and how to inspect/delete entries, honest security boundaries, troubleshooting (multiple Pythons, macOS keychain prompts, key rotation), and a full configuration-resolution reference table.

### Changed — `mcp/README.md`
- Retitled to version-neutral "Cohesity MCP — Setup Guide"; added "Which Version Should I Use?" section recommending v4 with a quick start, keeping the v2 env-var guide for headless/legacy setups.

---

## [2026-04-25] feat(protection_group_report): v5.0 — Helios bearer token auth mode

### Added — `reporting/protection_group_report.py` (v5.0)
- **New auth mode: Helios bearer token** (`--helios-user`). Instead of an API key,
  the script exchanges Helios username + password for a Bearer token via
  `POST https://helios.cohesity.com/irisservices/api/v1/public/mcm/createAccessToken`.
  The token is then used in the `Authorization: Bearer <token>` header for all
  subsequent Helios API calls (cluster list and all proxied cluster requests).
- **New CLI flags**: `--helios-user <email>`, `--helios-password <pass>` (optional),
  `--helios-domain <domain>` (default: `cohesity.com`).
- **Keychain integration**: Helios passwords stored/retrieved under keyring service
  `cohesity_helios_password`. `--clear-credentials --helios-user <email>` removes them.
- **Existing modes unchanged**: `--apikey` (API key) and `--cluster` (direct cluster)
  continue to work as before.

### Added — `utils/cohesity_auth.py` (v1.8)
- `get_helios_bearer_token(username, domain, cli_password, verify)` — shared helper.
- `make_helios_bearer_headers(token, cluster_id)` — shared helper.
- `clear_stored_credentials()` gains `helios_user` / `helios_domain` params to clear
  stored Helios bearer-token passwords from the keychain.

---

## [2026-04-21] feat(reporting): fortknox_vault_report v4.22 — remove TiB, group Bytes/TB columns, About Notes

### Changed — `reporting/fortknox_vault_report.py`
- **Removed** `Remote Storage Consumed (TiB)` column.
- **Column grouping**: all Bytes values now grouped together (Remote Storage Consumed, Data Read, Data Written, Logical Transferred, Physical Transferred), followed by the same five fields in TB — making side-by-side comparison straightforward.
- **About tab — Notes section** added between Report Fields and APIs Used:
  > "Activities: Data Read and Data Written show data sent to the **local cluster** (Backup activity). Activities: Logical Transferred and Physical Transferred show data sent to the **remote FortKnox vault** (Vault activity). The equivalent report on the Helios UI is the **Protection Activities Report**, filtered to Activity Type = Backup (for Data Read/Written) and Activity Type = Vault (for Logical/Physical Transferred), with the Cloud Vault filter applied."
- About tab field reference updated to match new column order.

---

## [2026-04-21] feat(reporting): fortknox_vault_report v4.21 — rename Storage Consumed → Remote; add Activities TB columns

### Changed — `reporting/fortknox_vault_report.py`
- **Renamed** `Storage Consumed (Bytes/TB/TiB)` → `Remote Storage Consumed (Bytes/TB/TiB)` to clarify these values reflect remote vault (FortKnox) storage, not local cluster storage.
- **Added TB companion columns** (÷ 1,000,000,000,000, 4 dp) for all four Activities fields, interleaved after each Bytes column: `Activities: Data Read (TB)`, `Activities: Data Written (TB)`, `Activities: Logical Transferred (TB)`, `Activities: Physical Transferred (TB)`.
- Source banner in the Report sheet shows `Computed` for all TB columns; About tab field reference updated.

---

## [2026-04-20] feat(reporting): fortknox_vault_report v4.20 — remove Logical/Physical Transferred columns and protectionRuns API

### Removed — `reporting/fortknox_vault_report.py`
- **Columns G and H** (`Logical Transferred (Bytes)` and `Physical Transferred (Bytes)`) and their data source `GET /public/protectionRuns`.
- `get_protection_runs_transfer()` function and all callers in summary and trend modes.
- `vault_id` from row dicts (was only used for protectionRuns matching).
- Two `_FIELD_REF` entries and the `protectionRuns` `_API_REF` entry from the About tab.

---

## [2026-04-20] fix(reporting): fortknox_vault_report v4.19 — paginate protection runs to fix 10-day data limit

### Fixed — `reporting/fortknox_vault_report.py`
- **Activities data cut off after ~10 days**: The v2 `/data-protect/protection-groups/{id}/runs` endpoint caps results to a small server-side page (typically 10 runs) regardless of `numRuns`. `get_activities_transfer()` now loops using `paginationCookie` until no further cookie is returned.
- **Protection runs data cut off after ~10 days**: `get_protection_runs_transfer()` (v1 API) now uses time-based pagination — fetches pages of 1,000 runs descending by time, advancing `endTimeUsecs` to one microsecond before the earliest run on the previous page until a partial page or the start boundary is reached.
- `--debug` now prints total run counts for both paths to confirm all pages are fetched.

---

## [2026-04-20] feat(reporting): fortknox_vault_report v4.18 — standardise all fills to #70AD47

### Changed — `reporting/fortknox_vault_report.py`
- All green bar fills in the About tab (title, section headers, table headers) and the Report tab (source-banner row, column-header row) changed to `#70AD47`.

---

## [2026-04-20] fix(reporting): fortknox_vault_report v4.17 — collect Backup and Vault activities independently by date

### Fixed — `reporting/fortknox_vault_report.py`
- **Data Read / Data Written date alignment**: Backup and Vault are separate run entries in the Protection Activities API — a backup-only run has `localBackupInfo` but no `archivalInfo`; a vault-only run has `archivalInfo` but no `localBackupInfo`. The prior implementation tried to combine them from a single run object, leaving Data Read/Written empty on days where only vault activity was recorded. Rewrote `get_activities_transfer()` to collect backup stats by backup start date and vault stats by vault start date independently, then merge results by `(group, date)`. On most days all four fields populate the same row; on days where only one activity type runs the other pair is 0.

---

## [2026-04-20] feat(reporting): fortknox_vault_report v4.15 — source-banner row on Report tab; About tab color unification

### Added — `reporting/fortknox_vault_report.py`
- **Source-banner row** (row 1 of the Report sheet): merged cells spanning each consecutive group of same-source columns, labelled with the originating API or `Computed`. Column headers moved to row 2; data starts at row 3; both rows frozen.
- About tab: all fills unified to `#70AD47`; section labels use the same green bar style as the title.

---

## [2026-04-20] fix(reporting): fortknox_vault_report v4.14 — Activities Logical/Physical Transferred always 0

### Fixed — `reporting/fortknox_vault_report.py`
- **Wrong field names**: v2 API uses `logicalBytesTransferred` / `physicalBytesTransferred`, not `logicalTransferredBytes` / `physicalTransferredBytes`.
- **Vault name match failures**: added `_FK_TARGET_TYPES` frozenset (`krpaas`, `kfortknox`, `kfort_knox`, `krpaasarchival` and plain equivalents) as a `targetType` fallback when vault name matching fails. `_xfer_bytes()` checks the `stats` sub-dict then the top-level of each `archivalTargetResult`.
- `--debug` now prints raw `archivalTargetResult` keys and all `(targetName, targetType)` pairs observed for easy diagnosis.

---

## [2026-04-20] feat(reporting): fortknox_vault_report v4.13 — Activities columns from Protection Activities report

### Added — `reporting/fortknox_vault_report.py`
- Four new columns sourced from the Protection Activities report, filtered to FortKnox cloud vault archival runs:
  - `Activities: Data Read (Bytes)` — `localSnapshotStats.bytesRead` from Backup activity
  - `Activities: Data Written (Bytes)` — `localSnapshotStats.bytesWritten` from Backup activity
  - `Activities: Logical Transferred (Bytes)` — `archivalTargetResults[].stats.logicalBytesTransferred`
  - `Activities: Physical Transferred (Bytes)` — `archivalTargetResults[].stats.physicalBytesTransferred`
- Fetched via `GET /v2/data-protect/protection-groups/{id}/runs`; matched to FortKnox vaults by name, with `targetType` fallback. Summary mode aggregates all days per group; trend mode aligns per day.
- `get_fortknox_vaults()` now returns `(ids, names)` tuple; `get_protection_group_map()` added.
- Two new API entries added to the About sheet API reference table.

---

## [2026-04-20] feat(reporting): fortknox_vault_report v4.12 — About tab with field reference and API table

### Added — `reporting/fortknox_vault_report.py`
- **About tab** (first sheet): command run, report parameters (version, mode, date range, cluster/vault filters, timezone), full field reference table (column name, description, source field path, API endpoint), and APIs Used table.
- Output filename now includes the script version: `fortknox_vault_report_v4.12_YYYYMMDD_HHMMSS.xlsx`.

---

## [2026-04-20] feat(reporting): fortknox_vault_report v4.11 — TLS hardening; formula injection protection

### Added — `reporting/fortknox_vault_report.py`
- TLS verification enabled by default. `--ca-bundle PATH` accepts a corporate proxy CA certificate. `--insecure` disables verification with a startup warning.
- `_safe_cell()` — prefixes any cell value starting with `=`, `+`, `-`, `@`, tab, or CR with a single quote to prevent Excel formula injection.

---

## [2026-04-17] fix(utils): cohesity_auth.py security hardening (v1.7)

### Security — `utils/cohesity_auth.py`

- **TLS certificate validation**: Removed all hardcoded `verify=False` from every HTTP call site. Added `verify=` parameter threaded from `main()` so `get_auth_token()` and `get_helios_clusters()` inherit the caller's TLS setting.
- **Default is now `verify=True`**: TLS certificates are validated against the system CA bundle by default.
- **`--ca-bundle PATH`**: path to a CA bundle `.pem` for self-signed or corporate proxy certs. Accepted by all scripts that use `cohesity_auth`.
- **`--insecure`**: explicit opt-in to disable validation (prints a startup warning; suppresses urllib3 warning only when this flag is set).
- **Credential scrub from error messages**: Raw API response bodies removed from auth error output — `r.text[:N]` fragments were printed on authentication failure, leaking token/credential hints into terminal history and CI logs. Error output now shows only the HTTP status code.

---

## [2026-04-10] fix(utils): cohesity_auth.py auth improvements (v1.5–1.6) + formatters.py column-width fix

### Fixed — `utils/cohesity_auth.py`

- **v1 accessTokens endpoint for direct cluster auth** (v1.5): `get_auth_token()` now tries the v1 endpoint first (`POST /irisservices/api/v1/public/accessTokens`) which has widest compatibility across Cohesity versions and account types. Falls back to v2 (`POST /v2/users/sessions`) if v1 fails. The v2 endpoint returned `400 KValidationError Access denied` on some clusters.
- **v1 auth failure now visible** (v1.6): Previously silently caught the v1 `HTTPError` and fell through to v2 with no output. Now prints `[!] v1 endpoint failed (HTTP NNN) — trying v2 ...` so the user can see that both endpoints were tried.
- **Improved final error message**: When both v1 and v2 fail, the error output now shows both responses and three actionable troubleshooting hints: `--clear-credentials` to wipe a stale stored password (most common cause), `--domain <AD_DOMAIN>` for Active Directory accounts, and a note to check account status in Cohesity Access Management if locked/disabled.

### Fixed — `utils/formatters.py`

- **`auto_fit_columns()` multiline fix**: Multiline cell values (newline-separated) are now measured by the longest line instead of the full concatenated string — prevents columns from being over-widened by wrapped text. Default `min_width` raised from 10 → 12; `max_width` lowered from 60 → 45.

---

## [2026-04-08] feat: add version number to auto-generated output filenames across all scripts

### Changed
All 13 scripts that auto-generate output filenames now embed the script version:

```
fortknox_vault_report_v4.10_20260408_1430.xlsx
protection_group_report_v4.5_20260408_1430.xlsx
alert_to_csv_v1.0_20260408_143022.csv
```

Scripts affected: `alert_summary_report.py`, `alert_to_csv.py`, `cluster_info_report.py`, `node_status.py`, `upgrade_readiness.py`, `list_policies.py`, `policy_compliance_report.py`, `list_snapshots.py`, `fortknox_vault_report.py`, `protection_group_report.py`, `capacity_report.py`, `list_views.py`, `storage_domain_report.py`. Scripts that accept an explicit `--output` path are unaffected.

---

## [2026-04-08] style: update all green to #70AD47 (RGB 112, 173, 71) across every script output

### Changed
Replaced `#00B388` with `#70AD47` in every Excel header fill, chart line/marker color, and cell background fill across all scripts in this repository. Light green cell fills (`#C6EFCE`) updated to `#E2EFDA` (lighter derivative of `#70AD47`).

---

## [2026-04-07] fix(reporting): fortknox_vault_report v4.10 — fix blank Logical/Physical Transferred

### Fixed — `reporting/fortknox_vault_report.py`
- Vault ID in `protectionRuns` copy-run is nested at `target.archivalTarget.vaultId`, not `target.vaultId`. The ID match always failed, leaving all Logical/Physical Transferred rows at 0.

---

## [2026-04-07] fix(reporting): fortknox_vault_report v4.9 — time-windowed Logical/Physical Transferred via protectionRuns

### Fixed — `reporting/fortknox_vault_report.py`
- `numLogicalBytesTransferred` / `numPhysicalBytesTransferred` in `dataTransferPerProtectionJob` are cumulative lifetime totals — not scoped to the query window. Replaced with `GET /public/protectionRuns` per-run archival copy stats (`logicalBytesTransferred` / `physicalBytesTransferred`), which are inherently time-windowed. Only `kSuccess`/`kWarning` archival runs targeting FortKnox vault IDs are counted. In trend mode, runs are bucketed by calendar day so each row reflects only that day's transfers.

---

## [2026-04-07] feat(reporting): fortknox_vault_report v4.8 — Storage Consumed (TiB); calendar-day date boundaries

### Added — `reporting/fortknox_vault_report.py`
- `Storage Consumed (TiB)` column (÷ 1,099,511,627,776) alongside the existing TB column — matches the unit displayed in the Helios UI.
- `--days` / default date range now snaps to `00:00:00` of the start day (calendar-day boundary) instead of a rolling window. End of day computed as `23:59:59.999` to include the full final second.

---

## [2026-04-07] fix(infrastructure): v1 API field names — node health, disk, and capacity checks

### Changed
- `infrastructure/cluster_info_report.py` — Switched from v2 to v1 API endpoints. Fixed: `Healthy Nodes`, `NTP Servers`, and `Helios Connected` columns now populate correctly from v1 response structure.
- `infrastructure/node_status.py` — Complete rewrite to use v1 API field names. Now correctly parses node IP, disk status, fault counts, and capacity from v1 response schema.
- `infrastructure/upgrade_readiness.py` — Fixed node health check, disk health check, and disk capacity margin detection. Added `--debug` flag to inspect raw cluster API response for troubleshooting. Fixed active runs detection and disk capacity fallback logic.

### Note
All infrastructure scripts now use Cohesity v1 API endpoints instead of v2. v1 is more reliable for cluster infrastructure data via Helios.

---

## [2026-04-07] fix(alerts): switch to v1 public alerts API endpoint

### Changed
- `alerts/alert_summary_report.py` — Switched from v2 to v1 public alerts API endpoint (`GET /irisservices/api/v1/public/alerts`). More reliable alert data retrieval via Helios.
- `alerts/resolve_alerts.py` — Switched from v2 to v1 public alerts API endpoint.
- `alerts/alert_to_csv.py` — Switched from v2 to v1 public alerts API endpoint.

---

## [2026-04-07] fix(storage): switch to v1 API endpoints

### Changed
- `storage/capacity_report.py` — Switched from v2 to v1 API endpoints for cluster capacity and stat reporting.
- `storage/list_views.py` — Switched from v2 to v1 API endpoints for view enumeration.
- `storage/storage_domain_report.py` — Switched from v2 to v1 API endpoints for domain utilization.

---

## [2026-04-06] fix(reporting): fortknox_vault_report v4.7 — fix x-axis dates not displaying on trend charts

### Fixed — `reporting/fortknox_vault_report.py`
- `openpyxl.set_categories()` emits `<c:numRef>` in chart XML; Excel treats string date values as invalid numbers and renders nothing. Fixed by assigning `DataSource(strRef=StrRef(f=...))` directly to each series' `.cat`, forcing `<c:strRef>` so Excel renders the date strings as text labels on the x-axis.

---

## [2026-04-06] feat: initial scripts for all planned folders (alerts, infrastructure, policies, protection, recovery, storage, utils)

### Added
- `utils/cohesity_auth.py` v1.0 — Shared auth module: Helios API key + direct cluster Bearer token,
  OS keychain credential storage, `get_helios_clusters`. Extracted from existing reporting scripts
  for reuse across new scripts.
- `utils/formatters.py` v1.0 — Shared formatting helpers: `usecs_to_datetime`, `bytes_to_gb`,
  `bytes_to_tb`, `format_duration`, `style_worksheet` (Cohesity green headers), `auto_fit_columns`.
- `alerts/alert_summary_report.py` v1.0 — Active alerts by severity and category across all clusters.
  Dual-sheet Excel output (Summary counts + Alerts detail). Supports `--days` window and `--cluster` filter.
- `alerts/resolve_alerts.py` v1.0 — Bulk resolve open alerts matching `--severity`, `--category`,
  or `--code` filter. Previews matches and confirms before resolving. `--yes` skips prompt.
- `alerts/alert_to_csv.py` v1.0 — Export alert history for a date range to CSV. Supports
  `--days`, `--start`/`--end`, severity and category filters.
- `infrastructure/cluster_info_report.py` v1.0 — Software version, node count, cluster ID,
  DNS/NTP/timezone config across all Helios clusters. Excel output.
- `infrastructure/node_status.py` v1.0 — Node health, disk status, fault counts, and capacity.
  Dual-sheet Excel (Nodes + Disks). `--unhealthy-only` flag for issue-focused view.
- `infrastructure/upgrade_readiness.py` v1.0 — Pre-upgrade readiness check: node health, disk health,
  active runs in progress, software version vs minimum, and disk capacity margin. PASS/WARN/FAIL per
  check with console table and color-coded Excel output.
- `policies/list_policies.py` v1.0 — List all policies with retention, replication targets, archival
  targets, and schedule settings. Excel output.
- `policies/policy_compliance_report.py` v1.0 — Flag protection groups whose policy doesn't meet
  configurable SLA targets (`--min-retention`, `--require-replication`, `--require-archival`,
  `--rpo-hours`). Color-coded Excel (green/yellow/red).
- `policies/clone_policy.py` v1.0 — Clone a policy by name or ID to a new name. Strips server-managed
  fields before POST. `--dry-run` mode.
- `protection/run_protection_group.py` v1.0 — Trigger on-demand backup run (kRegular/kFull/kLog)
  by group name. `--wait` flag polls for completion with configurable timeout.
- `protection/pause_resume_groups.py` v1.0 — Bulk pause or resume groups by name pattern.
  Confirmation prompt (bypass with `--yes`). `--resume-after N` auto-resumes after N minutes.
- `protection/clone_protection_group.py` v1.0 — Clone a protection group's full config to a new
  group name. Created in paused state. `--no-objects` clears object list; `--dry-run` mode.
- `recovery/list_snapshots.py` v1.0 — List available snapshots for any protected object with
  timestamp, run type, status, expiry, and local/replication/archival flags.
- `recovery/restore_vm.py` v1.0 — Restore a VMware VM to original location with optional suffix rename,
  snapshot selection, and `--power-on` flag. `--dry-run` mode.
- `recovery/recover_files.py` v1.0 — File/folder level recovery from any protected physical or NAS
  source. Multiple paths per run, optional alternate destination, overwrite flag, `--dry-run` mode.
- `storage/capacity_report.py` v1.0 — Cluster capacity summary: usable/used/free in TB, data
  reduction ratio, dedup ratio, compression ratio, and savings. Excel output.
- `storage/list_views.py` v1.0 — Enumerate all NAS views with quota, logical usage, protocol
  (SMB/NFS/S3), share paths, and creation date. Optional protocol filter. Excel output.
- `storage/storage_domain_report.py` v1.0 — Storage domain utilization: physical/logical capacity,
  data reduction ratio, dedup and compression config, encryption, cloud tiering, view count. Excel output.

### Changed
- All folder `README.md` files updated: planned scripts moved to the Scripts table with descriptions,
  auth pattern, and per-script usage examples.
- Root `README.md` updated with full quick-reference command list for all new scripts.

---

## [2026-04-06] fix(reporting): protection_group_report v4.4 — revert Storage Consumed to stats/consumers endpoint

### Changed
- `reporting/protection_group_report.py` — Reverted trend mode Storage Consumed to `GET /stats/consumers` (same source as summary mode). Previous snapshot-summing approach (`sum(bytesWritten)` across active snapshots per day) grossly exceeded actual cluster physical capacity because `bytesWritten` is measured before cross-snapshot deduplication. Column renamed to **"Storage Consumed (As of End Date, Bytes/TB)"** to make clear it is a current-state value stamped identically on all date rows for a group. Startup caveat message added. README updated with limitation note.

---

## [2026-04-05] feat(reporting): protection_group_report v4.3 — full column parity in trend mode

### Added
- `reporting/protection_group_report.py` — `--mode trend` now includes all summary mode columns alongside the daily storage data. Run detail columns (Run Type, SLA Violated, Run Status, Run Start/End, Duration, Snapshot Expiry, Object counts, Data Read/Written/Logical Size, Replication and Archive status/targets/expiry) are aggregated across all runs that started on each day: counts and sizes are summed; status, type, targets, and expiry reflect the last run; SLA is "Yes" if any run violated. A new `Runs On Day` column shows how many runs occurred (`0` = no backup ran that day, storage consumed is still shown).

---

## [2026-04-05] feat(reporting): protection_group_report v4.2 — add Policy Name, Is Active, Is Paused in trend mode

### Added
- `reporting/protection_group_report.py` — `--mode trend` output now includes **Policy Name**, **Is Active**, and **Is Paused** alongside the daily storage data. These are group-level attributes (same across all date rows for a group) that provide the context needed to filter and interpret the trend chart.

---

## [2026-04-05] fix(reporting): protection_group_report v4.1 — replace timeSeriesStats (400 error) with runs-based storage estimate

### Changed
- `reporting/protection_group_report.py` — Replaced `timeSeriesStats` approach (returned 400 — `storageConsumedBytes` requires an unknown `schemaName` when routed via Helios) with a runs-based estimate. For each protection group, queries all runs going back 1 year, then for each calendar day sums `bytesWritten` across every snapshot active on that day (`startTime ≤ day AND expiryTime > day`). Values are now genuinely day-varying and show real storage growth/expiration trends. `bytesWritten` is post-dedup/compression; may slightly overestimate vs `stats/consumers` due to shared dedup blocks between runs — trend direction is accurate.

---

## [2026-04-05] feat(reporting): protection_group_report v4.0 — --mode trend with daily storage history and chart per cluster

### Added
- `reporting/protection_group_report.py` — `--mode trend`: queries daily storage consumed per protection group via `GET /irisservices/api/v1/public/statistics/timeSeriesStats` (metric: `storageConsumedBytes`, `rollupIntervalSecs=86400`, `rollupFunction=Last`). Produces one row per group per day and adds a trend chart sheet per cluster — identical format to `fortknox_vault_report.py` (bottom legend, labeled axes, auto tick density). Default date range in trend mode is last 30 days. `--debug` prints raw `timeSeriesStats` response for metric name verification.

### Changed
- `Storage Consumed (GB)` renamed to `Storage Consumed (Current, GB)` in summary/historical mode to make explicit that `GET /stats/consumers` is always a point-in-time snapshot with no date parameter. Trend mode uses the historical time-series endpoint instead.

---

## [2026-04-05] refactor(reporting): rename cohesity_protection_report → protection_group_report

### Changed
- `reporting/cohesity_protection_report.py` renamed to `reporting/protection_group_report.py`
- `reporting/cohesity_protection_report.md` renamed to `reporting/protection_group_report.md`
- All internal references, docstrings, usage examples, README index, and CHANGELOG entries updated.

---

## [2026-04-05] fix(reporting): fortknox_vault_report v4.6 — fix axis labels; auto-reduce x-axis tick density

### Fixed
- `reporting/fortknox_vault_report.py` — X-axis date labels were invisible because `numFmt = "yyyy-mm-dd"` was set on a category axis whose values are strings — Excel tried to parse them as date serials and rendered nothing. Removed `numFmt` from the category axis; strings now display as-is. Added `delete = False` on both axes to guarantee labels always render.

### Added
- X-axis tick label density auto-adjusts to prevent overlapping dates: ≤14 days shows every date, 15–60 days shows every 2nd date, >60 days shows every 7th date (weekly).

---

## [2026-04-05] feat(reporting): fortknox_vault_report v4.5 — one chart sheet per cluster

### Changed
- `reporting/fortknox_vault_report.py` — `--mode trend` now creates one chart sheet per cluster (sheet named after the cluster) instead of a single combined sheet. Each chart shows all protection groups for that cluster as separate series, with the same styling (16 pt bold title, bottom legend, labeled axes).

---

## [2026-04-05] feat(reporting): Cohesity green header branding across fortknox and protection_group reports

### Changed
- `reporting/fortknox_vault_report.py` v4.4 — Trend chart improvements: 16 pt bold chart title, legend repositioned to bottom (no overlap with chart area), x-axis labeled "Date" with `yyyy-mm-dd` tick format, y-axis labeled "Storage Consumed (TB)" with 4-decimal number format. Report header row color changed from dark blue to Cohesity green (`#00B388`).
- `reporting/protection_group_report.py` v3.4 — Report header row color changed to Cohesity green (`#00B388`).

---

## [2026-04-04] feat(reporting): protection_group_report v3.3 — OS keychain credential storage; Excel output

### Changed
- `reporting/protection_group_report.py` — `--apikey` is now optional. Helios API key is stored in the OS keychain (`keyring`) after the first run and retrieved automatically on subsequent runs. Direct cluster passwords are also stored in the keychain (keyed by `cluster:domain:username`). `--clear-credentials` removes the stored Helios key; `--clear-credentials --cluster <ip>` removes a cluster password. Output changed from CSV to Excel (`.xlsx`) with styled headers (bold white on dark blue), frozen header row, and auto-fit columns. Output filename is auto-generated as `<script_name>_YYYYMMDD_HHMMSS.xlsx` in the current working directory. Requires `pip install keyring openpyxl`.

---

## [2026-04-04] feat(reporting): fortknox_vault_report v4.3 — OS keychain credential storage; auto output filename

### Changed
- `reporting/fortknox_vault_report.py` — `--apikey` is now optional. API key is stored in the OS keychain (`keyring`) after the first run and retrieved automatically on subsequent runs — never visible in the process list or written to disk in plaintext. `--clear-credentials` removes the stored key. Output filename is auto-generated as `<script_name>_YYYYMMDD_HHMMSS.xlsx` in the current working directory. Requires `pip install keyring`.

---

## [2026-04-04] feat(reporting): fortknox_vault_report v4.2 — add Storage Consumed (TB) column; TB trend chart

### Added
- `reporting/fortknox_vault_report.py` — `Storage Consumed (TB)` column added to every row (raw bytes ÷ 1,000,000,000,000, 4 decimal places). Trend chart Y axis now plots TB values instead of raw bytes, giving a human-readable scale.

---

## [2026-04-04] refactor(reporting): fortknox_vault_report v4.1 — chart references Report sheet directly; no data duplication

### Changed
- `reporting/fortknox_vault_report.py` — Trend chart no longer duplicates data on the chart sheet. Report rows are sorted by (group, cluster, vault, date) so each series is a contiguous range that the chart references directly in the Report sheet. All-zero Storage Consumed series are excluded. Chart fixed to 18 × 28 cm to fit a single screen. "Trend Charts" sheet contains only the chart.

---

## [2026-04-04] feat(reporting): fortknox_vault_report v4.0 — single combined trend chart with column filter dropdowns

### Changed
- `reporting/fortknox_vault_report.py` — Replaced per-group chart sheets with a single **"Trend Charts"** sheet. Contains a pivot table (Date × Protection Group) formatted as an Excel Table with ▼ column-filter dropdowns on every header for group/date navigation. Chart below shows all protection groups as series. Storage is summed across vaults per group per day. Native slicers require pivot tables (not supported by openpyxl); the column-filter dropdowns provide equivalent navigation.

---

## [2026-04-04] feat(reporting): fortknox_vault_report v3.9 — switch to Excel output; add per-group trend charts

### Changed
- `reporting/fortknox_vault_report.py` — Output changed from CSV to Excel (`.xlsx`) via `openpyxl`. Header row styled (bold white on dark blue), columns auto-fitted, header frozen. In `--mode trend`, one chart sheet is added per protection group showing `Storage Consumed (Bytes)` as a line chart over time. If a protection group archives to multiple vaults, each vault is a separate series on the same chart. Requires `pip install openpyxl`.

---

## [2026-04-04] docs(reporting): fortknox_vault_report v3.8 — add zero-transfer caveat and primary field note

### Added
- `reporting/fortknox_vault_report.py` — Startup note explaining that `Storage Consumed` is the primary metric, and that zero `Logical/Physical Transferred` values are expected for groups with no archival run in the queried window.
- `reporting/fortknox_vault_report.md` — Same caveat and primary field note added to the Report Fields section.

---

## [2026-04-04] feat(reporting): fortknox_vault_report v3.7 — --mode trend with per-day rows

### Added
- `reporting/fortknox_vault_report.py` — `--mode trend` flag. In trend mode the full date range is split into individual calendar days (using the auto-detected cluster timezone); each day is queried separately and produces its own set of rows. `Period Start` and `Period End` both show the specific date (YYYY-MM-DD) so rows from multiple runs can be stacked in Excel or a BI tool to plot a daily storage utilization trend per protection group. Vault IDs are fetched once per cluster and reused across all daily calls. Default behavior (`--mode summary`) is unchanged.

---

## [2026-04-04] feat(reporting): fortknox_vault_report v3.6 — auto-detect cluster timezone; remove manual --timezone flag

### Changed
- `reporting/fortknox_vault_report.py` — Timezone is now queried automatically from the cluster via `GET /irisservices/api/v1/public/cluster` before computing date boundaries. Removed the manual `--timezone` flag. When `--cluster` is specified the matching cluster is queried; otherwise the first available cluster is used. Startup banner shows the detected timezone and resolved start/end timestamps for verification.

---

## [2026-04-04] feat(reporting): fortknox_vault_report v3.5 — timezone-aware date boundaries via --timezone flag

### Added
- `reporting/fortknox_vault_report.py` — `--timezone TZ` flag (default: `UTC`). When `--start`/`--end`/`--days` are used, start-of-day and end-of-day boundaries are now computed in the given timezone instead of UTC. The Helios UI uses the cluster's local time (e.g. `America/Chicago` for US Central / GMT-6) for its date range — setting `--timezone` to match eliminates the window mismatch. Startup banner now shows the resolved timestamps with timezone label for verification.

---

## [2026-04-04] feat(reporting): fortknox_vault_report v3.4 — --start-msecs/--end-msecs flags; --debug prints request URL

### Added
- `reporting/fortknox_vault_report.py` — `--start-msecs` / `--end-msecs` flags accept exact millisecond timestamps pasted from the Chrome DevTools request URL, eliminating time boundary differences between script and UI. `--debug` now prints the exact request URL before the call so it can be compared directly against the Chrome network request.

---

## [2026-04-04] fix(reporting): fortknox_vault_report v3.3 — output raw bytes; remove unit conversions that didn't match UI

### Changed
- `reporting/fortknox_vault_report.py` — Size columns now report raw bytes (no conversion). Column headers updated from `(TB)` to `(Bytes)`. API payload confirmed via Chrome DevTools to return values in bytes; all prior unit conversions produced values that did not match the UI.

---

## [2026-04-04] fix(reporting): fortknox_vault_report v3.2 — switch size columns from GB to TB

### Changed
- `reporting/fortknox_vault_report.py` — Size columns changed from GB (÷ 1e9) to TB (÷ 1e12). Helios UI shows binary TiB; 1 TiB = 1.09951 TB so script TB values are ~9.95% higher than UI TiB figures — correct decimal equivalent. Column headers updated from `(GB)` to `(TB)`. Precision increased to 4 decimal places.

---

## [2026-04-04] fix(reporting): fortknox_vault_report v3.1 — correct GB conversion to decimal SI; add Period Start/End columns

### Changed
- `reporting/fortknox_vault_report.py` — GB conversion corrected from binary (÷ 1024³) to decimal SI (÷ 1,000,000,000). Added `Period Start` and `Period End` columns (YYYY-MM-DD) to every row so stacked CSV runs from periodic executions can be used to build a storage utilization trend in Excel or similar tools.

---

## [2026-04-04] refactor(reporting): fortknox_vault_report v3.0 — rewrite using dataTransferToVaults as single source of truth

### Changed
- `reporting/fortknox_vault_report.py` — Full rewrite. Removed Helios activity API and Reporting API component calls (data was inaccurate). Single source of truth is now `GET /irisservices/api/v1/public/reports/dataTransferToVaults` per cluster, the same report shown in the Helios single-cluster "Data Transferred to External Targets" UI. One row per protection group per vault. Columns: Cluster, Vault Name, Vault Type, Protection Group, Logical Transferred (GB), Physical Transferred (GB), Storage Consumed (GB).
- `reporting/fortknox_vault_report.md` — Updated for v3.0: new endpoint table, simplified field list, available-but-not-reported fields.

---

## [2026-04-04] fix(reporting): SSL verification disabled across all calls — was causing ConnectionError on corporate networks

### Fixed
- `reporting/protection_group_report.py` — Set `verify=False` on all requests. Previously only the cluster discovery call had SSL disabled; the protection groups and runs calls still used `verify=True`, causing `SSLCertVerificationError` on corporate networks with SSL inspection proxies. Removed `verify_ssl` parameter entirely — it's always `False` in this environment.

---

## [2026-04-04] fix(reporting): fortknox_vault_report v2.4 — per-group retained storage via v1 dataTransferToVaults report API

### Changed
- `reporting/fortknox_vault_report.py` — Vault Storage Consumed now sourced from `GET /irisservices/api/v1/public/reports/dataTransferToVaults` (Helios-routed per-cluster via `accessClusterId`), filtered by FortKnox vault IDs and the report date range. Vault IDs fetched from `GET /vaults?includeFortKnoxVault=true`. This is the same endpoint used by the Helios single-cluster "Data Transferred to External Targets" Protection Group tab.

---

## [2026-04-04] fix(reporting): fortknox_vault_report v2.3 — use Reporting API component 1601 for retained storage

### Changed
- `reporting/fortknox_vault_report.py` — Vault Storage Consumed now sourced from Helios Reporting API component 1601 ("Data Transferred to External Targets"), filtered by `systemId` (clusterId:clusterIncarnationId) per cluster. This is the same data shown in the Helios UI report. Lookup keyed by (cluster, group, vault). `--debug` flag shows first component 1601 record to confirm field names.

---

## [2026-04-04] fix(reporting): fortknox_vault_report v2.2 — per-group vault storage from cloudTotalPhysicalUsageBytes

### Fixed
- `reporting/fortknox_vault_report.py` — `Vault Storage Consumed` now uses `cloudTotalPhysicalUsageBytes` from `GET /irisservices/api/v1/public/stats/consumers?consumerType=kProtectionRuns`, fetched once per unique cluster in the activity results and keyed by `(clusterName, groupName)`. Previously used the vault-level `scRetainedBytes` aggregate which was wildly inaccurate at the group level.

---

## [2026-04-04] fix(reporting): fortknox_vault_report v2.1 — fix storage stats parsing; add vault-level columns

### Fixed
- `reporting/fortknox_vault_report.py` — Storage stats parsed from `resp['component']['data']` (was falling back to the `filters` echo list, giving 0 results and empty Storage Consumed column).

### Added
- New vault-level columns from Helios Reporting API: `Vault Target Type`, `Vault Storage Consumed (GB)`, `Vault Storage Growth (GB)`, `Vault Cumul Transferred (GB)`, `Vault Period Transferred (GB)`, `Vault Daily Growth Pct`.

---

## [2026-04-04] refactor(reporting): fortknox_vault_report v2.0 — rewrite using Helios activity API

### Changed
- `reporting/fortknox_vault_report.py` — Full rewrite. Replaced per-cluster protection-group-runs approach with the Helios-native activity API (`POST /v2/mcm/data-protect/protection-group/activity` with `isRpaas=true`). One call returns all FortKnox runs across all clusters. Vault storage (`scRetainedBytes`) now sourced from Helios Reporting API (`POST /heliosreporting/api/v1/public/components/1600/preview` filtered to `managedBy=FortKnox`). Approach identified from ELF script (cohesityInv.ps1). Script is now Helios-only (`--apikey` required). Default date range is last 7 days.
- `reporting/fortknox_vault_report.md` — Updated with new API endpoints, field table, and CLI options.

---

## [2026-04-04] docs(reporting): add per-script READMEs for fortknox and protection_group reports

### Changed
- `reporting/README.md` — Slimmed to an index table linking to per-script docs.
- `reporting/protection_group_report.md` — New dedicated README for the protection report: usage, CLI options, all 27 report fields, available-but-not-reported fields.
- `reporting/fortknox_vault_report.md` — New dedicated README for the FortKnox report: usage, CLI options, all 24 report fields, available-but-not-reported fields.

---

## [2026-04-04] feat(reporting): fortknox_vault_report v1.0 — initial FortKnox vault activity report

### Added
- `reporting/fortknox_vault_report.py` — FortKnox (RPaaS) vault activity report. One row per FortKnox archival target per run per protection group. Filters on `targetType = kRpaas`. Reports: Vault Name, Run Type, SLA Violated, Vault Status, Run Start, Vault Start/End, Duration, Vault Expiry, Objects Succeeded/Failed/Canceled, Data Read (GB), Logical Size (GB), Logical Transferred (GB), Physical Transferred (GB), Vault Storage Consumed (GB). Supports `--vault` filter, `--days`/`--start`/`--end` date range, Helios and direct cluster auth. Cluster-local timestamps, GB-only size columns.
- `reporting/README.md` — Added `fortknox_vault_report.py` section: usage, CLI options, 24-column field table, available-but-not-reported fields table.

---

## [2026-04-04] feat(reporting): protection_group_report v3.2 — Policy Name column replaces Policy ID

### Changed
- `reporting/protection_group_report.py` — `Policy ID` column replaced with `Policy Name`. Name is resolved via `GET /v2/data-protect/policies/{id}` with in-process caching so each unique policy is only fetched once per script run. Falls back to the raw ID if the lookup fails.
- `reporting/README.md` — Updated field table (Policy Name); moved Policy ID to the unreported fields table.

---

## [2026-04-04] fix(reporting): protection_group_report v3.1 — GB-only size columns for pivot table compatibility

### Changed
- `reporting/protection_group_report.py` — Size values (Data Read, Data Written, Logical Size, Storage Consumed) now report as plain GB decimal numbers. The `(GB)` label is in the column header only, keeping cell values numeric for Excel pivot tables and sorting. Replaces the previous auto-scaling TB/GB/MB format.
- `reporting/README.md` — Full rewrite: added per-field descriptions table for all 25 columns, and a separate table listing available-but-not-reported API fields with their locations and notes.

---

## [2026-04-04] feat(reporting): protection_group_report v2.0 — historical run data with --days/--start/--end flags

### Added
- `reporting/protection_group_report.py` — `__version__ = "2.0"`. Version history block in docstring. Version number shown in `--help`.
- `--days N` — pull last N days of runs (e.g. `--days 30`). Convenient shortcut.
- `--start YYYY-MM-DD` / `--end YYYY-MM-DD` — explicit date range. `--end` defaults to today if omitted. Each run in the range becomes its own CSV row.
- `Run #` column added in historical mode (sequential per protection group).
- `Logical Size Gb` column added (from `localSnapshotStats.logicalSizeBytes`).
- Column rename: `Last Run *` → `Run *` (accurate for both single and historical rows).

### Changed
- `get_last_run` replaced by `get_runs(start_usecs, end_usecs)` — returns a list. Default (no dates) returns last run as a single-element list. Date range returns all matching runs (capped at 1000 per group).
- `build_report` accepts `start_usecs`/`end_usecs` and emits one row per run.

---

## [2026-04-04] fix(reporting): protection_group_report — bytes read/written always empty; wrong stats nesting path

### Fixed
- `reporting/protection_group_report.py` — Bytes read/written (columns N, O) were always N/A. Stats are nested at `localBackupInfo.localSnapshotStats`, not `localBackupInfo.stats`. Fixed by changing `backup.get("stats", {})` to `backup.get("localSnapshotStats", {})`.

---

## [2026-04-04] fix(reporting): protection_group_report — wrong Helios cluster routing header (clusterIdentifier → accessClusterId)

### Fixed
- `reporting/protection_group_report.py` — The Helios cluster routing header was wrong: `clusterIdentifier` does not exist. The correct header confirmed in `cohesity-api.ps1` (Brian Seltzer) is **`accessClusterId`**. This was causing Helios to proxy to no cluster, returning an empty result for every API call — hence "Found 0 active protection jobs".
- Reverted from v1 back to v2 API for both protection groups and runs, matching the `ClusterV2` pattern used in `cohesityInv.ps1`. Added `includeObjectDetails=true` to the runs call to get per-object success/failure/warning counts.
- Updated `extract_run_stats` to parse v2 response structure (`localBackupInfo`, plain-string status values).

---

## [2026-04-04] refactor(reporting): protection_group_report — switch to v1 API for job listing and run stats

### Changed
- `reporting/protection_group_report.py` — Replaced v2 `data-protect/protection-groups` and `protection-groups/{id}/runs` endpoints with v1 `protectionJobs` and `protectionRuns` endpoints (same approach used in Brian Seltzer's reporting scripts).
  - v2 runs endpoint returned sparse/empty data through Helios; errors were silently swallowed returning "No runs yet" for all jobs
  - v1 `protectionRuns` returns complete stats in a single call: start/end times, bytes read from source, bytes written to Cohesity, and per-task success/failure/warning counts
  - Added explicit error logging on failed API calls (status code + truncated response body) instead of silent `return {}`
- Added `objects_warnings` column to CSV output

---

## [2026-04-04] fix(reporting): protection_group_report — wrong Helios cluster list endpoint (/mcm/clusters/info → /mcm/clusters/connectionStatus)

### Fixed
- `reporting/protection_group_report.py` — Corrected Helios cluster discovery endpoint from `/mcm/clusters/info` (404) to `/mcm/clusters/connectionStatus`.

---

## [2026-04-04] fix(reporting): protection_group_report — SSL verification error on corporate networks with proxy CA

### Fixed
- `reporting/protection_group_report.py` — Changed `verify=True` to `verify=False` for Helios cluster discovery call. Corporate laptops with SSL inspection (proxy CA) caused a `ConnectionError` on the initial Helios connection even when the API key was valid.

---

## [2026-04-04] feat(reporting): protection_group_report v1.0 — initial protection group last-run status report

### Added
- `reporting/protection_group_report.py` — Initial script. Generates a CSV report of all protection group last-run status, timing, object counts, and data written/read. Supports Cohesity 7.1 and 7.3.

### Added — Auth
- Direct cluster auth: username/password → Bearer token via `/v2/users/sessions`
- Helios API key auth: `apiKey` header + `clusterIdentifier` routing via `helios.cohesity.com`. Auto-discovers all connected clusters or filters to a named cluster with `--cluster`.

### Added — Repo structure
- Organized repo into topic folders: `reporting/`, `protection/`, `recovery/`, `storage/`, `policies/`, `alerts/`, `infrastructure/`, `utils/`
- Added per-folder `README.md` with planned scripts for each area
- Added root `README.md` and this `CHANGELOG.md`
