# Changelog

All notable changes to this repository are documented here.
Format: `[YYYY-MM-DD] type(scope): description`

Commit types: `feat` (new feature), `fix` (bug fix), `refactor` (restructure), `docs` (docs only), `chore` (maintenance)

---

## [2026-04-09] fix(health_check): switch disk API from private v1 endpoint to GET /v2/disks (v1.26)

### Fixed — `health_check/health_check_report.py`
- `_disks()` now calls `GET /v2/disks` first (works through Helios proxy, same endpoint used by `node_status.py` and `upgrade_readiness.py`). The private `/irisservices/api/v1/disks/local?nodeId=X` endpoint is kept as a fallback but always returned 400 via Helios.
- Fixed `isinstance` guard on the `usageStats` sub-object in `_sheet_disk_health()`.

---

## [2026-04-09] chore(health_check): standardize on US English throughout (v1.25)

### Changed — `health_check/health_check_report.py` and docs
- Replaced all British English spellings with US equivalents: utilization, prioritized, normalized, color-coded, behavior, organization, authorized, summarizes, maximize, minimize, penalized, labeled.
- Affected locations: sheet title strings, Word document narrative, docstring version history, inline comments. No logic changes.

---

## [2026-04-09] fix(health_check): Source Coverage sheet empty on older clusters — v1 fallback for /protectionSources (v1.24)

### Fixed — `health_check/health_check_report.py`
- **Source Coverage tab was empty** on clusters where `/v2/data-protect/sources` returns 404 (pre-7.x firmware). `_sources()` now tries v2 first, then falls back to v1 `/irisservices/api/v1/public/protectionSources/registrationInfo`. The v1 response (`rootNodes[].stats.protectedCount` / `unprotectedCount`) is normalized into the same dict shape (`protectedObjectsCount`, `unprotectedObjectsCount`) so `_sheet_source_coverage` and `_build_recommendations` work without changes.
- Fixed `isinstance` guards on `stats` and `objectCount` sub-object lookups in `_sheet_source_coverage` (same `(x or {}).get()` class of bug).

---

## [2026-04-09] fix(health_check): suppress debug body noise for best-effort endpoints (v1.23)

### Fixed — `health_check/health_check_report.py`
- Added `silent` parameter to `_get()`. When `silent=True`, non-200 responses still print the status code in `--debug` mode but skip the response body. Callers: `_disks()` (`/disks/local` → 400 when not proxied), `_sources()` (`/v2/data-protect/sources` → 404 on older firmware), `_tenants()` (`/tenants` → 403 on restricted API keys).

---

## [2026-04-09] fix(health_check): replace unsafe (x or {}).get() patterns with isinstance guards throughout (v1.22)

### Fixed — `health_check/health_check_report.py`
- **`healthStatus` field** — in some cluster versions this is a string (e.g. `"kHealthy"`) not a nested dict; `(x or {}).get("status")` raised `AttributeError: 'str' object has no attribute 'get'`. Fixed in `_build_recommendations()`, `_sheet_agent_health()`, and Word doc `_ag_st()` helper.
- **`auditLogConfig` / `clusterAuditConfig`** — could be returned as a string on certain cluster versions; replaced `(x or {}).get("enabled")` with `isinstance(x, dict)` guards in both `_build_recommendations()` and `_sheet_security()`.
- **`encryptionConfig`** — same pattern; fixed in `_score_security()`, `_build_recommendations()`, `_sheet_infrastructure()`, `_sheet_security()`, and two Word doc sections.
- **`ntpSettings`** — replaced double `(info.get("ntpSettings") or {}).get(...)` calls (one in `_sheet_infrastructure()`, one in `_sheet_security()`) with a single `_ntp_s` variable guarded by `isinstance`.
- **`remoteSupportConfig` / `mfaConfig`** — same pattern in `_sheet_security()`.
- **Disk comprehensions** — added `isinstance(d, dict)` guard to `failed_disks` and `high_wear` list comprehensions in `_build_recommendations()`.
- **Source `stats` sub-object** — replaced `(src.get("stats") or {}).get(...)` with `isinstance` guard.

---

## [2026-04-08] feat(health_check): ELF inventory additions — disk health, agent health, source coverage, user security, expanded security controls (v1.21)

### Added — `health_check/health_check_report.py`
- **New data collection**: `_certificates()`, `_agents()`, `_disks()`, `_sources()`, `_users()`, `_idps()`, `_tenants()` — 7 new API calls per cluster.
- **Disk Health sheet** (sheet 4) — per-disk status, type, model, serial, SSD wear %, capacity, encryption, storage tier. CRITICAL on failed/missing disks, HIGH on SSD wear ≥80%.
- **Agent Health sheet** (sheet 11) — per-host agent version, health status, upgradability, cert expiry, last upgrade error.
- **Source Coverage sheet** (sheet 12) — registered sources with protected/unprotected object counts and coverage %. RED < 75%, ORANGE < 95%.
- **User Security sheet** (sheet 16) — user accounts, domain, roles, active/locked, MFA status, last login. RED for admin accounts without MFA.
- **Security sheet expanded** from 11 to 17 columns: added Audit Log, NTP Authentication, Remote Support Tunnel, Cluster MFA, SSO/IDP, Soonest TLS Cert Expiry.
- **New threshold constants**: `SSD_WEAR_WARN_PCT=70`, `SSD_WEAR_CRIT_PCT=80`, `CERT_WARN_DAYS=90`, `CERT_CRIT_DAYS=30`.
- **Recommendations engine** — new findings: TLS cert expiry (CRITICAL/HIGH), unhealthy agents (HIGH), failed disks (CRITICAL), SSD wear (HIGH), unprotected sources (HIGH), admins without MFA (MEDIUM), audit log disabled (MEDIUM), no SSO (LOW).
- **Word document** — new sections 6 (Agent Health & Source Coverage) and 7 (User Security) with narrative and tables.
- Sheet count: 14 → 18.

---

## [2026-04-08] fix(health_check): correct 6.8.1 EOS classification and expand hardware EOL table (v1.20)

### Fixed — `health_check/health_check_report.py`
- `6.8.1` LTS reached EOS **November 15, 2024** and is already out of support. The old `"6.8"` prefix catch-all incorrectly matched it as "In Support — LTS (Jun 2026)". Table now maps `6.8.2` → In Support and `6.8.0`/`6.8.1` → Out of Support.
- Hardware EOL table expanded with model variants confirmed in the official Cohesity Products EOL document (Oct 2025): **C2100**, **C2200** (C2xx0 series), **C2510** (C2xx5 series), **C4500** (C4xx0 series).

---

## [2026-04-08] feat(health_check): software and hardware lifecycle tracking (v1.19)

### Added — `health_check/health_check_report.py`
- **`SOFTWARE_EOS` table** — version prefix → lifecycle status, EOS date, label. Covers 7.3+ (Current), 7.2.x / 7.1.2 / 6.8.2 (In Support — EOS Jun 8 2026), 7.1.0 / 7.1.1 / 7.0 / 6.8.1 / 6.8.0 (Out of Support).
- **`HARDWARE_EOL` table** — model substring → EOL date. Covers C2xx0 (Jan 2023), C2xx5 (May 2024), C3500 (Feb 2025), C4xx0/CN (Feb 2027). Source: Cohesity Products EOL document, Oct 2025.
- **`_sw_eos(version)`** helper — returns `(status, eos_date, label)` via prefix matching.
- **`_hw_eol(model)`** helper — returns `eol_date` or `None` via substring matching.
- **Infrastructure sheet** — 3 new columns: `SW Lifecycle Status` (RED = Out of Support, ORANGE = <90 days, green = current), `SW EOS Date`, `Days to EOS`.
- **New "Node Hardware" sheet** (sheet 3, 14 sheets total) — per-node inventory: model, serial, node type, software version, raw capacity, disk count, storage tiers, HW EOL date, HW EOL status (color-coded).
- **Recommendations engine** — CRITICAL for OOS software or past-EOL hardware; HIGH for hardware EOL within 180 days or software EOS within 90 days; MEDIUM for software EOS within 180 days.
- **Word document** — new "Software & Hardware Lifecycle" section with SW EOS table per cluster and hardware EOL narrative.
- Sheet count: 13 → 14.

---

## [2026-04-08] docs: update README, add health check quick-start; CHANGELOG updated

### Changed
- `README.md` — Health check report promoted to top-level featured section with full quick-start: blurb, requirements, where to run from, common usage examples, all options table, and output filename format. `health_check/` added to the repository structure table. `python-docx` added to global requirements.
- `health_check/health_check_report.md` — Version banner (`v1.18`), EBRs → CBRs, cross-reference link to root README, `GET /public/protectionRuns` added to APIs Used table.
- `CHANGELOG.md` — Health check script history added.

---

## [2026-04-08] feat: add version number to auto-generated output filenames across all scripts

### Changed
All 14 scripts that auto-generate output filenames now embed the script version:

```
health_check_report_v1.18_AcmeCorp_20260408_1430.xlsx
fortknox_vault_report_v4.10_20260408_1430.xlsx
protection_group_report_v4.5_20260408_1430.xlsx
alert_to_csv_v1.0_20260408_143022.csv
```

Scripts affected: `health_check_report.py`, `alert_summary_report.py`, `alert_to_csv.py`, `cluster_info_report.py`, `node_status.py`, `upgrade_readiness.py`, `list_policies.py`, `policy_compliance_report.py`, `list_snapshots.py`, `fortknox_vault_report.py`, `protection_group_report.py`, `capacity_report.py`, `list_views.py`, `storage_domain_report.py`. Scripts that accept an explicit `--output` path are unaffected.

---

## [2026-04-08] style: update all green to #70AD47 (RGB 112, 173, 71) across every script output

### Changed
Replaced `#00B388` with `#70AD47` in every Excel header fill, chart line/marker color, and cell background fill across all 14 scripts. Light green cell fills (`#C6EFCE`) updated to `#E2EFDA` (lighter derivative of `#70AD47`).

---

## [2026-04-08] fix(health_check): Trends tab pagination — remove page-size assumption (v1.18)

### Fixed
- `health_check/health_check_report.py` — Trends tab still showed fewer than 30 days on busy clusters after v1.17. Root cause: `if len(page) < PAGE_SIZE: break` assumed the server caps at exactly 1000 runs per response. The actual server-side cap varies by cluster version (often 100–200); the check fired on the first page and stopped pagination immediately. Fix: removed the page-size check entirely. Pagination now advances purely by timestamp cursor and stops when the response is empty, the oldest run predates the window start, or no progress is made. Safety limit raised to 200 pages.

---

## [2026-04-08] fix(health_check): Trends tab — paginate v1 protectionRuns to cover full 30-day window (v1.17)

### Fixed
- `health_check/health_check_report.py` — Trends tab showed only 10–20 days on clusters with many protection groups. Root cause: the v1 `protectionRuns` API returns runs newest-first and caps per page; a single call would only return the most recent N days. Fixed by paginating: advance `endTimeUsecs` to just before the oldest run on each page and repeat until the full window is covered.

---

## [2026-04-08] feat(health_check): Trends chart — visible axes, data-point markers (v1.16)

### Changed
- `health_check/health_check_report.py` — Trends sheet line chart now has visible X and Y axes with tick marks, axis titles, and data-point circle markers on every day. Y axis fixed to 0–100 scale with `0.0` format. Date labels use StrRef to force string rendering in Excel. Auto tick-density reduction for ranges >14 days and >60 days.

---

## [2026-04-08] feat(health_check): include kWarning in Success % calculation (v1.15)

### Changed
- `health_check/health_check_report.py` — `kWarning` runs (complete with data protected, minor issues) now count toward Success %. Trends col H renamed "Success % (incl. Warnings)". `_score_protection()` and `_success_stats()` quick mode updated to match non-quick mode. Word doc Methodology appendix documents this behavior.

---

## [2026-04-08] fix(health_check): Trends tab — switch to v1 protectionRuns API for accurate data (v1.14)

### Changed
- `health_check/health_check_report.py` — Trends tab primary data source switched from per-group v2 runs to the v1 `GET /public/protectionRuns` endpoint — the same source powering the Helios protection-group-summary reporting page. Single call per cluster; `backupRun.slaViolated` (col I) and `backupRun.stats.totalLogicalBackupSizeBytes` (col J) are accurate. v2 per-group runs retained as fallback. Works in `--quick` mode.

---

## [2026-04-08] fix(health_check): Trends tab cols I and J always blank (v1.13)

### Fixed
- `health_check/health_check_report.py` — Col I (SLA Violations): used `run.get("isSlaViolated")` but the field lives inside `localBackupInfo`; fixed to `lb.get("isSlaViolated")`. Col J (Logical GB): used `lb.get("stats").get("logicalSizeBytes")` but the correct sub-object is `localSnapshotStats`; fixed to `lb.get("localSnapshotStats").get("logicalSizeBytes")`.

---

## [2026-04-08] fix(health_check): Replication & Archive col E and K empty (v1.12)

### Fixed
- `health_check/health_check_report.py` — `policy_arch_targets` stayed empty because `_arch_name()` returns `""` when v2 policies store archival targets by vault ID only (no `targetName`/`name` field). With no policy-chain rows, `all_arch_targets` was empty and both col E (protection group count) and col K (group names) were blank for all rows. Fixed by supplementing `arch_groups` directly from `fk_data.dataTransferPerProtectionJob[].protectionJobName` per vault, and building `all_arch_targets` as the union of three independent sources: policy chain + fk_data vault names + vault list names.

---

## [2026-04-08] feat(health_check): Replication & Archive — protection group count and group names (v1.11)

### Changed
- `health_check/health_check_report.py` — Col E changed from policy count to protection group count using each archival target. New col K "Protection Group Names" lists every group name associated with each vault target, comma-separated. Status shows "Configured" (groups present) or "No groups assigned". Target discovery rewired through policy→group chain.

---

## [2026-04-08] fix(health_check): Replication & Archive — vault ID resolution, col E blank (v1.10)

### Fixed
- `health_check/health_check_report.py` — v2 policies may reference archival targets by vault ID only (no name/targetName field), causing `_arch_name()` to return `""` and col E to show 0. Fixed by building a `vaultId` → name lookup from `/public/vaults` and resolving IDs to names before building target maps. Cols G/H renamed to include API field names and "cumulative total" note.

---

## [2026-04-08] fix(health_check): Replication & Archive — vaultIds filter required for dataTransferToVaults (v1.9)

### Fixed
- `health_check/health_check_report.py` — `dataTransferToVaults` API returns empty results without a `vaultIds` filter. Fixed by extracting vault IDs from `_vaults()` and passing them to `_fortknox_data()`. Vault type (col D) now read from `archivalTargetConfig.targetType` inside the policy target rather than a name-based vault list lookup.

---

## [2026-04-08] fix(health_check): Replication & Archive cols D–H blank (v1.8)

### Fixed
- `health_check/health_check_report.py` — Three root causes: (1) `_fortknox_data()` used `startTimeUsecs`/`endTimeUsecs` param names — API requires `startTimeMsecs`/`endTimeMsecs`; (2) archival rows driven by vault list name mismatches; (3) cols G/H (Logical/Physical Transferred) hardcoded as empty — now populated from `numLogicalBytesTransferred`/`numPhysicalBytesTransferred`.

---

## [2026-04-08] fix(health_check): FortKnox status — false-negatives from lastRun.archivalInfo (v1.7)

### Fixed
- `health_check/health_check_report.py` — `_fk_status()` was returning "idle" for clusters actively sending to FortKnox. Root cause: `lastRun.archivalInfo` is only populated when the most recent primary run included archival — groups with daily local backups and less-frequent archival appeared idle even when active. Fixed by scanning full `group_runs` history (non-quick mode) for successful archival runs matched by target type and vault name. `lastRun` remains the fallback in `--quick` mode.

---

## [2026-04-08] feat(health_check): FortKnox idle detection in Security tab (v1.6)

### Added
- `health_check/health_check_report.py` — Security tab now distinguishes three FortKnox states: "Active" (green, data sent recently), "Configured — Idle" (orange, vault configured but no recent transfers), "No" (red, not configured). Idle state scores 10/20 points instead of the full 20. Triggers a HIGH-priority recommendation. Detected via `group_runs` archival history and `dataTransferToVaults` transfer records.

---

## [2026-04-08] fix(health_check): drop Alerts columns K and L; fix Policy Audit col N; add Policy → Groups sheet (v1.4–1.5)

### Fixed
- `health_check/health_check_report.py` — Alerts tab: removed columns K (Node ID) and L (Entity) — `propertyList` fields are only populated on hardware alerts; they were always blank for software/backup alerts.
- Policy Audit col N (Groups Using): `p.get("numProtectionGroups")` does not exist in the v2 API. Fixed by counting groups whose `policyId` matches each policy from `cd["groups"]`.

### Added
- New "Policy → Groups" sheet (sheet 6): lists every protection group with its governing policy name, sorted by policy then group name. Paused groups highlighted in yellow, failed groups in red.

---

## [2026-04-08] fix(health_check): infrastructure disk count and NTP columns (v1.3)

### Fixed
- `health_check/health_check_report.py` — Disk count (col G): previous helper looked for `diskVec`/`disks`/`diskInfo` list fields which don't exist in the v1 `/public/nodes` response. Fixed to use `n.get("diskCount")` integer per node. NTP servers (col J): empty list produced a blank cell instead of "Not configured". Added `ntpServersList` and `ntpServerIps` fallback paths.

---

## [2026-04-08] feat(health_check): initial release — 13-sheet Excel workbook + Word document (v1.0–1.2)

### Added
- `health_check/health_check_report.py` v1.0 — Multi-cluster Cohesity health check for CBRs and SE engagements. Connects to Helios, discovers all clusters, and collects: cluster info, nodes, alerts, protection groups, policies, storage domains, views, vaults, FortKnox transfer data, and run history. Produces:
  - **Excel workbook** (13 sheets): Executive Summary, Infrastructure, Protection Health, Storage & Capacity, Policy Audit, Policy → Groups, Alerts, Security, Replication & Archive, Data Services, Coverage Gaps, Trends (30d), Recommendations
  - **Word document**: cover page, executive scorecard, environment overview, protection health, storage, security posture, recommendations, and scoring methodology appendix
- Scoring engine: six weighted dimensions (protection 25%, SLA 20%, infrastructure 15%, storage 15%, security 15%, alerts 10%). Grades: Excellent / Good / Fair / At Risk / Critical.
- Recommendations engine: CRITICAL / HIGH / MEDIUM / LOW prioritized findings with business impact text.
- `--quick` mode: uses last-run only (no per-group run history) for fast scans of large installs.
- `--debug` flag: prints HTTP status for every API call.
- Output filenames include version number, customer name, and local timestamp.
- `health_check/health_check_report.md` — Full reference: sheet descriptions, scoring model, recommendations priorities, all API endpoints used, performance notes, and version history table.

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
