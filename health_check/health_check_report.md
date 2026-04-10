# health_check_report.py

**Current version: 1.34**

Multi-cluster Cohesity health check designed for enterprise customer business reviews (CBRs) and SE engagements. Gathers live data from Cohesity Helios and produces a **19-sheet Excel workbook** and a **Word document**.

> See the root [README.md](../README.md) for quick-start instructions and common usage examples.

---

## Output

### Excel Workbook (19 sheets)

| # | Sheet | Contents |
|---|-------|----------|
| 1 | Executive Summary | Per-cluster health score (0–100), grade, success %, capacity used %, open criticals, top finding |
| 2 | Infrastructure | Software version, node count, healthy nodes, cluster encryption, SD encryption, DNS, NTP, timezone, SW lifecycle status, SW EOS date, days to EOS |
| 3 | Node Hardware | Per-node model, serial, node type, software version, raw capacity, disk count, storage tiers, HW EOL date, HW EOL status |
| 4 | **Disk Health** | Per-disk status, type, model, serial, **SSD wear %**, capacity, encryption, storage tier — CRITICAL on failed/missing, HIGH on wear ≥80% |
| 5 | Protection Health | Per-group last-run status, SLA violations, RPO gap, object counts, logical/physical bytes, **DataLock (WORM) per group** |
| 6 | Storage & Capacity | Cluster and per-domain usable/used/free, data reduction ratio, dedup ratio, compression ratio |
| 7 | Policy Audit | Retention schedules, replication targets, archival targets, **DataLock (WORM)** mode + duration per policy, compliance gaps, group count |
| 8 | Policy → Groups | Every protection group with the policy that governs it; sorted by policy then group name |
| 9 | Alerts | All open alerts sorted by severity, with age, description, and entity |
| 10 | Security | **Expanded 19-column checklist**: cluster encryption, SD encryption, vault, FortKnox/indelible, replication, archival, audit log, NTP auth, remote tunnel, cluster MFA, SSO/IDP, **Quorum**, **DataLock (WORM)**, TLS cert expiry |
| 11 | **Agent Health** | Per-host agent version, health status, upgradability, cert expiry, last upgrade error |
| 12 | **Source Coverage** | Registered sources with protected/unprotected object counts and coverage % |
| 13 | Replication & Archive | Replication targets, vault names and types, FortKnox storage consumed (TB) |
| 14 | **FortKnox Data Transfer** | Per-protection-group transfer to every external vault: logical TB, physical TB, storage consumed TB, snapshots. Full=`--days` window; quick=last 1 day |
| 15 | Data Services | NAS views with protocol, quota, usage %, near-quota warnings |
| 16 | Coverage Gaps | Protection groups with failed last run, paused state, or RPO gap > threshold |
| 17 | **User Security** | User accounts, domain, roles, MFA status, locked status, last login — RED for admins without MFA |
| 18 | Trends (30d) | Daily backup success rate aggregated per cluster with embedded line chart |
| 19 | Recommendations | Prioritized action list (CRITICAL / HIGH / MEDIUM / LOW) with business impact |

### Word Document

Narrative report suitable for printing or sharing with customers:

- Cover page (customer name, date, confidential notice)
- Executive Summary with per-cluster scorecard table
- Environment Overview — infrastructure table (cluster enc, SD enc) + software lifecycle + **Environment Topology** table (clusters → replication → archival → FortKnox)
- Protection & Recovery Health
- Storage & Capacity
- Security Posture — updated table: cluster enc, SD enc, indelible (FK/DL), Quorum; FIPS removed
- **Agent Health & Source Coverage** — agent status summary with compatibility doc reference; unprotected source count
- **User Security** — per-cluster user table with MFA and locked account counts
- Prioritized Recommendations table
- Appendix: scoring methodology and thresholds (updated 6-dimension security score)

---

## Usage

Two authentication modes are supported:

```bash
# Helios — API key (key saved to OS keychain after first prompt)
python3 health_check_report.py --apikey <key>

# Direct cluster — username/password + optional MFA (bypasses Helios)
python3 health_check_report.py --cluster-host 10.1.2.3 --username admin [--domain LOCAL]
```

> **Helios note:** Helios does not expose a scriptable username/password login endpoint. Its web login uses OIDC/OAuth2 and cannot be driven from a script. Generate an API key at **helios.cohesity.com → Settings → Access Management → API Keys**.

### Options

| Flag | Default | Description |
|------|---------|-------------|
| `--apikey KEY` | keychain | Helios API key |
| `--cluster-host HOST` | — | Bypass Helios; connect directly to this cluster hostname or IP |
| `--username USER` | — | Cluster username (required with `--cluster-host`) |
| `--password PASS` | keychain / prompt | Password (prompted securely if omitted; saved to keychain) |
| `--domain DOMAIN` | `LOCAL` | Auth domain for cluster login (`LOCAL` or AD name) |
| `--mfa-code CODE` | — | TOTP/OTP code for MFA-enabled cluster accounts |
| `--cluster NAME` | all | (Helios mode) Limit to one cluster by partial name match |
| `--clear-credentials` | — | Remove stored credentials from keychain and exit |
| `--days N` | 30 | Lookback window for alerts and run history |
| `--customer NAME` | *(blank)* | Customer name printed on Word cover page |
| `--output PATH` | `cohesity_health_check` | Base path for output files (no extension) |
| `--quick` | off | Use last-run only; skip per-group run history (faster for large installs) |
| `--excel-only` | off | Skip Word document generation |
| `--word-only` | off | Skip Excel generation |
| `--debug` | off | Print HTTP status for every API call |

### Examples

```bash
# Full report — all clusters, 30-day window
python3 health_check_report.py --apikey abc123 --customer "Acme Corp"

# Quick scan of a single cluster
python3 health_check_report.py --apikey abc123 --cluster prod-east --quick

# Excel only, custom output path and filename
python3 health_check_report.py --apikey abc123 --output /reports/q2_review --excel-only

# 90-day lookback with debug logging
python3 health_check_report.py --apikey abc123 --days 90 --debug
```

---

## Requirements

### Install all dependencies

```bash
# Recommended — works on macOS/Linux when pip is not in PATH
python3 -m pip install requests openpyxl python-docx

# Alternative if pip3 is available
pip3 install requests openpyxl python-docx
```

If you get a "not writeable" warning about site-packages, add `--user`:

```bash
python3 -m pip install --user requests openpyxl python-docx
```

### Package notes

| Package | Required for | Install if missing |
|---------|-------------|-------------------|
| `requests` | All API calls | `python3 -m pip install requests` |
| `openpyxl` | Excel output | `python3 -m pip install openpyxl` |
| `python-docx` | Word output | `python3 -m pip install python-docx` |

`python-docx` is required only for Word output.  If it is not installed, Excel is still produced and Word is skipped with a warning.  `requests` and `openpyxl` are hard requirements — the script will exit with an error if either is missing.

---

## Health Scoring

Each cluster receives a score from 0 to 100 and a grade:

| Score | Grade |
|-------|-------|
| 90–100 | Excellent |
| 75–89 | Good |
| 60–74 | Fair |
| 40–59 | At Risk |
| < 40 | Critical |

**Score components:**

| Dimension | Weight | Measure |
|-----------|--------|---------|
| Protection success rate | 25% | % of runs in the lookback window that succeeded |
| SLA compliance | 20% | % of runs that did not violate SLA |
| Infrastructure health | 15% | % of nodes in healthy state |
| Storage capacity | 15% | Penalized above 70% used; critical above 85% |
| Security posture | 15% | 20 pts each: encryption, FIPS, vault, replication, immutability |
| Alert health | 10% | −10 pts per open critical alert, −2 pts per warning |

---

## Recommendations

The engine evaluates all collected data and generates a prioritized finding list:

| Priority | Meaning | Expected response |
|----------|---------|-------------------|
| CRITICAL | Active risk — data loss, capacity exhaustion, or node failure | Immediate |
| HIGH | Significant gap — unprotected data, no replication, high capacity | Within 30 days |
| MEDIUM | Best-practice gap — FIPS, archival, quota governance | Within 90 days |
| LOW | Housekeeping — quota hygiene, minor policy improvements | Next review cycle |

---

## APIs Used

| Endpoint | Data |
|----------|------|
| `GET /mcm/clusters/connectionStatus` | Cluster list |
| `GET /irisservices/api/v1/public/cluster?fetchStats=true` | Version, config, capacity stats |
| `GET /irisservices/api/v1/public/nodes` | Node health, disk inventory |
| `GET /irisservices/api/v1/public/alerts` | Open alerts with severity and age |
| `GET /v2/data-protect/protection-groups` | Groups with last-run info |
| `GET /v2/data-protect/protection-groups/{id}/runs` | Run history (full mode) |
| `GET /v2/data-protect/policies` | Policy retention and targets |
| `GET /v2/storage-domains?includeStats=true` | Per-domain capacity and reduction stats |
| `GET /v2/file-services/views?includeStats=true` | NAS views, quotas, usage |
| `GET /irisservices/api/v1/public/vaults` | Vault / archival target inventory |
| `GET /irisservices/api/v1/public/protectionRuns` | All protection run history — primary source for Trends tab (paginated) |
| `GET /irisservices/api/v1/public/reports/dataTransferToVaults` | FortKnox transfer data |
| `GET /irisservices/api/v1/public/certificates/webServer` | TLS certificate expiry (Disk Health, Security sheet) |
| `GET /irisservices/api/v1/public/reports/agents` | Agent health, version, upgradability per host |
| `GET /v2/disks/local?nodeId={id}` | Per-disk status, SSD wear %, encryption, storage tier — queried per node (matches Cohesity ELF tool approach) |
| `GET /v2/data-protect/sources` | Registered sources with protected/unprotected object counts (falls back to v1 `/protectionSources/registrationInfo` on older firmware) |
| `GET /irisservices/api/v1/public/users` | User accounts, MFA status, roles, last login |
| `GET /irisservices/api/v1/public/idps` | SSO / IDP configuration |
| `GET /irisservices/api/v1/public/tenants` | Tenant / organization inventory |

---

## Performance Notes

For large installs (many clusters, many groups) the full mode can be slow because it fetches run history for every protection group individually.  Use `--quick` for a fast scan that uses only the last-run status embedded in the protection-group response.

Typical runtimes:

| Mode | 5 clusters / 50 groups each | 10 clusters / 200 groups each |
|------|----------------------------|-------------------------------|
| `--quick` | ~30 seconds | ~60 seconds |
| Full | ~5–10 minutes | ~20–30 minutes |

---

## Version History

| Version | Date | Change |
|---------|------|--------|
| 1.32 | 2026-04-09 | feat: DataLock (WORM) column added to Protection Health (col 18, per group, color-coded); new FortKnox Data Transfer sheet (sheet 14) — per-protection-group view of logical/physical TB transferred and storage consumed per vault; full mode uses `--days` window, quick mode uses 1-day window; sheet count 18 → 19 |
| 1.31 | 2026-04-09 | polish: Word doc headings/cover title color unified to `#70AD47`; Excel header rows gain thin black border + height 32; `auto_fit_columns()` handles multiline values, min=12, max=45; Trends charts moved to col L (right of data), height=16, legend at bottom, title 14 pt bold with overlay disabled |
| 1.30 | 2026-04-09 | feat: DataLock (WORM) tracking — new `DataLock (WORM)` column in Policy Audit (per-policy mode+duration, col 15) and Security sheet (cluster rollup col 16: Compliance/Administrative/Mixed/None). "No DataLock" added to Gaps. MEDIUM recommendation if no DataLock; LOW if Administrative-only. No new API calls — reads `backupPolicy.regular.retention.dataLockConfig` from already-fetched v2 policies |
| 1.29 | 2026-04-09 | fix: removed Helios username/password auth — `helios.cohesity.com` login endpoints require OIDC/OAuth2 (`401 "No open ID token found in header"`). Using `--username` without `--cluster-host` now prints a clear error with instructions to generate an API key. `cohesity_auth.py` v1.4 |
| 1.28 | 2026-04-09 | feat: direct cluster mode (`--cluster-host` + `--username` → Bearer token via `POST /v2/users/sessions`). Cluster name auto-detected from cluster's own API. `--clear-credentials` flag. `_get()` uses `session.base_url`. `cohesity_auth.py` v1.2 |
| 1.27 | 2026-04-09 | fix: Disk Health sheet still empty after v1.26 — `GET /v2/disks` is not a valid Helios-proxied endpoint. Rewrote `_disks()` to use `GET /v2/disks/local?nodeId={id}` per node (matches Cohesity ELF tool `cohesityInv.ps1` approach). Updated field fallbacks: `ssdUsedPercentage` (ELF name), `capacityInBytes` (ELF name), `encryptionStatus` string (ELF name) |
| 1.26 | 2026-04-09 | fix: Disk Health sheet empty — switched `_disks()` from private `/v1/disks/local?nodeId=X` (always 400 via Helios) to supported `GET /v2/disks`; private endpoint kept as fallback; added `isinstance` guard on `usageStats` in `_sheet_disk_health()` |
| 1.25 | 2026-04-09 | chore: standardize on US English throughout — utilization, prioritized, normalized, color-coded, behavior, organization, authorized, summarizes, maximize, minimize, penalized, labeled; affects sheet titles, Word doc narrative, docstring, comments |
| 1.24 | 2026-04-09 | fix: Source Coverage sheet empty on older clusters — `/v2/data-protect/sources` returns 404 on pre-7.x firmware; `_sources()` now falls back to v1 `/protectionSources/registrationInfo` and normalizes `rootNodes[].stats.protectedCount/unprotectedCount` into standard dict shape; also added `isinstance` guards to `stats`/`objectCount` lookups in `_sheet_source_coverage` |
| 1.23 | 2026-04-09 | fix: suppress debug body output for best-effort endpoints (`/disks/local`, `/v2/data-protect/sources`, `/tenants`) — added `silent` param to `_get()`; these callers pass `silent=True` so `--debug` output is not cluttered with expected 400/403/404 bodies |
| 1.22 | 2026-04-09 | fix: replace all unsafe `(x or {}).get()` patterns with `isinstance(x, dict)` guards — affected: `healthStatus` (string in older clusters), `auditLogConfig`, `clusterAuditConfig`, `encryptionConfig`, `ntpSettings`, `remoteSupportConfig`, `mfaConfig`, source `stats`; added `isinstance(d, dict)` guards to disk comprehensions in `_build_recommendations()` |
| 1.21 | 2026-04-08 | feat: ELF inventory additions — new sheets: Disk Health (per-disk status/SSD wear/encryption), Agent Health (agent version/status/upgradability), Source Coverage (protected vs unprotected objects per source), User Security (MFA/locked/roles/last login); Security sheet expanded to 17 controls (+ audit log, NTP auth, remote tunnel, cluster MFA, SSO, cert expiry); 7 new API calls; new recommendations for cert expiry, disk failures, SSD wear, agents, source coverage, user MFA, audit log, SSO; sheet count 14→18 |
| 1.20 | 2026-04-08 | fix: 6.8.1 LTS EOS was November 15 2024 (already OOS) — old `"6.8"` prefix catch-all incorrectly reported it as in-support; split table so only 6.8.2+ maps to In Support — LTS. Hardware EOL table expanded: added C2100, C2200, C2510, C4500 model variants (Cohesity EOL doc Oct 2025) |
| 1.19 | 2026-04-08 | feat: software and hardware lifecycle tracking — `SOFTWARE_EOS` and `HARDWARE_EOL` tables; `_sw_eos()` / `_hw_eol()` helpers; Infrastructure sheet gains SW Lifecycle Status (color-coded), SW EOS Date, Days to EOS; new "Node Hardware" sheet (14th sheet) with per-node model/serial/type/capacity/tiers/EOL; recommendations engine updated with CRITICAL/HIGH/MEDIUM lifecycle findings; Word doc gains Software & Hardware Lifecycle section |
| 1.18 | 2026-04-08 | fix: Trends pagination still broken after v1.17 — `len(page) < PAGE_SIZE` assumed server cap == 1000, but actual cap varies by cluster version and is often lower; check fired on page 1 and stopped early. Fix: removed page-size check; page purely by timestamp cursor until empty response, window start reached, or no progress. Safety limit 200 pages |
| 1.17 | 2026-04-08 | fix: Trends tab showed fewer than 30 days on busy clusters — added pagination by advancing `endTimeUsecs` cursor |
| 1.16 | 2026-04-08 | feat: Trends chart — visible X/Y axes with tick marks and titles, Y axis fixed 0–100 with `0.0` format, circle data-point markers (size 5, Cohesity green), 2 pt line, StrRef category fix for date labels in Excel, auto tick-density reduction for >14d and >60d ranges |
| 1.15 | 2026-04-08 | feat: include `kWarning` runs in Success % — Trends col H renamed "Success % (incl. Warnings)"; `_score_protection()` and `_success_stats()` quick mode now count warnings as successful; Word doc Methodology appendix documents that warnings complete successfully but appear separately for visibility |
| 1.14 | 2026-04-08 | fix: Trends tab inaccurate — switched primary data source from per-group v2 runs to v1 `/public/protectionRuns` (same source as Helios reporting page); single call per cluster; `backupRun.slaViolated` and `backupRun.stats.totalLogicalBackupSizeBytes` are accurate; v2 per-group runs kept as fallback; works in `--quick` mode |
| 1.13 | 2026-04-08 | fix: Trends tab cols I (SLA Violations) and J (Logical GB) always blank — col I used `run.get("isSlaViolated")` but field lives in `localBackupInfo`; col J used `lb.get("stats").logicalSizeBytes` but correct sub-object is `localSnapshotStats` |
| 1.12 | 2026-04-08 | fix: Replication & Archive col E and col K still empty after v1.11 — `policy_arch_targets` stayed empty because `_arch_name()` returns `""` for ID-only vault references; supplement `arch_groups` directly from `fk_data.dataTransferPerProtectionJob[].protectionJobName` per vault; build `all_arch_targets` as union of policy chain + fk_data vault names + vault list names |
| 1.11 | 2026-04-08 | feat: Replication & Archive col E changed from policy count to protection group count; new col K "Protection Group Names" lists every group using each target; Status shows "Configured" / "No groups assigned"; target discovery rewired through policy→group chain |
| 1.10 | 2026-04-08 | fix: Replication & Archive col E blank — v2 policies may reference archival targets by vault ID only; added `vaultId` → name resolution via `/public/vaults`; cols G/H renamed to include API field names (`numLogicalBytesTransferred`, `numPhysicalBytesTransferred`) and cumulative note; "Vault only — no policy" → "Vault exists — not referenced by any policy" |
| 1.9 | 2026-04-08 | fix: Replication & Archive still blank after v1.8 — `dataTransferToVaults` requires `vaultIds` filter to return rows; extract vault IDs from `_vaults()` and pass to `_fortknox_data()`; vault type (col D) now read from `archivalTargetConfig.targetType` inside policy target, not a name-based vault list lookup |
| 1.8 | 2026-04-08 | fix: Replication & Archive cols D–H blank — `_fortknox_data` used wrong time param names (`Usecs` → `Msecs`); archival rows now driven from policy `arch_use` targets (not vault list) with name-then-ID FK stats lookup; G/H (Logical/Physical Transferred) populated from `numLogicalBytesTransferred`/`numPhysicalBytesTransferred` |
| 1.7 | 2026-04-08 | fix: `_fk_status()` false-negatives — `lastRun.archivalInfo` is only set when the most recent run included archival; daily-backup + weekly-archive groups showed as idle. Fix scans full `group_runs` history (non-quick mode) by archival target type and vault name; `lastRun` remains fallback for `--quick` |
| 1.6 | 2026-04-08 | feat: FortKnox idle detection — `_fk_status()` inspects group `lastRun.archivalInfo` to distinguish "Active" / "Configured — Idle" / "No"; idle shown in orange on Security sheet, triggers HIGH recommendation, scores 10/20 |
| 1.5 | 2026-04-08 | fix: remove Alerts columns K (Node ID) and L (Entity) — `propertyList` fields only populated on hardware alerts; description + alert code already capture the context |
| 1.4 | 2026-04-08 | fix: Policy Audit column N (Groups Using) computed from group policyId cross-reference — v2 API has no `numProtectionGroups`; feat: new "Policy → Groups" sheet (sheet 6) mapping every group to its governing policy |
| 1.3 | 2026-04-08 | fix: infrastructure disk count uses `diskCount` integer per node (v1 API field); NTP servers shows "Not configured" instead of blank; added `ntpServersList`/`ntpServerIps` fallback paths |
| 1.2 | 2026-04-08 | fix: correct API field names — firstTimestampUsecs, localSnapshotStats, successfulObjectsCount, policyId→name lookup, v2 schedule/target paths, FortKnox detection, N/A capacity fallback; feat: output filename includes customer name + UTC timestamp |
| 1.1 | 2026-04-08 | fix: show N/A for capacity when cluster stats return null; fix: clusterSoftwareVersion field name; fix: MergedCell crash in auto_fit_columns; fix: deprecated utcfromtimestamp/utcnow |
| 1.0 | 2026-04-08 | feat: initial release — 12-sheet Excel, Word doc, scoring engine, recommendations |
