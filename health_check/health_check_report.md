# health_check_report.py

**Current version: 1.73**

Multi-cluster Cohesity health check designed for enterprise customer business reviews (CBRs) and SE engagements. Gathers live data from Cohesity Helios and produces a **31-tab Excel workbook** (About + Guide + 29 data sheets), a **Word document**, and a comprehensive **~27-slide PowerPoint deck** (requires `python-pptx`).

> See the root [README.md](../README.md) for quick-start instructions and common usage examples.

---

## Output

### Excel Workbook (31 tabs: About + Guide + 29 data sheets)

| # | Sheet | Contents |
|---|-------|----------|
| — | **About** | First tab — full CLI command, run parameters (mode, days, customer, clusters), 4 explanatory notes (Quick Mode, Storage Consumed, DataLock limit, Health Score), and reference table of all 30 API endpoints with routing and purpose |
| — | **Guide** | Sheet descriptions, health/security/ransomware/workload scoring methodology, recommendation priority levels, color/RAG legend |
| 1 | Executive Summary | Per-cluster health score (0–100), grade, success %, capacity used %, open criticals, ransomware readiness score, capacity runway, top finding |
| 2 | Infrastructure | Software version, node count, healthy nodes, cluster encryption, SD encryption, DNS, NTP, timezone, SW lifecycle status, SW EOS date, days to EOS, **Fault Tolerance margin** |
| 3 | Node Hardware | Per-node model, serial, node type, software version, raw capacity, disk count, storage tiers, HW EOL date, HW EOL status, **SW Version Match** (MISMATCH flag) |
| 4 | **Disk Health** | Per-disk status, type, model, serial, **SSD wear %**, capacity, encryption, storage tier — CRITICAL on failed/missing, HIGH on wear ≥80% |
| 5 | Protection Health | Per-group last-run status, SLA violations, RPO gap, object counts, logical/physical bytes, **DataLock (WORM) per group** |
| 6 | Storage & Capacity | Cluster and per-domain usable/used/free, data reduction ratio, dedup ratio, compression ratio, **predictive runway to 80% utilization** |
| 7 | Policy Audit | Retention schedules, replication targets, archival targets, **DataLock (WORM)** mode + duration per policy; FortKnox/RPaaS policies show **"FortKnox (Indelible)"** in green — not flagged as "No DataLock" |
| 8 | **DataLock Verification** | Snapshot-level DataLock check — VERIFIED / PARTIAL / NOT LOCKED per group; CRITICAL recommendation if policy says locked but snapshots are not |
| 9 | Policy → Groups | Every protection group with the policy that governs it; sorted by policy then group name |
| 10 | **Retention Health** | Three-section retention security analysis: (1) per-cluster summary — min/max/avg local retention, min/max archive, min replication, DataLock coverage, FortKnox status, Risk Level (HIGH/MEDIUM/LOW); (2) per-policy detail — local retention, WORM mode, log retention, replication targets + min retention, archival targets + min retention, archive DataLock, per-policy Flags; (3) prioritized P1/P2/P3 security recommendations per cluster |
| 11 | Alerts | All open alerts sorted by severity, with age, description, and entity |
| 12 | Security | **Expanded checklist** with security & anomaly alert sub-section and **Encryption Key Management** sub-section (KMS type, server, risk) |
| 13 | **Agent Health** | Per-host agent version, health status, upgradability, cert expiry, last upgrade error |
| 14 | **Source Coverage** | Registered sources with protected/unprotected object counts and coverage % |
| 15 | **Recovery Audit** | Per-recovery rows (status, duration, type, objects); cluster summary with days-since-last-recovery; HIGH recommendation if no recovery found |
| 16 | **Unprotected Objects** | Named list of unprotected objects per source — red if count >10, amber if count >0 |
| 17 | Replication & Archive | Replication targets, vault names and types, FortKnox storage consumed (TB), **Replication Lag (hrs)** |
| 18 | **FortKnox Data Transfer** | Per-protection-group transfer to every external vault: storage consumed TB (cumulative, from dataTransferToVaults API), **Last FK Archival date**, **Days Since FK Transfer** |
| 19 | **Data Exposure** | NAS view exposure analysis — SMB discovery, NFS open mount, S3 access, quota presence; per-view Risk Level (HIGH/MEDIUM/LOW) |
| 20 | Data Services | NAS views with protocol, quota, usage %, near-quota warnings |
| 21 | Coverage Gaps | Protection groups with failed last run, paused state, or RPO gap > threshold |
| 22 | **User Security** | User accounts, domain, roles, MFA status, locked status, last login, **Risk Flag** (Stale Admin / Admin No MFA / Locked Admin); **Custom Role Definitions** sub-section |
| 23 | Trends (30d) | Daily backup success rate + **Avg Duration (min)** column; dual line charts (success rate + duration trend) |
| 24 | Recommendations | Prioritized action list (CRITICAL / HIGH / MEDIUM / LOW) with business impact |
| 25 | **Workload Risk Heatmap** | All protection groups scored 0–100 by recovery risk (last-run status 35pts, SLA 25pts, RPO gap 25pts, DataLock 15pts); sorted worst-first with full-row RAG coloring |
| 26 | **Audit Log** | Last 30 days of configuration changes from the cluster audit trail; categorized by type; high-risk events highlighted in red; shows actionable message if API key lacks audit-log permission (HTTP 403) |
| 27 | **AD & Identity Health** | Active Directory domain connection status with preferred DC; LDAP provider status, server, port, base DN — disconnected identity sources highlighted red |
| 28 | **Certificate Inventory** | Full TLS certificate inventory: name, type, subject, issuer, issued/expiry dates, days remaining — CRITICAL (<30d), HIGH (<90d), OK |
| 29 | **Capacity & Perf Trends** | 30-day dedup ratio trend (ransomware indicator — sudden drop = data incompressible) and daily read/write I/O throughput (GB/day) with embedded line charts |

### Word Document

Narrative report suitable for printing or sharing with customers:

- Cover page (customer name, date, confidential notice)
- Executive Summary with per-cluster scorecard table
- Environment Overview — infrastructure table (cluster enc, SD enc) + software lifecycle + **Environment Topology** table (clusters → replication → archival → FortKnox)
- Protection & Recovery Health
- Storage & Capacity
- Security Posture — table: cluster enc, SD enc, indelible (FK/DL), Quorum (live from Helios quorum API); FIPS removed
- **Agent Health & Source Coverage** — agent status summary with compatibility doc reference; unprotected source count
- **User Security** — per-cluster user table with MFA and locked account counts; admins without MFA flagged HIGH
- Prioritized Recommendations table
- Appendix: scoring methodology and thresholds (security score: 6 dimensions + −10 admin MFA penalty)

### PowerPoint Deck (v1.50+)

Each run produces a **`.pptx` file** (requires `pip install python-pptx`) with ~27 slides organised into 5 native PowerPoint sections:

| Section | Slides |
|---------|--------|
| **Cover** | Title, customer name, date |
| **Contents** | Hyperlinked table of contents |
| **Executive Summary** | KPI snapshot, health scorecard, capacity & growth, protection, ransomware readiness, top at-risk workloads, priority actions |
| **Environment Topology** | Cluster/replication/archival/FortKnox diagram with connector labels and legend |
| **Security Deep-Dive** | Security posture overview, user security & MFA, audit & governance activity, security recommendations |
| **Backup Engineering** | SW/HW lifecycle, coverage gaps, agent health, source coverage, workload risk heatmap, engineering recommendations |

All data slides use green/amber/red RAG colour-coding consistent with the Excel workbook. DataLock is RAG-coloured (red = no DataLock, amber = partial, green = full) across all slides where it appears. Scoring methodology for all four models (Overall Health Score, Security Score, Ransomware Readiness, Workload Risk) is embedded inline at the bottom of each respective data slide — compact tables separated by a thin gray rule so they don't distract from the main data. Table body font is 10pt throughout. The topology diagram is also embedded as editable native Word shapes in the Word document's Environment section. Install `matplotlib` as a fallback for the Word topology if native shapes fail: `pip install matplotlib`.

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
| `--ca-bundle PATH` | *(system bundle)* | Path to a PEM-format CA bundle (typically `.pem`; extension doesn't matter — file must be PEM-encoded). Use for self-signed cluster certs or corporate HTTPS-inspection proxies (e.g. Zscaler) |
| `--insecure` | off | Disable TLS certificate validation entirely — prints a startup warning; use only on isolated/trusted networks |

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

# Corporate HTTPS-inspection proxy (e.g. Zscaler) or self-signed cluster cert
python3 health_check_report.py --apikey abc123 --ca-bundle /path/to/ca.pem

# Direct cluster with self-signed certificate
python3 health_check_report.py --cluster-host 10.1.2.3 --username admin --ca-bundle /path/to/ca.pem
```

### TLS certificate validation

TLS verification is **on by default** as of v1.53. All HTTPS connections validate against the system CA bundle.

| Scenario | How to run |
|----------|-----------|
| Helios (public cloud) | No extra flags — `helios.cohesity.com` uses a publicly-trusted cert |
| Direct cluster with trusted cert | No extra flags |
| Direct cluster with self-signed cert | `--ca-bundle /path/to/cluster-ca.pem` |
| Corporate HTTPS proxy (e.g. Zscaler) | `--ca-bundle /path/to/proxy-ca.pem` |
| Isolated lab / legacy environment | `--insecure` *(prints a warning)* |

To get the cluster CA certificate:
```bash
openssl s_client -connect <cluster-host>:443 -showcerts </dev/null 2>/dev/null \
  | awk '/BEGIN CERTIFICATE/,/END CERTIFICATE/' > cluster-ca.pem
```

Validate your bundle before running:
```bash
curl --cacert cluster-ca.pem https://<cluster-host>
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
| `GET /irisservices/api/v1/public/statistics/timeSeriesStats` | 30-day daily capacity time series (kMorphedUsageBytes) for runway regression |
| `GET /irisservices/api/v1/public/auditLog` | Last 30 days of configuration change events |
| `GET /v2/data-protect/recoveries` | Recovery task history — status, duration, type, objects recovered (Recovery Audit sheet) |
| `GET /irisservices/api/v1/public/kmsConfig` | KMS / encryption key management configuration — type, server, status (Security sheet) |
| `GET /irisservices/api/v1/public/roles` | All role definitions (built-in and custom) — privilege sets and risk analysis (User Security sheet) |
| `GET /v2/data-protect/snapshots` | Per-group snapshot DataLock verification — up to 5 snapshots × 20 groups per cluster (DataLock Verification sheet) |
| `GET /irisservices/api/v1/public/activeDirectory` | Active Directory domain list with connection status and preferred DCs (AD / Identity Health sheet) |
| `GET /irisservices/api/v1/public/ldapProvider` | LDAP provider list with status, server, port, base DN (AD / Identity Health sheet) |
| `GET /irisservices/api/v1/public/certificates` | Full TLS certificate inventory — all cert types with expiry dates (Certificate Inventory sheet) |
| `GET /irisservices/api/v1/public/statistics/timeSeriesStats` (kLogicalCapacityBytes) | 30-day logical capacity time series for dedup ratio computation (Capacity & Perf Trends sheet) |
| `GET /irisservices/api/v1/public/statistics/timeSeriesStats` (kNumBytesRead / kNumBytesWritten) | 30-day daily I/O throughput time series (Capacity & Perf Trends sheet) |

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
| 1.73 | 2026-04-23 | fix: About tab Command row — strip directory path from script name; mask `--apikey` and `--password` values with `xxxxx` so credentials are never written in plain text to the Excel output. |
| 1.72 | 2026-04-23 | feat: added **About tab** as the first tab in the Excel workbook. Sections: Title (version + run timestamp), Command (full CLI command), Parameters (mode, days, customer, cluster filter, clusters), Notes (Quick Mode scope, Storage Consumed cumulative caveat, DataLock Verification 20-group limit, Health Score formula), APIs Used (30 endpoints: 2 Helios MCM + 21 Cluster v1 + 7 Cluster v2). Tab count updated 30 → 31. |
| 1.71 | 2026-04-22 | fix: FortKnox Data Transfer sheet — removed Logical Transferred, Physical Transferred, Data Read, and Data Written columns. Sheet now shows: Cluster, Vault Name, Vault Type, Protection Group, Storage Consumed (TB), Period, Last FK Archival, Days Since FK Transfer. |
| 1.70 | 2026-04-22 | fix: clarify quick-mode CLI output — `Lookback` line now reads "N days (alerts, audit log, summaries)" so it's clear the window applies to non-run-history calls; `Mode` line explicitly states per-group run history is skipped and FortKnox Data Transfer covers last 1 day. |
| 1.69 | 2026-04-22 | fix: FortKnox Data Transfer sheet — data sources aligned with FortKnox vault report. Logical/Physical Transferred now from group run `archivalTargetResults` (not the absent `numLogicalBytesTransferred` dataTransferToVaults field). Added **Data Read (TB)** and **Data Written (TB)** from `localSnapshotStats`. Removed Snapshots column. Last FK Archival / Days Since now from full run history (was `lastRun` only — caused "Unknown" for groups whose most recent run was not an archival run). |
| 1.68 | 2026-04-22 | fix: six post-v1.67 bug fixes — Key Rotation displays correctly in Security tab; Recovery Audit populates Objects Recovered, Source Cluster, and Target columns; Vault Type column D populated in Replication & Archive tab; topology excludes inactive targets and uses active vault count; keyring `NoKeyringError` handled on headless Linux (Rocky Linux); auth failure message includes shell quoting hint. |
| 1.67 | 2026-04-22 | feat: new "Retention Health" tab (sheet 9) — cluster retention summary, per-policy detail, and P1/P2/P3 security recommendations covering short retention, no off-site copy, missing WORM, archive-shorter-than-local, and FortKnox gaps; `AMBER` color constant added (was referenced but undefined); sheet count 28 → 29. |
| 1.66 | 2026-04-22 | fix: prereq security hardening — `python-pptx` promoted to required (PPT generated on every run); `keyring` description updated to explain credential exposure risk (shell history / process lists); security note printed at startup when keyring absent; single `pip install` command covers all missing packages in one block. |
| 1.65 | 2026-04-21 | feat: AD/LDAP identity health sheet; full certificate inventory sheet with expiry RAG; Capacity & Perf Trends sheet (dedup ratio + I/O throughput charts); audit log 403 now surfaces actionable permission message; cert-expiry and AD/LDAP-disconnect recommendations added. Sheet count 25 → 28. |
| 1.64 | 2026-04-21 | feat: 4 new API-backed features — **Recovery Audit sheet** (`/v2/data-protect/recoveries`): per-recovery status/duration/type with days-since-last-recovery RAG; **DataLock Verification sheet** (`/v2/data-protect/snapshots`): snapshot-level VERIFIED/PARTIAL/NOT LOCKED per DataLock group; **KMS section** in Security sheet (`/v1/public/kmsConfig`): key management type and risk; **Custom Role Definitions** in User Security (`/v1/public/roles`): privilege analysis with cluster/user/security-modify flags. Sheet count 23 → 25. CRITICAL recommendation if DataLock policy set but snapshots not locked. HIGH if no recovery in lookback window. |
| 1.63 | 2026-04-21 | feat: 9 security & recoverability analytics using existing collected data — Risk Flag column (Stale Admin / Admin No MFA / Locked Admin) in User Security; Data Exposure sheet (NAS SMB/NFS/S3 risk); Security & Anomaly Alerts sub-section in Security sheet; Node SW Version Match column with MISMATCH flag; Cluster Fault Tolerance margin column; FortKnox Last Archival date + Days Since Transfer columns; Replication Lag (hrs) column; Unprotected Objects sheet with named object list; Avg Duration (min) column + duration trend chart in Trends. Sheet count 21 → 23. |
| 1.62 | 2026-04-17 | fix(security): Input validation and credential hygiene. `--days` bounded 1–3650; `--cluster-host` validated as hostname/IPv4; `--output` rejected if path contains `..` traversal; `--ca-bundle` existence check added. `--password` / `--mfa-code` print a process-list exposure warning when used on the CLI. Raw API response bodies (`r.text[:N]`) removed from auth error messages — HTTP status only — to prevent token/credential leakage to terminal history and CI logs. `cohesity_auth.py` v1.7 — same error-message scrub. |
| 1.61 | 2026-04-17 | feat(pptx): Topology — legend removed; FortKnox gap widened to 1.60" (was 0.80"); R/A/F columns vertically centred vs source cluster column; diagram uses full height (DIAG_BOT H−0.30"). |
| 1.60 | 2026-04-17 | fix: True content-driven column widths across all outputs. Excel: min_width 12→6, max_width 45→60, wrap_text removed. Word: new _word_fit_widths() measures all cells after population and sets proportional column widths (all 10 tables). PPT: _table() auto-computes column widths from content, overrides hardcoded col_w. |
| 1.59 | 2026-04-17 | feat: Excel 130% zoom + autofit + left-align; Word tables autofit; PPT copyright 2026, page numbers on data slides, table headers 10pt left-aligned, Contents bars dynamically sized, Protection Summary methodology panel, topology legend horizontal row (active columns only, 0.60" bottom margin). |
| 1.58 | 2026-04-17 | feat(pptx): Topology diagram wider spacing — T_NODE_W 2.50→2.30", T_NODE_H 0.80→0.85", T_V_GAP 0.18→0.32", T_COL_GAP 0.40→0.80" (double connector gap), connector line weight 0.025→0.030". |
| 1.57 | 2026-04-17 | fix(pptx): All data tables top-aligned at y=0.90" (0.13" below header bar); mixed slides top-align table with methodology immediately below — eliminates bottom whitespace. |
| 1.56 | 2026-04-17 | feat(pptx): KPI tiles — solid #D9D9D9 fill + accent side bar + outer shadow. All data tables vertically centred (`_tbl_y`). Mixed slides (Scorecard, Ransomware, Security, Risk) table+methodology block centred as unit (`_layout_mixed`). Topology diagram vertically centred with auto-scaling; all topology shapes grouped into a single PowerPoint group for easy user arrangement. |
| 1.55 | 2026-04-17 | feat(pptx): Environment Topology diagram now horizontally centred on the slide — column X positions computed dynamically from the active column set so left/right margins are always equal regardless of how many node types are present. |
| 1.54 | 2026-04-17 | feat(pptx): PowerPoint visual improvements — green header bar height increased to 0.77" on all slides, content rows shifted to y=0.82". Contents slide serial numbers without leading zero. Environment Snapshot KPI tiles use a two-stop linear gradient (light tint → accent, 120°) with white text. Added "Thank You" end slide matching cover/section-divider style. |
| 1.53 | 2026-04-17 | fix(security): TLS certificate validation hardened — removed all `verify=False`, added `--ca-bundle PATH` and `--insecure` flags, default is now `verify=True`. Excel formula injection protection via `_safe_cell()` applied to all API-sourced worksheet writes. |
| 1.52 | 2026-04-16 | feat: Inline scoring methodology panels on Scorecard and Security Posture PPT slides; removed 4 standalone Scoring Methodology reference slides. |
| 1.51 | 2026-04-16 | fix: Storage Capacity & Growth PPT slide now uses correct API field paths — was empty because it read non-existent `storage_domains` key. feat: DataLock RAG coloring (red/amber/green) added to Protection Summary and Ransomware Readiness PPT slides. feat: Two Scoring Methodology Reference slides appended to deck (Health Score + Security Score; Ransomware Score + Workload Risk Score + RAG legend). PPT grows from ~29 to ~31 slides. |
| 1.50 | 2026-04-15 | feat: Comprehensive PowerPoint deck (`write_pptx()`). ~29 slides across 4 sections: Cover, Executive Summary (Snapshot, Scorecard, Capacity, Protection, Ransomware, At-Risk Workloads, Priority Actions), Environment Topology (full diagram with legend), Security Deep-Dive (Posture, User Security, Audit, Recommendations), Backup Engineering (Lifecycle, Coverage Gaps, Agents, Source Coverage, Workload Risk, Recommendations). All slides RAG colour-coded. draw.io output removed entirely. |
| 1.49 | 2026-04-15 | perf: Parallel API data collection (ThreadPoolExecutor wave-1 + wave-2 + per-group runs), cached `fk_status` / `policy_by_id` / `audit_enabled` in `cd`, `end_usecs` replaces repeated `_now_usecs()` calls, merged dual agent/disk loops, isinstance filtering at source, `_FONT_NORMAL`/`_FONT_BOLD` singletons. No output changes. |
| 1.48 | 2026-04-13 | feat: Topology diagram overhaul — fixed missing FK vault connections (vaults list now always emits edges), draw.io arrow labels (Replication/Archive/FortKnox/Indelible), cluster nodes show software version + group/source counts, legend block added. New optional PowerPoint output (.pptx via python-pptx) — 16:9 slide with colored nodes, STRAIGHT connectors, arrow labels, legend. python-pptx added to optional prereq check table. |
| 1.47 | 2026-04-13 | feat: Guide tab added as the first tab in every generated workbook — covers all 22 tabs with descriptions, overall health scoring (weights + grade table), security score deductions, ransomware readiness scoring, workload risk score factors, recommendation priorities, and color/RAG legend |
| 1.46 | 2026-04-12 | feat: Ransomware Readiness Score (0–100) per cluster — scored on DataLock, FortKnox activity, recovery window, quorum, admin MFA; shown in Executive Summary, Security sheet, and Word security table (11 cols). feat: Predictive Capacity Runway — 30-day linear regression via `/v1/public/statistics/timeSeriesStats`; shown in Storage sheet, Executive Summary, and Word Storage section with forecast paragraph. feat: Workload Risk Heatmap (sheet 20) — all protection groups scored worst-first with RAG coloring; top 5 Critical/High in Word Executive Summary. feat: Audit Log sheet (sheet 21) — 30-day change summary by category with high-risk event highlighting; governance paragraph in Word Security section. Sheet count 19→21 |
| 1.45 | 2026-04-10 | fix: `get_auth_token()` v1 failure no longer silently swallowed — prints `[!] v1 endpoint failed (HTTP NNN) — trying v2...`. Final error shows both v1 and v2 responses plus troubleshooting hints: `--clear-credentials`, `--domain`, account locked. `cohesity_auth.py` v1.6 |
| 1.44 | 2026-04-10 | feat: Prereq check redesigned — always prints a full status table for all 5 packages (requests, openpyxl, python-docx, matplotlib, keyring) with installed/NOT FOUND and required/optional labels. Previously matplotlib and keyring only appeared as small NOTE lines when absent |
| 1.43 | 2026-04-10 | fix: `ModuleNotFoundError: No module named 'urllib3'` on Windows — `requests` now imported first with try/except; warning suppression falls back to `requests.packages.urllib3`. fix: `cohesity_auth` and `formatters` imports wrapped with inline fallbacks so script runs as a single downloaded file without the repo's `utils/` directory |
| 1.42 | 2026-04-10 | fix: Direct cluster auth — `get_auth_token()` tries v1 `/irisservices/api/v1/public/accessTokens` first (widest compat), falls back to v2. Fixes 400 "KValidationError Access denied" on some clusters. Direct mode `"mode"` key now propagated to `cd` dict for correct quorum/security handling. `cohesity_auth.py` v1.5 |
| 1.41 | 2026-04-10 | feat: Startup prerequisite check — validates required packages at launch, reports missing with exact pip commands. README.md rewritten with full prerequisite table, all CLI flags, updated Word/topology descriptions |
| 1.40 | 2026-04-10 | feat: Security Posture table (Word) — all 10 columns color-coded (red/amber/green). Environment Topology replaced with native Word DrawingML shapes: green (#96C73D) rounded rects for clusters, cloud for FortKnox, cylinders for archive, connector arrows. Cluster color updated globally. draw.io fonts + box sizes increased. matplotlib dark header bar removed |
| 1.39 | 2026-04-10 | fix: Executive Summary "30d Success %" now equals the average of Trends tab column H — `_success_stats()` rebuilt to use same source priority (`v1_runs` → `group_runs`) and same per-day formula as `_sheet_trends()` |
| 1.38 | 2026-04-10 | fix: Environment Topology table (Word) FortKnox Vaults column — replaced narrow `kFortKnox`/`kS3Compatible` vault-type check with `_fk_status(cd)` for accurate active/idle/none classification; vault name collection now uses the same comprehensive `_is_fk` logic (covers `krpaas`, `kfortknox`, `kfort_knox`, `krpaasarchival`, `"fortknox"` substring). Active = vault names on green; idle = vault names + `(Configured — not sending data)` on amber; none = `No` |
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
