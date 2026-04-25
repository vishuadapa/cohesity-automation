# Cohesity Automation Scripts

Personal collection of Cohesity REST API automation scripts for reporting, operations, and customer use cases. Built and maintained by Vishu Adapa.

Targets Cohesity 7.1 and 7.3. Uses the [Cohesity REST API](https://developer.cohesity.com/) (v1 and v2 endpoints as appropriate).
Inspired by [Brian Seltzer's scripts](https://github.com/bseltz-cohesity/scripts).

---

## Health Check Report ★

> Multi-cluster Cohesity health check designed for enterprise customer business reviews (CBRs) and SE engagements. Gathers live data from Cohesity Helios and produces a **31-tab Excel workbook** (About tab + Guide tab + 29 data sheets), a **Word document**, and a **comprehensive ~27-slide PowerPoint deck**. The About tab Command row automatically strips the script directory path and masks `--apikey`, `--password`, and `--ca-bundle` values with `xxxxx`.

### What it produces

| Output | Description |
|--------|-------------|
| Excel workbook | **31 tabs** — an **About tab** (first tab, with run metadata, parameters, notes, and full API reference) + a **Guide tab** (sheet descriptions, scoring methodology, and color legend) + **29 data sheets** covering infrastructure (**fault tolerance margin**), node hardware with EOL status and **SW version skew flag**, **disk health with SSD wear %**, protection health, storage capacity with **predictive runway forecast**, policy audit (with DataLock/FortKnox status), **Retention Health** (cluster retention summary, per-policy detail, P1/P2/P3 security recommendations), **DataLock Verification** (snapshot-level WORM proof), alerts, **expanded security checklist** (audit log, MFA, NTP auth, remote tunnel, SSO/IDP, TLS cert expiry, quorum, **KMS type & risk**, **security & anomaly alert counts**), **agent health**, **source coverage**, **Recovery Audit** (recoverability proof — days since last test), **Unprotected Objects** (named object list), FortKnox/replication with **replication lag hours**, **vault type**, and **last FK transfer date**, **Data Exposure** (NAS SMB/NFS/S3 risk), data services, coverage gaps, **user security** (MFA/locked/roles, **Risk Flag**, **custom role privilege audit**), 30-day trends with **duration trend chart**, recommendations, **per-workload risk heatmap**, **30-day audit log**, **AD & Identity Health** (AD/LDAP connection status), **Certificate Inventory** (TLS cert expiry RAG), and **Capacity & Perf Trends** (dedup ratio trend + I/O throughput) |
| Word document | Narrative report with cover page, executive scorecard with **top at-risk workloads table**, environment overview, software & hardware lifecycle, protection health, storage with **capacity runway forecast**, **fully color-coded security posture table** (11 columns including **ransomware readiness score**, red/amber/green per finding) with **governance & change activity summary**, **agent health & source coverage**, **user security table**, **native Word topology diagram** (editable shapes + connector arrows), recommendations, and scoring methodology appendix — ready to share with customers |
| PowerPoint deck | **~27-slide `.pptx` file** (requires `python-pptx`) — comprehensive presentation across 5 native PowerPoint sections: **Cover**, **Contents** (hyperlinked TOC), **Executive Summary** (KPI snapshot, health scorecard, capacity, protection, ransomware readiness, top at-risk workloads, priority actions), **Environment Topology** (cluster/replication/archival/FortKnox diagram with connector labels and legend), **Security Deep-Dive** (posture overview, user security & MFA, audit activity, security recommendations), **Backup Engineering** (SW/HW lifecycle, coverage gaps, agent health, source coverage, workload risk heatmap, engineering recommendations). Scoring methodology tables embedded inline at the bottom of each data slide. All data slides RAG colour-coded; 10pt body font. Topology also embedded as editable native shapes in the Word document. |

### Sheet reference (29 data sheets + About + Guide = 31 tabs total)

| # | Sheet | Key content |
|---|-------|-------------|
| — | **About** | First tab — full CLI command used, run parameters (mode, days, customer, clusters), notes (Quick Mode, Storage Consumed, DataLock limit, Health Score), and reference table of all 30 API endpoints |
| — | **Guide** | Sheet descriptions, health/security/ransomware/workload scoring methodology, recommendation priorities, color legend |
| 1 | Executive Summary | Per-cluster health score (0–100), grade, success %, capacity used %, open criticals, ransomware readiness score, capacity runway, top finding |
| 2 | Infrastructure | Software version, node count, healthy nodes, cluster encryption, SD encryption, DNS, NTP, timezone, SW lifecycle status, fault tolerance margin |
| 3 | Node Hardware | Per-node model, serial, node type, software version, raw capacity, disk count, storage tiers, HW EOL date, SW version match flag |
| 4 | Disk Health | Per-disk status, type, model, serial, SSD wear %, capacity, encryption, storage tier — CRITICAL on failed/missing, HIGH on wear ≥80% |
| 5 | Protection Health | Per-group last-run status, SLA violations, RPO gap, object counts, logical/physical bytes, DataLock (WORM) per group |
| 6 | Storage & Capacity | Cluster and per-domain usable/used/free, data reduction ratio, dedup ratio, compression ratio, predictive runway to 80% utilization |
| 7 | Policy Audit | Retention schedules, replication targets, archival targets, DataLock (WORM) mode + duration per policy; FortKnox/RPaaS policies flagged "FortKnox (Indelible)" |
| 8 | DataLock Verification | Snapshot-level DataLock check — VERIFIED / PARTIAL / NOT LOCKED per group; CRITICAL recommendation if policy says locked but snapshots are not |
| 9 | Policy → Groups | Every protection group with the policy that governs it; sorted by policy then group name |
| 10 | Retention Health | Three-section retention analysis: cluster summary (min/max/avg local, archive, replication retention + DataLock + FortKnox + Risk Level); per-policy detail with flags; P1/P2/P3 security recommendations |
| 11 | Alerts | All open alerts sorted by severity, with age, description, and entity |
| 12 | Security | Expanded checklist: encryption, FIPS, audit log, MFA, NTP, quorum, TLS, ransomware score, anomaly alerts, KMS/encryption key management |
| 13 | Agent Health | Per-host agent version, health status, upgradability, cert expiry, last upgrade error |
| 14 | Source Coverage | Registered sources with protected/unprotected object counts and coverage % |
| 15 | Recovery Audit | Per-recovery rows (status, duration, type, objects); cluster summary with days-since-last-recovery; HIGH recommendation if no recovery found |
| 16 | Unprotected Objects | Named list of unprotected objects per source — red if count >10, amber if count >0 |
| 17 | Replication & Archive | Replication targets, vault names and types, FortKnox storage consumed (TB), replication lag (hrs) |
| 18 | FortKnox Data Transfer | Per-protection-group transfer to every external vault: storage consumed TB (cumulative), last FK archival date, days since FK transfer |
| 19 | Data Exposure | NAS view exposure analysis — SMB discovery, NFS open mount, S3 access, quota presence; per-view Risk Level (HIGH/MEDIUM/LOW) |
| 20 | Data Services | NAS views with protocol, quota, usage %, near-quota warnings |
| 21 | Coverage Gaps | Protection groups with failed last run, paused state, or RPO gap > threshold |
| 22 | User Security | Per-user MFA status, locked state, roles, last login, risk flag (Stale Admin / Admin No MFA / Locked Admin); custom role definitions sub-section |
| 23 | Trends (30d) | Daily backup success rate and avg duration; dual line charts for success rate and duration trend |
| 24 | Recommendations | Prioritized action list (CRITICAL / HIGH / MEDIUM / LOW) with business impact |
| 25 | Workload Risk Heatmap | All protection groups scored 0–100 by recovery risk (last-run 35pts, SLA 25pts, RPO 25pts, DataLock 15pts); sorted worst-first with RAG coloring |
| 26 | Audit Log | Last 30 days of configuration changes from the cluster audit trail; categorized by type; high-risk events highlighted in red; shows actionable permission message on HTTP 403 |
| 27 | AD & Identity Health | Active Directory connection status with preferred DC; LDAP provider status, server, port, base DN — disconnected sources flagged red |
| 28 | Certificate Inventory | Full TLS certificate inventory with expiry dates and days remaining — CRITICAL (<30d), HIGH (<90d), OK |
| 29 | Capacity & Perf Trends | 30-day dedup ratio trend (ransomware indicator) and daily I/O throughput (read/write GB) with embedded charts |

### Prerequisites

**Python 3.7** or later.

Install all packages before running:

```bash
pip install requests openpyxl python-docx matplotlib keyring python-pptx
```

Full package list:

| Package | Required? | Purpose | Install |
|---------|-----------|---------|---------|
| `requests` | **Yes** | HTTP client for Cohesity REST API calls | `pip install requests` |
| `openpyxl` | **Yes** (unless `--word-only`) | Excel workbook generation (31 tabs: About + Guide + 29 data sheets) | `pip install openpyxl` |
| `python-docx` | **Yes** (unless `--excel-only`) | Word document generation (narrative report + topology shapes) | `pip install python-docx` |
| `matplotlib` | Optional | Topology PNG fallback if native Word shapes fail | `pip install matplotlib` |
| `python-pptx` | Optional | Comprehensive ~27-slide PowerPoint deck (.pptx) | `pip install python-pptx` |
| `keyring` | Optional | Securely store API key / passwords in OS keychain between runs | `pip install keyring` |

At startup the script prints a status table for **all six packages** — installed vs. NOT FOUND, required vs. optional. If a required package is missing, it exits with a clear `pip install` command. Optional packages that are absent are flagged but do not block execution.

> **Standalone / single-file use**: `health_check_report.py` can be downloaded and run as a single file — all auth and formatting helpers are bundled inside it as fallbacks. The only external requirement is `pip install requests` (plus `openpyxl` / `python-docx` for report output).

### Where to run from

Run the script from the `health_check/` directory, from the repo root, or as a single standalone file — the working directory does not matter:

```bash
cd health_check
python3 health_check_report.py --apikey <your-helios-api-key> --customer "Acme Corp"
```

Or from the repo root:

```bash
python3 health_check/health_check_report.py --apikey <your-helios-api-key> --customer "Acme Corp"
```

### Helios API key

Generate your key in Helios → Settings → Access Management → API Keys. The key can be passed via `--apikey` or set as the environment variable `HELIOS_API_KEY`.

### Common usage

```bash
# Full report — all clusters, 30-day window, Excel + Word + topology diagram
python3 health_check_report.py --apikey abc123 --customer "Acme Corp"

# Target a single cluster
python3 health_check_report.py --apikey abc123 --cluster prod-east --customer "Acme Corp"

# Quick scan (uses last-run only — no per-group run history, much faster)
python3 health_check_report.py --apikey abc123 --quick

# Excel only, custom output path
python3 health_check_report.py --apikey abc123 --output /reports/q2_review --excel-only

# Extended lookback with debug logging
python3 health_check_report.py --apikey abc123 --days 90 --debug

# With corporate HTTPS proxy CA cert (e.g. Zscaler)
python3 health_check_report.py --apikey abc123 --ca-bundle ~/.cohesity/zscaler-ca.pem

# Direct cluster with self-signed certificate
python3 health_check_report.py --cluster-host 10.1.2.3 --username admin --ca-bundle ./cluster-ca.pem
```

### All options

| Flag | Default | Description |
|------|---------|-------------|
| **Helios auth** | | |
| `--apikey KEY` | env `HELIOS_API_KEY` | Helios API key (stored in keychain after first use) |
| `--cluster NAME` | all | Limit to one cluster (partial name match) |
| **Direct cluster auth** | | |
| `--cluster-host HOST` | *(none)* | Bypass Helios — connect directly to a cluster IP/hostname |
| `--username USER` | *(none)* | Cluster username (required with `--cluster-host`) |
| `--password PASS` | *(prompted)* | Password (prompted securely if omitted; saved to keychain) |
| `--domain DOMAIN` | `LOCAL` | Auth domain — `LOCAL` or AD domain name |
| `--mfa-code CODE` | *(none)* | TOTP code for MFA-enabled cluster accounts |
| **Report options** | | |
| `--days N` | 30 | Lookback window for alerts and run history |
| `--customer NAME` | *(blank)* | Customer name on Word cover page and output filename |
| `--output PATH` | auto-generated | Base path for output files (no extension added) |
| `--quick` | off | Use last-run only; skip per-group run history |
| `--excel-only` | off | Skip Word document generation |
| `--word-only` | off | Skip Excel generation |
| `--debug` | off | Print HTTP status for every API call |
| `--clear-credentials` | off | Remove stored credentials from OS keychain and exit |
| **TLS options** | | |
| `--ca-bundle PATH` | *(system bundle)* | CA certificate `.pem` for self-signed cluster certs or corporate HTTPS proxies (e.g. Zscaler) |
| `--insecure` | off | Disable TLS validation — prints a startup warning; avoid in production |

### Output files

Auto-generated filenames include the version number and timestamp for easy tracking:

```
cohesity_health_check_v1.71_AcmeCorp_20260422_1430.xlsx
cohesity_health_check_v1.71_AcmeCorp_20260422_1430.docx
cohesity_health_check_v1.71_AcmeCorp_20260422_1430.pptx
```

The `.pptx` file is a comprehensive ~27-slide deck ready to present to customers. Requires `pip install python-pptx`. Sections: Cover · Contents (hyperlinked TOC) · Executive Summary · Environment Topology · Security Deep-Dive · Backup Engineering — all with RAG colour-coding and inline scoring methodology tables.

### Security scoring

Each cluster receives a score out of 100. Key deductions include:

| Finding | Severity | Points |
|---------|----------|--------|
| No cluster encryption | HIGH | −20 |
| No storage domain encryption | HIGH | −15 |
| Admin accounts without MFA | HIGH | −10 |
| No DataLock (WORM) on any policies | HIGH | −10 |
| No FortKnox vault | MEDIUM | −5 |
| No replication | MEDIUM | −5 |
| Quorum not enabled | MEDIUM | −5 |
| Audit logging disabled | MEDIUM | −5 |
| TLS certificate expired/near expiry | MEDIUM | −5 |

See [health_check/health_check_report.md](health_check/health_check_report.md) for the full sheet reference, scoring model, and version history.

---

## FortKnox Vault Report

> Per-protection-group FortKnox (RPaaS) vault reporting script. Connects to Helios, iterates every connected cluster, and produces a **multi-tab Excel workbook** with a data report, trend charts, and a built-in About tab explaining every field and API used.

### What it produces

| Output | Description |
|--------|-------------|
| **About tab** | Command run, report parameters, full field reference (name, description, source field path, API endpoint), APIs used table, and a Notes section explaining local vs. vault activity fields |
| **Report tab** | One row per protection group per vault. Source-banner row groups columns by originating API. Columns frozen at row 3. |
| **Chart tabs** | One line-chart sheet per cluster (trend mode only) showing Remote Storage Consumed (TB) over time, one series per protection group |

### Report columns

All Bytes values are grouped together, followed by TB equivalents in the same order:

| Column | Source | Notes |
|--------|--------|-------|
| Period Start / Period End | Computed | Date range covered by this row |
| Cluster | Helios MCM | Cluster name |
| Vault Name | dataTransferToVaults | FortKnox vault name |
| Vault Type | dataTransferToVaults | e.g. kFortKnox, kRPaaS |
| Protection Group | dataTransferToVaults | Protection job name |
| **Remote Storage Consumed (Bytes)** | dataTransferToVaults | Total retained bytes in the vault — cumulative, not windowed |
| **Activities: Data Read (Bytes)** | Protection Activities | Data read from source during Backup activity — sent to **local cluster** |
| **Activities: Data Written (Bytes)** | Protection Activities | Data written to local storage during Backup activity — sent to **local cluster** |
| **Activities: Logical Transferred (Bytes)** | Protection Activities | Logical bytes sent to the **remote FortKnox vault** (Vault activity) |
| **Activities: Physical Transferred (Bytes)** | Protection Activities | Physical (post-dedup/compress) bytes sent to the **remote FortKnox vault** |
| Remote Storage Consumed (TB) | Computed | ÷ 1,000,000,000,000 |
| Activities: Data Read (TB) | Computed | ÷ 1,000,000,000,000 |
| Activities: Data Written (TB) | Computed | ÷ 1,000,000,000,000 |
| Activities: Logical Transferred (TB) | Computed | ÷ 1,000,000,000,000 |
| Activities: Physical Transferred (TB) | Computed | ÷ 1,000,000,000,000 |

> **Note**: Data Read / Data Written reflect the Protection Activities report **Activity Type = Backup** (local cluster). Logical / Physical Transferred reflect **Activity Type = Vault** (remote FortKnox). This is equivalent to the **Protection Activities Report** on Helios, filtered to cloud vault (FortKnox).

### Field value types — cumulative vs. point-in-time

| Column | Value type | Explanation |
|--------|-----------|-------------|
| **Remote Storage Consumed** | **Cumulative** | Total retained bytes across all snapshots ever stored in the vault for this protection group. The `dataTransferToVaults` API ignores your `--start`/`--end`/`--days` window for this field — the value is always the current vault footprint. In trend mode every day row shows the same value. |
| **Activities: Data Read** | **Point-in-time** | Per-run stat from the runs API, filtered to runs that started within your queried window. In summary mode it sums all qualifying runs across the full range; in trend mode it shows only runs that started on that specific day. |
| **Activities: Data Written** | **Point-in-time** | Same filtering as Data Read. |
| **Activities: Logical Transferred** | **Point-in-time** | Same filtering — only vault archival runs that started within your window. |
| **Activities: Physical Transferred** | **Point-in-time** | Same filtering as Logical Transferred. |

**Practical implications**

- Do not use **Remote Storage Consumed** to measure ingest activity or growth within a period — it is always the current total and will be the same value regardless of the date range you query.
- Use **Logical / Physical Transferred** to understand how much data was sent to the vault during a specific window.
- Use **Data Read / Data Written** to understand local backup workload (how much data was read from source and written to the local cluster) during a specific window.
- In **trend mode**, plotting Logical/Physical Transferred per day gives a true daily vault ingest view. Plotting Remote Storage Consumed per day shows the running vault footprint (all values will be identical within a single report run).

### APIs used

| API | Endpoint | Purpose |
|-----|----------|---------|
| Helios MCM | `GET /mcm/clusters/connectionStatus` | List all connected clusters |
| Cluster v1 | `GET /public/cluster` | Detect cluster timezone |
| Cluster v1 | `GET /public/vaults?includeFortKnoxVault=true` | Fetch FortKnox vault IDs and names |
| Cluster v1 | `GET /public/reports/dataTransferToVaults` | Storage consumed per protection group (cumulative) |
| Cluster v2 | `GET /v2/data-protect/protection-groups` | List protection groups to resolve group IDs |
| Cluster v2 | `GET /v2/data-protect/protection-groups/{id}/runs` | Per-run backup and archival stats for Activities columns |

All cluster API calls are routed via the Helios proxy using the `accessClusterId` header — a single Helios API key is all that is required.

### Prerequisites

```bash
pip install requests openpyxl keyring
```

| Package | Required? | Purpose |
|---------|-----------|---------|
| `requests` | **Yes** | HTTP client |
| `openpyxl` | **Yes** | Excel workbook output |
| `keyring` | Optional | Store API key in OS keychain between runs |

### Usage

```bash
# First run — prompts for API key, saves to keychain
python3 reporting/fortknox_vault_report.py

# Summary report — last 30 days, all clusters
python3 reporting/fortknox_vault_report.py --days 30

# Trend mode — one row per day per group, with charts
python3 reporting/fortknox_vault_report.py --mode trend --days 30

# Specific date range, single cluster
python3 reporting/fortknox_vault_report.py --start 2026-03-01 --end 2026-03-31 --cluster prod-east

# Corporate proxy with custom CA cert
python3 reporting/fortknox_vault_report.py --ca-bundle /path/to/ca.pem

# Remove stored API key
python3 reporting/fortknox_vault_report.py --clear-credentials
```

### All options

| Flag | Default | Description |
|------|---------|-------------|
| `--apikey KEY` | keychain / prompted | Helios API key |
| `--clear-credentials` | off | Remove stored key from OS keychain and exit |
| `--cluster NAME` | all | Filter to one cluster (partial name match) |
| `--vault NAME` | all | Filter to one vault (partial name match) |
| `--mode` | `summary` | `summary`: one row per group for the full range. `trend`: one row per group per day with charts |
| `--days N` | 7 | Last N days |
| `--start YYYY-MM-DD` | — | Start date (inclusive) |
| `--end YYYY-MM-DD` | today | End date (inclusive) |
| `--start-msecs MS` | — | Exact start timestamp in milliseconds |
| `--end-msecs MS` | — | Exact end timestamp in milliseconds |
| `--output PATH` | auto-generated | Output `.xlsx` path |
| `--ca-bundle PATH` | system bundle | CA cert for corporate proxy or self-signed cluster |
| `--insecure` | off | Disable TLS verification (prints warning) |
| `--debug` | off | Print raw API response samples and request URLs |

### Output filename

Auto-generated with version and timestamp:

```
fortknox_vault_report_v4.22_20260421_143000.xlsx
```

---

## Repository Structure

| Folder | Purpose |
|--------|---------|
| [health_check/](health_check/) | **Multi-cluster health check** — 30-tab Excel workbook (Guide + 29 data sheets) + Word report + ~27-slide PowerPoint deck for CBRs and SE engagements |
| [reporting/](reporting/) | Protection group run reports and FortKnox vault data transfer reports |
| [protection/](protection/) | Protection group management — run, pause, clone jobs |
| [recovery/](recovery/) | Restore automation — VMs, files, file/folder recovery |
| [storage/](storage/) | Capacity planning, views, storage domain reporting |
| [policies/](policies/) | Policy audit, SLA compliance, and cloning |
| [alerts/](alerts/) | Alert queries, summaries, bulk resolve, and CSV export |
| [infrastructure/](infrastructure/) | Cluster health, node status, version inventory, upgrade readiness |
| [utils/](utils/) | Shared auth and helper modules |

---

## Protection Group Report

> Per-protection-group run status report. Connects to Helios or a direct cluster and produces an Excel workbook with run stats, SLA, replication, archival, and storage data. Supports summary (last run per group), historical (all runs in a date range), and trend (daily storage consumed per group with charts) modes.

### Prerequisites

```bash
pip install requests openpyxl keyring
```

### Usage

```bash
# ── Bearer token (Helios) ────────────────────────────────────────────────────
# First run — prompts once (hidden input), saves to OS keychain as 'helios'
python3 reporting/protection_group_report.py --bearer

# Subsequent runs — token retrieved silently, no prompt
python3 reporting/protection_group_report.py --bearer --days 30
python3 reporting/protection_group_report.py --bearer --start 2026-03-01 --end 2026-04-01
python3 reporting/protection_group_report.py --bearer --mode trend --days 14

# Clear stored Helios bearer token (e.g. after token expiry)
python3 reporting/protection_group_report.py --clear-credentials --bearer

# ── Bearer token (direct cluster) ───────────────────────────────────────────
# Token stored per-cluster — each cluster name/IP has its own keychain entry
python3 reporting/protection_group_report.py --bearer --cluster 10.1.2.100
python3 reporting/protection_group_report.py --bearer --cluster 10.1.2.100 --days 7

# Clear stored token for a specific cluster
python3 reporting/protection_group_report.py --clear-credentials --bearer --cluster 10.1.2.100

# ── Helios API key ────────────────────────────────────────────────────────────
# First run — prompts for API key, saves to keychain
python3 reporting/protection_group_report.py

# With explicit key (not saved)
python3 reporting/protection_group_report.py --apikey <key> --days 30

# ── Helios username/password → token exchange ────────────────────────────────
python3 reporting/protection_group_report.py --helios-user user@example.com
python3 reporting/protection_group_report.py --helios-user user@example.com --days 30 --cluster my-cluster

# ── Direct cluster (username/password) ──────────────────────────────────────
python3 reporting/protection_group_report.py --cluster 10.1.2.100 --username admin --domain LOCAL
```

### All options

| Flag | Default | Description |
|------|---------|-------------|
| **Bearer token auth** | | |
| `--bearer` | off | Pre-existing bearer token mode. Prompts once via hidden input; token stored in OS keychain keyed by cluster IP/name (or `'helios'`). Reused silently on future runs. |
| **Helios auth (API key)** | | |
| `--apikey KEY` | keychain / prompted | Helios API key |
| **Helios auth (username/password)** | | |
| `--helios-user EMAIL` | *(none)* | Helios email — exchanges credentials for a Bearer token via `POST /mcm/createAccessToken` |
| `--helios-password PASS` | keychain / prompted | Helios password (optional; keychain used if omitted) |
| `--helios-domain DOMAIN` | `cohesity.com` | Helios auth domain |
| **Direct cluster auth** | | |
| `--cluster HOST` | *(none)* | Cluster IP/hostname (direct mode) or name filter (Helios mode) |
| `--username USER` | `admin` | Cluster username |
| `--domain DOMAIN` | `LOCAL` | Auth domain — `LOCAL` or AD domain name |
| **Report options** | | |
| `--mode` | `summary` | `summary`: last run per group or all runs in range. `trend`: daily storage per group with charts. |
| `--days N` | *(none)* | Last N days |
| `--start YYYY-MM-DD` | — | Start date (inclusive) |
| `--end YYYY-MM-DD` | today | End date (inclusive) |
| `--output PATH` | auto-generated | Output `.xlsx` path |
| `--clear-credentials` | off | Remove stored credentials from OS keychain and exit |
| **TLS options** | | |
| `--ca-bundle PATH` | system bundle | CA cert for corporate proxy or self-signed cluster |
| `--insecure` | off | Disable TLS verification (prints warning) |
| `--debug` | off | Print raw API responses |

---

## Authentication

Scripts support four authentication modes:

**Bearer token — pre-existing token (recommended for automation):**
```bash
# Helios — token stored as 'helios' in OS keychain
python3 reporting/protection_group_report.py --bearer

# Direct cluster — token stored keyed by cluster IP/name
python3 reporting/protection_group_report.py --bearer --cluster <ip>
```
First run prompts for the token (hidden input, never echoed). Every subsequent run retrieves it silently from the OS keychain. Use `--clear-credentials --bearer` to remove a stored token.

**Helios API key:**
```bash
python3 <script>.py --apikey <your-helios-api-key>
```
Key is stored in the OS keychain after the first use — `--apikey` is optional on subsequent runs.

**Helios username/password → token exchange:**
```bash
python3 reporting/protection_group_report.py --helios-user user@example.com
```
Exchanges credentials for a Bearer token via `POST /irisservices/api/v1/public/mcm/createAccessToken`. Password stored in OS keychain.

**Direct cluster — username/password:**
```bash
python3 <script>.py --cluster <ip-or-hostname> --username admin --domain LOCAL
```

---

## Requirements

All scripts require `requests`. The health check report has additional dependencies. Install everything:

```bash
pip3 install requests openpyxl python-docx matplotlib keyring python-pptx
```

| Package | Scripts | Required? |
|---------|---------|-----------|
| `requests` | All scripts | **Yes** — core HTTP client |
| `openpyxl` | health_check | **Yes** — Excel workbook output |
| `python-docx` | health_check | **Yes** — Word document output (skip with `--excel-only`) |
| `matplotlib` | health_check | Optional — topology PNG fallback |
| `python-pptx` | health_check | Optional — comprehensive ~31-slide PowerPoint deck (.pptx) |
| `keyring` | All scripts | Optional — OS keychain credential storage |

The health check script validates all dependencies at startup and tells you exactly what to install if anything is missing.

---

## Quick Reference

### Health Check
```bash
python3 health_check/health_check_report.py --apikey <key> --customer "Customer Name"
python3 health_check/health_check_report.py --apikey <key> --quick --excel-only
```

### Reporting
```bash
# Protection group report — bearer token (Helios)
python3 reporting/protection_group_report.py --bearer
python3 reporting/protection_group_report.py --bearer --days 30
python3 reporting/protection_group_report.py --bearer --mode trend --days 14

# Protection group report — bearer token (direct cluster)
python3 reporting/protection_group_report.py --bearer --cluster <ip>

# Protection group report — Helios API key
python3 reporting/protection_group_report.py --days 7
python3 reporting/protection_group_report.py --apikey <key> --start 2026-03-01 --end 2026-04-01

# FortKnox vault report
python3 reporting/fortknox_vault_report.py --days 30
python3 reporting/fortknox_vault_report.py --mode trend --days 30
python3 reporting/fortknox_vault_report.py --start 2026-03-01 --end 2026-03-31 --cluster prod-east
```

### Alerts
```bash
python3 alerts/alert_summary_report.py --days 1
python3 alerts/resolve_alerts.py --severity kCritical --cluster prod --yes
python3 alerts/alert_to_csv.py --days 30
```

### Infrastructure
```bash
python3 infrastructure/cluster_info_report.py
python3 infrastructure/node_status.py --unhealthy-only
python3 infrastructure/upgrade_readiness.py --min-version 7.1
```

### Policies
```bash
python3 policies/list_policies.py
python3 policies/policy_compliance_report.py --min-retention 14 --require-replication
python3 policies/clone_policy.py --source "Gold 30-Day" --name "Gold 60-Day" --cluster prod
```

### Protection
```bash
python3 protection/run_protection_group.py --group "VMware Daily" --cluster prod
python3 protection/pause_resume_groups.py --action pause --pattern "VMware" --cluster prod
python3 protection/clone_protection_group.py --source "SQL Prod" --name "SQL DR" --cluster prod
```

### Recovery
```bash
python3 recovery/list_snapshots.py --object "vm-prod-01" --cluster prod
python3 recovery/restore_vm.py --vm "vm-prod-01" --cluster prod --suffix "-restored"
python3 recovery/recover_files.py --source "linux-host" --paths /var/log --cluster prod
```

### Storage
```bash
python3 storage/capacity_report.py
python3 storage/list_views.py --protocol nfs
python3 storage/storage_domain_report.py
```

---

## API Versions

- **v1 API** — Used for infrastructure, alerts, storage, and other operations for better Helios reliability
- **v2 API** — Used for protection group reporting, policies, protection operations, and recovery

Individual `README.md` files in each folder document which API version is used by those scripts.

---

## Disclaimer

These scripts are personal tools, not official Cohesity products. Use at your own risk in production environments.
