# health_check_report.py

Multi-cluster Cohesity health check designed for enterprise customer business reviews (EBRs) and SE trusted-advisor engagements.  Gathers live data from Cohesity Helios and produces a **13-sheet Excel workbook** and a **Word document**.

---

## Output

### Excel Workbook (13 sheets)

| Sheet | Contents |
|-------|----------|
| Executive Summary | Per-cluster health score (0–100), grade, success %, capacity used %, open criticals, top finding |
| Infrastructure | Software version, node count, healthy nodes, encryption, FIPS, DNS, NTP, timezone |
| Protection Health | Per-group last-run status, SLA violations, RPO gap, object counts, logical/physical bytes |
| Storage & Capacity | Cluster and per-domain usable/used/free, data reduction ratio, dedup ratio, compression ratio |
| Policy Audit | Retention schedules, replication targets, archival targets, compliance gaps per policy, group count |
| Policy → Groups | Every protection group with the policy that governs it; sorted by policy then group name |
| Alerts | All open alerts sorted by severity, with age, description, and entity |
| Security | Six-control checklist per cluster: encryption, FIPS, vault, FortKnox, replication, archival |
| Replication & Archive | Replication targets, vault names and types, FortKnox storage consumed (TB) |
| Data Services | NAS views with protocol, quota, usage %, near-quota warnings |
| Coverage Gaps | Protection groups with failed last run, paused state, or RPO gap > threshold |
| Trends (30d) | Daily backup success rate aggregated per cluster with embedded line chart |
| Recommendations | Prioritised action list (CRITICAL / HIGH / MEDIUM / LOW) with business impact |

### Word Document

Narrative report suitable for printing or sharing with customers:

- Cover page (customer name, date, confidential notice)
- Executive Summary with per-cluster scorecard table
- Environment Overview (infrastructure table)
- Protection & Recovery Health
- Storage & Capacity
- Security Posture
- Prioritised Recommendations table
- Appendix: scoring methodology and thresholds

---

## Usage

```bash
python health_check_report.py --apikey <key> [options]
```

### Options

| Flag | Default | Description |
|------|---------|-------------|
| `--apikey KEY` | env / keychain | Helios API key |
| `--cluster NAME` | all | Limit to one cluster (partial name match) |
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
python health_check_report.py --apikey abc123 --customer "Acme Corp"

# Quick scan of a single cluster
python health_check_report.py --apikey abc123 --cluster prod-east --quick

# Excel only, custom output path and filename
python health_check_report.py --apikey abc123 --output /reports/q2_review --excel-only

# 90-day lookback with debug logging
python health_check_report.py --apikey abc123 --days 90 --debug
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
| Storage capacity | 15% | Penalised above 70% used; critical above 85% |
| Security posture | 15% | 20 pts each: encryption, FIPS, vault, replication, immutability |
| Alert health | 10% | −10 pts per open critical alert, −2 pts per warning |

---

## Recommendations

The engine evaluates all collected data and generates a prioritised finding list:

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
| `GET /irisservices/api/v1/public/reports/dataTransferToVaults` | FortKnox transfer data |

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
