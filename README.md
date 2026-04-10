# Cohesity Automation Scripts

Personal collection of Cohesity REST API automation scripts for reporting, operations, and customer use cases. Built and maintained by Vishu Adapa.

Targets Cohesity 7.1 and 7.3. Uses the [Cohesity REST API](https://developer.cohesity.com/) (v1 and v2 endpoints as appropriate).
Inspired by [Brian Seltzer's scripts](https://github.com/bseltz-cohesity/scripts).

---

## Health Check Report ★

> Multi-cluster Cohesity health check designed for enterprise customer business reviews (CBRs) and SE engagements. Gathers live data from Cohesity Helios and produces a **19-sheet Excel workbook**, a **Word document**, and an **editable topology diagram**.

### What it produces

| Output | Description |
|--------|-------------|
| Excel workbook | **19 sheets** covering infrastructure, node hardware with EOL status, **disk health with SSD wear %**, protection health, storage capacity, policy audit (with DataLock/FortKnox status), alerts, **expanded security checklist** (audit log, MFA, NTP auth, remote tunnel, SSO/IDP, TLS cert expiry, **quorum**), **agent health**, **source coverage ratio**, FortKnox/replication, data services, coverage gaps, **user security** (MFA/locked/roles), 30-day trends with charts, a prioritized recommendations list, and **environment topology** |
| Word document | Narrative report with cover page, executive scorecard, environment overview, software & hardware lifecycle, protection health, storage, **10-column security posture table** (cluster enc, SD enc, vault, FortKnox, DataLock WORM, replication, quorum, **admins without MFA**, score), **agent health & source coverage**, **user security table**, topology PNG preview, recommendations, and scoring methodology appendix — ready to share with customers |
| Topology diagram | **Editable `.drawio` file** (opens in diagrams.net / draw.io) with clusters, workload counts, frontend capacity, replication partners, archive targets, and FortKnox vaults — arrows color-coded by target type. PNG preview embedded in the Word document. |

### Sheet reference (19 sheets)

| # | Sheet | Key content |
|---|-------|-------------|
| 1 | Executive Summary | Per-cluster health score, grade, and scorecard |
| 2 | Infrastructure | Cluster versions, nodes, software lifecycle |
| 3 | Node Hardware | Per-node hardware model, serial, EOL status |
| 4 | Disk Health | Disk state, SSD wear %, failed/degraded flags |
| 5 | Protection Health | Per-group pass/fail/miss rates, SLA status, DataLock |
| 6 | Storage & Capacity | Utilization, data reduction, runway by storage domain |
| 7 | Policy Audit | Policy list, retention, DataLock WORM + FortKnox, replication |
| 8 | Policy → Groups | Every protection group with the policy governing it |
| 9 | Alerts | Open critical/warning alerts with age |
| 10 | Security | Expanded checklist: encryption, FIPS, audit log, MFA, NTP, quorum, TLS, … |
| 11 | Replication & Archive | Replication targets, last transfer, archival vaults |
| 12 | FortKnox Data Transfer | Per-group transfer to FortKnox/external targets (30d / 1d) |
| 13 | Data Services | NAS views, quota utilization |
| 14 | Coverage Gaps | Protection groups with no recent successful backup |
| 15 | Recommendations | Prioritized action list (HIGH / MEDIUM / LOW) |
| 16 | Disk Health | Per-disk status and SSD endurance |
| 17 | Agent Health | Per-host agent version, status, upgrade readiness |
| 18 | Source Coverage | Registered sources with protected/unprotected counts |
| 19 | User Security | Per-user MFA status, locked state, roles, last login |

### Requirements

Python 3.7 or later. Install core dependencies with:

```bash
pip3 install requests openpyxl python-docx
```

`python-docx` is required only for Word output. If it is not installed the script still produces the Excel file and skips Word with a warning.

`matplotlib` is required for the topology PNG preview embedded in the Word document. If not installed the PNG is skipped (the `.drawio` file is always produced):

```bash
pip3 install matplotlib
```

### Where to run from

Run the script from the `health_check/` directory, or from the repo root — it uses relative imports only through the standard library so the working directory does not matter:

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
```

### All options

| Flag | Default | Description |
|------|---------|-------------|
| `--apikey KEY` | env `HELIOS_API_KEY` | Helios API key |
| `--cluster NAME` | all | Limit to one cluster (partial name match) |
| `--days N` | 30 | Lookback window for alerts and run history |
| `--customer NAME` | *(blank)* | Customer name on Word cover page and output filename |
| `--output PATH` | auto-generated | Base path for output files (no extension added) |
| `--quick` | off | Use last-run only; skip per-group run history |
| `--excel-only` | off | Skip Word document generation |
| `--word-only` | off | Skip Excel generation |
| `--debug` | off | Print HTTP status for every API call |

### Output files

Auto-generated filenames include the version number and timestamp for easy tracking:

```
cohesity_health_check_v1.37_AcmeCorp_20260410_1430.xlsx
cohesity_health_check_v1.37_AcmeCorp_20260410_1430.docx
cohesity_health_check_v1.37_AcmeCorp_20260410_1430.drawio
```

The `.drawio` file opens directly in [diagrams.net](https://app.diagrams.net/) (free, browser-based) or the draw.io desktop app. All elements are fully editable — clusters, arrows, labels, colors.

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

## Repository Structure

| Folder | Purpose |
|--------|---------|
| [health_check/](health_check/) | **Multi-cluster health check** — 19-sheet Excel + Word report + editable topology diagram for CBRs and SE engagements |
| [reporting/](reporting/) | Protection group run reports and FortKnox vault data transfer reports |
| [protection/](protection/) | Protection group management — run, pause, clone jobs |
| [recovery/](recovery/) | Restore automation — VMs, files, file/folder recovery |
| [storage/](storage/) | Capacity planning, views, storage domain reporting |
| [policies/](policies/) | Policy audit, SLA compliance, and cloning |
| [alerts/](alerts/) | Alert queries, summaries, bulk resolve, and CSV export |
| [infrastructure/](infrastructure/) | Cluster health, node status, version inventory, upgrade readiness |
| [utils/](utils/) | Shared auth and helper modules |

---

## Authentication

All scripts support two modes:

**Helios (recommended)** — single API key covers all connected clusters:
```bash
python3 <script>.py --apikey <your-helios-api-key>
```

The key is stored in the OS keychain after the first use — `--apikey` is optional on subsequent runs.

**Direct cluster** — username/password against a specific cluster:
```bash
python3 <script>.py --cluster <ip-or-hostname> --username admin --domain LOCAL
```

---

## Requirements

```bash
pip3 install requests openpyxl keyring
```

For the health check report, also install:
```bash
# Word document output
pip3 install python-docx

# Topology PNG preview in Word (optional — .drawio file is always produced)
pip3 install matplotlib
```

---

## Quick Reference

### Health Check
```bash
python3 health_check/health_check_report.py --apikey <key> --customer "Customer Name"
python3 health_check/health_check_report.py --apikey <key> --quick --excel-only
```

### Reporting
```bash
python3 reporting/protection_group_report.py --days 7
python3 reporting/fortknox_vault_report.py --mode trend --days 30
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
