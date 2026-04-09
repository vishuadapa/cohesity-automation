# Cohesity Automation Scripts

Personal collection of Cohesity REST API automation scripts for reporting, operations, and customer use cases. Built and maintained by Vishu Adapa.

Targets Cohesity 7.1 and 7.3. Uses the [Cohesity REST API](https://developer.cohesity.com/) (v1 and v2 endpoints as appropriate).
Inspired by [Brian Seltzer's scripts](https://github.com/bseltz-cohesity/scripts).

---

## Health Check Report ★

> Multi-cluster Cohesity health check designed for enterprise customer business reviews (CBRs) and SE engagements. Gathers live data from Cohesity Helios and produces an **18-sheet Excel workbook** and a **Word document**.

### What it produces

| Output | Description |
|--------|-------------|
| Excel workbook | **18 sheets** covering infrastructure, node hardware with EOL status, **disk health with SSD wear %**, protection health, storage capacity, policy audit, alerts, **expanded security checklist** (audit log, MFA, NTP auth, remote tunnel, SSO/IDP, TLS cert expiry), **agent health**, **source coverage ratio**, FortKnox/replication, data services, coverage gaps, **user security** (MFA/locked/roles), 30-day trends with charts, and a prioritized recommendations list |
| Word document | Narrative report with cover page, executive scorecard, environment overview, software & hardware lifecycle, protection health, storage, security posture, **agent health & source coverage**, **user security table**, recommendations, and scoring methodology appendix — ready to share with customers |

### Requirements

Python 3.7 or later. Install dependencies with:

```bash
pip3 install requests openpyxl python-docx
```

`python-docx` is required only for Word output. If it is not installed the script still produces the Excel file and skips Word with a warning.

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
# Full report — all clusters, 30-day window, Excel + Word
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

### Output filename

Auto-generated filenames include the version number and timestamp for easy tracking:

```
cohesity_health_check_v1.21_AcmeCorp_20260408_1430.xlsx
cohesity_health_check_v1.21_AcmeCorp_20260408_1430.docx
```

See [health_check/health_check_report.md](health_check/health_check_report.md) for the full sheet reference, scoring model, and version history.

---

## Repository Structure

| Folder | Purpose |
|--------|---------|
| [health_check/](health_check/) | **Multi-cluster health check** — 18-sheet Excel + Word report for CBRs and SE engagements |
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
pip3 install python-docx
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
