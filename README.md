# Cohesity Automation Scripts

Personal collection of Cohesity REST API automation scripts for reporting, operations, and customer use cases. Built and maintained by Vishu Adapa.

Targets Cohesity 7.1 and 7.3. Uses the [Cohesity REST API](https://developer.cohesity.com/) (v1 and v2 endpoints as appropriate).
Inspired by [Brian Seltzer's scripts](https://github.com/bseltz-cohesity/scripts).

---

## API Versions

- **v1 API** — Used for infrastructure, alerts, storage, and other operations for better Helios reliability
- **v2 API** — Used for protection group reporting, policies, protection operations, and recovery

Individual `README.md` files in each folder document which API version is used by those scripts.

---

## Repository Structure

| Folder | Purpose |
|--------|---------|
| [reporting/](reporting/) | Health check and status reports — CSV/Excel output for Excel and customer decks |
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
Generate your API key in Helios → Settings → Access Management → API Keys.

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

---

## Quick Reference

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

## Disclaimer

These scripts are personal tools, not official Cohesity products. Use at your own risk in production environments.
