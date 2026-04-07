# policy_compliance_report.py

Flags protection groups whose policies don't meet configurable SLA targets. Generates an Excel report with color-coded status (green = compliant, yellow = warning, red = fail) for retention, replication, archival, and RPO checks.

Requires a **Helios API key**. The key is stored securely in the OS keychain after the first run.

Required packages:
```bash
pip install requests keyring openpyxl
```

**Version:** 1.0

---

## Usage

```bash
# Basic compliance check (all clusters, no specific SLA targets):
python3 policy_compliance_report.py

# Check one cluster:
python3 policy_compliance_report.py --cluster prod-cluster

# Require at least 14 days local retention:
python3 policy_compliance_report.py --min-retention 14

# Require replication and archival to be configured:
python3 policy_compliance_report.py --require-replication --require-archival

# Combine multiple SLA checks:
python3 policy_compliance_report.py --min-retention 30 --require-replication \
        --require-archival --rpo-hours 24 --cluster prod-cluster

# Custom output filename:
python3 policy_compliance_report.py --min-retention 14 --output compliance.xlsx

# First run — prompts for Helios API key:
python3 policy_compliance_report.py --apikey <key> --min-retention 7

# Remove stored credentials:
python3 policy_compliance_report.py --clear-credentials
```

Output file is created automatically as `policy_compliance_report_YYYYMMDD_HHMMSS.xlsx` in the directory where the script is run (unless `--output` is specified).

---

## CLI Options

| Option | Description |
|--------|-------------|
| `--apikey` | Helios API key (optional — keychain used if omitted, or you will be prompted) |
| `--clear-credentials` | Remove stored API key from the OS keychain and exit |
| `--cluster` | Filter to a specific cluster (partial match supported) |
| `--output` | Output Excel filename (default: `policy_compliance_report_YYYYMMDD_HHMMSS.xlsx` in current directory) |
| `--min-retention` | Fail protection groups with local retention less than N days |
| `--require-replication` | Fail groups with no replication target configured in their policy |
| `--require-archival` | Fail groups with no archival/external target configured in their policy |
| `--rpo-hours` | Warn if the last successful backup run was more than N hours ago (helps identify stale/failing groups) |

---

## Report Fields

| Column | Description |
|--------|-------------|
| Cluster | Cluster name |
| Protection Group | Protection group name |
| Policy Name | Assigned protection policy name |
| Status | Compliance status: `Pass` (green), `Warning` (yellow), `Fail` (red) |
| Local Retention | Configured local retention (days) |
| Replication Configured | `Yes` / `No` |
| Archival Configured | `Yes` / `No` |
| Last Run | Date and time of the most recent backup run |
| Last Run Status | Success status of the last run |
| Issues Found | List of specific SLA violations (e.g., "Retention 7 days < min 14", "No replication configured") |

---

## Behavior

- **SLA flags:** At least one SLA flag (`--min-retention`, `--require-replication`, `--require-archival`, `--rpo-hours`) must be specified, or all groups will report "Pass" with no violations.
- **Status calculation:**
  - **Pass (green):** All configured SLA checks are met.
  - **Warning (yellow):** `--rpo-hours` check failed (last run was too long ago).
  - **Fail (red):** Retention, replication, or archival requirements not met.
- **Output:** Excel file with color-coded row backgrounds and detailed issue descriptions in the last column for easy filtering in spreadsheets or customer reports.

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.0 | 2026-04-06 | feat: initial release |
