# list_policies.py

Lists all protection policies across Cohesity clusters with retention, replication targets, archival targets, and schedule settings. Outputs to Excel with styled headers.

Requires a **Helios API key**. The key is stored securely in the OS keychain after the first run.

Required packages:
```bash
pip install requests keyring openpyxl
```

**Version:** 1.3

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.3 | 2026-04-07 | Groups Using Policy now counted from `GET /v2/data-protect/protection-groups` (groups per `policyId`) — v2 policies API does not return `numProtectionGroups` |
| 1.2 | 2026-04-07 | Fixed empty columns D, E, H, I, J, L: wrong field names for the v2 API — `policyCategory` → `type`, `fullBackups` → `full`, nested schedule frequency (`daySchedule`/`weekSchedule`/`hourSchedule`), improved replication and archival target name lookup with fallbacks |
| 1.1 | 2026-04-07 | Fixed `AttributeError` crash when API returns explicit `null` for policy sub-fields |
| 1.0 | 2026-04-06 | Initial release |

---

## Usage

```bash
# List all policies across all clusters:
python3 list_policies.py

# List policies for one cluster:
python3 list_policies.py --cluster prod-cluster

# Save to a custom Excel filename:
python3 list_policies.py --output policies.xlsx

# Combine cluster filter and custom output:
python3 list_policies.py --cluster prod-cluster --output prod-policies.xlsx

# First run — prompts for Helios API key:
python3 list_policies.py --apikey <key>

# Remove stored credentials:
python3 list_policies.py --clear-credentials
```

Output file is created automatically as `list_policies_YYYYMMDD_HHMMSS.xlsx` in the directory where the script is run (unless `--output` is specified).

---

## CLI Options

| Option | Description |
|--------|-------------|
| `--apikey` | Helios API key (optional — keychain used if omitted, or you will be prompted) |
| `--clear-credentials` | Remove stored API key from the OS keychain and exit |
| `--cluster` | Filter to a specific cluster (partial match supported) |
| `--output` | Output Excel filename (default: `list_policies_YYYYMMDD_HHMMSS.xlsx` in current directory) |

---

## Report Fields

| Column | Description |
|--------|-------------|
| Cluster | Cluster name |
| Policy Name | Protection policy name |
| Policy ID | Internal Cohesity policy identifier |
| Policy Type | Workload type: `VMware`, `PhysicalFiles`, `Oracle`, `MongoDB`, etc. |
| Full Backup Schedule | Schedule for full backups (e.g., `Weekly on Sunday`) |
| Incremental Schedule | Schedule for incremental backups (e.g., `Daily at 2:00 AM`) |
| Local Retention | How long snapshots are retained on the cluster (e.g., `30 days`) |
| Log Retention | How long transaction logs are retained (for database policies) |
| Replication Targets | Comma-separated list of clusters this policy replicates to |
| Archival Targets | Comma-separated list of external vaults (e.g., FortKnox, S3) this policy archives to |
| Is Active | `True` if the policy is currently active and in use |
| Groups Using Policy | Number of protection groups using this policy |

---

## Behavior

- **Cluster filter:** Uses partial name matching. If omitted, all clusters are listed.
- **Output:** Excel file with styled headers (bold white text on dark blue background), frozen header row, and auto-fit columns.
