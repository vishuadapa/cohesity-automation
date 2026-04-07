# resolve_alerts.py

Bulk-resolve open Cohesity alerts matching a filter (severity, category, alert code,
or cluster name) via Helios routing.

Prints matching alerts and prompts for confirmation before resolving. Use `--yes`
to skip the prompt for scripting/automation. At least one filter must be specified —
resolving all alerts without a filter is blocked as a safety guard.

Requires a **Helios API key** stored securely in the OS keychain after the first run.

Required packages:
```bash
pip install keyring
```

**Version:** 1.0

---

## API Endpoints Used

| Purpose | Method | Endpoint |
|---------|--------|----------|
| Cluster list | `GET` | `https://helios.cohesity.com/mcm/clusters/connectionStatus` |
| Fetch open alerts | `GET` | `https://helios.cohesity.com/irisservices/api/v1/public/alerts` |
| Resolve alerts | `POST` | `https://helios.cohesity.com/irisservices/api/v1/public/alerts/resolve` |

---

## Usage

```bash
# Resolve all critical alerts (with confirmation prompt):
python3 resolve_alerts.py --severity kCritical

# Resolve disk alerts on one cluster without prompting:
python3 resolve_alerts.py --category kDisk --cluster <name> --yes

# Resolve by specific alert code:
python3 resolve_alerts.py --code CE01516014 --yes

# Remove stored credentials:
python3 resolve_alerts.py --clear-credentials
```

---

## CLI Options

| Option | Description |
|--------|-------------|
| `--apikey` | Helios API key (optional — keychain is used if omitted) |
| `--cluster` | Limit to one cluster by name |
| `--severity` | Filter by severity: `kCritical`, `kWarning`, `kInfo` |
| `--category` | Filter by category: `kDisk`, `kNode`, `kBackupRestore`, etc. |
| `--code` | Filter by exact alert code (e.g. `CE01516014`) |
| `--yes` | Skip confirmation prompt — resolve immediately |
| `--clear-credentials` | Remove the stored API key from the OS keychain and exit |

> **Safety:** At least one of `--severity`, `--category`, `--code`, or `--cluster` must be
> specified. The script will not resolve all alerts cluster-wide without a filter.

---

## Workflow

1. Fetches all open alerts from each matching cluster
2. Applies the specified filters
3. Prints the matching alerts with severity, category, code, and timestamp
4. Prompts for confirmation (`y/N`) unless `--yes` is set
5. POSTs the resolve request with the list of alert IDs per cluster
6. Reports the total count of successfully resolved alerts

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.0 | 2026-04-06 | feat: initial release |
