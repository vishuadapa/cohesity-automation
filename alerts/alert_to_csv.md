# alert_to_csv.py

Exports Cohesity alert history for a date range to CSV. Includes both open and
resolved alerts across all Helios-connected clusters (or a single named cluster).

Output is plain CSV — suitable for Excel pivot tables, SIEM ingestion, or
ticketing system imports.

Requires a **Helios API key** stored securely in the OS keychain after the first run.

Required packages:
```bash
pip install keyring
```

Default window is the last 7 days.

**Version:** 1.0

---

## API Endpoints Used

| Purpose | Method | Endpoint |
|---------|--------|----------|
| Cluster list | `GET` | `https://helios.cohesity.com/mcm/clusters/connectionStatus` |
| Alerts | `GET` | `https://helios.cohesity.com/irisservices/api/v1/public/alerts` |

Parameters: `startDateUsecs`, `endDateUsecs`, `maxAlerts=1000`

---

## Usage

```bash
# Last 30 days:
python3 alert_to_csv.py --days 30

# Specific date range:
python3 alert_to_csv.py --start 2026-03-01 --end 2026-04-01

# Filter by severity:
python3 alert_to_csv.py --days 7 --severity kCritical

# One cluster, filtered by category:
python3 alert_to_csv.py --cluster <name> --category kDisk --days 14

# Remove stored credentials:
python3 alert_to_csv.py --clear-credentials
```

Output is auto-generated as `alert_to_csv_YYYYMMDD_HHMMSS.csv` in the current directory.

---

## CLI Options

| Option | Description |
|--------|-------------|
| `--apikey` | Helios API key (optional — keychain is used if omitted) |
| `--cluster` | Filter to one cluster by name |
| `--days N` | Last N days (default: 7) |
| `--start YYYY-MM-DD` | Start date (UTC) |
| `--end YYYY-MM-DD` | End date (UTC, defaults to today) |
| `--severity` | Filter by severity: `kCritical`, `kWarning`, `kInfo` |
| `--category` | Filter by category: `kDisk`, `kNode`, `kBackupRestore`, etc. |
| `--output` | Output `.csv` path (auto-generated if omitted) |
| `--debug` | Print API request details |
| `--clear-credentials` | Remove the stored API key from the OS keychain and exit |

---

## CSV Fields

| Column | Source field | Description |
|--------|-------------|-------------|
| Cluster | _(script)_ | Cluster name from Helios |
| Alert ID | `id` | Unique alert identifier |
| Alert Code | `alertCode` | Alert type code (e.g. `CE01516014`) |
| Severity | `severity` | `kCritical`, `kWarning`, or `kInfo` |
| Category | `alertCategory` | Alert category |
| State | `alertState` | `kOpen` or `kResolved` |
| Alert Name | `alertDocument.alertName` | Human-readable alert name |
| Description | `alertDocument.alertDescription` | Alert description text |
| First Occurred (UTC) | `firstTimestampUsecs` | When the alert was first raised |
| Last Updated (UTC) | `latestTimestampUsecs` | When the alert was last updated |
| Resolved (UTC) | `resolvedTimestampUsecs` | When the alert was resolved (empty if still open) |
| Acknowledged | `acknowledged` | `Yes` or `No` |
| Node ID | `propertyList[nodeId]` | Node that raised the alert, if applicable |
| Entity Name | `propertyList[entityName]` | Entity (disk, volume, etc.) related to the alert |
