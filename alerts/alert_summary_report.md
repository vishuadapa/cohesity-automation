# alert_summary_report.py

Reports active Cohesity alerts grouped by severity and category across all
Helios-connected clusters (or a single named cluster).

Requires a **Helios API key** stored securely in the OS keychain after the first run.

Required packages:
```bash
pip install openpyxl keyring
```

Default window is the last 1 day.

**Version:** 1.0

---

## API Endpoints Used

| Purpose | Method | Endpoint |
|---------|--------|----------|
| Cluster list | `GET` | `https://helios.cohesity.com/mcm/clusters/connectionStatus` |
| Alerts | `GET` | `https://helios.cohesity.com/irisservices/api/v1/public/alerts` |

Parameters: `alertStateList=kOpen`, `startDateUsecs`, `endDateUsecs`, `maxAlerts=1000`

---

## Usage

```bash
# First run — prompts for API key and saves to OS keychain:
python3 alert_summary_report.py

# Last 7 days of open alerts:
python3 alert_summary_report.py --days 7

# One cluster only:
python3 alert_summary_report.py --cluster <name>

# Remove stored credentials:
python3 alert_summary_report.py --clear-credentials
```

Output is auto-generated as `alert_summary_report_YYYYMMDD_HHMMSS.xlsx` in the current directory.

---

## CLI Options

| Option | Description |
|--------|-------------|
| `--apikey` | Helios API key (optional — keychain is used if omitted) |
| `--cluster` | Filter to one cluster by name |
| `--days N` | Look back N days for open alerts (default: 1) |
| `--output` | Output `.xlsx` path (auto-generated if omitted) |
| `--debug` | Print API request details |
| `--clear-credentials` | Remove the stored API key from the OS keychain and exit |

---

## Output Sheets

### Summary
One row per (Cluster × Severity × Category) combination with an alert count.

| Column | Description |
|--------|-------------|
| Cluster | Cluster name |
| Severity | Alert severity: `kCritical`, `kWarning`, `kInfo` |
| Category | Alert category: `kDisk`, `kNode`, `kBackupRestore`, `kNetwork`, etc. |
| Alert Count | Number of open alerts matching this combination |

### Alerts
One row per alert with full detail.

| Column | Source field | Description |
|--------|-------------|-------------|
| Cluster | _(script)_ | Cluster name from Helios |
| Alert ID | `id` | Unique alert identifier |
| Alert Code | `alertCode` | Alert type code (e.g. `CE01516014`) |
| Severity | `severity` | `kCritical`, `kWarning`, or `kInfo` |
| Category | `alertCategory` | Alert category |
| State | `alertState` | `kOpen`, `kResolved`, etc. |
| Alert Name | `alertDocument.alertName` | Human-readable alert name |
| Description | `alertDocument.alertDescription` | Alert description text |
| First Occurred | `firstTimestampUsecs` | When the alert was first raised (UTC) |
| Last Updated | `latestTimestampUsecs` | When the alert was last updated (UTC) |
| Node ID | `propertyList[nodeId]` | Node that raised the alert, if applicable |
| Acknowledged | `acknowledged` | Whether the alert has been acknowledged |
