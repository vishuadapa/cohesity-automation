# Alerts

Scripts for querying, summarizing, and routing Cohesity alerts. Useful for proactive customer health monitoring and integrating with ticketing systems.

## Scripts

| Script | Description | Auth |
|--------|-------------|------|
| `alert_summary_report.py` | Active alerts by severity, category, and cluster — dual-sheet Excel (Summary + detail) | Helios |
| `resolve_alerts.py` | Bulk resolve open alerts matching a severity / category / alert code filter | Helios |
| `alert_to_csv.py` | Export full alert history for a date range to CSV | Helios |

---

## Common auth pattern

```bash
# Helios (all clusters):
python3 <script>.py --apikey <key>

# Helios (one cluster):
python3 <script>.py --apikey <key> --cluster <cluster-name>
```

API key is stored in the OS keychain after the first use — `--apikey` is optional on subsequent runs.

---

## alert_summary_report.py

```bash
python3 alert_summary_report.py                        # last 1 day (default)
python3 alert_summary_report.py --days 7               # last 7 days
python3 alert_summary_report.py --cluster prod-cluster # one cluster
python3 alert_summary_report.py --output alerts.xlsx
```

**Output:** `alert_summary_report_YYYYMMDD_HHMMSS.xlsx`
- **Summary sheet** — alert count by cluster × severity × category
- **Alerts sheet** — one row per alert with ID, code, severity, category, state, description, timestamps, node ID

---

## resolve_alerts.py

```bash
python3 resolve_alerts.py --severity kCritical --cluster prod-cluster
python3 resolve_alerts.py --category kDisk --yes
python3 resolve_alerts.py --code CE01516014 --cluster prod-cluster --yes
```

Prints matching alerts and prompts for confirmation before resolving. Use `--yes` to skip the prompt.

**Filters:** `--severity` (kCritical, kWarning, kInfo), `--category`, `--code`, `--cluster`
At least one filter is required — resolving all alerts with no filter is blocked.

---

## alert_to_csv.py

```bash
python3 alert_to_csv.py --days 30
python3 alert_to_csv.py --start 2026-03-01 --end 2026-04-01
python3 alert_to_csv.py --days 7 --severity kCritical
```

**Output:** `alert_to_csv_YYYYMMDD_HHMMSS.csv`
Includes both open and resolved alerts. Fields: Cluster, Alert ID, Code, Severity, Category, State, Name, Description, First Occurred, Last Updated, Resolved, Acknowledged, Node ID, Entity Name.
