# Infrastructure

Scripts for cluster-level operations — node status, software version inventory, upgrade readiness, and configuration audits. Useful for presales assessments and health checks.

**API:** Uses Cohesity v1 REST API endpoints for cluster infrastructure data via Helios.

## Scripts

| Script | README | Description | Auth |
|--------|--------|-------------|------|
| `cluster_info_report.py` | [docs](cluster_info_report.md) | Software version, node count, cluster ID, DNS/NTP/timezone config across all clusters | Helios |
| `node_status.py` | [docs](node_status.md) | Node health, disk status, fault counts, and capacity — dual-sheet Excel (Nodes + Disks) | Helios |
| `upgrade_readiness.py` | [docs](upgrade_readiness.md) | Pre-upgrade checks: node health, disk health, active runs, version, and capacity margin | Helios |

---

## Common auth pattern

```bash
# Helios (all clusters):
python3 <script>.py --apikey <key>

# Helios (one cluster):
python3 <script>.py --apikey <key> --cluster <cluster-name>
```

API key is stored in the OS keychain after the first use.

---

## TLS / certificate options

All scripts support the following flags for environments with custom PKI or self-signed certificates:

| Flag | Description |
|------|-------------|
| `--ca-bundle PATH` | Path to a CA bundle (PEM) to verify TLS (e.g. corporate proxy cert) |
| `--insecure` | Disable TLS certificate verification — **not recommended for production** |

---

## cluster_info_report.py

```bash
python3 cluster_info_report.py
python3 cluster_info_report.py --cluster prod-cluster
python3 cluster_info_report.py --output inventory.xlsx
```

**Output:** `cluster_info_report_YYYYMMDD_HHMMSS.xlsx`
Fields: Cluster Name, Cluster ID, Incarnation ID, Software Version, Cluster Type, Node Count, Healthy Nodes, Domain Name, DNS Servers, NTP Servers, Timezone, Encryption Enabled, FIPS Enabled, Helios Connected.

---

## node_status.py

```bash
python3 node_status.py
python3 node_status.py --cluster prod-cluster
python3 node_status.py --unhealthy-only    # only nodes/disks with issues
```

**Output:** `node_status_YYYYMMDD_HHMMSS.xlsx`
- **Nodes sheet** — Cluster, Node ID, IP, Status, Role, Total/Healthy/Faulted Disks, Usable Capacity (TB), Software Version
- **Disks sheet** — Cluster, Node IP, Disk ID, Slot, Tier, Capacity (TB), Status, Mount Path

---

## upgrade_readiness.py

```bash
python3 upgrade_readiness.py
python3 upgrade_readiness.py --cluster prod-cluster
python3 upgrade_readiness.py --min-version 7.1
python3 upgrade_readiness.py --output readiness.xlsx
```

Runs PASS/WARN/FAIL checks per cluster: node health, disk health, active runs in progress, version meets minimum, and available capacity margin (warns below 20%, fails below 10%).

Output is printed to the console in a table and saved to Excel with color-coded rows.
