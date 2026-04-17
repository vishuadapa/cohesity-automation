# Policies

Scripts for auditing, creating, and modifying protection policies. Useful for enforcing RPO/RTO standards across a customer environment.

**API:** Uses Cohesity v2 REST API endpoints via Helios.

## Scripts

| Script | README | Description | Auth |
|--------|--------|-------------|------|
| `list_policies.py` | [docs](list_policies.md) | List all policies with retention, replication, and archival settings | Helios |
| `policy_compliance_report.py` | [docs](policy_compliance_report.md) | Flag protection groups whose policy doesn't meet configurable SLA targets | Helios |
| `clone_policy.py` | [docs](clone_policy.md) | Duplicate a policy as a starting point for a new one | Helios |

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

## list_policies.py

```bash
python3 list_policies.py
python3 list_policies.py --cluster prod-cluster
python3 list_policies.py --output policies.xlsx
```

**Output:** `list_policies_YYYYMMDD_HHMMSS.xlsx`
Fields: Cluster, Policy Name, Policy ID, Policy Type, Full Backup Schedule, Incremental Schedule, Local Retention, Log Retention, Replication Targets, Archival Targets, Is Active, Groups Using Policy.

---

## policy_compliance_report.py

```bash
python3 policy_compliance_report.py
python3 policy_compliance_report.py --min-retention 14
python3 policy_compliance_report.py --require-replication --require-archival
python3 policy_compliance_report.py --rpo-hours 24 --cluster prod-cluster
```

**Output:** `policy_compliance_report_YYYYMMDD_HHMMSS.xlsx` — color-coded (green/yellow/red) by status.

**SLA flags:**

| Flag | Description |
|------|-------------|
| `--min-retention N` | Fail groups with local retention < N days |
| `--require-replication` | Fail groups with no replication target in policy |
| `--require-archival` | Fail groups with no archival target in policy |
| `--rpo-hours N` | Warn if last successful run was more than N hours ago |

---

## clone_policy.py

```bash
python3 clone_policy.py --source "Gold 30-Day" --name "Gold 60-Day" --cluster prod-cluster
python3 clone_policy.py --source-id <id> --name "New Policy" --cluster prod-cluster
python3 clone_policy.py --source "Silver" --name "Silver DR" --cluster prod-cluster --dry-run
```

Fetches the source policy, strips server-managed fields (ID, timestamps, group counts), renames it, and POSTs a new policy. Use `--dry-run` to preview the payload before creating.
