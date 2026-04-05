# Reporting

Scripts that pull data from Cohesity clusters and generate CSV reports for customer health checks, capacity planning, and executive summaries.

## Scripts

| Script | README | Description | Auth |
|--------|--------|-------------|------|
| `protection_group_report.py` | [docs](protection_group_report.md) | Protection group run status, timing, object counts, data sizes, replication/archive status, storage consumed | Direct / Helios |
| `fortknox_vault_report.py` | [docs](fortknox_vault_report.md) | FortKnox (RPaaS) vault activity by protection group — transfer sizes, run status, vault expiry, object counts, vault storage consumed | Direct / Helios |

---

## Common auth pattern

```bash
# Helios (all clusters):
python3 <script>.py --apikey <key>

# Helios (one cluster):
python3 <script>.py --apikey <key> --cluster <cluster-name>

# Direct cluster:
python3 <script>.py --cluster <ip> --username admin --domain LOCAL
```
