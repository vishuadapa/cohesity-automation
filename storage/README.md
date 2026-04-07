# Storage

Scripts for capacity management, storage domains, views (SMB/NFS shares), and data reduction reporting. Useful for capacity planning conversations and storage tier optimization.

## Scripts

| Script | Description | Auth |
|--------|-------------|------|
| `capacity_report.py` | Cluster capacity, used/free/usable in TB, data reduction ratio, dedup, compression, savings | Helios |
| `list_views.py` | Enumerate all NAS views with quota, logical usage, protocol, and share paths | Helios |
| `storage_domain_report.py` | Storage domain utilization — physical/logical usage, DR ratio, dedup and encryption config | Helios |

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

## capacity_report.py

```bash
python3 capacity_report.py
python3 capacity_report.py --cluster prod-cluster
python3 capacity_report.py --output capacity.xlsx
```

**Output:** `capacity_report_YYYYMMDD_HHMMSS.xlsx`
Fields: Cluster, Software Version, Usable Capacity (TB), Used (TB), Free (TB), Used %, Logical Data (TB), Data Reduction Ratio, Dedup Ratio, Compression Ratio, Savings (TB), Node Count.

---

## list_views.py

```bash
python3 list_views.py
python3 list_views.py --cluster prod-cluster
python3 list_views.py --protocol nfs          # filter by SMB / NFS / S3
python3 list_views.py --output views.xlsx
```

**Output:** `list_views_YYYYMMDD_HHMMSS.xlsx`
Fields: Cluster, View Name, Storage Domain, Protocols, Quota Limit (GB), Logical Usage (GB), Usage %, Case Insensitive, Created, SMB Paths, NFS Paths.

---

## storage_domain_report.py

```bash
python3 storage_domain_report.py
python3 storage_domain_report.py --cluster prod-cluster
python3 storage_domain_report.py --output domains.xlsx
```

**Output:** `storage_domain_report_YYYYMMDD_HHMMSS.xlsx`
Fields: Cluster, Domain Name, Domain ID, Physical Capacity (TB), Logical Used (TB), Physical Used (TB), Data Reduction Ratio, Dedup Enabled, Compression Policy, Encryption Enabled, Cloud Tiering Enabled, View Count, Default Domain.
