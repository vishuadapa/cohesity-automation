# list_views.py

Enumerates all NAS views (SMB/NFS/S3 shares) in Cohesity clusters with quota, logical usage, protocol support, and share paths. Useful for storage capacity planning and data governance.

Requires a **Helios API key**. The key is stored securely in the OS keychain after the first run.

Required packages:
```bash
pip install requests keyring openpyxl
```

**Version:** 1.1

---

## Usage

```bash
# List all views across all clusters:
python3 list_views.py

# Filter to one cluster:
python3 list_views.py --cluster prod-cluster

# Filter by protocol (SMB, NFS, or S3):
python3 list_views.py --protocol nfs

# Combine cluster and protocol filters:
python3 list_views.py --cluster prod-cluster --protocol smb

# Custom output filename:
python3 list_views.py --output views.xlsx

# First run — prompts for Helios API key:
python3 list_views.py --apikey <key>

# Remove stored credentials:
python3 list_views.py --clear-credentials
```

Output file is created automatically as `list_views_YYYYMMDD_HHMMSS.xlsx` in the directory where the script is run (unless `--output` is specified).

---

## CLI Options

| Option | Description |
|--------|-------------|
| `--apikey` | Helios API key (optional — keychain used if omitted, or you will be prompted) |
| `--clear-credentials` | Remove stored API key from the OS keychain and exit |
| `--cluster` | Filter to a specific cluster (partial match supported) |
| `--protocol` | Filter by protocol: `smb`, `nfs`, or `s3` (case-insensitive) |
| `--output` | Output Excel filename (default: `list_views_YYYYMMDD_HHMMSS.xlsx` in current directory) |

---

## Report Fields

| Column | Description |
|--------|-------------|
| Cluster | Cluster name |
| View Name | Name of the NAS view/share |
| Storage Domain | Storage domain where the view is hosted |
| Protocols | Comma-separated list of enabled protocols: SMB, NFS, S3 |
| Quota Limit (GB) | Maximum capacity allocated to this view (if set), otherwise "Unlimited" |
| Logical Usage (GB) | Current logical data usage of the view |
| Usage % | Percentage of quota in use (N/A if quota is unlimited) |
| Case Insensitive | `True` if the view uses case-insensitive naming |
| Created | View creation date and time |
| SMB Paths | SMB share path(s) if protocol is enabled, otherwise empty |
| NFS Paths | NFS mount path(s) if protocol is enabled, otherwise empty |

---

## Behavior

- **Cluster filter:** Uses partial name matching. If omitted, all clusters are listed.
- **Protocol filter:** Returns only views that support the specified protocol. A view supporting multiple protocols will appear once even if it supports SMB, NFS, and S3.
- **Output:** Excel file with styled headers (bold white text on dark blue background), frozen header row, and auto-fit columns.
- **Quota:** If no quota is set on a view, the `Quota Limit` column shows "Unlimited" and `Usage %` is N/A.

---

## Notes

- Logical usage reflects the current data in the view and does not account for deduplication across other views or protection groups.
- Share paths show the complete access paths for each protocol (e.g., `\\cluster-name\view-name` for SMB, `/cluster-name/view-name` for NFS).

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.1 | 2026-04-07 | fix: AttributeError — API returns explicit `null` for `logicalQuota`, `storagePolicy`, `stats`, `smbMountPaths`, `nfsMountPaths`; replace `.get(k, {})` with `(... or {})` guards; improve error reporting in views fetch |
| 1.0 | 2026-04-06 | feat: initial release |
