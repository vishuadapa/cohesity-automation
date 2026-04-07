# list_snapshots.py

Lists all available recovery points (snapshots) for any protected object in Cohesity. Displays snapshot details including timestamp, run type, status, expiry, and flags indicating whether local, replicated, or archived copies exist.

Requires a **Helios API key**. The key is stored securely in the OS keychain after the first run.

Required packages:
```bash
pip install requests keyring openpyxl
```

**Version:** 1.1

---

## Usage

```bash
# List all snapshots for a protected object:
python3 list_snapshots.py --object "vm-prod-01" --cluster prod-cluster

# List snapshots and save to Excel:
python3 list_snapshots.py --object "NAS-Share" --cluster prod-cluster --output snaps.xlsx

# Search for an object with partial name match:
python3 list_snapshots.py --object "prod" --cluster prod-cluster

# First run — prompts for Helios API key:
python3 list_snapshots.py --apikey <key> --object "test-vm" --cluster prod-cluster

# Remove stored credentials:
python3 list_snapshots.py --clear-credentials
```

---

## CLI Options

| Option | Description |
|--------|-------------|
| `--apikey` | Helios API key (optional — keychain used if omitted, or you will be prompted) |
| `--clear-credentials` | Remove stored API key from the OS keychain and exit |
| `--object` | Protected object name to search for (partial match supported; case-insensitive) |
| `--cluster` | Cluster name (required; partial match supported) |
| `--output` | Optional Excel filename to save snapshot list (format: `.xlsx`). If omitted, output is printed to console only. |

---

## Output

Console table and optional Excel file with the following columns:

| Column | Description |
|--------|-------------|
| Snapshot ID | Internal Cohesity snapshot identifier |
| Timestamp | Snapshot creation date and time (cluster local timezone) |
| Run Type | Backup type: `Regular`, `Full`, `Log`, `System` |
| Status | `Succeeded`, `Failed`, `SucceededWithWarning` |
| Expiry | Snapshot expiration date (when it will be deleted from cluster) |
| Local Copy | `Yes` if the snapshot exists on this cluster |
| Replicated Copy | `Yes` if the snapshot has been replicated to another cluster |
| Archived Copy | `Yes` if the snapshot has been archived to external target (e.g., FortKnox) |

---

## Behavior

- **Object search:** Uses partial name matching. If multiple objects match, all snapshots across all matches are listed.
- **Output:** Printed to console as a formatted table. Use `--output` to also save to an Excel file with styled headers and auto-fit columns.
- **Snapshot availability:** Snapshots are listed with their expiry dates. Snapshots past expiry will no longer be available for recovery.

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.1 | 2026-04-07 | refactor: remove unused `targets` variable (`indexingStatus` assigned but never referenced) |
| 1.0 | 2026-04-06 | feat: initial release |
