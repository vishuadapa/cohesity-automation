# recover_files.py

Recovers files or folders from a protected physical host or NAS snapshot to the original or alternate location. Supports multiple paths per recovery, overwrite options, and snapshot selection.

Requires a **Helios API key**. The key is stored securely in the OS keychain after the first run.

Required packages:
```bash
pip install requests keyring
```

**Version:** 1.1

---

## Usage

```bash
# Recover a single path to its original location:
python3 recover_files.py --source "linux-host-01" --paths /var/log --cluster prod-cluster

# Recover multiple paths in one run:
python3 recover_files.py --source "linux-host-01" --paths /var/log /etc /home \
        --cluster prod-cluster

# Recover to an alternate destination:
python3 recover_files.py --source "nas-share" --paths /data/reports \
        --cluster prod-cluster --dest-path /tmp/restore

# Recover from a specific snapshot (index 1 = 2nd most recent):
python3 recover_files.py --source "win-host" --paths "C:\\Users\\jdoe" \
        --cluster prod-cluster --snapshot 1

# Overwrite existing files at destination:
python3 recover_files.py --source "linux-host" --paths /var/log \
        --cluster prod-cluster --dest-path /archive --overwrite

# Validate the recovery payload without submitting:
python3 recover_files.py --source "nas-share" --paths /data \
        --cluster prod-cluster --dry-run

# First run — prompts for Helios API key:
python3 recover_files.py --apikey <key> --source "test-host" --paths /home --cluster prod-cluster

# Remove stored credentials:
python3 recover_files.py --clear-credentials
```

---

## CLI Options

| Option | Description |
|--------|-------------|
| `--apikey` | Helios API key (optional — keychain used if omitted, or you will be prompted) |
| `--clear-credentials` | Remove stored API key from the OS keychain and exit |
| `--source` | Protected physical host or NAS share name (partial match supported; case-insensitive) |
| `--paths` | One or more file/folder paths to recover (space-separated). Example: `/var/log /etc /home` |
| `--cluster` | Cluster name (required; partial match supported) |
| `--snapshot` | Snapshot index to recover from (0 = latest, 1 = 2nd most recent, etc.). Default: 0 |
| `--dest-path` | Alternate destination path for recovery. Omit to restore to the original location. |
| `--overwrite` | Overwrite existing files at the destination if they already exist. Without this flag, conflicts will cause the recovery to fail. |
| `--dry-run` | Print the recovery payload without submitting the request. Use this to inspect the recovery configuration before executing. |

---

## Behavior

- **Source search:** Uses partial name matching. If multiple protected objects match, the first match is used.
- **Path handling:**
  - Linux/NAS paths: `/var/log`, `/data/reports`
  - Windows paths: `C:\Users\jdoe` (use backslashes for Windows paths)
  - Multiple paths: space-separated list, all recovered in a single operation
- **Destination:**
  - Without `--dest-path`: files/folders are restored to their original location.
  - With `--dest-path`: all recovered paths are restored under this alternate root.
- **Overwrite:**
  - By default, recovery fails if files exist at the destination (fail-safe for production).
  - Use `--overwrite` to allow replacement of existing files.
- **Snapshot selection:** Snapshots are ordered most-recent-first (index 0 = latest). Use `--snapshot N` to select older snapshots.
- **`--dry-run`:** Prints the full recovery API payload without submitting it. Use this to validate paths and configuration before executing.

---

## Output

Confirmation message with recovery details: source host/share, paths being recovered, destination location, and overwrite status. If `--dry-run` is specified, prints the JSON recovery payload.

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.1 | 2026-04-07 | Fixed alternate destination path field name: `absolutePath` → `alternateRestoreBaseDirectory` (v2 API). Removed redundant `restore_params` intermediary and unused `targetEnvironment` variable. |
| 1.0 | 2026-04-06 | Initial release |
