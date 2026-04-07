# Recovery

Scripts for automating restore operations — VMs, files, databases, NAS. Useful for DR testing automation and customer recovery runbooks.

**API:** Uses Cohesity v2 REST API endpoints via Helios.

## Scripts

| Script | README | Description | Auth |
|--------|--------|-------------|------|
| `list_snapshots.py` | [docs](list_snapshots.md) | List available recovery points for any protected object | Helios |
| `restore_vm.py` | [docs](restore_vm.md) | Restore a VMware VM to original or alternate location | Helios |
| `recover_files.py` | [docs](recover_files.md) | File/folder level recovery to original or alternate path | Helios |

---

## Common auth pattern

```bash
python3 <script>.py --apikey <key> --cluster <cluster-name> [options]
```

API key is stored in the OS keychain after the first use.

---

## list_snapshots.py

```bash
python3 list_snapshots.py --object "vm-prod-01" --cluster prod-cluster
python3 list_snapshots.py --object "NAS-Share" --cluster prod-cluster --output snaps.xlsx
```

Searches for the protected object by name (partial match), then lists all available snapshots with timestamp, run type, status, expiry, and whether local / replicated / archived copies exist.

Console table output. Use `--output` to also save to Excel.

---

## restore_vm.py

```bash
python3 restore_vm.py --vm "vm-prod-01" --cluster prod-cluster
python3 restore_vm.py --vm "vm-prod-01" --cluster prod-cluster --suffix "-restored"
python3 restore_vm.py --vm "vm-prod-01" --cluster prod-cluster --snapshot 2 --power-on
python3 restore_vm.py --vm "vm-prod-01" --cluster prod-cluster --dry-run
```

Restores a VMware VM to its original location. `--suffix` appends a string to the VM name so the original is preserved. `--snapshot N` selects a specific snapshot by index (0 = latest).

| Flag | Description |
|------|-------------|
| `--snapshot N` | Snapshot index to restore (0 = latest, default: 0) |
| `--suffix` | Append suffix to restored VM name (e.g. `-restored`). No suffix = overwrite |
| `--power-on` | Power on VM after restore completes |
| `--dry-run` | Print the recovery payload without submitting |

---

## recover_files.py

```bash
python3 recover_files.py --source "linux-host-01" --paths /var/log /etc \
        --cluster prod-cluster
python3 recover_files.py --source "nas-share" --paths /data/reports \
        --cluster prod-cluster --dest-path /tmp/restore
python3 recover_files.py --source "win-host" --paths "C:\\Users\\jdoe" \
        --cluster prod-cluster --snapshot 1 --dry-run
```

Recovers files or folders from a physical host or NAS snapshot. Multiple `--paths` can be specified in a single run.

| Flag | Description |
|------|-------------|
| `--paths` | One or more file/folder paths to recover (space-separated) |
| `--snapshot N` | Snapshot index (0 = latest) |
| `--dest-path` | Alternate destination path. Omit to restore to original location |
| `--overwrite` | Overwrite existing files at the destination |
| `--dry-run` | Print the recovery payload without submitting |
