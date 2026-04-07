# restore_vm.py

Restores a VMware VM to its original location, with optional naming suffix to preserve the original VM, snapshot selection, and automatic power-on. Includes dry-run mode for validation before execution.

Requires a **Helios API key**. The key is stored securely in the OS keychain after the first run.

Required packages:
```bash
pip install requests keyring
```

**Version:** 1.1

---

## Usage

```bash
# Restore a VM to its original location (overwrites the original):
python3 restore_vm.py --vm "vm-prod-01" --cluster prod-cluster

# Restore with a suffix to preserve the original:
python3 restore_vm.py --vm "vm-prod-01" --cluster prod-cluster --suffix "-restored"

# Restore a specific snapshot (index 2 = 3rd most recent):
python3 restore_vm.py --vm "vm-prod-01" --cluster prod-cluster --snapshot 2

# Restore latest snapshot and power on the VM:
python3 restore_vm.py --vm "vm-prod-01" --cluster prod-cluster --power-on

# Validate the recovery payload without submitting:
python3 restore_vm.py --vm "vm-prod-01" --cluster prod-cluster --suffix "-test" --dry-run

# First run — prompts for Helios API key:
python3 restore_vm.py --apikey <key> --vm "test-vm" --cluster prod-cluster

# Remove stored credentials:
python3 restore_vm.py --clear-credentials
```

---

## CLI Options

| Option | Description |
|--------|-------------|
| `--apikey` | Helios API key (optional — keychain used if omitted, or you will be prompted) |
| `--clear-credentials` | Remove stored API key from the OS keychain and exit |
| `--vm` | VM name to restore (partial match supported; case-insensitive) |
| `--cluster` | Cluster name (required; partial match supported) |
| `--snapshot` | Snapshot index to restore (0 = latest, 1 = 2nd most recent, etc.). Default: 0 |
| `--suffix` | Append a suffix to the VM name during restore (e.g., `-restored`, `-test`). If omitted, the restored VM overwrites the original. |
| `--power-on` | Automatically power on the restored VM after recovery completes |
| `--dry-run` | Print the recovery payload without submitting the request. Use this to inspect the recovery configuration before executing. |

---

## Behavior

- **VM search:** Uses partial name matching. If multiple VMs match, the first match is restored.
- **Snapshot selection:** Snapshots are ordered most-recent-first (index 0 = latest). Use `--snapshot N` to select older snapshots.
- **Suffix vs. overwrite:**
  - Without `--suffix`: restored VM **overwrites** the original (dangerous for production VMs).
  - With `--suffix`: restored VM is created with a new name (original is preserved).
- **Power-on:** By default, restored VMs are powered off. Use `--power-on` to power them on immediately after recovery.
- **`--dry-run`:** Prints the full recovery API payload without submitting it. Use this to validate the recovery configuration before executing it.

---

## Output

Confirmation message with recovery details: source VM, snapshot date, destination VM name, and power-on status. If `--dry-run` is specified, prints the JSON recovery payload.

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.1 | 2026-04-07 | fix: rename params — eliminate None/delete pattern by building `originalSourceConfig` conditionally before payload construction |
| 1.0 | 2026-04-06 | feat: initial release |
