# pause_resume_groups.py

Bulk pause or resume protection groups by name pattern matching. Prints matching groups and requests confirmation before making changes. Supports automatic resume after a specified delay for scripted maintenance windows.

Requires a **Helios API key**. The key is stored securely in the OS keychain after the first run.

Required packages:
```bash
pip install requests keyring
```

**Version:** 1.0

---

## Usage

```bash
# Pause all groups matching a pattern (confirmation required):
python3 pause_resume_groups.py --action pause --pattern "VMware" --cluster prod-cluster

# Pause and skip confirmation:
python3 pause_resume_groups.py --action pause --pattern "SQL" --cluster prod-cluster --yes

# Resume groups:
python3 pause_resume_groups.py --action resume --pattern "Daily" --cluster prod-cluster --yes

# Pause, auto-resume after 60 minutes (useful for maintenance windows):
python3 pause_resume_groups.py --action pause --pattern "NAS" --cluster prod-cluster \
        --resume-after 60 --yes

# First run — prompts for Helios API key:
python3 pause_resume_groups.py --apikey <key> --action pause --pattern "test" --cluster prod-cluster

# Remove stored credentials:
python3 pause_resume_groups.py --clear-credentials
```

---

## CLI Options

| Option | Description |
|--------|-------------|
| `--apikey` | Helios API key (optional — keychain used if omitted, or you will be prompted) |
| `--clear-credentials` | Remove stored API key from the OS keychain and exit |
| `--action` | `pause` or `resume` (required) |
| `--pattern` | Protection group name pattern to match (partial match supported; case-insensitive) |
| `--cluster` | Cluster name (required; partial match supported) |
| `--yes` | Skip confirmation prompt and proceed immediately |
| `--resume-after` | Minutes to wait before automatically sending resume command. Useful for timed maintenance windows. Implies `--yes` (skips pause confirmation) |

---

## Behavior

- **Name matching:** Both `--pattern` and `--cluster` support partial name matching (case-insensitive).
- **Confirmation:** Prints matching groups and prompts for confirmation unless `--yes` is set.
- **`--resume-after N`:** Pauses groups, waits N minutes, then automatically resumes them. Useful for orchestrating maintenance windows that have known duration. Implies `--yes` so the pause is not confirmed.
- **Multiple matches:** If multiple groups match, all matching groups are paused/resumed together.

---

## Output

Lists matched protection groups with their current state and requests confirmation. After the action completes, prints summary of updated groups.

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.0 | 2026-04-06 | feat: initial release |
