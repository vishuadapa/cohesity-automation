# run_protection_group.py

Triggers on-demand backup runs for Cohesity protection groups. Supports run type selection (Regular, Full, Log), polling for completion with timeout, and batch operations via name pattern matching.

Requires a **Helios API key**. The key is stored securely in the OS keychain after the first run.

Required packages:
```bash
pip install requests keyring
```

**Version:** 1.0

---

## Usage

```bash
# Trigger a regular run for a protection group:
python3 run_protection_group.py --group "VMware Daily" --cluster prod-cluster

# Full backup run:
python3 run_protection_group.py --group "SQL Prod" --run-type kFull --cluster prod-cluster

# Trigger and wait for completion (default 60 min timeout):
python3 run_protection_group.py --group "NAS" --cluster prod-cluster --wait

# Wait with custom timeout (90 minutes):
python3 run_protection_group.py --group "VMware Daily" --cluster prod-cluster --wait --timeout 90

# Batch: match multiple groups by pattern:
python3 run_protection_group.py --group "daily" --cluster prod-cluster

# First run — prompts for Helios API key and saves to keychain:
python3 run_protection_group.py --apikey <key> --group "test" --cluster prod-cluster

# Remove stored credentials:
python3 run_protection_group.py --clear-credentials
```

---

## CLI Options

| Option | Description |
|--------|-------------|
| `--apikey` | Helios API key (optional — keychain used if omitted, or you will be prompted) |
| `--clear-credentials` | Remove stored API key from the OS keychain and exit |
| `--group` | Protection group name (partial match supported; triggers all matching groups) |
| `--cluster` | Cluster name (required; partial match supported) |
| `--run-type` | `kRegular` (default), `kFull`, or `kLog` |
| `--wait` | Poll for run completion; print status until done or timeout reached |
| `--timeout` | Poll timeout in minutes (default: 60) |

---

## Behavior

- **Name matching:** Both `--group` and `--cluster` support partial name matching. If multiple groups match, all are triggered.
- **Dry-run:** No dry-run mode — runs are triggered immediately after confirmation (no confirmation prompt by default; use a script wrapper if you need one).
- **Run types:**
  - `kRegular` — Incremental backup
  - `kFull` — Full backup
  - `kLog` — Log backup (for database protection groups)
- **`--wait` mode:** Polls the run status every 10 seconds until completion or timeout. Prints run status, objects, and bytes processed.

---

## Output

Summary of triggered runs with run ID and status. If `--wait` is set, continuous polling output until completion.

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.0 | 2026-04-06 | feat: initial release |
