# Protection

Scripts for managing protection groups (jobs), policies, and schedules. Use cases include bulk operations, pause/resume during maintenance windows, and on-demand backup triggers.

**API:** Uses Cohesity v2 REST API endpoints via Helios.

## Scripts

| Script | README | Description | Auth |
|--------|--------|-------------|------|
| `run_protection_group.py` | [docs](run_protection_group.md) | Trigger an on-demand backup run (kRegular / kFull / kLog) | Helios |
| `pause_resume_groups.py` | [docs](pause_resume_groups.md) | Bulk pause or resume groups by name pattern with optional timed auto-resume | Helios |
| `clone_protection_group.py` | [docs](clone_protection_group.md) | Copy a group's config to a new group (policy, schedule, objects) | Helios |

---

## Common auth pattern

```bash
python3 <script>.py --apikey <key> --cluster <cluster-name> [options]
```

API key is stored in the OS keychain after the first use.

---

## TLS / certificate options

All scripts support the following flags for environments with custom PKI or self-signed certificates:

| Flag | Description |
|------|-------------|
| `--ca-bundle PATH` | Path to a PEM-format CA bundle (`.pem`; extension does not matter, file must be PEM-encoded). Use for self-signed cluster certs or corporate HTTPS-inspection proxies |
| `--insecure` | Disable TLS certificate verification — **not recommended for production** |

---

## run_protection_group.py

```bash
python3 run_protection_group.py --group "VMware Daily" --cluster prod-cluster
python3 run_protection_group.py --group "SQL Prod" --run-type kFull --cluster prod-cluster
python3 run_protection_group.py --group "NAS" --cluster prod-cluster --wait --timeout 90
python3 run_protection_group.py --group "VMware Daily" --cluster prod-cluster --ca-bundle /path/to/ca.pem
```

`--cluster` is required. `--group` uses partial name matching — if multiple groups match, all are triggered.

| Flag | Description |
|------|-------------|
| `--run-type` | kRegular (default), kFull, or kLog |
| `--wait` | Poll and display run status until completion |
| `--timeout` | Poll timeout in minutes when `--wait` is set (default: 60) |

---

## pause_resume_groups.py

```bash
python3 pause_resume_groups.py --action pause --pattern "VMware" --cluster prod-cluster
python3 pause_resume_groups.py --action resume --pattern "SQL" --cluster prod-cluster --yes
python3 pause_resume_groups.py --action pause --pattern "NAS" --cluster prod-cluster \
        --resume-after 60
python3 pause_resume_groups.py --action pause --pattern "VMware" --cluster prod-cluster \
        --ca-bundle /path/to/ca.pem
```

Prints matching groups and prompts for confirmation before changing state. Use `--yes` to skip.

`--resume-after N` pauses groups, waits N minutes, then automatically sends resume — useful for scripted maintenance windows.

---

## clone_protection_group.py

```bash
python3 clone_protection_group.py --source "VMware Prod" --name "VMware DR" \
        --cluster prod-cluster
python3 clone_protection_group.py --source "SQL Daily" --name "SQL Weekly" \
        --cluster prod-cluster --no-objects --dry-run
python3 clone_protection_group.py --source "VMware Prod" --name "VMware DR" \
        --cluster prod-cluster --ca-bundle /path/to/ca.pem
```

Clones a group's full configuration (policy, schedule, settings) to a new group. The clone is created in a **paused** state so it doesn't run immediately.

| Flag | Description |
|------|-------------|
| `--no-objects` | Clear the object list — clone policy/schedule only, add objects in the UI |
| `--dry-run` | Print the payload without creating the group |
