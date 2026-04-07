# clone_protection_group.py

Clones a protection group's full configuration (policy, schedule, retention, replication, archival settings) to a new protection group. The cloned group is created in a **paused state** and does not run immediately.

Requires a **Helios API key**. The key is stored securely in the OS keychain after the first run.

Required packages:
```bash
pip install requests keyring
```

**Version:** 1.0

---

## Usage

```bash
# Clone a protection group with default settings (objects included):
python3 clone_protection_group.py --source "VMware Prod" --name "VMware DR" \
        --cluster prod-cluster

# Clone policy and schedule only, clear the object list (add objects in UI later):
python3 clone_protection_group.py --source "SQL Daily" --name "SQL Weekly" \
        --cluster prod-cluster --no-objects

# Preview the payload without creating the group:
python3 clone_protection_group.py --source "NAS Backup" --name "NAS Backup 2" \
        --cluster prod-cluster --dry-run

# Clone with full options:
python3 clone_protection_group.py --source "Oracle Prod" --name "Oracle DR" \
        --cluster prod-cluster --no-objects --dry-run

# First run — prompts for Helios API key:
python3 clone_protection_group.py --apikey <key> --source "test" --name "test2" --cluster prod-cluster

# Remove stored credentials:
python3 clone_protection_group.py --clear-credentials
```

---

## CLI Options

| Option | Description |
|--------|-------------|
| `--apikey` | Helios API key (optional — keychain used if omitted, or you will be prompted) |
| `--clear-credentials` | Remove stored API key from the OS keychain and exit |
| `--source` | Source protection group name to clone from (partial match supported) |
| `--name` | Name for the new protection group (required) |
| `--cluster` | Cluster name (required; partial match supported) |
| `--no-objects` | Clear the protected object list from the clone. Use this to clone policy/schedule only without copying the object list. Objects can be added in the UI after creation. |
| `--dry-run` | Print the JSON payload that would be sent to create the group without actually creating it. Useful for inspection and validation. |

---

## Behavior

- **Source matching:** Uses partial name matching for `--source` and `--cluster`. If multiple groups match, the first match is used.
- **Cloned state:** The new group is created in a **paused** state — it will not run immediately.
- **Cloned configuration includes:**
  - Protection policy (retention, replication, archival settings)
  - Backup schedule (full, incremental, log backup times)
  - Protected object list (unless `--no-objects` is specified)
  - All policy-level settings (encryption, dedup, compression, etc.)
- **`--no-objects`:** Strips the protected object list from the clone. Useful when you want to reuse a policy/schedule for a different set of objects. Add objects in the Cohesity UI after creation.
- **`--dry-run`:** Prints the API payload JSON without submitting the POST request. Use this to inspect the configuration before creating the group.

---

## Output

Summary of cloned protection group with new name, policy, and object count. If `--dry-run` is specified, prints the JSON payload instead of creating the group.
