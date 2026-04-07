# clone_policy.py

Duplicates a protection policy to a new name. Useful for creating similar policies with minor variations or as a starting point for new policy development. Includes dry-run mode for validation.

Requires a **Helios API key**. The key is stored securely in the OS keychain after the first run.

Required packages:
```bash
pip install requests keyring
```

**Version:** 1.0

---

## Usage

```bash
# Clone a policy to a new name:
python3 clone_policy.py --source "Gold 30-Day" --name "Gold 60-Day" --cluster prod-cluster

# Clone by source policy ID (if name is ambiguous):
python3 clone_policy.py --source-id <policy-id> --name "New Policy" --cluster prod-cluster

# Validate the payload before creating:
python3 clone_policy.py --source "Silver" --name "Silver DR" --cluster prod-cluster --dry-run

# First run — prompts for Helios API key:
python3 clone_policy.py --apikey <key> --source "test" --name "test2" --cluster prod-cluster

# Remove stored credentials:
python3 clone_policy.py --clear-credentials
```

---

## CLI Options

| Option | Description |
|--------|-------------|
| `--apikey` | Helios API key (optional — keychain used if omitted, or you will be prompted) |
| `--clear-credentials` | Remove stored API key from the OS keychain and exit |
| `--source` | Source policy name to clone from (partial match supported) |
| `--source-id` | Source policy ID (use if policy name is ambiguous or unavailable) |
| `--name` | Name for the new policy (required) |
| `--cluster` | Cluster name (required; partial match supported) |
| `--dry-run` | Print the JSON payload that would be sent to create the policy without actually creating it. Useful for inspection and validation. |

---

## Behavior

- **Source matching:** Uses partial name matching for `--source`. If multiple policies match, the first match is cloned. Use `--source-id` to specify exactly.
- **Cloned configuration includes:**
  - All retention settings (local retention duration)
  - All replication targets and settings
  - All archival targets and settings
  - Schedule information (backup frequency, timing)
  - Policy-level settings (dedup, compression, encryption)
- **Server-managed fields are stripped:** ID, timestamps, and protection group usage counts are generated fresh by Cohesity when the new policy is created.
- **`--dry-run`:** Prints the full API payload JSON without submitting the POST request. Use this to inspect the policy configuration before creating it.

---

## Output

Confirmation message with cloned policy details: source policy name, new policy name, and basic configuration summary. If `--dry-run` is specified, prints the JSON payload instead of creating the policy.
