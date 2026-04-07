# cohesity_auth.py

Unified authentication helper module for Cohesity scripts. Provides shared credential handling for both direct cluster authentication and Helios (SaaS) API key authentication. Credentials are stored securely in the OS keychain and are never written to disk in plaintext.

Designed to be imported by other scripts to keep authentication logic in one place. All scripts in this repository use this module.

Required packages:
```bash
pip install requests keyring
```

**Version:** 1.1

---

## Usage

Import in your scripts:

```python
from utils.cohesity_auth import (
    get_api_key,
    get_auth_token,
    make_headers,
    make_helios_headers,
    get_helios_clusters,
    clear_stored_credentials,
    HELIOS_HOST,
)
```

---

## Functions

### `get_api_key(cli_key=None) -> str`

Returns the Helios API key with priority:
1. `--apikey` CLI argument (one-time use, not stored)
2. Key stored in OS keychain under `service="cohesity_helios"` / `user="apikey"`
3. Interactive prompt — key is saved to keychain for future runs

**Example:**
```python
api_key = get_api_key(cli_args.apikey)  # from argparse
headers = make_helios_headers(api_key)
resp = requests.get("https://helios.cohesity.com/mcm/clusters/connectionStatus",
                    headers=headers, verify=False)
```

---

### `get_auth_token(cluster, username, password, domain="LOCAL") -> str`

Authenticates against a direct Cohesity cluster and returns a Bearer token.

- **cluster:** Cluster IP or hostname
- **username:** Username (default: `admin`)
- **password:** Password (will prompt interactively if None)
- **domain:** Auth domain — `LOCAL` or AD domain name (default: `LOCAL`)

Passwords are stored in the OS keychain under `service="cohesity_cluster"` / `user="<cluster>:<domain>:<username>"` for future runs.

**Example:**
```python
token = get_auth_token("192.168.1.100", "admin", password=None, domain="LOCAL")
headers = make_headers(token)
resp = requests.get("https://192.168.1.100/irisservices/api/v1/public/cluster",
                    headers=headers, verify=False)
```

---

### `make_headers(token: str) -> dict`

Returns standard headers for direct cluster API calls.

```python
headers = make_headers(token)
# Returns: {
#   "Authorization": "Bearer <token>",
#   "Accept": "application/json",
#   "Content-Type": "application/json"
# }
```

---

### `make_helios_headers(api_key: str, cluster_id: str = None) -> dict`

Returns standard headers for Helios API calls. Optionally routes to a specific cluster via the `accessClusterId` header.

```python
# For cluster list (no cluster routing):
headers = make_helios_headers(api_key)

# For cluster-specific operations (route to a specific cluster):
headers = make_helios_headers(api_key, cluster_id="cluster-123-abc")
# Returns: {
#   "apiKey": "<api_key>",
#   "accessClusterId": "cluster-123-abc",
#   "Accept": "application/json",
#   "Content-Type": "application/json"
# }
```

---

### `get_helios_clusters(api_key: str) -> list[dict]`

Queries Helios for the list of all connected clusters. Returns a list of cluster dictionaries with fields like `name`, `clusterId`, `incarnationId`, `status`, etc.

**Example:**
```python
clusters = get_helios_clusters(api_key)
for cluster in clusters:
    print(f"Cluster: {cluster['name']} ({cluster['clusterId']})")
```

---

### `clear_stored_credentials(cluster: str = None) -> None`

Removes stored credentials from the OS keychain.

- **No arguments:** Removes the stored Helios API key
- **`cluster` argument:** Removes the stored password for that cluster (format: IP or hostname)

**Example:**
```python
clear_stored_credentials()          # Remove Helios key
clear_stored_credentials("192.168.1.100")  # Remove cluster password
```

---

## Keychain Storage

Credentials are stored securely in the OS keychain:

| What | Service | User | Notes |
|-----|---------|------|-------|
| Helios API key | `cohesity_helios` | `apikey` | Never written to disk or process list |
| Cluster password | `cohesity_cluster` | `<cluster>:<domain>:<username>` | Keyed by cluster, domain, and user |

If `keyring` is unavailable or nothing is stored, the scripts fall back to interactive prompts.

---

## Constants

- **`HELIOS_HOST`** = `"helios.cohesity.com"` — Helios SaaS endpoint (read-only)

---

## Error Handling

Functions raise exceptions for:
- Invalid cluster credentials (403 Forbidden from cluster API)
- Invalid Helios API key (401 Unauthorized)
- Network errors (connection failures)

Callers should catch and handle these as needed.

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.1 | 2026-04-07 | fix: parse response JSON once in `get_auth_token`; validate access token non-empty before returning |
| 1.0 | 2026-04-06 | feat: initial module — extracted shared auth patterns from reporting scripts into a reusable module |
