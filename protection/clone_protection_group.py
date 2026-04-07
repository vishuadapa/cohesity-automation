#!/usr/bin/env python3
# =============================================================================
# DISCLAIMER: This code is provided on a best-effort basis and is not in any
# way officially supported or sanctioned by Cohesity. The code is intentionally
# kept simple to retain value as example code. The code in this repository is
# provided as-is and the author accepts no liability for damages resulting
# from its use.
# =============================================================================
"""
clone_protection_group.py
--------------------------
Copy a Cohesity protection group configuration to a new group on the same
cluster. Fetches the source group's full config, strips server-managed fields,
renames it, and POSTs a new group.

Useful for creating similar groups for new workloads without retyping all
settings. The cloned group is created in a paused state so it doesn't
immediately run on creation.

API endpoints used:
  Cluster list:   GET   https://helios.cohesity.com/mcm/clusters/connectionStatus
  List groups:    GET   https://helios.cohesity.com/v2/data-protect/protection-groups
  Create group:   POST  https://helios.cohesity.com/v2/data-protect/protection-groups

Note: The object list (VMs, volumes, shares) from the source group is included
in the clone by default. Use --no-objects to create the group with an empty
object list so you can add new objects in the UI.

Version history:
  1.0 (2026-04-06) — Initial release. Find source by name, clone to new name,
                     --no-objects flag, --dry-run mode.

Usage:
  python3 clone_protection_group.py --source "VMware Prod" --name "VMware DR" \
          --cluster <name>
  python3 clone_protection_group.py --source "SQL Daily" --name "SQL Weekly" \
          --cluster <name> --no-objects --dry-run
  python3 clone_protection_group.py --clear-credentials
"""

__version__ = "1.0"

import argparse
import copy
import getpass
import sys
import urllib3

import requests

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

HELIOS_HOST     = "helios.cohesity.com"
_KR_SVC_HELIOS  = "cohesity_helios"
_KR_USER_HELIOS = "apikey"

# Fields the server generates / manages — must be stripped before POST
_STRIP_KEYS = {
    "id", "clusterId", "clusterIncarnationId", "regionId",
    "storageDomainId",      # keep policyId but not storageDomainId if cross-domain
    "createTime", "modificationTime", "version",
    "lastRun", "lastRunStatus", "numRuns",
    "isActive",             # will default to true on create
    "isDeleted", "isProtected",
    "permissions", "physicalParams",   # re-add manually if needed
}


# ---------------------------------------------------------------------------
# Credential helpers
# ---------------------------------------------------------------------------

def _keyring_available() -> bool:
    try:
        import keyring  # noqa: F401
        return True
    except ImportError:
        return False


def get_api_key(cli_key: str = None) -> str:
    if cli_key:
        return cli_key
    if _keyring_available():
        import keyring
        stored = keyring.get_password(_KR_SVC_HELIOS, _KR_USER_HELIOS)
        if stored:
            print("[*] Using Helios API key stored in system keychain.")
            return stored
        print("[*] No stored Helios API key found.")
    else:
        print("    NOTE: 'keyring' not installed — key will not be saved.")
        print("          Install with:  pip install keyring")
    key = getpass.getpass("    Enter Helios API key (will be saved to keychain): ").strip()
    if not key:
        print("ERROR: API key cannot be empty.")
        sys.exit(1)
    if _keyring_available():
        import keyring
        keyring.set_password(_KR_SVC_HELIOS, _KR_USER_HELIOS, key)
        print("[*] Helios API key saved to system keychain.")
    return key


def clear_stored_credentials():
    if not _keyring_available():
        print("ERROR: 'keyring' package not installed. Nothing to clear.")
        sys.exit(1)
    import keyring
    if keyring.get_password(_KR_SVC_HELIOS, _KR_USER_HELIOS):
        keyring.delete_password(_KR_SVC_HELIOS, _KR_USER_HELIOS)
        print("[*] Stored Helios API key removed from system keychain.")
    else:
        print("[*] No stored Helios API key found — nothing to clear.")


# ---------------------------------------------------------------------------
# Helios helpers
# ---------------------------------------------------------------------------

def helios_headers(api_key: str, cluster_id: int = None) -> dict:
    h = {"apiKey": api_key, "Accept": "application/json",
         "Content-Type": "application/json"}
    if cluster_id is not None:
        h["accessClusterId"] = str(cluster_id)
    return h


def get_helios_clusters(api_key: str) -> list:
    url = f"https://{HELIOS_HOST}/mcm/clusters/connectionStatus"
    try:
        r = requests.get(url, headers=helios_headers(api_key), verify=False, timeout=30)
        r.raise_for_status()
    except requests.exceptions.ConnectionError:
        print("ERROR: Cannot connect to Helios. Check your network/VPN.")
        sys.exit(1)
    except requests.exceptions.HTTPError:
        print(f"ERROR: Helios auth failed ({r.status_code}). Check your API key.")
        sys.exit(1)
    clusters = r.json()
    if not isinstance(clusters, list):
        clusters = clusters.get("clusterList", [])
    return [{"name": c.get("name", ""), "clusterId": c.get("clusterId")}
            for c in clusters]


def list_groups(api_key: str, cluster_id: int) -> list:
    url    = f"https://{HELIOS_HOST}/v2/data-protect/protection-groups"
    params = {"includeTenants": "true", "includeLastRunInfo": "false"}
    try:
        r = requests.get(url, headers=helios_headers(api_key, cluster_id),
                         params=params, verify=False, timeout=20)
        r.raise_for_status()
        return r.json().get("protectionGroups", [])
    except requests.exceptions.HTTPError:
        print(f"ERROR: Could not list groups ({r.status_code}).")
        sys.exit(1)


def create_group(api_key: str, cluster_id: int, payload: dict) -> dict:
    url = f"https://{HELIOS_HOST}/v2/data-protect/protection-groups"
    try:
        r = requests.post(url, headers=helios_headers(api_key, cluster_id),
                          json=payload, verify=False, timeout=20)
        r.raise_for_status()
        return r.json()
    except requests.exceptions.HTTPError:
        print(f"ERROR: Group creation failed ({r.status_code}): {r.text[:300]}")
        sys.exit(1)


def strip_server_fields(group: dict) -> dict:
    cloned = copy.deepcopy(group)
    for key in _STRIP_KEYS:
        cloned.pop(key, None)
    return cloned


def clear_objects(group: dict) -> dict:
    """Remove all protected objects from the group config."""
    for section in ("vmwareParams", "physicalParams", "oracleParams",
                    "mssqlParams", "genericNasParams", "nasParams"):
        if section in group:
            obj_key = next((k for k in group[section]
                            if "Objects" in k or "objects" in k), None)
            if obj_key:
                group[section][obj_key] = []
    return group


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description=f"Cohesity clone protection group v{__version__}",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--apikey",              help="Helios API key")
    parser.add_argument("--cluster",  required=True,
                        help="Cluster name (as shown in Helios)")
    parser.add_argument("--source",   required=True,
                        help="Source protection group name (partial match OK)")
    parser.add_argument("--name",     required=True,
                        help="Name for the new cloned group")
    parser.add_argument("--no-objects", action="store_true",
                        help="Clone policy/schedule but clear the object list")
    parser.add_argument("--dry-run",  action="store_true",
                        help="Print the payload without creating the group")
    parser.add_argument("--clear-credentials", action="store_true",
                        help="Remove stored Helios API key and exit")
    args = parser.parse_args()

    if args.clear_credentials:
        clear_stored_credentials()
        sys.exit(0)

    api_key  = get_api_key(args.apikey)
    clusters = get_helios_clusters(api_key)
    match_c  = [c for c in clusters
                if c["name"].lower() == args.cluster.lower()]
    if not match_c:
        print(f"ERROR: Cluster '{args.cluster}' not found in Helios.")
        sys.exit(1)
    cluster_id = match_c[0]["clusterId"]

    groups = list_groups(api_key, cluster_id)
    term   = args.source.lower()
    found  = [g for g in groups if term in g.get("name", "").lower()]

    if not found:
        print(f"ERROR: No group matching '{args.source}' on {args.cluster}.")
        print("  Available groups:")
        for g in groups:
            print(f"    {g.get('name')}")
        sys.exit(1)

    if len(found) > 1:
        print(f"[*] Multiple groups match '{args.source}':")
        for g in found:
            print(f"    {g.get('id')}  {g.get('name')}")
        print("    Using the first match.")

    source = found[0]
    print(f"\n[*] Source group: {source.get('name')} (id={source.get('id')})")
    print(f"[*] Clone name:   {args.name}")

    payload = strip_server_fields(source)
    payload["name"]     = args.name
    payload["isPaused"] = True   # start paused so it doesn't run immediately

    if args.no_objects:
        payload = clear_objects(payload)
        print("[*] Object list cleared (--no-objects).")

    if args.dry_run:
        import json
        print("\n[DRY RUN] Payload that would be POSTed:")
        print(json.dumps(payload, indent=2))
        sys.exit(0)

    result = create_group(api_key, cluster_id, payload)
    print(f"\n[+] Protection group cloned successfully.")
    print(f"    Name: {result.get('name')}")
    print(f"    ID:   {result.get('id')}")
    print(f"    Note: Group created in paused state — resume in the UI or "
          f"with pause_resume_groups.py --action resume.")


if __name__ == "__main__":
    main()
