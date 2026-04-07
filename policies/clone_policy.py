#!/usr/bin/env python3
# =============================================================================
# DISCLAIMER: This code is provided on a best-effort basis and is not in any
# way officially supported or sanctioned by Cohesity. The code is intentionally
# kept simple to retain value as example code. The code in this repository is
# provided as-is and the author accepts no liability for damages resulting
# from its use.
# =============================================================================
"""
clone_policy.py
---------------
Duplicate an existing Cohesity protection policy as a starting point for
a new one. Fetches the source policy, strips its ID/metadata, optionally
renames it, and POSTs it back to the same cluster.

API endpoints used:
  Cluster list:   GET   https://helios.cohesity.com/mcm/clusters/connectionStatus
  List policies:  GET   https://helios.cohesity.com/v2/data-protect/policies
  Create policy:  POST  https://helios.cohesity.com/v2/data-protect/policies

Version history:
  1.0 (2026-04-06) — Initial release. Find source by name or ID, clone with
                     new name, optional Helios or direct cluster auth.
                     Prints new policy ID on success.

Usage:
  python3 clone_policy.py --source "Gold 30-Day" --name "Gold 60-Day" --cluster <name>
  python3 clone_policy.py --source-id <policy-id> --name "Cloned Policy" --cluster <name>
  python3 clone_policy.py --clear-credentials
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


def get_policies(api_key: str, cluster_id: int) -> list:
    url = f"https://{HELIOS_HOST}/v2/data-protect/policies"
    try:
        r = requests.get(url, headers=helios_headers(api_key, cluster_id),
                         verify=False, timeout=20)
        r.raise_for_status()
        return r.json().get("policies", [])
    except Exception as e:
        print(f"ERROR: Could not fetch policies: {e}")
        sys.exit(1)


def create_policy(api_key: str, cluster_id: int, payload: dict) -> dict:
    url = f"https://{HELIOS_HOST}/v2/data-protect/policies"
    try:
        r = requests.post(url, headers=helios_headers(api_key, cluster_id),
                          json=payload, verify=False, timeout=20)
        r.raise_for_status()
        return r.json()
    except requests.exceptions.HTTPError:
        print(f"ERROR: Policy creation failed ({r.status_code}): {r.text[:300]}")
        sys.exit(1)


# Server-side read-only / auto-generated fields to strip before cloning
_STRIP_KEYS = {
    "id", "createdTimeMsecs", "lastModifiedTimeMsecs",
    "numProtectionGroups", "numProtectionPolicies", "isReplicated",
    "isDeleted", "templateId",
}


def strip_server_fields(policy: dict) -> dict:
    cloned = copy.deepcopy(policy)
    for key in _STRIP_KEYS:
        cloned.pop(key, None)
    return cloned


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description=f"Cohesity clone policy v{__version__}",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--apikey",            help="Helios API key")
    parser.add_argument("--cluster", required=True,
                        help="Target cluster name (where source and clone both live)")
    parser.add_argument("--source",            help="Source policy name (partial match OK)")
    parser.add_argument("--source-id",         help="Source policy ID (exact)")
    parser.add_argument("--name",   required=True,
                        help="Name for the new cloned policy")
    parser.add_argument("--dry-run", action="store_true",
                        help="Print the payload that would be sent — do not create")
    parser.add_argument("--clear-credentials", action="store_true",
                        help="Remove stored Helios API key and exit")
    args = parser.parse_args()

    if args.clear_credentials:
        clear_stored_credentials()
        sys.exit(0)

    if not args.source and not args.source_id:
        print("ERROR: Specify either --source <name> or --source-id <id>.")
        sys.exit(1)

    api_key  = get_api_key(args.apikey)
    clusters = get_helios_clusters(api_key)
    match    = [c for c in clusters
                if c["name"].lower() == args.cluster.lower()]
    if not match:
        print(f"ERROR: Cluster '{args.cluster}' not found in Helios.")
        sys.exit(1)
    cluster_id = match[0]["clusterId"]

    policies = get_policies(api_key, cluster_id)

    # Find source
    source = None
    if args.source_id:
        for p in policies:
            if p.get("id") == args.source_id:
                source = p
                break
    else:
        term = args.source.lower()
        for p in policies:
            if term in p.get("name", "").lower():
                source = p
                break

    if not source:
        hint = args.source_id or args.source
        print(f"ERROR: Source policy '{hint}' not found on {args.cluster}.")
        print("       Available policies:")
        for p in policies:
            print(f"         {p.get('id',''):40s}  {p.get('name','')}")
        sys.exit(1)

    print(f"\n[*] Source policy: {source.get('name')} (id={source.get('id')})")
    print(f"[*] Clone name:    {args.name}")

    payload = strip_server_fields(source)
    payload["name"] = args.name

    if args.dry_run:
        import json
        print("\n[DRY RUN] Payload that would be POSTed:")
        print(json.dumps(payload, indent=2))
        sys.exit(0)

    result = create_policy(api_key, cluster_id, payload)
    print(f"\n[+] Policy created successfully.")
    print(f"    Name: {result.get('name')}")
    print(f"    ID:   {result.get('id')}")


if __name__ == "__main__":
    main()
