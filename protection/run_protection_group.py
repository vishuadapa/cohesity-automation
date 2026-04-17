#!/usr/bin/env python3
# =============================================================================
# DISCLAIMER: This code is provided on a best-effort basis and is not in any
# way officially supported or sanctioned by Cohesity. The code is intentionally
# kept simple to retain value as example code. The code in this repository is
# provided as-is and the author accepts no liability for damages resulting
# from its use.
# =============================================================================
"""
run_protection_group.py
-----------------------
Trigger an on-demand backup run for one or more Cohesity protection groups.

Finds the group by name (partial match) on a specified cluster and POSTs a
run request. Supports specifying the run type (kRegular / kFull / kLog).
Optionally waits and polls for completion.

API endpoints used:
  Cluster list:    GET   https://helios.cohesity.com/mcm/clusters/connectionStatus
  Find group:      GET   https://helios.cohesity.com/v2/data-protect/protection-groups
                         ?names=<name>
  Trigger run:     POST  https://helios.cohesity.com/v2/data-protect/protection-groups/{id}/runs
  Poll run status: GET   https://helios.cohesity.com/v2/data-protect/protection-groups/{id}/runs
                         ?numRuns=1

Version history:
  1.1 (2026-04-17) — Security: TLS verification enabled by default; added
                     --ca-bundle and --insecure CLI flags.
  1.0 (2026-04-06) — Initial release. Find by name, trigger run, optional
                     --wait polling loop with configurable timeout.

Usage:
  python3 run_protection_group.py --group "VMware Daily" --cluster <cluster-name>
  python3 run_protection_group.py --group "SQL Prod" --run-type kFull --cluster <name>
  python3 run_protection_group.py --group "NAS Weekly" --cluster <name> --wait
  python3 run_protection_group.py --clear-credentials
"""

__version__ = "1.1"

import argparse
import getpass
import sys
import time
from datetime import datetime, timezone

import requests

HELIOS_HOST     = "helios.cohesity.com"
_KR_SVC_HELIOS  = "cohesity_helios"
_KR_USER_HELIOS = "apikey"

_verify = True   # overridden in main() via --insecure / --ca-bundle


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
        r = requests.get(url, headers=helios_headers(api_key), verify=_verify, timeout=30)
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


def find_groups(api_key: str, cluster_id: int, name_filter: str) -> list:
    url    = f"https://{HELIOS_HOST}/v2/data-protect/protection-groups"
    params = {"isActive": "true", "includeTenants": "true"}
    try:
        r = requests.get(url, headers=helios_headers(api_key, cluster_id),
                         params=params, verify=_verify, timeout=20)
        r.raise_for_status()
    except requests.exceptions.HTTPError:
        print(f"ERROR: Could not list protection groups ({r.status_code}).")
        sys.exit(1)
    groups = r.json().get("protectionGroups", [])
    term   = name_filter.lower()
    return [g for g in groups if term in g.get("name", "").lower()]


def trigger_run(api_key: str, cluster_id: int, group_id: str,
                run_type: str = "kRegular") -> bool:
    url     = (f"https://{HELIOS_HOST}/v2/data-protect/protection-groups"
               f"/{group_id}/runs")
    payload = {"runType": run_type}
    try:
        r = requests.post(url, headers=helios_headers(api_key, cluster_id),
                          json=payload, verify=_verify, timeout=20)
        r.raise_for_status()
        return True
    except requests.exceptions.HTTPError:
        print(f"    ERROR: trigger failed ({r.status_code}): {r.text[:200]}")
        return False


def poll_run(api_key: str, cluster_id: int, group_id: str,
             timeout_mins: int = 60) -> str:
    """Poll until the latest run is no longer Running. Returns final status."""
    deadline = time.time() + timeout_mins * 60
    url      = (f"https://{HELIOS_HOST}/v2/data-protect/protection-groups"
                f"/{group_id}/runs")
    params   = {"numRuns": 1, "includeObjectDetails": "false"}

    while time.time() < deadline:
        try:
            r = requests.get(url, headers=helios_headers(api_key, cluster_id),
                             params=params, verify=_verify, timeout=20)
            r.raise_for_status()
            runs = r.json().get("runs", [])
            if not runs:
                time.sleep(15)
                continue
            run    = runs[0]
            bk     = run.get("localBackupInfo", run.get("backupInfo", {}))
            status = bk.get("status", "")
            now    = datetime.now(tz=timezone.utc).strftime("%H:%M:%S")
            print(f"    [{now}] Run status: {status}", end="\r")
            if status not in ("Running", "Accepted", ""):
                print()
                return status
        except Exception:
            pass
        time.sleep(20)

    print()
    return "Timeout"


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description=f"Cohesity on-demand run trigger v{__version__}",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--apikey",              help="Helios API key")
    parser.add_argument("--cluster",  required=True,
                        help="Cluster name (as shown in Helios)")
    parser.add_argument("--group",    required=True,
                        help="Protection group name (partial match)")
    parser.add_argument("--run-type", default="kRegular",
                        choices=["kRegular", "kFull", "kLog"],
                        help="Run type (default: kRegular)")
    parser.add_argument("--wait",     action="store_true",
                        help="Wait and poll for run completion")
    parser.add_argument("--timeout",  type=int, default=60,
                        help="Poll timeout in minutes when --wait is set (default: 60)")
    parser.add_argument("--clear-credentials", action="store_true",
                        help="Remove stored Helios API key and exit")
    parser.add_argument("--ca-bundle", dest="ca_bundle", default=None, metavar="PATH",
                        help="Path to CA bundle for TLS verification (e.g. corporate proxy cert)")
    parser.add_argument("--insecure", action="store_true",
                        help="Disable TLS certificate verification (NOT recommended)")
    args = parser.parse_args()

    global _verify
    if args.insecure:
        _verify = False
        print("=" * 65)
        print("  WARN: --insecure — TLS certificate validation DISABLED.")
        print("        Credentials and tokens may be intercepted (MITM).")
        print("=" * 65)
        try:
            import urllib3 as _u3
            _u3.disable_warnings(_u3.exceptions.InsecureRequestWarning)
        except ImportError:
            requests.packages.urllib3.disable_warnings()
    elif args.ca_bundle:
        _verify = args.ca_bundle

    if args.clear_credentials:
        clear_stored_credentials()
        sys.exit(0)

    api_key  = get_api_key(args.apikey)
    clusters = get_helios_clusters(api_key)
    match    = [c for c in clusters
                if c["name"].lower() == args.cluster.lower()]
    if not match:
        print(f"ERROR: Cluster '{args.cluster}' not found in Helios.")
        sys.exit(1)
    cluster_id = match[0]["clusterId"]

    groups = find_groups(api_key, cluster_id, args.group)
    if not groups:
        print(f"ERROR: No protection group matching '{args.group}' found on {args.cluster}.")
        sys.exit(1)

    if len(groups) > 1:
        print(f"[*] Found {len(groups)} groups matching '{args.group}':")
        for g in groups:
            print(f"      {g.get('id')}  {g.get('name')}")
        print("    Triggering all matching groups.\n")

    for g in groups:
        gname = g.get("name")
        gid   = g.get("id")
        print(f"[*] Triggering {args.run_type} run on: {gname}")
        ok = trigger_run(api_key, cluster_id, gid, args.run_type)
        if not ok:
            continue
        print(f"    [+] Run triggered successfully.")

        if args.wait:
            print(f"    Waiting for completion (timeout={args.timeout}m)…")
            final = poll_run(api_key, cluster_id, gid, args.timeout)
            print(f"    [+] Final status: {final}")


if __name__ == "__main__":
    main()
