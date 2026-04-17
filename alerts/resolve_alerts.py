#!/usr/bin/env python3
# =============================================================================
# DISCLAIMER: This code is provided on a best-effort basis and is not in any
# way officially supported or sanctioned by Cohesity. The code is intentionally
# kept simple to retain value as example code. The code in this repository is
# provided as-is and the author accepts no liability for damages resulting
# from its use.
# =============================================================================
"""
resolve_alerts.py
-----------------
Bulk-resolve open Cohesity alerts matching a filter (severity, category,
alert code, or cluster name) via Helios routing.

Prints matching alerts and prompts for confirmation before resolving.
Use --yes to skip the confirmation prompt (for scripting/automation).

API endpoints used:
  Cluster list:  GET  https://helios.cohesity.com/mcm/clusters/connectionStatus
  Fetch alerts:  GET  https://helios.cohesity.com/v2/alerts
                      ?alertStateList=kOpen&maxAlerts=1000
  Resolve:       POST https://helios.cohesity.com/v2/alerts/resolve
                      body: {"alertIdList": [...]}

Version history:
  1.1 (2026-04-17) — Security: TLS verification enabled by default; added
                     --ca-bundle and --insecure CLI flags.
  1.0 (2026-04-06) — Initial release. Filter by severity, category, alert
                     code, or cluster name. Dry-run by default — explicit
                     --yes required to resolve. Prints match count and list
                     before any write action.

Usage:
  python3 resolve_alerts.py --severity kCritical
  python3 resolve_alerts.py --category kDisk --cluster <cluster-name> --yes
  python3 resolve_alerts.py --code CE01516014 --yes
  python3 resolve_alerts.py --ca-bundle /path/to/ca.crt  # corporate proxy cert
  python3 resolve_alerts.py --clear-credentials
"""

__version__ = "1.1"

import argparse
import getpass
import sys
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
    result = [{"name": c.get("name", ""), "clusterId": c.get("clusterId")}
              for c in clusters]
    print(f"[+] Connected to Helios — {len(result)} cluster(s)")
    return result


def get_open_alerts(api_key: str, cluster_id: int) -> list:
    url    = f"https://{HELIOS_HOST}/irisservices/api/v1/public/alerts"
    params = {"alertStateList": "kOpen", "maxAlerts": 1000}
    try:
        r = requests.get(url, headers=helios_headers(api_key, cluster_id),
                         params=params, verify=_verify, timeout=30)
        r.raise_for_status()
    except requests.exceptions.HTTPError:
        print(f"    WARN: alerts fetch failed ({r.status_code}) — skipping")
        return []
    data = r.json()
    return data if isinstance(data, list) else data.get("alerts", [])


def resolve_alerts(api_key: str, cluster_id: int, alert_ids: list) -> bool:
    url     = f"https://{HELIOS_HOST}/irisservices/api/v1/public/alerts/resolve"
    payload = {"alertIdList": alert_ids}
    try:
        r = requests.post(url, headers=helios_headers(api_key, cluster_id),
                          json=payload, verify=_verify, timeout=30)
        r.raise_for_status()
        return True
    except requests.exceptions.HTTPError:
        print(f"    ERROR: resolve failed ({r.status_code}): {r.text[:200]}")
        return False


def usecs_to_str(usecs: int) -> str:
    if not usecs:
        return ""
    return datetime.fromtimestamp(usecs / 1_000_000, tz=timezone.utc).strftime(
        "%Y-%m-%d %H:%M:%S UTC")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description=f"Bulk-resolve Cohesity alerts v{__version__}",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--apikey",   help="Helios API key (stored in keychain after first use)")
    parser.add_argument("--cluster",  help="Limit to one cluster by name")
    parser.add_argument("--severity", help="Filter by severity  (e.g. kCritical, kWarning, kInfo)")
    parser.add_argument("--category", help="Filter by category  (e.g. kDisk, kNode, kBackupRestore)")
    parser.add_argument("--code",     help="Filter by alert code (e.g. CE01516014)")
    parser.add_argument("--yes",      action="store_true",
                        help="Skip confirmation prompt and resolve immediately")
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

    if not any([args.severity, args.category, args.code, args.cluster]):
        print("ERROR: Specify at least one filter: --severity, --category, --code, or --cluster.")
        print("       Resolving ALL alerts without a filter is not supported (safety guard).")
        sys.exit(1)

    api_key  = get_api_key(args.apikey)
    clusters = get_helios_clusters(api_key)

    if args.cluster:
        clusters = [c for c in clusters
                    if c["name"].lower() == args.cluster.lower()]
        if not clusters:
            print(f"ERROR: Cluster '{args.cluster}' not found in Helios.")
            sys.exit(1)

    # Collect matching alerts per cluster
    to_resolve: dict = {}   # cluster_id -> list of alert IDs
    match_count = 0

    for c in clusters:
        alerts = get_open_alerts(api_key, c["clusterId"])
        matched = []
        for a in alerts:
            if args.severity and a.get("severity") != args.severity:
                continue
            if args.category and a.get("alertCategory") != args.category:
                continue
            if args.code and a.get("alertCode") != args.code:
                continue
            matched.append(a)

        if matched:
            to_resolve[c["clusterId"]] = (c["name"], matched)
            match_count += len(matched)

    if not match_count:
        print("\n[*] No matching open alerts found.")
        sys.exit(0)

    # Preview
    print(f"\n{'─'*70}")
    print(f"  Matching open alerts: {match_count}")
    print(f"{'─'*70}")
    for cid, (cname, alerts) in to_resolve.items():
        print(f"\n  Cluster: {cname}")
        for a in alerts:
            print(f"    [{a.get('severity',''):12s}] [{a.get('alertCategory',''):20s}]"
                  f"  {a.get('alertCode',''):14s}  {usecs_to_str(a.get('latestTimestampUsecs'))}")
    print(f"\n{'─'*70}")

    if not args.yes:
        ans = input(f"\nResolve {match_count} alert(s)? [y/N] ").strip().lower()
        if ans != "y":
            print("[*] Aborted — no alerts resolved.")
            sys.exit(0)

    # Resolve
    resolved_total = 0
    for cid, (cname, alerts) in to_resolve.items():
        ids = [a["id"] for a in alerts]
        print(f"\n[*] Resolving {len(ids)} alert(s) on {cname}…")
        if resolve_alerts(api_key, cid, ids):
            print(f"    [+] Done.")
            resolved_total += len(ids)
        else:
            print(f"    [-] Failed — check output above.")

    print(f"\n[+] Resolved {resolved_total}/{match_count} alert(s).")


if __name__ == "__main__":
    main()
