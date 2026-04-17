#!/usr/bin/env python3
# =============================================================================
# DISCLAIMER: This code is provided on a best-effort basis and is not in any
# way officially supported or sanctioned by Cohesity. The code is intentionally
# kept simple to retain value as example code. The code in this repository is
# provided as-is and the author accepts no liability for damages resulting
# from its use.
# =============================================================================
"""
pause_resume_groups.py
----------------------
Bulk pause or resume Cohesity protection groups by name pattern.

Useful for maintenance windows — pause all groups matching a tag or name
prefix before a host reboot, then resume them after.

Prints matching groups and confirms before making changes (bypass with --yes).
Supports an optional --resume-after flag to automatically resume after a
specified delay (minutes).

API endpoints used:
  Cluster list:    GET   https://helios.cohesity.com/mcm/clusters/connectionStatus
  List groups:     GET   https://helios.cohesity.com/v2/data-protect/protection-groups
  Pause/resume:    POST  https://helios.cohesity.com/v2/data-protect/protection-groups/states
                         body: {"action": "kPause"|"kResume", "ids": [...]}

Version history:
  1.1 (2026-04-17) — Security: TLS verification enabled by default; added
                     --ca-bundle and --insecure CLI flags.
  1.0 (2026-04-06) — Initial release. Bulk pause/resume by name pattern,
                     confirmation prompt, optional timed auto-resume.

Usage:
  python3 pause_resume_groups.py --action pause --pattern "VMware" --cluster <name>
  python3 pause_resume_groups.py --action resume --pattern "SQL" --cluster <name> --yes
  python3 pause_resume_groups.py --action pause --pattern "NAS" --cluster <name> \
          --resume-after 60
  python3 pause_resume_groups.py --clear-credentials
"""

__version__ = "1.1"

import argparse
import getpass
import sys
import time

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


def list_groups(api_key: str, cluster_id: int) -> list:
    url    = f"https://{HELIOS_HOST}/v2/data-protect/protection-groups"
    params = {"includeTenants": "true"}
    try:
        r = requests.get(url, headers=helios_headers(api_key, cluster_id),
                         params=params, verify=_verify, timeout=20)
        r.raise_for_status()
        return r.json().get("protectionGroups", [])
    except Exception:
        return []


def set_group_state(api_key: str, cluster_id: int, action: str,
                    group_ids: list) -> bool:
    """action: 'kPause' or 'kResume'"""
    url     = (f"https://{HELIOS_HOST}/v2/data-protect/protection-groups/states")
    payload = {"action": action, "ids": group_ids}
    try:
        r = requests.post(url, headers=helios_headers(api_key, cluster_id),
                          json=payload, verify=_verify, timeout=20)
        r.raise_for_status()
        return True
    except requests.exceptions.HTTPError:
        print(f"    ERROR: state change failed ({r.status_code}): {r.text[:200]}")
        return False


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description=f"Cohesity bulk pause/resume protection groups v{__version__}",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--apikey",              help="Helios API key")
    parser.add_argument("--cluster",  required=True,
                        help="Cluster name (as shown in Helios)")
    parser.add_argument("--action",   required=True,
                        choices=["pause", "resume"],
                        help="Action to perform: pause or resume")
    parser.add_argument("--pattern",  required=True,
                        help="Name pattern to match (case-insensitive substring)")
    parser.add_argument("--yes",      action="store_true",
                        help="Skip confirmation prompt")
    parser.add_argument("--resume-after", type=int, metavar="MINUTES",
                        help="After pausing, automatically resume after N minutes")
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

    if args.resume_after and args.action != "pause":
        print("ERROR: --resume-after only makes sense with --action pause.")
        sys.exit(1)

    api_key    = get_api_key(args.apikey)
    clusters   = get_helios_clusters(api_key)
    match_c    = [c for c in clusters
                  if c["name"].lower() == args.cluster.lower()]
    if not match_c:
        print(f"ERROR: Cluster '{args.cluster}' not found in Helios.")
        sys.exit(1)
    cluster_id = match_c[0]["clusterId"]

    all_groups = list_groups(api_key, cluster_id)
    term       = args.pattern.lower()
    matched    = [g for g in all_groups
                  if term in g.get("name", "").lower()]

    if not matched:
        print(f"[*] No groups matching '{args.pattern}' found on {args.cluster}.")
        sys.exit(0)

    action_str = "PAUSE" if args.action == "pause" else "RESUME"
    print(f"\n{'─'*60}")
    print(f"  {action_str} — {len(matched)} group(s) matching '{args.pattern}':")
    print(f"{'─'*60}")
    for g in matched:
        paused = " [currently paused]" if g.get("isPaused") else ""
        print(f"    {g.get('name')}{paused}")
    print(f"{'─'*60}")

    if not args.yes:
        ans = input(f"\n{action_str} {len(matched)} group(s)? [y/N] ").strip().lower()
        if ans != "y":
            print("[*] Aborted.")
            sys.exit(0)

    group_ids  = [g["id"] for g in matched]
    api_action = "kPause" if args.action == "pause" else "kResume"

    print(f"\n[*] Sending {api_action} for {len(group_ids)} group(s)…")
    if set_group_state(api_key, cluster_id, api_action, group_ids):
        print(f"[+] {action_str} applied successfully.")
    else:
        sys.exit(1)

    # Auto-resume after delay
    if args.resume_after:
        print(f"\n[*] Waiting {args.resume_after} minute(s) before resuming…")
        time.sleep(args.resume_after * 60)
        print(f"[*] Sending kResume for {len(group_ids)} group(s)…")
        if set_group_state(api_key, cluster_id, "kResume", group_ids):
            print(f"[+] RESUME applied successfully.")
        else:
            sys.exit(1)


if __name__ == "__main__":
    main()
