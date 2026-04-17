#!/usr/bin/env python3
# =============================================================================
# DISCLAIMER: This code is provided on a best-effort basis and is not in any
# way officially supported or sanctioned by Cohesity. The code is intentionally
# kept simple to retain value as example code. The code in this repository is
# provided as-is and the author accepts no liability for damages resulting
# from its use.
# =============================================================================
"""
recover_files.py
----------------
File/folder level recovery from a Cohesity snapshot to the original host
or an alternate target path.

Searches for a protected physical host or NAS source by name, selects a
snapshot, and submits a file recovery task for specified paths.

API endpoints used:
  Cluster list:    GET   https://helios.cohesity.com/mcm/clusters/connectionStatus
  Object search:   GET   https://helios.cohesity.com/v2/data-protect/search/protected-objects
                         ?searchString=<host>
  Snapshots:       GET   https://helios.cohesity.com/v2/data-protect/objects/{id}/snapshots
  Initiate restore: POST  https://helios.cohesity.com/v2/data-protect/recoveries

Version history:
  1.1 (2026-04-07) — Fixed alternate destination path field name:
                     absolutePath → alternateRestoreBaseDirectory (v2 API).
                     Removed redundant restore_params intermediary and
                     unused targetEnvironment variable.
  1.0 (2026-04-06) — Initial release. File/folder recovery to original or
                     alternate path, snapshot selection, dry-run mode.

Usage:
  python3 recover_files.py --source "linux-host-01" --paths /var/log /etc \
          --cluster <name>
  python3 recover_files.py --source "nas-share" --paths /data/reports \
          --cluster <name> --dest-path /tmp/restore
  python3 recover_files.py --source "win-host" --paths "C:\\Users\\jdoe" \
          --cluster <name> --snapshot 1 --dry-run
  python3 recover_files.py --clear-credentials
"""

__version__ = "1.2"

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
    return [{"name": c.get("name", ""), "clusterId": c.get("clusterId")}
            for c in clusters]


def usecs_to_str(usecs: int) -> str:
    if not usecs:
        return ""
    return datetime.fromtimestamp(usecs / 1_000_000, tz=timezone.utc).strftime(
        "%Y-%m-%d %H:%M:%S UTC")


def search_objects(api_key: str, cluster_id: int, name: str) -> list:
    url    = f"https://{HELIOS_HOST}/v2/data-protect/search/protected-objects"
    params = {"searchString": name, "count": 20}
    try:
        r = requests.get(url, headers=helios_headers(api_key, cluster_id),
                         params=params, verify=_verify, timeout=20)
        r.raise_for_status()
        return r.json().get("objects", [])
    except Exception:
        return []


def get_snapshots(api_key: str, cluster_id: int, object_id: int) -> list:
    url = f"https://{HELIOS_HOST}/v2/data-protect/objects/{object_id}/snapshots"
    try:
        r = requests.get(url, headers=helios_headers(api_key, cluster_id),
                         verify=_verify, timeout=20)
        r.raise_for_status()
        return r.json().get("snapshots", [])
    except Exception:
        return []


def initiate_file_recovery(api_key: str, cluster_id: int, payload: dict) -> dict:
    url = f"https://{HELIOS_HOST}/v2/data-protect/recoveries"
    try:
        r = requests.post(url, headers=helios_headers(api_key, cluster_id),
                          json=payload, verify=_verify, timeout=30)
        r.raise_for_status()
        return r.json()
    except requests.exceptions.HTTPError:
        print(f"ERROR: Recovery initiation failed ({r.status_code}): {r.text[:300]}")
        sys.exit(1)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description=f"Cohesity file/folder recovery v{__version__}",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--apikey",              help="Helios API key")
    parser.add_argument("--cluster",  required=True,
                        help="Cluster name (as shown in Helios)")
    parser.add_argument("--source",   required=True,
                        help="Protected source name to recover from (partial match)")
    parser.add_argument("--paths",    nargs="+", required=True,
                        help="File/folder paths to recover (space-separated)")
    parser.add_argument("--snapshot", type=int, default=0,
                        help="Snapshot index (0 = latest, default: 0)")
    parser.add_argument("--dest-path",
                        help="Alternate destination path. "
                             "Omit to restore to original location.")
    parser.add_argument("--overwrite", action="store_true",
                        help="Overwrite existing files at destination")
    parser.add_argument("--dry-run",  action="store_true",
                        help="Print payload without submitting")
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
    match_c  = [c for c in clusters
                if c["name"].lower() == args.cluster.lower()]
    if not match_c:
        print(f"ERROR: Cluster '{args.cluster}' not found in Helios.")
        sys.exit(1)
    cluster_id = match_c[0]["clusterId"]

    print(f"\n[*] Searching for '{args.source}' on {args.cluster}…")
    objects = search_objects(api_key, cluster_id, args.source)
    if not objects:
        print(f"ERROR: No protected object matching '{args.source}' found.")
        sys.exit(1)

    if len(objects) > 1:
        exact = [o for o in objects
                 if o.get("name", "").lower() == args.source.lower()]
        objects = exact if exact else objects

    obj     = objects[0]
    obj_name = obj.get("name", "")
    obj_id   = obj.get("id")
    print(f"[+] Found object: {obj_name} (id={obj_id})  "
          f"[{obj.get('objectType','')}]")

    snaps = get_snapshots(api_key, cluster_id, obj_id)
    if not snaps:
        print("ERROR: No snapshots found for this object.")
        sys.exit(1)

    snap_idx  = min(args.snapshot, len(snaps) - 1)
    snap      = snaps[snap_idx]
    snap_time = usecs_to_str(snap.get("snapshotTimestampUsecs"))

    print(f"\n[*] Snapshot selected: [{snap_idx}] {snap_time}  "
          f"{snap.get('runType','')}")
    print(f"    Paths to recover:  {args.paths}")
    dest = args.dest_path or "[original location]"
    print(f"    Destination:       {dest}")
    print(f"    Overwrite:         {'Yes' if args.overwrite else 'No'}")

    # Build recovery payload
    recover_file_params = {
        "filesAndFolders": [{"absolutePath": p} for p in args.paths],
        "targetEntity": {"id": obj.get("sourceId", obj_id)},
        "overwriteExisting": args.overwrite,
        "preserveAttributes": True,
    }
    if args.dest_path:
        recover_file_params["alternateRestoreBaseDirectory"] = args.dest_path

    payload = {
        "name":         (f"FileRestore-{obj_name}-"
                         f"{datetime.now().strftime('%Y%m%d%H%M%S')}"),
        "snapshotEnvironment": obj.get("environment", "kPhysical"),
        "physicalParams": {
            "objects": [{"snapshotId": snap.get("id")}],
            "recoveryAction": "RecoverFiles",
            "recoverFileAndFolderParams": recover_file_params,
        },
    }

    if args.dry_run:
        import json
        print("\n[DRY RUN] Payload that would be POSTed:")
        print(json.dumps(payload, indent=2))
        sys.exit(0)

    print("\n[*] Submitting file recovery task…")
    result = initiate_file_recovery(api_key, cluster_id, payload)
    print(f"[+] Recovery task created.")
    print(f"    Task ID: {result.get('id')}")
    print(f"    Status:  {result.get('status')}")
    print(f"\n    Monitor progress in the Cohesity UI: "
          f"Data Protection → Recoveries")


if __name__ == "__main__":
    main()
