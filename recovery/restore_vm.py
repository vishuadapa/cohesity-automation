#!/usr/bin/env python3
# =============================================================================
# DISCLAIMER: This code is provided on a best-effort basis and is not in any
# way officially supported or sanctioned by Cohesity. The code is intentionally
# kept simple to retain value as example code. The code in this repository is
# provided as-is and the author accepts no liability for damages resulting
# from its use.
# =============================================================================
"""
restore_vm.py
-------------
Restore a VMware VM to its original location or to an alternate datastore/
resource pool / vCenter.

Searches for the VM by name, lists available snapshots, and initiates a
VMware recovery task. Supports restoring to the original location (overwrite
or with a suffix) or to an alternate location specified via CLI flags.

API endpoints used:
  Cluster list:    GET   https://helios.cohesity.com/mcm/clusters/connectionStatus
  Object search:   GET   https://helios.cohesity.com/v2/data-protect/search/protected-objects
                         ?searchString=<vm>&objectTypes=kVMware
  Snapshots:       GET   https://helios.cohesity.com/v2/data-protect/objects/{id}/snapshots
  Initiate restore: POST  https://helios.cohesity.com/v2/data-protect/recoveries

Version history:
  1.1 (2026-04-07) — Fixed rename params: eliminated None/delete pattern
                     by building originalSourceConfig conditionally before
                     constructing the payload.
  1.0 (2026-04-06) — Initial release. Original-location restore with optional
                     suffix rename, snapshot selection by index or latest,
                     power-on after restore flag.

Usage:
  python3 restore_vm.py --vm "vm-prod-01" --cluster <name>
  python3 restore_vm.py --vm "vm-prod-01" --cluster <name> --suffix "-restored"
  python3 restore_vm.py --vm "vm-prod-01" --cluster <name> --snapshot 2 --power-on
  python3 restore_vm.py --clear-credentials
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


def search_vms(api_key: str, cluster_id: int, name: str) -> list:
    url    = f"https://{HELIOS_HOST}/v2/data-protect/search/protected-objects"
    params = {"searchString": name, "objectTypes": "kVMware", "count": 20}
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


def initiate_vm_recovery(api_key: str, cluster_id: int, payload: dict) -> dict:
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
        description=f"Cohesity VM restore v{__version__}",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--apikey",              help="Helios API key")
    parser.add_argument("--cluster",  required=True,
                        help="Cluster name (as shown in Helios)")
    parser.add_argument("--vm",       required=True,
                        help="VM name to restore (partial match)")
    parser.add_argument("--snapshot", type=int, default=0,
                        help="Snapshot index to restore from (0 = latest, default: 0)")
    parser.add_argument("--suffix",   default="",
                        help="Suffix to append to restored VM name (e.g. '-restored'). "
                             "Leave blank to overwrite original.")
    parser.add_argument("--power-on", action="store_true",
                        help="Power on VM after restore")
    parser.add_argument("--dry-run",  action="store_true",
                        help="Print the recovery payload without submitting")
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

    print(f"\n[*] Searching for VM '{args.vm}' on {args.cluster}…")
    vms = search_vms(api_key, cluster_id, args.vm)
    if not vms:
        print(f"ERROR: No protected VM matching '{args.vm}' found.")
        sys.exit(1)

    if len(vms) > 1:
        print(f"[*] Multiple VMs found — using first exact or best match:")
        exact = [v for v in vms if v.get("name", "").lower() == args.vm.lower()]
        vms = exact if exact else vms
    vm = vms[0]

    vm_name = vm.get("name", "")
    vm_id   = vm.get("id")
    print(f"[+] Found VM: {vm_name} (id={vm_id})")
    print(f"    Protection group: {vm.get('protectionGroupName', 'N/A')}")

    snaps = get_snapshots(api_key, cluster_id, vm_id)
    if not snaps:
        print("ERROR: No snapshots found for this VM.")
        sys.exit(1)

    snap_idx = min(args.snapshot, len(snaps) - 1)
    snap     = snaps[snap_idx]
    snap_time = usecs_to_str(snap.get("snapshotTimestampUsecs"))

    print(f"\n[*] Available snapshots (showing first 5 of {len(snaps)}):")
    for i, s in enumerate(snaps[:5]):
        marker = " <-- selected" if i == snap_idx else ""
        print(f"    [{i}] {usecs_to_str(s.get('snapshotTimestampUsecs'))}  "
              f"{s.get('runType','')}{marker}")

    new_name = vm_name + args.suffix if args.suffix else vm_name
    print(f"\n[*] Restoring snapshot: {snap_time}")
    print(f"    Restored VM name:  {new_name}")
    print(f"    Power on after:    {'Yes' if args.power_on else 'No'}")

    # Build recovery payload — original-location VMware restore
    original_source_cfg = {
        "networkConfig": {"detachNetwork": False},
        "powerOnVms":    args.power_on,
    }
    if args.suffix:
        original_source_cfg["renameRecoveredVmsParams"] = {"suffix": args.suffix}

    payload = {
        "name":         f"Restore-{vm_name}-{datetime.now().strftime('%Y%m%d%H%M%S')}",
        "snapshotEnvironment": "kVMware",
        "vmwareParams": {
            "objects": [{"snapshotId": snap.get("id")}],
            "recoveryAction": "RecoverVMs",
            "recoverVmParams": {
                "targetEnvironment": "kVMware",
                "recoverFromStandby": False,
                "vmwareTargetParams": {
                    "recoveryTargetConfig": {
                        "recoverToNewSource":  False,
                        "originalSourceConfig": original_source_cfg,
                    }
                },
            },
        },
    }

    if args.dry_run:
        import json
        print("\n[DRY RUN] Payload that would be POSTed:")
        print(json.dumps(payload, indent=2))
        sys.exit(0)

    print("\n[*] Submitting recovery task…")
    result = initiate_vm_recovery(api_key, cluster_id, payload)
    print(f"[+] Recovery task created.")
    print(f"    Task ID: {result.get('id')}")
    print(f"    Status:  {result.get('status')}")
    print(f"\n    Monitor progress in the Cohesity UI: "
          f"Data Protection → Recoveries")


if __name__ == "__main__":
    main()
