#!/usr/bin/env python3
"""
fortknox_vault_report.py
------------------------
Reports FortKnox (RPaaS) vault data transfer by protection group.

Iterates every Helios-connected cluster, fetches the "Data Transferred to
External Targets" report (GET /irisservices/api/v1/public/reports/dataTransferToVaults)
via Helios routing, and emits one row per protection group per vault.

API endpoints used:
  Cluster list:     GET  https://helios.cohesity.com/mcm/clusters/connectionStatus
  FortKnox vaults:  GET  https://helios.cohesity.com/irisservices/api/v1/public/vaults
                         ?includeFortKnoxVault=true
  Transfer report:  GET  https://helios.cohesity.com/irisservices/api/v1/public/reports/dataTransferToVaults
                         ?startTimeMsecs=<ms>&endTimeMsecs=<ms>&vaultIds=<id>&vaultIds=<id>

Requires Helios API key — FortKnox is a Helios-managed feature.

Version history:
  3.2 (2026-04-04) — Unit changed from GB to TB (÷ 1,000,000,000,000). UI
                     displays binary TiB; 1 TiB = 1.09951 TB so script values
                     are ~9.95% higher than UI figures — correct decimal
                     equivalent. Column headers updated to (TB). 4 decimal
                     places to preserve precision at TB scale.
  3.1 (2026-04-04) — GB conversion corrected to decimal SI (÷ 1,000,000,000).
                     Added Period Start / Period End columns so stacked CSV
                     runs can be used to build a storage utilization trend.
  3.0 (2026-04-04) — Full rewrite. Single source of truth: dataTransferToVaults
                     report API only. Removed Helios activity API and Reporting
                     API component calls — data was inaccurate. One row per
                     protection group per vault per cluster.
  2.4 (2026-04-04) — Per-group retained storage via dataTransferToVaults; nested
                     response structure fix.
  2.0 (2026-04-04) — Helios activity API + Reporting API component 1600.
  1.0 (2026-04-04) — Initial release.

Usage:
  python3 fortknox_vault_report.py --apikey <key>
  python3 fortknox_vault_report.py --apikey <key> --days 30
  python3 fortknox_vault_report.py --apikey <key> --start 2026-03-01 --end 2026-04-01
  python3 fortknox_vault_report.py --apikey <key> --cluster <cluster-name>
  python3 fortknox_vault_report.py --apikey <key> --vault <vault-name>
"""

__version__ = "3.2"

import argparse
import csv
import sys
import urllib3
from datetime import datetime, timedelta, timezone

import requests

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

HELIOS_HOST = "helios.cohesity.com"


# ---------------------------------------------------------------------------
# Auth / cluster helpers
# ---------------------------------------------------------------------------

def helios_headers(api_key: str, cluster_id: int = None) -> dict:
    h = {
        "apiKey":       api_key,
        "Accept":       "application/json",
        "Content-Type": "application/json",
    }
    if cluster_id is not None:
        h["accessClusterId"] = str(cluster_id)
    return h


def get_helios_clusters(api_key: str) -> list:
    """Return [{name, clusterId}, ...] for all Helios-connected clusters."""
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

    result = [{"name": c.get("name", ""), "clusterId": c.get("clusterId")}
              for c in clusters]
    print(f"[+] Connected to Helios — {len(result)} cluster(s)")
    return result


def get_fortknox_vault_ids(api_key: str, cluster_id: int) -> list:
    """Return [vault_id, ...] for FortKnox vaults on a cluster."""
    url = f"https://{HELIOS_HOST}/irisservices/api/v1/public/vaults"
    try:
        r = requests.get(url, headers=helios_headers(api_key, cluster_id),
                         params={"includeFortKnoxVault": "true"},
                         verify=False, timeout=30)
        r.raise_for_status()
        vaults = r.json() or []
        return [v["id"] for v in vaults if "id" in v]
    except Exception as e:
        print(f"    WARNING: Could not fetch vault IDs: {e}")
        return []


# ---------------------------------------------------------------------------
# Report API
# ---------------------------------------------------------------------------

def get_data_transfer_report(api_key: str, cluster_id: int, cluster_name: str,
                              vault_ids: list, start_msecs: int, end_msecs: int,
                              debug: bool = False) -> list:
    """
    Fetch "Data Transferred to External Targets" report for a cluster.

    Returns a flat list of row dicts, one per protection group per vault:
      cluster, vault_name, vault_type, protection_group,
      logical_bytes, physical_bytes, storage_consumed_bytes
    """
    if not vault_ids:
        return []

    url = f"https://{HELIOS_HOST}/irisservices/api/v1/public/reports/dataTransferToVaults"
    params = [("startTimeMsecs", start_msecs), ("endTimeMsecs", end_msecs)]
    for vid in vault_ids:
        params.append(("vaultIds", vid))

    try:
        r = requests.get(url, headers=helios_headers(api_key, cluster_id),
                         params=params, verify=False, timeout=60)
        r.raise_for_status()
        resp = r.json()
    except Exception as e:
        print(f"    WARNING: Report fetch failed for {cluster_name}: {e}")
        return []

    if debug:
        import json
        print(f"\n[DEBUG] {cluster_name} — keys: {list(resp.keys()) if isinstance(resp, dict) else type(resp)}")
        for k, v in (resp.items() if isinstance(resp, dict) else []):
            if isinstance(v, list) and v:
                print(f"[DEBUG] {k}[0] = {json.dumps(v[0], indent=2)[:1000]}")
                break

    rows = []
    for vault_summary in (resp.get("dataTransferSummary") or []):
        vault_name = vault_summary.get("vaultName", "")
        vault_type = vault_summary.get("vaultType", "N/A")
        for job in (vault_summary.get("dataTransferPerProtectionJob") or []):
            rows.append({
                "cluster":               cluster_name,
                "vault_name":            vault_name,
                "vault_type":            vault_type,
                "protection_group":      job.get("protectionJobName", ""),
                "logical_bytes":         job.get("numLogicalBytesTransferred", 0),
                "physical_bytes":        job.get("numPhysicalBytesTransferred", 0),
                "storage_consumed_bytes": job.get("storageConsumed", 0),
            })
    return rows


# ---------------------------------------------------------------------------
# Size / date helpers
# ---------------------------------------------------------------------------

def to_tb(byte_count) -> str:
    """Convert bytes to TB (decimal SI: 1 TB = 1,000,000,000,000 bytes).
    The UI displays binary TiB; 1 TiB = 1.09951 TB, so script values will be
    ~9.95% higher than UI figures — this is the correct decimal equivalent."""
    if not byte_count:
        return "0.00"
    return f"{byte_count / 1_000_000_000_000:.4f}"


def now_msecs() -> int:
    return int(datetime.now(tz=timezone.utc).timestamp() * 1000)


def date_to_msecs(date_str: str) -> int:
    try:
        dt = datetime.strptime(date_str, "%Y-%m-%d").replace(tzinfo=timezone.utc)
    except ValueError:
        print(f"ERROR: Invalid date '{date_str}'. Use YYYY-MM-DD.")
        sys.exit(1)
    return int(dt.timestamp() * 1000)


def end_of_day_msecs(date_str: str) -> int:
    try:
        dt = datetime.strptime(date_str, "%Y-%m-%d").replace(
            hour=23, minute=59, second=59, tzinfo=timezone.utc)
    except ValueError:
        print(f"ERROR: Invalid date '{date_str}'. Use YYYY-MM-DD.")
        sys.exit(1)
    return int(dt.timestamp() * 1000)


def resolve_date_range(args) -> tuple:
    """Return (start_msecs, end_msecs). Default: last 7 days."""
    if args.days:
        start = datetime.now(tz=timezone.utc) - timedelta(days=args.days)
        return int(start.timestamp() * 1000), now_msecs()
    if args.start:
        return date_to_msecs(args.start), \
               (end_of_day_msecs(args.end) if args.end else now_msecs())
    if args.end:
        print("WARNING: --end given without --start. Ignoring.")
    start = datetime.now(tz=timezone.utc) - timedelta(days=7)
    return int(start.timestamp() * 1000), now_msecs()


# ---------------------------------------------------------------------------
# CSV output
# ---------------------------------------------------------------------------

COLUMNS = [
    ("Period Start",               "period_start"),
    ("Period End",                 "period_end"),
    ("Cluster",                    "cluster"),
    ("Vault Name",                 "vault_name"),
    ("Vault Type",                 "vault_type"),
    ("Protection Group",           "protection_group"),
    ("Logical Transferred (TB)",   "logical_bytes"),
    ("Physical Transferred (TB)",  "physical_bytes"),
    ("Storage Consumed (TB)",      "storage_consumed_bytes"),
]


def write_csv(rows: list, output_file: str):
    if not rows:
        print("No FortKnox data found for the requested scope.")
        return
    headers = [col for col, _ in COLUMNS]
    with open(output_file, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=headers)
        writer.writeheader()
        for row in rows:
            writer.writerow({
                col: (to_tb(row[key]) if key.endswith("_bytes") else row[key])
                for col, key in COLUMNS
            })
    print(f"\n[+] Report saved to: {output_file}  ({len(rows)} rows)")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def parse_args():
    parser = argparse.ArgumentParser(
        description=f"Cohesity FortKnox data transfer report  v{__version__}",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python3 fortknox_vault_report.py --apikey <key>
  python3 fortknox_vault_report.py --apikey <key> --days 30
  python3 fortknox_vault_report.py --apikey <key> --start 2026-03-01 --end 2026-04-01
  python3 fortknox_vault_report.py --apikey <key> --cluster <name>
  python3 fortknox_vault_report.py --apikey <key> --vault <name>
        """
    )
    parser.add_argument("--apikey",  required=True, help="Helios API key")
    parser.add_argument("--cluster", help="Filter to a specific cluster (partial name match)")
    parser.add_argument("--vault",   help="Filter to a specific vault (partial name match)")
    parser.add_argument("--output",  default="fortknox_report.csv",
                        help="Output CSV filename (default: fortknox_report.csv)")
    parser.add_argument("--debug",   action="store_true",
                        help="Print raw API response samples")

    dg = parser.add_argument_group("Date range (default: last 7 days)")
    dg.add_argument("--days",  type=int, metavar="N", help="Last N days")
    dg.add_argument("--start", metavar="YYYY-MM-DD", help="Start date (inclusive)")
    dg.add_argument("--end",   metavar="YYYY-MM-DD",
                    help="End date (inclusive, defaults to today)")
    return parser.parse_args()


def main():
    args = parse_args()
    start_msecs, end_msecs = resolve_date_range(args)

    s = datetime.fromtimestamp(start_msecs / 1000, tz=timezone.utc).strftime("%Y-%m-%d")
    e = datetime.fromtimestamp(end_msecs   / 1000, tz=timezone.utc).strftime("%Y-%m-%d")
    print(f"\n[*] FortKnox data transfer report  v{__version__}")
    print(f"[*] Date range: {s} → {e}")
    if args.cluster:
        print(f"[*] Cluster filter: '{args.cluster}'")
    if args.vault:
        print(f"[*] Vault filter:   '{args.vault}'")

    clusters = get_helios_clusters(args.apikey)

    all_rows = []
    for c in clusters:
        cname = c["name"]
        if args.cluster and args.cluster.lower() not in cname.lower():
            continue

        print(f"  [{cname}] fetching vault IDs...")
        vault_ids = get_fortknox_vault_ids(args.apikey, c["clusterId"])
        if not vault_ids:
            print(f"  [{cname}] no FortKnox vaults — skipping")
            continue

        print(f"  [{cname}] {len(vault_ids)} vault(s) — fetching report...")
        rows = get_data_transfer_report(args.apikey, c["clusterId"], cname,
                                        vault_ids, start_msecs, end_msecs,
                                        debug=args.debug)

        if args.vault:
            rows = [r for r in rows if args.vault.lower() in r["vault_name"].lower()]

        for r in rows:
            r["period_start"] = s
            r["period_end"]   = e

        print(f"  [{cname}] {len(rows)} protection group row(s)")
        all_rows.extend(rows)

    write_csv(all_rows, args.output)


if __name__ == "__main__":
    main()
