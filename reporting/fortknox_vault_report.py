#!/usr/bin/env python3
"""
fortknox_vault_report.py
------------------------
Generates a CSV report of FortKnox (RPaaS) vault activity by protection group.

Uses the Helios-native activity API — a single POST call returns all FortKnox
archival runs across every connected cluster without iterating per-cluster APIs.

Each row in the report represents one FortKnox archival run.

API endpoints used:
  Activity:      POST https://helios.cohesity.com/v2/mcm/data-protect/protection-group/activity
                 body: {"activityTypes": ["ArchivalRun"], "archivalRunParams": {"isRpaas": true}}
  Vault storage: POST https://helios.cohesity.com/heliosreporting/api/v1/public/components/1600/preview
                 body: {"filters": [{"attribute": "managedBy", ...}], ...}

Identified from ELF script (cohesityInv.ps1 — Eduardo Figueredo / Brian Seltzer).
Requires Helios API key — FortKnox is a Helios-managed feature, direct cluster
auth is not applicable.

Version history:
  1.0 (2026-04-04) — Initial release. Per-cluster approach using protection
                     group runs filtered to kRpaas targets.
  1.1 (2026-04-04) — Removed invalid kArchiveRuns consumer type that caused
                     400 errors; vault storage showed N/A.
  2.4 (2026-04-04) — Vault Storage Consumed now from v1 cluster report
                     GET /irisservices/api/v1/public/reports/dataTransferToVaults
                     filtered by FortKnox vault IDs and date range. Per-
                     protection-group breakdown. Source: Helios single-cluster
                     report → Data Transferred to External Targets → PG tab.
  2.3 (2026-04-04) — Vault Storage Consumed now from component 1601
                     (Data Transferred to External Targets report) via the
                     Helios Reporting API, filtered per cluster by systemId
                     (clusterId:clusterIncarnationId). One call per cluster.
  2.2 (2026-04-04) — Vault Storage Consumed now per-protection-group using
                     cloudTotalPhysicalUsageBytes from stats/consumers
                     (consumerType=kProtectionRuns), fetched once per cluster.
                     Growth/transfer columns still from Helios Reporting API.
  2.1 (2026-04-04) — Fixed storage stats parsing (data at resp['component']['data']).
                     Added Vault Storage Growth, Vault Cumul Transferred, Vault
                     Period Transferred, Vault Daily Growth Pct, Vault Target Type
                     columns from the Helios Reporting API response.
  2.0 (2026-04-04) — Full rewrite using the Helios activity API
                     (POST /v2/mcm/data-protect/protection-group/activity with
                     isRpaas=true). Vault storage via Helios Reporting API
                     (/heliosreporting/api/v1/public/components/1600/preview).
                     Approach sourced from ELF script (cohesityInv.ps1).

Usage:
  python3 fortknox_vault_report.py --apikey <key>
  python3 fortknox_vault_report.py --apikey <key> --days 30
  python3 fortknox_vault_report.py --apikey <key> --start 2026-03-01 --end 2026-04-01
  python3 fortknox_vault_report.py --apikey <key> --cluster <cluster-name>
  python3 fortknox_vault_report.py --apikey <key> --vault <vault-name>
"""

__version__ = "2.4"

import argparse
import csv
import sys
import urllib3
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

import requests

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

HELIOS_HOST = "helios.cohesity.com"


# ---------------------------------------------------------------------------
# Auth helpers
# ---------------------------------------------------------------------------

def make_helios_headers(api_key: str, cluster_id: int = None) -> dict:
    h = {
        "apiKey":         api_key,
        "Accept":         "application/json",
        "Content-Type":   "application/json",
    }
    if cluster_id is not None:
        h["accessClusterId"] = str(cluster_id)
    return h


def get_helios_clusters(api_key: str) -> list:
    """
    Returns list of cluster dicts for all Helios-connected clusters.
    Each dict has: name, clusterId, clusterIncarnationId, systemId (id:incarnationId).
    """
    url = f"https://{HELIOS_HOST}/mcm/clusters/connectionStatus"
    try:
        r = requests.get(url, headers=make_helios_headers(api_key), verify=False, timeout=30)
        r.raise_for_status()
    except requests.exceptions.ConnectionError:
        print("ERROR: Cannot connect to Helios. Check your network/VPN.")
        sys.exit(1)
    except requests.exceptions.HTTPError as e:
        print(f"ERROR: Helios auth failed: {e}\n       Check your API key. Response: {r.text[:200]}")
        sys.exit(1)
    clusters = r.json()
    if not isinstance(clusters, list):
        clusters = clusters.get("clusterList", [])
    result = []
    for c in clusters:
        cid  = c.get("clusterId", "")
        incr = c.get("clusterIncarnationId", "")
        result.append({
            "name":       c.get("name", ""),
            "clusterId":  cid,
            "systemId":   f"{cid}:{incr}",
        })
    print(f"[+] Connected to Helios ({HELIOS_HOST}) — {len(result)} cluster(s)")
    return result


def get_fortknox_vault_ids(api_key: str, cluster_id: int) -> list:
    """
    Fetch FortKnox vault IDs configured on a cluster.
    Endpoint: GET /irisservices/api/v1/public/vaults?includeFortKnoxVault=true
    Returns: [vault_id, ...]
    """
    url = f"https://{HELIOS_HOST}/irisservices/api/v1/public/vaults"
    params = {"includeFortKnoxVault": "true"}
    try:
        r = requests.get(url, headers=make_helios_headers(api_key, cluster_id),
                         params=params, verify=False, timeout=30)
        r.raise_for_status()
        vaults = r.json() or []
        ids = [v["id"] for v in vaults if "id" in v]
        return ids
    except Exception as e:
        print(f"    WARNING: Could not fetch vault IDs for cluster {cluster_id}: {e}")
        return []


def get_pg_retained_storage(api_key: str, cluster_id: int, cluster_name: str,
                             vault_ids: list, start_msecs: int, end_msecs: int,
                             debug: bool = False) -> dict:
    """
    Fetch Storage Consumed for Retained Data per protection group.

    Endpoint: GET /irisservices/api/v1/public/reports/dataTransferToVaults
    Parameters:
      startTimeMsecs / endTimeMsecs — date range in milliseconds
      vaultIds                      — FortKnox vault IDs (one param per ID)

    Source: Helios UI single-cluster report → "Data Transferred to External
    Targets" → Protection Group tab.

    Returns: {(cluster_name, group_name, vault_name): retained_bytes}
    """
    if not vault_ids:
        return {}

    url = f"https://{HELIOS_HOST}/irisservices/api/v1/public/reports/dataTransferToVaults"
    params = [
        ("startTimeMsecs", start_msecs),
        ("endTimeMsecs",   end_msecs),
    ]
    for vid in vault_ids:
        params.append(("vaultIds", vid))

    try:
        r = requests.get(url, headers=make_helios_headers(api_key, cluster_id),
                         params=params, verify=False, timeout=60)
        r.raise_for_status()
        resp = r.json()

        if debug:
            import json
            print(f"\n[DEBUG] dataTransferToVaults for {cluster_name} — top-level keys: {list(resp.keys()) if isinstance(resp, dict) else type(resp)}")
            if isinstance(resp, dict):
                for k, v in resp.items():
                    if isinstance(v, list) and v:
                        print(f"[DEBUG] resp['{k}'][0] = {json.dumps(v[0], indent=2)[:800]}")
                        break

        lookup = {}
        # Response structure:
        #   dataTransferSummary[]:              vault-level summaries
        #     .vaultName                        vault name
        #     .dataTransferPerProtectionJob[]:  per-group breakdown
        #       .protectionJobName              protection group name
        #       .storageConsumed                retained bytes for this group+vault
        for vault_summary in (resp.get("dataTransferSummary") or []):
            vault_name = vault_summary.get("vaultName", "")
            for job in (vault_summary.get("dataTransferPerProtectionJob") or []):
                grp      = job.get("protectionJobName", "")
                retained = job.get("storageConsumed", 0)
                if grp:
                    lookup[(cluster_name, grp, vault_name)] = retained

        return lookup

    except Exception as e:
        print(f"    WARNING: Could not fetch retained storage for {cluster_name}: {e}")
        return {}


# ---------------------------------------------------------------------------
# Size / time helpers
# ---------------------------------------------------------------------------

def to_gb(byte_count) -> str:
    """
    Convert bytes to GB as a plain decimal string.
    Column headers carry the (GB) label so cell values stay numeric in Excel.
    """
    if not byte_count:
        return "N/A"
    return f"{byte_count / 1024 ** 3:.2f}"


def format_time(usecs: int, tz_name: str = "UTC") -> str:
    if not usecs:
        return "N/A"
    try:
        tz = ZoneInfo(tz_name)
    except ZoneInfoNotFoundError:
        tz = timezone.utc
    return datetime.fromtimestamp(usecs / 1_000_000, tz=tz).strftime("%Y-%m-%d %H:%M %Z")


# ---------------------------------------------------------------------------
# Date helpers
# ---------------------------------------------------------------------------

def now_usecs() -> int:
    return int(datetime.now(tz=timezone.utc).timestamp() * 1_000_000)


def date_to_usecs(date_str: str) -> int:
    try:
        dt = datetime.strptime(date_str, "%Y-%m-%d").replace(tzinfo=timezone.utc)
    except ValueError:
        print(f"ERROR: Invalid date '{date_str}'. Use YYYY-MM-DD.")
        sys.exit(1)
    return int(dt.timestamp() * 1_000_000)


def end_of_day_usecs(date_str: str) -> int:
    try:
        dt = datetime.strptime(date_str, "%Y-%m-%d").replace(
            hour=23, minute=59, second=59, tzinfo=timezone.utc)
    except ValueError:
        print(f"ERROR: Invalid date '{date_str}'. Use YYYY-MM-DD.")
        sys.exit(1)
    return int(dt.timestamp() * 1_000_000)


def resolve_date_range(args) -> tuple:
    """Return (start_usecs, end_usecs). Default: last 7 days."""
    if args.days:
        start = datetime.now(tz=timezone.utc) - timedelta(days=args.days)
        return int(start.timestamp() * 1_000_000), now_usecs()
    if args.start:
        return date_to_usecs(args.start), \
               (end_of_day_usecs(args.end) if args.end else now_usecs())
    if args.end:
        print("WARNING: --end given without --start. Ignoring.")
    # Default: last 7 days
    start = datetime.now(tz=timezone.utc) - timedelta(days=7)
    return int(start.timestamp() * 1_000_000), now_usecs()


# ---------------------------------------------------------------------------
# Helios Activity API  (source: ELF cohesityInv.ps1 HeliosFortKnoxActivity)
# ---------------------------------------------------------------------------

def get_fortknox_activity(api_key: str, start_usecs: int, end_usecs: int) -> list:
    """
    Fetch all FortKnox archival runs across all Helios-connected clusters.

    Endpoint: POST https://helios.cohesity.com/v2/mcm/data-protect/protection-group/activity
    Body:
      activityTypes:        ["ArchivalRun"]
      archivalRunParams:    {"isRpaas": true}
      startTimeUsecs:       <epoch microseconds>
      endTimeUsecs:         <epoch microseconds>

    Returns list of activity objects. Each has top-level fields:
      clusterName, protectionGroupName, policyName, type
    and nested archivalRunParams:
      startTimeUsecs, endTimeUsecs, status, archivalTargetName,
      rpaasRegionId, protectionEnvironmentType,
      logicalSizeBytes, physicalBytesTransferred, logicalBytesTransferred
    """
    url = f"https://{HELIOS_HOST}/v2/mcm/data-protect/protection-group/activity"
    body = {
        "activityTypes":     ["ArchivalRun"],
        "archivalRunParams": {"isRpaas": True},
        "startTimeUsecs":    start_usecs,
        "endTimeUsecs":      end_usecs,
    }
    try:
        r = requests.post(url, json=body, headers=make_helios_headers(api_key),
                          verify=False, timeout=60)
        r.raise_for_status()
    except requests.exceptions.HTTPError as e:
        print(f"ERROR: Activity API call failed ({r.status_code}): {r.text[:300]}")
        sys.exit(1)

    activity = r.json().get("activity") or []
    print(f"[+] Retrieved {len(activity)} FortKnox archival run(s)")
    return activity


# ---------------------------------------------------------------------------
# Helios Reporting API  (source: ELF cohesityInv.ps1 fortKnoxVaultStats)
# ---------------------------------------------------------------------------

def get_fortknox_storage_stats(api_key: str) -> list:
    """
    Fetch FortKnox vault storage consumed via the Helios Reporting API.

    Endpoint: POST https://helios.cohesity.com/heliosreporting/api/v1/public/components/1600/preview
    Filter:   managedBy = "FortKnox"

    Returns list of stat objects. Each has:
      managedBy, systemName (cluster), targetName (vault), scRetainedBytes

    Used to look up storage consumed per cluster+vault combination.
    """
    url = f"https://{HELIOS_HOST}/heliosreporting/api/v1/public/components/1600/preview"
    body = {
        "filters": [
            {
                "attribute":    "managedBy",
                "filterType":   "In",
                "inFilterParams": {
                    "attributeDataType":  "String",
                    "stringFilterValues": ["FortKnox"],
                },
            },
            {
                "attribute":   "date",
                "filterType":  "TimeRange",
                "timeRangeFilterParams": {"dateRange": "Last7Days"},
            },
        ],
        "sort":     None,
        "timezone": "UTC",
        "limit":    {"size": 10000},
    }
    try:
        r = requests.post(url, json=body, headers=make_helios_headers(api_key),
                          verify=False, timeout=60)
        r.raise_for_status()
        resp = r.json()

        # The reporting API wraps results inside resp["component"]
        component = resp.get("component") or {}

        # Results live at resp["component"]["data"]
        stats = component.get("data") or [] if isinstance(component, dict) else []
        print(f"[+] Retrieved storage stats for {len(stats)} vault record(s)")
        return stats
    except requests.exceptions.HTTPError as e:
        print(f"    WARNING: Vault storage stats unavailable ({r.status_code}): {r.text[:200]}")
        return []
    except Exception as e:
        print(f"    WARNING: Vault storage stats unavailable: {e}")
        return []


def build_storage_lookup(stats: list) -> dict:
    """
    Build a lookup dict keyed by (systemName, targetName) → full stat dict.
    Fields available per entry:
      scRetainedBytes, cumulativeDataTransferredBytes, dataTransferredBytes,
      scRetainedBytesGrowth, scRetainedDailyGrowthPercent, targetType
    """
    lookup = {}
    for s in stats:
        if isinstance(s, dict):
            cluster = s.get("systemName", "")
            vault   = s.get("targetName", "")
            if cluster and vault:
                lookup[(cluster, vault)] = s
    return lookup


# ---------------------------------------------------------------------------
# Report builder
# ---------------------------------------------------------------------------

# Size columns that get a "(GB)" suffix in the header
_GB_COL_RENAMES = {
    "Logical Size":               "Logical Size (GB)",
    "Logical Transferred":        "Logical Transferred (GB)",
    "Physical Transferred":       "Physical Transferred (GB)",
    "Vault Storage Consumed":     "Vault Storage Consumed (GB)",
    "Vault Storage Growth":       "Vault Storage Growth (GB)",
    "Vault Cumul Transferred":    "Vault Cumul Transferred (GB)",
    "Vault Period Transferred":   "Vault Period Transferred (GB)",
}


def build_report(activity: list, vault_stats_lookup: dict, pg_vault_lookup: dict,
                 cluster_filter: str = None, vault_filter: str = None) -> list:
    """
    Convert raw activity records into report rows.

    One row per activity record (each record = one FortKnox archival run).

    Storage sources:
      pg_vault_lookup   — {(cluster_name, group_name): cloudTotalPhysicalUsageBytes}
                          Per-protection-group vault storage from the cluster stats API.
                          Source: stats/consumers?consumerType=kProtectionRuns,
                                  field cloudTotalPhysicalUsageBytes
      vault_stats_lookup — {(cluster_name, vault_name): vault_stat_dict}
                          Vault-level aggregate stats from the Helios Reporting API.
                          Used for growth, cumulative transfer, and daily growth %.
    """
    rows = []

    for result in activity:
        cluster_name = result.get("clusterName", "Unknown")
        group_name   = result.get("protectionGroupName", "Unknown")
        policy_name  = result.get("policyName", "N/A")

        if cluster_filter and cluster_filter.lower() not in cluster_name.lower():
            continue

        arp = result.get("archivalRunParams", {})
        if not arp:
            continue

        vault_name = arp.get("archivalTargetName", "Unknown")

        if vault_filter and vault_filter.lower() not in vault_name.lower():
            continue

        start_usecs = arp.get("startTimeUsecs", 0)
        end_usecs   = arp.get("endTimeUsecs", 0)
        duration    = (
            round((end_usecs - start_usecs) / 1_000_000 / 60, 1)
            if start_usecs and end_usecs else "N/A"
        )

        # Retained storage: per-group if available, otherwise vault-level fallback
        pg_vault_bytes = (pg_vault_lookup.get((cluster_name, group_name, vault_name))
                          or pg_vault_lookup.get((cluster_name, group_name, ""))
                          or pg_vault_lookup.get((cluster_name, "", "")))

        # Vault-level aggregate stats (growth, cumulative transfer, etc.)
        vs = vault_stats_lookup.get((cluster_name, vault_name), {})

        row = {
            "Cluster":                  cluster_name,
            "Protection Group":         group_name,
            "Environment":              arp.get("protectionEnvironmentType", "N/A"),
            "Policy":                   policy_name,
            "Vault Name":               vault_name,
            "Vault Target Type":        vs.get("targetType", "N/A"),
            "Region":                   arp.get("rpaasRegionId", "N/A"),
            "Status":                   arp.get("status", "Unknown"),
            "Start Time":               format_time(start_usecs),
            "End Time":                 format_time(end_usecs),
            "Duration Mins":            duration,
            "Logical Size":             to_gb(arp.get("logicalSizeBytes")),
            "Logical Transferred":      to_gb(arp.get("logicalBytesTransferred")),
            "Physical Transferred":     to_gb(arp.get("physicalBytesTransferred")),
            # Vault retained storage — currently vault-level total (component 1601)
            # Will be updated to per-group once correct component ID is confirmed
            "Vault Storage Consumed":   to_gb(pg_vault_bytes),
            # Vault-level aggregate stats (same value for all rows on this cluster+vault)
            "Vault Storage Growth":     to_gb(vs.get("scRetainedBytesGrowth")),
            "Vault Cumul Transferred":  to_gb(vs.get("cumulativeDataTransferredBytes")),
            "Vault Period Transferred": to_gb(vs.get("dataTransferredBytes")),
            "Vault Daily Growth Pct":   (
                f"{vs['scRetainedDailyGrowthPercent']:.4f}"
                if vs.get("scRetainedDailyGrowthPercent") else "N/A"
            ),
        }

        rows.append({_GB_COL_RENAMES.get(k, k): v for k, v in row.items()})

    return rows


def write_csv(rows: list, output_file: str):
    if not rows:
        print("No FortKnox archival activity found for the requested scope.")
        return
    with open(output_file, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)
    print(f"\n[+] Report saved to: {output_file}")
    print(f"    Total rows: {len(rows)}")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def parse_args():
    parser = argparse.ArgumentParser(
        description=f"Cohesity FortKnox vault activity report  v{__version__}",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  Last 7 days (default):
    python3 fortknox_vault_report.py --apikey <key>

  Last 30 days:
    python3 fortknox_vault_report.py --apikey <key> --days 30

  Specific date range:
    python3 fortknox_vault_report.py --apikey <key> --start 2026-03-01 --end 2026-04-01

  One cluster only:
    python3 fortknox_vault_report.py --apikey <key> --cluster <cluster-name>

  One vault only:
    python3 fortknox_vault_report.py --apikey <key> --vault <vault-name>
        """
    )
    parser.add_argument("--apikey", required=True, help="Helios API key")
    parser.add_argument("--cluster",
                        help="Filter output to a specific cluster (partial name match)")
    parser.add_argument("--vault",
                        help="Filter output to a specific vault (partial name match)")
    parser.add_argument("--output", default="fortknox_report.csv",
                        help="Output CSV filename (default: fortknox_report.csv)")
    parser.add_argument("--debug", action="store_true",
                        help="Print first record from component 1601 to confirm field names")

    dg = parser.add_argument_group("Date range (default: last 7 days)")
    dg.add_argument("--days", type=int, metavar="N",
                    help="Last N days")
    dg.add_argument("--start", metavar="YYYY-MM-DD",
                    help="Start date (inclusive)")
    dg.add_argument("--end",   metavar="YYYY-MM-DD",
                    help="End date (inclusive, defaults to today)")
    return parser.parse_args()


def main():
    args = parse_args()
    start_usecs, end_usecs = resolve_date_range(args)

    s = datetime.fromtimestamp(start_usecs / 1e6, tz=timezone.utc).strftime("%Y-%m-%d")
    e = datetime.fromtimestamp(end_usecs   / 1e6, tz=timezone.utc).strftime("%Y-%m-%d")
    print(f"\n[*] FortKnox activity report  v{__version__}")
    print(f"[*] Date range: {s} → {e}")
    if args.cluster:
        print(f"[*] Cluster filter: '{args.cluster}'")
    if args.vault:
        print(f"[*] Vault filter:   '{args.vault}'")

    clusters = get_helios_clusters(args.apikey)
    cluster_by_name = {c["name"]: c for c in clusters}

    print("\n[*] Fetching FortKnox archival activity...")
    activity = get_fortknox_activity(args.apikey, start_usecs, end_usecs)

    # Per-group retained storage from /reports/dataTransferToVaults (cluster v1 API)
    print("[*] Fetching per-group retained storage (dataTransferToVaults)...")
    pg_vault_lookup = {}   # {(cluster_name, group_name, vault_name): retained_bytes}
    # Convert usecs → msecs for the v1 reports endpoint
    start_msecs = start_usecs // 1000
    end_msecs   = end_usecs   // 1000
    seen_clusters = {r.get("clusterName") for r in activity if r.get("clusterName")}
    for cname in sorted(seen_clusters):
        c = cluster_by_name.get(cname)
        if not c:
            print(f"    WARNING: No cluster info found for '{cname}', skipping")
            continue
        print(f"    {cname}")
        vault_ids = get_fortknox_vault_ids(args.apikey, c["clusterId"])
        if not vault_ids:
            print(f"      No FortKnox vaults found on {cname}")
            continue
        pg_stats = get_pg_retained_storage(args.apikey, c["clusterId"], cname,
                                           vault_ids, start_msecs, end_msecs,
                                           debug=args.debug)
        pg_vault_lookup.update(pg_stats)

    print("[*] Fetching vault-level aggregate stats (component 1600)...")
    vault_stats          = get_fortknox_storage_stats(args.apikey)
    vault_stats_lookup   = build_storage_lookup(vault_stats)

    rows = build_report(activity, vault_stats_lookup, pg_vault_lookup,
                        cluster_filter=args.cluster, vault_filter=args.vault)

    write_csv(rows, args.output)


if __name__ == "__main__":
    main()
