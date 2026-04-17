#!/usr/bin/env python3
# =============================================================================
# DISCLAIMER: This code is provided on a best-effort basis and is not in any
# way officially supported or sanctioned by Cohesity. The code is intentionally
# kept simple to retain value as example code. The code in this repository is
# provided as-is and the author accepts no liability for damages resulting
# from its use.
# =============================================================================
"""
alert_to_csv.py
---------------
Export Cohesity alert history for a date range to CSV.

One row per alert across all Helios-connected clusters (or a single named
cluster). Includes both open and resolved alerts. Output is a plain CSV
suitable for Excel pivot tables, SIEM ingestion, or ticketing system imports.

API endpoints used:
  Cluster list:  GET  https://helios.cohesity.com/mcm/clusters/connectionStatus
  Alerts:        GET  https://helios.cohesity.com/v2/alerts
                      ?startDateUsecs=<us>&endDateUsecs=<us>&maxAlerts=1000

Version history:
  1.1 (2026-04-17) — Security: TLS verification enabled by default; added
                     --ca-bundle and --insecure CLI flags. CSV formula
                     injection hardening via _safe_cell() on all API data.
  1.0 (2026-04-06) — Initial release. Full alert history export to CSV with
                     configurable date range, severity and category filters,
                     and all key alert fields.

Usage:
  python3 alert_to_csv.py --days 30
  python3 alert_to_csv.py --start 2026-03-01 --end 2026-04-01
  python3 alert_to_csv.py --days 7 --severity kCritical
  python3 alert_to_csv.py --cluster <name> --days 14
  python3 alert_to_csv.py --ca-bundle /path/to/ca.pem  # corporate proxy cert
  python3 alert_to_csv.py --clear-credentials
"""

__version__ = "1.1"

import argparse
import csv
import getpass
import os
import sys
from datetime import datetime, timedelta, timezone

import requests

HELIOS_HOST     = "helios.cohesity.com"
_KR_SVC_HELIOS  = "cohesity_helios"
_KR_USER_HELIOS = "apikey"

_verify = True   # overridden in main() via --insecure / --ca-bundle

_FORMULA_PREFIXES = ("=", "+", "-", "@", "\t", "\r")


def _safe_cell(val):
    if isinstance(val, str) and val and val[0] in _FORMULA_PREFIXES:
        return "'" + val
    return val


CSV_FIELDS = [
    "Cluster", "Alert ID", "Alert Code", "Severity", "Category", "State",
    "Alert Name", "Description", "First Occurred (UTC)", "Last Updated (UTC)",
    "Resolved (UTC)", "Acknowledged", "Node ID", "Entity Name",
]


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


def default_output_path() -> str:
    timestamp   = datetime.now().strftime("%Y%m%d_%H%M%S")
    script_name = os.path.splitext(os.path.basename(__file__))[0]
    return os.path.join(os.getcwd(), f"{script_name}_v{__version__}_{timestamp}.csv")


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


def usecs_to_str(usecs: int) -> str:
    if not usecs:
        return ""
    return datetime.fromtimestamp(usecs / 1_000_000, tz=timezone.utc).strftime(
        "%Y-%m-%d %H:%M:%S")


def get_alerts(api_key: str, cluster_id: int, start_usecs: int,
               end_usecs: int, debug: bool = False) -> list:
    url    = f"https://{HELIOS_HOST}/irisservices/api/v1/public/alerts"
    params = {
        "startDateUsecs": start_usecs,
        "endDateUsecs":   end_usecs,
        "maxAlerts":      1000,
    }
    if debug:
        print(f"    GET {url}  params={params}")
    try:
        r = requests.get(url, headers=helios_headers(api_key, cluster_id),
                         params=params, verify=_verify, timeout=30)
        r.raise_for_status()
    except requests.exceptions.HTTPError:
        print(f"    WARN: alerts fetch failed ({r.status_code}) — skipping cluster")
        return []
    data = r.json()
    return data if isinstance(data, list) else data.get("alerts", [])


# ---------------------------------------------------------------------------
# Row builder
# ---------------------------------------------------------------------------

def alert_to_row(a: dict, cluster_name: str) -> dict:
    props      = {p.get("key"): p.get("value") for p in a.get("propertyList", [])}
    doc        = a.get("alertDocument", {})
    return {
        "Cluster":              cluster_name,
        "Alert ID":             a.get("id", ""),
        "Alert Code":           a.get("alertCode", ""),
        "Severity":             a.get("severity", ""),
        "Category":             a.get("alertCategory", ""),
        "State":                a.get("alertState", ""),
        "Alert Name":           doc.get("alertName", ""),
        "Description":          doc.get("alertDescription", ""),
        "First Occurred (UTC)": usecs_to_str(a.get("firstTimestampUsecs")),
        "Last Updated (UTC)":   usecs_to_str(a.get("latestTimestampUsecs")),
        "Resolved (UTC)":       usecs_to_str(a.get("resolvedTimestampUsecs")),
        "Acknowledged":         "Yes" if a.get("acknowledged") else "No",
        "Node ID":              props.get("nodeId", ""),
        "Entity Name":          props.get("entityName", ""),
    }


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description=f"Cohesity alert history CSV export v{__version__}",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--apikey",            help="Helios API key (stored in keychain after first use)")
    parser.add_argument("--cluster",           help="Filter to one Helios-connected cluster by name")
    parser.add_argument("--days",   type=int,  help="Last N days (default: 7)")
    parser.add_argument("--start",             help="Start date YYYY-MM-DD (UTC)")
    parser.add_argument("--end",               help="End date YYYY-MM-DD (UTC, defaults to today)")
    parser.add_argument("--severity",          help="Filter by severity (kCritical, kWarning, kInfo)")
    parser.add_argument("--category",          help="Filter by category (kDisk, kNode, kBackupRestore, …)")
    parser.add_argument("--output",            help="Output .csv path (auto-generated if omitted)")
    parser.add_argument("--debug",  action="store_true", help="Print API request details")
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

    # Resolve date range
    now = datetime.now(tz=timezone.utc)
    if args.start:
        start_dt = datetime.strptime(args.start, "%Y-%m-%d").replace(tzinfo=timezone.utc)
    elif args.days:
        start_dt = now - timedelta(days=args.days)
    else:
        start_dt = now - timedelta(days=7)

    end_dt = (datetime.strptime(args.end, "%Y-%m-%d").replace(tzinfo=timezone.utc)
              if args.end else now)

    start_usecs = int(start_dt.timestamp() * 1_000_000)
    end_usecs   = int(end_dt.timestamp() * 1_000_000)

    api_key = get_api_key(args.apikey)
    output  = args.output or default_output_path()

    print(f"\n[*] Alert window: {start_dt.strftime('%Y-%m-%d')} → {end_dt.strftime('%Y-%m-%d')} UTC")

    clusters = get_helios_clusters(api_key)
    if args.cluster:
        clusters = [c for c in clusters
                    if c["name"].lower() == args.cluster.lower()]
        if not clusters:
            print(f"ERROR: Cluster '{args.cluster}' not found in Helios.")
            sys.exit(1)

    rows = []
    for c in clusters:
        print(f"\n[*] Fetching alerts: {c['name']}")
        alerts = get_alerts(api_key, c["clusterId"], start_usecs, end_usecs,
                            debug=args.debug)
        for a in alerts:
            if args.severity and a.get("severity") != args.severity:
                continue
            if args.category and a.get("alertCategory") != args.category:
                continue
            rows.append(alert_to_row(a, c["name"]))
        print(f"    {len(alerts)} alert(s) fetched")

    with open(output, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_FIELDS)
        writer.writeheader()
        writer.writerows([{k: _safe_cell(v) for k, v in row.items()} for row in rows])

    print(f"\n[+] Report saved: {output}")
    print(f"    Total rows: {len(rows)}")


if __name__ == "__main__":
    main()
