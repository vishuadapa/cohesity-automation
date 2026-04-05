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
  3.9 (2026-04-04) — Output changed from CSV to Excel (.xlsx) via openpyxl.
                     In --mode trend, one chart sheet is added per protection
                     group showing Storage Consumed over time. Multiple vaults
                     for the same group appear as separate series. Header row
                     styled and columns auto-fitted. Requires: pip install openpyxl.
  3.8 (2026-04-04) — Added startup note explaining zero transfer values and
                     that Storage Consumed is the primary metric. Same caveat
                     added to README.
  3.7 (2026-04-04) — Added --mode trend. In trend mode the date range is split
                     into individual calendar days (cluster timezone); each day
                     gets its own API call and rows so users can plot a daily
                     storage utilization trend per protection group. Vault IDs
                     are fetched once per cluster and reused across all days.
                     Default remains --mode summary (existing behaviour).
  3.6 (2026-04-04) — Timezone auto-detected from cluster via GET /cluster.
                     Removed --timezone flag — no longer needed. Date
                     boundaries computed in cluster local time automatically.
                     Startup banner shows detected timezone and resolved
                     timestamps for verification.
  3.5 (2026-04-04) — Added --timezone flag. Date boundaries (--start/--end/
                     --days) are now computed in the specified timezone so the
                     script matches the Helios UI which uses the cluster's local
                     time (e.g. America/Chicago for US Central). Startup banner
                     shows resolved timestamps with timezone label.
  3.4 (2026-04-04) — Added --start-msecs / --end-msecs flags to pass exact
                     millisecond timestamps from Chrome DevTools, eliminating
                     time boundary mismatches vs the UI. --debug now prints
                     the exact request URL for direct comparison with Chrome.
  3.3 (2026-04-04) — Output raw bytes — no conversion. Column headers updated
                     to (Bytes). Allows direct comparison to API payload values
                     confirmed via Chrome DevTools.
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

__version__ = "3.9"

import argparse
import sys
import urllib3
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

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


def get_cluster_timezone(api_key: str, cluster_id: int, cluster_name: str) -> str:
    """Query the cluster's configured timezone via GET /irisservices/api/v1/public/cluster."""
    url = f"https://{HELIOS_HOST}/irisservices/api/v1/public/cluster"
    try:
        r = requests.get(url, headers=helios_headers(api_key, cluster_id),
                         verify=False, timeout=15)
        r.raise_for_status()
        tz = r.json().get("timezone", "")
        if tz:
            return tz
    except Exception:
        pass
    print(f"    WARNING: Could not query timezone for {cluster_name}, falling back to UTC")
    return "UTC"


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
        if debug:
            import json
            print(f"\n[DEBUG] {cluster_name} — exact request URL:\n        {r.url}")
        r.raise_for_status()
        resp = r.json()
    except Exception as e:
        print(f"    WARNING: Report fetch failed for {cluster_name}: {e}")
        return []

    if debug:
        import json
        print(f"[DEBUG] {cluster_name} — response keys: {list(resp.keys()) if isinstance(resp, dict) else type(resp)}")
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

def fmt_bytes(byte_count) -> str:
    if not byte_count:
        return "0"
    return str(int(byte_count))


def resolve_tz(tz_name: str):
    """Return a tzinfo object for tz_name, falling back to UTC on error."""
    try:
        return ZoneInfo(tz_name)
    except (ZoneInfoNotFoundError, KeyError):
        print(f"WARNING: Unknown timezone '{tz_name}', falling back to UTC.")
        return timezone.utc


def now_msecs() -> int:
    return int(datetime.now(tz=timezone.utc).timestamp() * 1000)


def date_to_msecs(date_str: str, tz) -> int:
    """Start of day (00:00:00) in tz."""
    try:
        dt = datetime.strptime(date_str, "%Y-%m-%d").replace(tzinfo=tz)
    except ValueError:
        print(f"ERROR: Invalid date '{date_str}'. Use YYYY-MM-DD.")
        sys.exit(1)
    return int(dt.timestamp() * 1000)


def end_of_day_msecs(date_str: str, tz) -> int:
    """End of day (23:59:59) in tz."""
    try:
        dt = datetime.strptime(date_str, "%Y-%m-%d").replace(
            hour=23, minute=59, second=59, tzinfo=tz)
    except ValueError:
        print(f"ERROR: Invalid date '{date_str}'. Use YYYY-MM-DD.")
        sys.exit(1)
    return int(dt.timestamp() * 1000)


def resolve_date_range(args, tz) -> tuple:
    """Return (start_msecs, end_msecs). Default: last 7 days."""
    if args.start_msecs or args.end_msecs:
        start = int(args.start_msecs) if args.start_msecs else now_msecs() - 7 * 86400 * 1000
        end   = int(args.end_msecs)   if args.end_msecs   else now_msecs()
        return start, end
    if args.days:
        start = datetime.now(tz=tz) - timedelta(days=args.days)
        return int(start.timestamp() * 1000), now_msecs()
    if args.start:
        return date_to_msecs(args.start, tz), \
               (end_of_day_msecs(args.end, tz) if args.end else now_msecs())
    if args.end:
        print("WARNING: --end given without --start. Ignoring.")
    start = datetime.now(tz=tz) - timedelta(days=7)
    return int(start.timestamp() * 1000), now_msecs()


# ---------------------------------------------------------------------------
# Trend helper
# ---------------------------------------------------------------------------

def iter_days(start_msecs: int, end_msecs: int, tz):
    """
    Yield (day_start_msecs, day_end_msecs, date_str) for each calendar day
    in the range [start_msecs, end_msecs], computed in tz.
    """
    # Snap to the start of the calendar day containing start_msecs
    start_dt = datetime.fromtimestamp(start_msecs / 1000, tz=tz).replace(
        hour=0, minute=0, second=0, microsecond=0)
    end_dt   = datetime.fromtimestamp(end_msecs   / 1000, tz=tz)

    current = start_dt
    while current <= end_dt:
        day_start = int(current.timestamp() * 1000)
        day_end   = int(current.replace(hour=23, minute=59, second=59).timestamp() * 1000)
        yield day_start, day_end, current.strftime("%Y-%m-%d")
        current += timedelta(days=1)


# ---------------------------------------------------------------------------
# Excel output
# ---------------------------------------------------------------------------

COLUMNS = [
    ("Period Start",                 "period_start"),
    ("Period End",                   "period_end"),
    ("Cluster",                      "cluster"),
    ("Vault Name",                   "vault_name"),
    ("Vault Type",                   "vault_type"),
    ("Protection Group",             "protection_group"),
    ("Logical Transferred (Bytes)",  "logical_bytes"),
    ("Physical Transferred (Bytes)", "physical_bytes"),
    ("Storage Consumed (Bytes)",     "storage_consumed_bytes"),
]


def _safe_sheet_name(name: str, existing: list) -> str:
    """Sanitise and deduplicate an Excel sheet name (max 31 chars)."""
    safe = name.translate(str.maketrans("", "", r":\/?*[]"))
    safe = safe[:31]
    if safe in existing:
        safe = safe[:28] + f"_{len(existing)}"
    return safe


def _add_trend_charts(wb, all_rows: list):
    """
    Add one chart sheet per (cluster, protection_group).
    Each sheet contains the date/storage-consumed data plus a line chart.
    If a protection group archives to multiple vaults, each vault becomes
    a separate series on the same chart.
    """
    try:
        from openpyxl.chart import LineChart, Reference
    except ImportError:
        print("WARNING: openpyxl chart module unavailable — charts skipped.")
        return

    # Build: {(cluster, group): {vault: {date: consumed_bytes}}}
    pg_data: dict = {}
    for r in all_rows:
        key   = (r["cluster"], r["protection_group"])
        vault = r["vault_name"]
        date  = r["period_start"]
        consumed = int(r.get("storage_consumed_bytes") or 0)
        pg_data.setdefault(key, {}).setdefault(vault, {})[date] = consumed

    for (cluster, group), vault_data in sorted(pg_data.items()):
        all_dates = sorted({d for vd in vault_data.values() for d in vd})
        vaults    = sorted(vault_data.keys())
        nrows     = len(all_dates)

        sheet_name = _safe_sheet_name(group, [s.title for s in wb.worksheets])
        ws = wb.create_sheet(title=sheet_name)

        # --- data table (top of sheet) ---
        ws.append(["Date"] + vaults)
        for d in all_dates:
            ws.append([d] + [vault_data[v].get(d, 0) for v in vaults])

        # Column widths
        ws.column_dimensions["A"].width = 14
        for i in range(len(vaults)):
            from openpyxl.utils import get_column_letter
            ws.column_dimensions[get_column_letter(i + 2)].width = 28

        # --- line chart ---
        chart = LineChart()
        chart.title  = f"{cluster}  |  {group}"
        chart.style  = 10
        chart.height = 15
        chart.width  = 32
        chart.y_axis.title = "Storage Consumed (Bytes)"
        chart.x_axis.title = "Date"

        data_ref = Reference(ws, min_col=2, min_row=1,
                             max_col=len(vaults) + 1, max_row=nrows + 1)
        chart.add_data(data_ref, titles_from_data=True)

        cats_ref = Reference(ws, min_col=1, min_row=2, max_row=nrows + 1)
        chart.set_categories(cats_ref)

        ws.add_chart(chart, f"A{nrows + 4}")

    print(f"    Trend charts added: {len(pg_data)} protection group(s)")


def write_excel(rows: list, output_file: str, mode: str):
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl.utils import get_column_letter
    except ImportError:
        print("ERROR: openpyxl is required.  Install with:  pip install openpyxl")
        sys.exit(1)

    if not rows:
        print("No FortKnox data found for the requested scope.")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Report"

    headers = [col for col, _ in COLUMNS]

    # Header row styling
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(fill_type="solid", fgColor="1F4E79")
    ws.append(headers)
    for cell in ws[1]:
        cell.font      = header_font
        cell.fill      = header_fill
        cell.alignment = Alignment(horizontal="center")

    # Data rows — byte columns written as integers so Excel can sort/chart them
    for row in rows:
        ws.append([
            int(row[key]) if key.endswith("_bytes") else row[key]
            for _, key in COLUMNS
        ])

    # Auto-fit column widths
    for col_idx, (col_name, _) in enumerate(COLUMNS, start=1):
        col_letter = get_column_letter(col_idx)
        max_len = max(
            len(col_name),
            max((len(str(ws.cell(r, col_idx).value or "")) for r in range(2, ws.max_row + 1)),
                default=0)
        )
        ws.column_dimensions[col_letter].width = min(max_len + 2, 45)

    # Freeze header row
    ws.freeze_panes = "A2"

    # Trend charts (trend mode only)
    if mode == "trend":
        _add_trend_charts(wb, rows)

    wb.save(output_file)
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
    parser.add_argument("--output",  default="fortknox_report.xlsx",
                        help="Output Excel filename (default: fortknox_report.xlsx)")
    parser.add_argument("--mode",    choices=["summary", "trend"], default="summary",
                        help="summary (default): one row per protection group covering the full "
                             "date range. trend: one row per protection group per day, enabling "
                             "storage utilization trend lines over time.")
    parser.add_argument("--debug",   action="store_true",
                        help="Print raw API response samples")

    dg = parser.add_argument_group("Date range (default: last 7 days)")
    dg.add_argument("--days",        type=int, metavar="N",        help="Last N days")
    dg.add_argument("--start",       metavar="YYYY-MM-DD",         help="Start date (inclusive)")
    dg.add_argument("--end",         metavar="YYYY-MM-DD",         help="End date (inclusive, defaults to today)")
    dg.add_argument("--start-msecs", metavar="MS", dest="start_msecs",
                    help="Exact start timestamp in milliseconds (paste from Chrome DevTools URL)")
    dg.add_argument("--end-msecs",   metavar="MS", dest="end_msecs",
                    help="Exact end timestamp in milliseconds (paste from Chrome DevTools URL)")
    return parser.parse_args()


def main():
    args = parse_args()

    print(f"\n[*] FortKnox data transfer report  v{__version__}")
    if args.cluster:
        print(f"[*] Cluster filter: '{args.cluster}'")
    if args.vault:
        print(f"[*] Vault filter:   '{args.vault}'")

    clusters = get_helios_clusters(args.apikey)

    # Pick the reference cluster for timezone detection:
    # use the first cluster matching --cluster filter, or the first overall.
    ref_clusters = [c for c in clusters
                    if not args.cluster or args.cluster.lower() in c["name"].lower()]
    if not ref_clusters:
        print("ERROR: No clusters match the given --cluster filter.")
        sys.exit(1)

    ref = ref_clusters[0]
    print(f"[*] Querying timezone from '{ref['name']}'...")
    tz_name = get_cluster_timezone(args.apikey, ref["clusterId"], ref["name"])
    tz      = resolve_tz(tz_name)
    print(f"[*] Cluster timezone: {tz_name}")

    start_msecs, end_msecs = resolve_date_range(args, tz)
    s = datetime.fromtimestamp(start_msecs / 1000, tz=tz).strftime("%Y-%m-%d %H:%M %Z")
    e = datetime.fromtimestamp(end_msecs   / 1000, tz=tz).strftime("%Y-%m-%d %H:%M %Z")
    print(f"[*] Date range: {s} → {e}")
    print(f"[*] Mode: {args.mode}")
    print()
    print("    NOTE: Storage Consumed (Bytes) is the primary metric — it reflects total")
    print("    retained storage in the vault for each protection group across all snapshots.")
    print("    Logical/Physical Transferred show 0 for groups that had no archival run")
    print("    within the queried window. This is expected; all rows are included.")
    print()

    # Filter cluster list once
    target_clusters = [c for c in clusters
                       if not args.cluster or args.cluster.lower() in c["name"].lower()]

    # Fetch vault IDs once per cluster (reused across all days in trend mode)
    cluster_vault_ids = {}
    for c in target_clusters:
        cname = c["name"]
        print(f"  [{cname}] fetching vault IDs...")
        vault_ids = get_fortknox_vault_ids(args.apikey, c["clusterId"])
        if not vault_ids:
            print(f"  [{cname}] no FortKnox vaults — skipping")
        else:
            cluster_vault_ids[cname] = (c["clusterId"], vault_ids)

    if not cluster_vault_ids:
        print("No clusters with FortKnox vaults found.")
        sys.exit(0)

    all_rows = []

    if args.mode == "summary":
        for cname, (cid, vault_ids) in cluster_vault_ids.items():
            print(f"  [{cname}] fetching summary report...")
            rows = get_data_transfer_report(args.apikey, cid, cname,
                                            vault_ids, start_msecs, end_msecs,
                                            debug=args.debug)
            if args.vault:
                rows = [r for r in rows if args.vault.lower() in r["vault_name"].lower()]
            for r in rows:
                r["period_start"] = s
                r["period_end"]   = e
            print(f"  [{cname}] {len(rows)} row(s)")
            all_rows.extend(rows)

    else:  # trend
        days = list(iter_days(start_msecs, end_msecs, tz))
        print(f"[*] Trend mode: {len(days)} day(s) × {len(cluster_vault_ids)} cluster(s) "
              f"= up to {len(days) * len(cluster_vault_ids)} report call(s)")
        for cname, (cid, vault_ids) in cluster_vault_ids.items():
            print(f"  [{cname}]")
            for day_start, day_end, date_str in days:
                rows = get_data_transfer_report(args.apikey, cid, cname,
                                                vault_ids, day_start, day_end,
                                                debug=args.debug)
                if args.vault:
                    rows = [r for r in rows if args.vault.lower() in r["vault_name"].lower()]
                for r in rows:
                    r["period_start"] = date_str
                    r["period_end"]   = date_str
                print(f"    {date_str}: {len(rows)} row(s)")
                all_rows.extend(rows)

    write_excel(all_rows, args.output, args.mode)


if __name__ == "__main__":
    main()
