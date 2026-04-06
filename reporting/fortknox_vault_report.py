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
  4.7 (2026-04-06) — Fixed x-axis dates not showing on trend charts. Root
                     cause: openpyxl set_categories() emits <c:numRef> in
                     chart XML; Excel discards string values passed via numRef.
                     Fix: assign DataSource(strRef=StrRef(...)) directly to
                     each series' .cat — forces <c:strRef> so Excel renders
                     date strings as text labels on the x-axis.
  4.6 (2026-04-05) — Fixed axis labels not rendering: removed numFmt from
                     category (string) x-axis, added delete=False on both
                     axes. X-axis tick density auto-reduces: every day for
                     ≤14 days, every 2 days for 15-60 days, weekly for >60.
  4.5 (2026-04-05) — One chart sheet per cluster (named after the cluster)
                     instead of a single combined sheet. Each chart shows all
                     protection groups for that cluster as separate series.
  4.4 (2026-04-05) — Trend chart: 16 pt bold title, legend moved to bottom
                     (no overlap), x-axis labelled "Date" with yyyy-mm-dd
                     format, y-axis labelled "Storage Consumed (TB)". Header
                     row color changed to Cohesity green (#00B388).
  4.3 (2026-04-04) — Secure credential storage via OS keychain (keyring).
                     --apikey is now optional; key is prompted once, saved to
                     the system keychain, and retrieved automatically on future
                     runs. --clear-credentials removes the stored key.
                     Output filename auto-generated as
                     <script_name>_YYYYMMDD_HHMMSS.xlsx in cwd.
  4.2 (2026-04-04) — Added Storage Consumed (TB) column (÷ 1,000,000,000,000,
                     4 decimal places). Trend chart now plots TB values instead
                     of raw bytes for human-readable Y axis.
  4.1 (2026-04-04) — Chart references Report sheet directly — no data duplication
                     on the chart sheet. Report rows sorted by (group, cluster,
                     vault, date) so each series is a contiguous range. All-zero
                     Storage Consumed series excluded from chart. Fixed chart
                     size (18 × 28 cm) to fit a single screen.
  4.0 (2026-04-04) — Combined all protection groups onto a single "Trend Charts"
                     sheet. Pivot table (Date × Group) formatted as an Excel
                     Table with ▼ dropdown filters on every column header for
                     group/date navigation (native slicers require pivot tables
                     not supported by openpyxl). Chart sums storage across
                     vaults per group per day. One series per protection group.
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
  python3 fortknox_vault_report.py                         # prompts for key on first run
  python3 fortknox_vault_report.py --days 30               # uses stored key
  python3 fortknox_vault_report.py --mode trend --days 30  # trend mode
  python3 fortknox_vault_report.py --clear-credentials     # remove stored key
"""

__version__ = "4.7"

import argparse
import getpass
import os
import sys
import urllib3
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

import requests

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

HELIOS_HOST     = "helios.cohesity.com"
_KEYRING_SVC    = "cohesity_helios"
_KEYRING_USER   = "apikey"


# ---------------------------------------------------------------------------
# Credential store  (OS keychain via keyring)
# ---------------------------------------------------------------------------

def _keyring_available() -> bool:
    try:
        import keyring          # noqa: F401
        return True
    except ImportError:
        return False


def get_api_key(cli_key: str = None) -> str:
    """
    Return the Helios API key using this priority order:
      1. --apikey CLI argument (one-time use, not stored)
      2. Key stored in the OS keychain from a previous run
      3. Interactive prompt — key is then saved to the keychain

    The key is never written to disk in plaintext and never visible
    in the process list.
    """
    if cli_key:
        return cli_key

    if _keyring_available():
        import keyring
        stored = keyring.get_password(_KEYRING_SVC, _KEYRING_USER)
        if stored:
            print("[*] Using API key stored in system keychain.")
            return stored
        print("[*] No stored API key found.")
        key = getpass.getpass("    Enter Helios API key (will be saved to keychain): ").strip()
        if not key:
            print("ERROR: API key cannot be empty.")
            sys.exit(1)
        keyring.set_password(_KEYRING_SVC, _KEYRING_USER, key)
        print("[*] API key saved to system keychain.")
        return key
    else:
        print("    NOTE: 'keyring' package not installed — key will not be saved.")
        print("          Install with:  pip install keyring")
        key = getpass.getpass("    Enter Helios API key: ").strip()
        if not key:
            print("ERROR: API key cannot be empty.")
            sys.exit(1)
        return key


def clear_stored_credentials():
    """Remove the stored API key from the OS keychain."""
    if not _keyring_available():
        print("ERROR: 'keyring' package not installed. Nothing to clear.")
        sys.exit(1)
    import keyring
    existing = keyring.get_password(_KEYRING_SVC, _KEYRING_USER)
    if existing:
        keyring.delete_password(_KEYRING_SVC, _KEYRING_USER)
        print("[*] Stored API key removed from system keychain.")
    else:
        print("[*] No stored API key found — nothing to clear.")


def default_output_path() -> str:
    """Auto-generate output filename: <script_name>_YYYYMMDD_HHMMSS.xlsx in cwd."""
    timestamp   = datetime.now().strftime("%Y%m%d_%H%M%S")
    script_name = os.path.splitext(os.path.basename(__file__))[0]
    return os.path.join(os.getcwd(), f"{script_name}_{timestamp}.xlsx")


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
            consumed_bytes = job.get("storageConsumed", 0) or 0
            rows.append({
                "cluster":                cluster_name,
                "vault_name":             vault_name,
                "vault_type":             vault_type,
                "protection_group":       job.get("protectionJobName", ""),
                "logical_bytes":          job.get("numLogicalBytesTransferred", 0),
                "physical_bytes":         job.get("numPhysicalBytesTransferred", 0),
                "storage_consumed_bytes": consumed_bytes,
                "storage_consumed_tb":    round(consumed_bytes / 1_000_000_000_000, 4),
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
    ("Storage Consumed (TB)",        "storage_consumed_tb"),
]

# Column index (1-based) of the TB column in the Report sheet — used by chart
_TB_COL_IDX = next(i for i, (_, k) in enumerate(COLUMNS, start=1)
                   if k == "storage_consumed_tb")


def _safe_sheet_name(name: str, existing: list) -> str:
    """Sanitise and deduplicate an Excel sheet name (max 31 chars)."""
    safe = name.translate(str.maketrans("", "", r":\/?*[]"))
    safe = safe[:31]
    if safe in existing:
        safe = safe[:28] + f"_{len(existing)}"
    return safe


def _make_chart(title: str, report_ws, series_list: list,
                LineChart, Reference, SeriesLabel, Legend):
    """
    Build a single LineChart with consistent styling.

    series_list: [(label, start_row, end_row), ...]
    X-axis tick density is reduced automatically for long date ranges:
      ≤14 days  → every day
      15-60 days → every 2 days
      >60 days   → every 7 days (weekly)
    """
    chart = LineChart()
    chart.style  = 10
    chart.height = 16   # cm — room for legend below without overlap
    chart.width  = 28   # cm

    # --- Chart title: 16 pt bold; plain string fallback ---
    try:
        from openpyxl.chart.title import Title
        from openpyxl.chart.text import RichText as ChartRichText, TxChoice
        from openpyxl.drawing.text import (
            Paragraph, RegularTextRun, RichTextProperties,
            ParagraphProperties, CharacterProperties,
        )
        chart.title = Title(
            tx=TxChoice(rich=ChartRichText(
                bodyPr=RichTextProperties(),
                p=[Paragraph(
                    pPr=ParagraphProperties(defRPr=CharacterProperties(sz=1600, b=True)),
                    r=[RegularTextRun(t=title)]
                )]
            )),
            overlay=False,
        )
    except Exception:
        chart.title = title

    # --- Y axis (value) ---
    chart.y_axis.title        = "Storage Consumed (TB)"
    chart.y_axis.numFmt       = "0.000"
    chart.y_axis.majorTickMark = "out"
    chart.y_axis.tickLblPos   = "nextTo"
    chart.y_axis.delete       = False

    # --- X axis (category — string dates, no numFmt) ---
    chart.x_axis.title        = "Date"
    chart.x_axis.majorTickMark = "out"
    chart.x_axis.tickLblPos   = "low"
    chart.x_axis.delete       = False

    # Reduce tick density for long ranges so dates don't overlap
    n_days = max((end - start + 1) for _, start, end in series_list) if series_list else 1
    if n_days > 60:
        chart.x_axis.tickLblSkip = 7    # weekly label
    elif n_days > 14:
        chart.x_axis.tickLblSkip = 2    # every-other-day label
    # else: ≤14 days — show every date, leave tickLblSkip at default

    # --- Legend at bottom, no overlap ---
    legend = Legend()
    legend.position = "b"
    legend.overlay  = False
    chart.legend = legend

    # Build all series first, then assign string categories to each one.
    # openpyxl's set_categories() generates <c:numRef> in the XML — Excel
    # treats string date values as invalid numbers and shows nothing.
    # Using StrRef forces <c:strRef> so Excel renders them as text labels.
    try:
        from openpyxl.chart.data_source import DataSource, StrRef
        from openpyxl.utils import get_column_letter as _gcl
        _use_strref = True
    except ImportError:
        _use_strref = False

    cat_formula = None
    for label, start_row, end_row in series_list:
        data_ref = Reference(report_ws, min_col=_TB_COL_IDX,
                             min_row=start_row, max_row=end_row)
        chart.add_data(data_ref)
        chart.series[-1].title = SeriesLabel(v=label)

        # Capture category formula from the first (longest) series
        if cat_formula is None:
            col_a = _gcl(1) if _use_strref else "A"
            cat_formula = (f"'{report_ws.title}'!${col_a}${start_row}"
                           f":${col_a}${end_row}")

    # Assign string categories to every series
    if cat_formula:
        if _use_strref:
            str_cat = DataSource(strRef=StrRef(f=cat_formula))
            for s in chart.series:
                s.cat = str_cat
        else:
            # Fallback: let openpyxl use numRef (dates may not display)
            chart.set_categories(Reference(report_ws, min_col=1,
                                           min_row=series_list[0][1],
                                           max_row=series_list[0][2]))

    return chart


def _add_trend_chart(wb, group_ranges: dict):
    """
    Add one chart sheet per cluster, each showing all protection groups
    for that cluster as separate series.

    References the Report sheet directly — no data is duplicated.
    All-zero series are excluded. Sheet names are truncated to 31 chars.

    group_ranges: {(cluster, group, vault): {"start": int, "end": int,
                                              "any_nonzero": bool}}
    """
    try:
        from openpyxl.chart import LineChart, Reference
        from openpyxl.chart.series import SeriesLabel
        from openpyxl.chart.legend import Legend
    except ImportError:
        print("WARNING: openpyxl chart module unavailable — chart skipped.")
        return

    report_ws = wb["Report"]

    # Group series by cluster
    clusters_seen = sorted({k[0] for k in group_ranges})
    sheet_names_used = [s.title for s in wb.worksheets]
    total_charts = 0

    for cluster in clusters_seen:
        # Collect non-zero series for this cluster, sorted by group name
        series_list = []
        for (c, group, vault), info in sorted(group_ranges.items(),
                                              key=lambda x: (x[0][1], x[0][2])):
            if c != cluster or not info["any_nonzero"]:
                continue
            vault_count = sum(1 for k in group_ranges
                              if k[0] == cluster and k[1] == group)
            label = group if vault_count == 1 else f"{group} [{vault}]"
            series_list.append((label, info["start"], info["end"]))

        if not series_list:
            print(f"    [{cluster}] all-zero — chart skipped")
            continue

        title = f"FortKnox \u2014 {cluster}"
        chart = _make_chart(title, report_ws, series_list,
                            LineChart, Reference, SeriesLabel, Legend)

        sheet_name = _safe_sheet_name(cluster, sheet_names_used)
        sheet_names_used.append(sheet_name)
        ws = wb.create_sheet(title=sheet_name)
        ws.add_chart(chart, "A1")
        total_charts += 1
        print(f"    Chart sheet '{sheet_name}': {len(series_list)} series")

    if total_charts == 0:
        print("    No non-zero Storage Consumed data — all charts skipped.")
    else:
        print(f"    {total_charts} chart sheet(s) added (one per cluster)")


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

    # Header row styling — Cohesity green
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(fill_type="solid", fgColor="00B388")
    ws.append(headers)
    for cell in ws[1]:
        cell.font      = header_font
        cell.fill      = header_fill
        cell.alignment = Alignment(horizontal="center")

    # Sort by (group, cluster, vault, date) so each series occupies contiguous rows
    # — required for the chart to reference Report sheet ranges directly.
    sorted_rows = sorted(rows, key=lambda r: (
        r["protection_group"], r["cluster"], r["vault_name"], r["period_start"]
    ))

    # Track Excel row range per (cluster, group, vault): header is row 1, data from row 2
    group_ranges: dict = {}
    for excel_row, r in enumerate(sorted_rows, start=2):
        key = (r["cluster"], r["protection_group"], r["vault_name"])
        consumed = int(r.get("storage_consumed_bytes") or 0)
        if key not in group_ranges:
            group_ranges[key] = {"start": excel_row, "end": excel_row,
                                 "any_nonzero": consumed > 0}
        else:
            group_ranges[key]["end"] = excel_row
            if consumed > 0:
                group_ranges[key]["any_nonzero"] = True

    # Data rows — byte columns written as integers so Excel can sort/chart them
    for row in sorted_rows:
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

    # Trend chart (trend mode only) — references Report sheet, no data duplication
    if mode == "trend":
        _add_trend_chart(wb, group_ranges)

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
  First run (prompts for API key, saves to keychain):
    python3 fortknox_vault_report.py

  Subsequent runs (key retrieved from keychain automatically):
    python3 fortknox_vault_report.py --days 30
    python3 fortknox_vault_report.py --start 2026-03-01 --end 2026-04-01 --mode trend
    python3 fortknox_vault_report.py --cluster <name>

  One-time override (key not saved):
    python3 fortknox_vault_report.py --apikey <key>

  Clear stored credentials:
    python3 fortknox_vault_report.py --clear-credentials
        """
    )
    parser.add_argument("--apikey",  default=None,
                        help="Helios API key (optional — if omitted the key stored in the "
                             "system keychain is used, or you will be prompted)")
    parser.add_argument("--clear-credentials", action="store_true", dest="clear_creds",
                        help="Remove the stored API key from the system keychain and exit")
    parser.add_argument("--cluster", help="Filter to a specific cluster (partial name match)")
    parser.add_argument("--vault",   help="Filter to a specific vault (partial name match)")
    parser.add_argument("--output",  default=None,
                        help="Output Excel filename (default: <script_name>_YYYYMMDD_HHMMSS.xlsx "
                             "in the current directory)")
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

    # Handle --clear-credentials before doing anything else
    if args.clear_creds:
        clear_stored_credentials()
        sys.exit(0)

    api_key     = get_api_key(args.apikey)
    output_path = args.output or default_output_path()

    print(f"\n[*] FortKnox data transfer report  v{__version__}")
    print(f"[*] Output: {output_path}")
    if args.cluster:
        print(f"[*] Cluster filter: '{args.cluster}'")
    if args.vault:
        print(f"[*] Vault filter:   '{args.vault}'")

    clusters = get_helios_clusters(api_key)

    # Pick the reference cluster for timezone detection:
    # use the first cluster matching --cluster filter, or the first overall.
    ref_clusters = [c for c in clusters
                    if not args.cluster or args.cluster.lower() in c["name"].lower()]
    if not ref_clusters:
        print("ERROR: No clusters match the given --cluster filter.")
        sys.exit(1)

    ref = ref_clusters[0]
    print(f"[*] Querying timezone from '{ref['name']}'...")
    tz_name = get_cluster_timezone(api_key, ref["clusterId"], ref["name"])
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
        vault_ids = get_fortknox_vault_ids(api_key, c["clusterId"])
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
            rows = get_data_transfer_report(api_key, cid, cname,
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
                rows = get_data_transfer_report(api_key, cid, cname,
                                                vault_ids, day_start, day_end,
                                                debug=args.debug)
                if args.vault:
                    rows = [r for r in rows if args.vault.lower() in r["vault_name"].lower()]
                for r in rows:
                    r["period_start"] = date_str
                    r["period_end"]   = date_str
                print(f"    {date_str}: {len(rows)} row(s)")
                all_rows.extend(rows)

    write_excel(all_rows, output_path, args.mode)


if __name__ == "__main__":
    main()
