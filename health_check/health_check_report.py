#!/usr/bin/env python3
# =============================================================================
# DISCLAIMER: This code is provided on a best-effort basis and is not in any
# way officially supported or sanctioned by Cohesity. The code is intentionally
# kept simple to retain value as example code. The code in this repository is
# provided as-is and the author accepts no liability for damages resulting
# from its use.
# =============================================================================
"""
health_check_report.py  v1.0

Multi-cluster Cohesity health check — 12-sheet Excel workbook + Word document.
Designed for enterprise customer business reviews (EBRs) and SE trusted-advisor
engagements.  Gathers live data from Cohesity Helios and produces:

  Excel sheets
  ────────────
  1  Executive Summary   — per-cluster health score & grade
  2  Infrastructure      — versions, nodes, disks, config
  3  Protection Health   — success rates, SLA, RPO gaps
  4  Storage & Capacity  — utilisation, data reduction, runway
  5  Policy Audit        — retention, replication, archival (groups count fixed)
  6  Policy → Groups     — every group with the policy that governs it
  7  Alerts              — open critical/warning with age
  8  Security            — encryption, FIPS, immutability, vault
  9  Replication/Archive — targets, last transfer, FortKnox
  10 Data Services       — NAS views, quota utilisation
  11 Coverage Gaps       — groups with no recent successful backup
  12 Trends (30d)        — daily success rate + storage growth charts
  13 Recommendations     — prioritised action list

  Word document
  ─────────────
  Cover, Executive Summary, Environment, Protection, Storage,
  Security, Recommendations, Methodology appendix.

Usage
─────
  python health_check_report.py --apikey <key> [options]

Options
───────
  --apikey KEY        Helios API key (or set HELIOS_API_KEY env var)
  --cluster NAME      Limit to a single cluster (partial name match)
  --days N            Lookback window for alerts/runs (default 30)
  --customer NAME     Customer name printed on report cover
  --output PATH       Base path for output files (no extension)
  --quick             Skip per-group run history; use last-run only (faster)
  --excel-only        Skip Word document generation
  --word-only         Skip Excel generation
  --debug             Print raw API payloads for troubleshooting

Requirements
────────────
  pip install requests openpyxl python-docx

Version history
───────────────
  1.17 (2026-04-08) — fix: Trends tab shows fewer than 30 days on clusters with
                     many protection groups. Root cause: v1 protectionRuns API
                     caps at 1000 runs per call (newest-first); clusters with
                     50+ groups hit the cap in ~20 days. Fixed by paginating:
                     advance endTimeUsecs to just before the oldest run on each
                     page and repeat until the full window is covered. Safety
                     limit of 50 pages (50 000 runs). --debug shows page count.
  1.16 (2026-04-08) — feat: Trends chart — visible X/Y axes with tick marks,
                     Y axis fixed 0-100 with "0.0" format, circle data-point
                     markers (size 5, Cohesity green), 2pt line weight, StrRef
                     category fix so date strings render as labels in Excel,
                     auto tick-density reduction for >14 and >60 day ranges,
                     chart title updated to note warnings are included.
  1.15 (2026-04-08) — feat: include kWarning runs in Success % calculation
                     throughout — Trends "Success % (incl. Warnings)" col H,
                     _score_protection() quick mode, _success_stats() quick
                     mode. Warning runs complete successfully (data protected)
                     but may have minor issues (skipped objects). They still
                     appear in the Warning column for visibility. Word doc
                     Methodology appendix now documents this behaviour.
  1.14 (2026-04-08) — fix: Trends tab inaccurate — switched primary data source
                     from per-group v2 runs to v1 /public/protectionRuns which
                     is the same source powering the Helios reporting page.
                     Single call per cluster; backupRun.slaViolated (col I) and
                     backupRun.stats.totalLogicalBackupSizeBytes (col J) are
                     accurate. v2 per-group runs retained as fallback. Works in
                     --quick mode (v1 call not gated behind the per-group loop).
  1.13 (2026-04-08) — fix: Trends tab cols I and J always blank — col I
                     (SLA Violations) used run.get("isSlaViolated") but the
                     field lives inside localBackupInfo, not the top-level run
                     object; fixed to lb.get("isSlaViolated"). Col J (Logical GB)
                     used lb.get("stats").get("logicalSizeBytes") but the correct
                     sub-object is localSnapshotStats; fixed to
                     lb.get("localSnapshotStats").get("logicalSizeBytes").
  1.12 (2026-04-08) — fix: Replication & Archive col E and col K still empty
                     after v1.11 — policy→group chain produced no rows because
                     _arch_name() returns "" when v2 policies store archival
                     targets by vault ID only (no name/targetName field), so
                     policy_arch_targets stayed empty and all_arch_targets was
                     empty. Fix: supplement arch_groups directly from
                     fk_data.dataTransferPerProtectionJob[].protectionJobName
                     per vault (reliable even without policy resolution); build
                     all_arch_targets as union of policy chain + fk_data vault
                     names + vault list names so at least one source always
                     produces rows and E/K are populated.
  1.11 (2026-04-08) — feat: Replication & Archive col E changed from policy
                     count to protection group count; new col K lists all group
                     names using each target. Status updated: "Configured" when
                     groups are assigned, "No groups assigned" when not.
                     Rewired target discovery through policy→group chain so E
                     and K reflect actual group assignment, not policy count.
  1.10 (2026-04-08) — fix: Replication & Archive col E (Policies Using) showed
                     0 because _arch_name() returned "" when the v2 policy stores
                     archival targets by vault ID only (no name/targetName field).
                     Added vault ID → name resolution via /public/vaults lookup.
                     Columns G/H renamed to include API field names and note that
                     values are cumulative totals. "Vault only — no policy" status
                     reworded to "Vault exists — not referenced by any policy".
  1.9 (2026-04-08) — fix: Replication & Archive still blank after v1.8.
                     dataTransferToVaults API requires vaultIds filter to
                     return any rows — added vault ID extraction and pass-
                     through from _vaults() to _fortknox_data(). Vault type
                     column D now read from archivalTargetConfig.targetType
                     inside the policy target (previously name-based lookup
                     against vault list failed on every mismatch).
  1.8 (2026-04-08) — fix: Replication & Archive sheet columns D-H blank.
                     Root causes: (1) _fortknox_data used startTimeUsecs/
                     endTimeUsecs — API expects startTimeMsecs/endTimeMsecs,
                     so fk_data was always empty; (2) archival rows driven by
                     cd["vaults"] list whose name may differ from policy target
                     names — flipped to drive rows from arch_use (policy targets)
                     and enrich vault type/FK stats via name-then-ID lookup;
                     (3) G/H (Logical/Physical Transferred) hardcoded "" — now
                     populated from numLogicalBytesTransferred/numPhysical.
                     Orphaned vaults (in API but not in any policy) still shown.
  1.7 (2026-04-08) — fix: _fk_status() false-negatives — lastRun.archivalInfo
                     is only populated when the most recent primary run included
                     archival; groups with daily local backups + less-frequent
                     archival showed as idle even when active. Fix: scan full
                     group_runs history (non-quick mode) for successful FK
                     archival results matched by target type AND vault name;
                     lastRun remains the fallback for --quick mode only.
  1.6 (2026-04-08) — feat: detect FortKnox configured-but-idle state. New
                     _fk_status() returns "active"/"idle"/"none" by inspecting
                     group lastRun.archivalInfo (dataTransferToVaults storageConsumed
                     is cumulative and not reliable for recency). Security sheet
                     column 5 shows "Active" (green) / "Configured — Idle" (orange) /
                     "No" (red). Idle triggers a HIGH recommendation and scores
                     10/20 points instead of the full 20.
  1.5 (2026-04-08) — fix: drop Alerts columns K (Node ID) and L (Entity) —
                     sparse propertyList fields only present on hardware alerts;
                     description + alert code already capture the context.
  1.4 (2026-04-08) — fix: Policy Audit column N (Groups Using) now computed by
                     counting groups whose policyId matches each policy — the v2
                     API does not return numProtectionGroups.
                     feat: new "Policy → Groups" sheet (sheet 6) listing every
                     protection group with the policy that governs it, sorted by
                     policy then group name, paused/failed rows highlighted.
  1.3 (2026-04-08) — fix: infrastructure disk count now reads diskCount integer
                     per node (v1 /public/nodes field) instead of looking for
                     non-existent list fields; NTP servers shows "Not configured"
                     instead of blank when the cluster has no NTP entries; added
                     ntpServersList and ntpServerIps fallback paths.
  1.2 (2026-04-08) — fix: correct field names throughout — firstTimestampUsecs,
                     clusterSoftwareVersion, localSnapshotStats, successfulObjectsCount,
                     policyId→name lookup, v2 schedule/target/retention paths,
                     broadened FortKnox detection, N/A capacity fallback.
                     feat: output filename includes customer name + local timestamp.
  1.0 (2026-04-08) — feat: initial release. 12-sheet Excel + Word document.
                     Health scoring, recommendations engine, trend charts.
"""

__version__ = "1.17"

import argparse
import datetime
import os
import sys
import time
import urllib3

import requests

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "utils"))
from cohesity_auth import (
    get_api_key, get_helios_clusters, make_helios_headers, HELIOS_HOST,
)
from formatters import (
    bytes_to_gb, bytes_to_tb, usecs_to_datetime, auto_fit_columns,
)

try:
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment
    from openpyxl.chart import LineChart, Reference
    from openpyxl.chart.series import SeriesLabel
    EXCEL_OK = True
except ImportError:
    EXCEL_OK = False

try:
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    DOCX_OK = True
except ImportError:
    DOCX_OK = False

# ── Thresholds / best-practice baselines ─────────────────────────────────────
CAPACITY_WARN_PCT    = 70       # cluster used % — warning
CAPACITY_CRIT_PCT    = 85       # cluster used % — critical
SUCCESS_WARN_PCT     = 95.0     # backup success rate below this = warning
SUCCESS_CRIT_PCT     = 90.0     # backup success rate below this = critical
RPO_WARN_HOURS       = 24       # hours since last successful run — warning
RPO_CRIT_HOURS       = 48       # hours since last successful run — critical
ALERT_AGE_CRIT_DAYS  = 7        # open critical alert older than N days = finding
VIEW_QUOTA_WARN_PCT  = 80       # view at this % of quota = warning
VIEW_QUOTA_CRIT_PCT  = 90       # view at this % of quota = critical
VERSION_WARN_PREFIX  = "7."     # clusters below this major version get a flag

# Scoring weights — must sum to 1.0
_W = {
    "protection": 0.25,
    "sla":        0.20,
    "infra":      0.15,
    "capacity":   0.15,
    "security":   0.15,
    "alerts":     0.10,
}

# ── Excel palette ─────────────────────────────────────────────────────────────
COH_GREEN   = "70AD47"
DARK_GREEN  = "4E7A2E"
LT_GREEN    = "E2EFDA"
RED         = "FF4C4C"
ORANGE      = "FF8C00"
YELLOW      = "FFD700"
LIGHT_GRAY  = "F2F2F2"
WHITE       = "FFFFFF"


# ── openpyxl style helpers (lazy — only called when EXCEL_OK is True) ─────────
def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def _font(bold=False, color="000000", size=10):
    return Font(bold=bold, color=color, size=size)

def _rag_fill(score):
    """Return fill based on 0-100 score."""
    if score >= 90: return _fill(LT_GREEN)
    if score >= 75: return _fill("FFFF99")
    if score >= 60: return _fill("FFD580")
    if score >= 40: return _fill("FFAA80")
    return _fill("FF8080")

def _hdr(ws, row, cols):
    """Write a bold Cohesity-green header row."""
    for c, val in enumerate(cols, 1):
        cell = ws.cell(row=row, column=c, value=val)
        cell.font      = Font(bold=True, color=WHITE, size=10)
        cell.fill      = _fill(COH_GREEN)
        cell.alignment = Alignment(wrap_text=True, vertical="center",
                                   horizontal="center")

def _title(ws, text, span):
    ws.merge_cells(f"A1:{span}1")
    c = ws["A1"]
    c.value     = text
    c.font      = Font(bold=True, color=WHITE, size=13)
    c.fill      = _fill(COH_GREEN)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22

def _reg(ws, row, col, val):
    cell = ws.cell(row=row, column=col, value=val)
    cell.font = _font()
    return cell

# ── Grade from score ──────────────────────────────────────────────────────────
def grade(score):
    if score >= 90: return "Excellent"
    if score >= 75: return "Good"
    if score >= 60: return "Fair"
    if score >= 40: return "At Risk"
    return "Critical"

# ── Time helpers ──────────────────────────────────────────────────────────────
def _now_usecs():
    return int(time.time() * 1_000_000)

def _days_ago_usecs(n):
    return int((time.time() - n * 86400) * 1_000_000)


# ─────────────────────────────────────────────────────────────────────────────
# API HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _get(session, headers, path, params=None, debug=False):
    """GET from Helios (cluster already routed via accessClusterId in headers)."""
    url = f"https://{HELIOS_HOST}{path}"
    try:
        r = session.get(url, headers=headers, params=params,
                        verify=False, timeout=60)
        if debug:
            print(f"    DEBUG {path} → {r.status_code}")
        if r.status_code == 200:
            return r.json()
        if debug:
            print(f"    DEBUG body: {r.text[:300]}")
        return None
    except Exception as exc:
        print(f"    WARN: GET {path} failed: {exc}")
        return None


# ─────────────────────────────────────────────────────────────────────────────
# DATA COLLECTION  (one function per resource type)
# ─────────────────────────────────────────────────────────────────────────────

def _cluster_info(session, h, debug):
    d = _get(session, h, "/irisservices/api/v1/public/cluster",
             params={"fetchStats": "true"}, debug=debug)
    return d or {}

def _nodes(session, h, debug):
    d = _get(session, h, "/irisservices/api/v1/public/nodes", debug=debug)
    return d or []

def _alerts(session, h, days, debug):
    d = _get(session, h, "/irisservices/api/v1/public/alerts",
             params={
                 "alertStateList": "kOpen",
                 "maxAlerts": 1000,
                 "startDateUsecs": _days_ago_usecs(days),
                 "endDateUsecs":   _now_usecs(),
             }, debug=debug)
    return d or []

def _protection_groups(session, h, debug):
    d = _get(session, h, "/v2/data-protect/protection-groups",
             params={
                 "isDeleted": "false",
                 "includeTenants": "true",
                 "includeLastRunInfo": "true",
             }, debug=debug)
    return (d or {}).get("protectionGroups") or []

def _group_runs(session, h, group_id, start_usecs, end_usecs, debug):
    d = _get(session, h,
             f"/v2/data-protect/protection-groups/{group_id}/runs",
             params={
                 "startTimeUsecs": start_usecs,
                 "endTimeUsecs":   end_usecs,
                 "numRuns": 500,
                 "includeTenants": "true",
             }, debug=debug)
    return (d or {}).get("runs") or []

def _v1_protection_runs(session, h, start_usecs, end_usecs, debug):
    """Paginated fetch of ALL protection runs in the time window via v1 API.

    The API returns runs newest-first and caps at 1000 per page. Clusters with
    many groups can easily exceed 1000 runs in 30 days, so a single call would
    only return the most recent N days. We page by advancing endTimeUsecs to
    just before the oldest run on each page until the full window is covered.
    """
    all_runs = []
    page_end  = end_usecs
    PAGE_SIZE = 1000

    for page_num in range(50):   # safety cap: 50 pages × 1000 = 50 000 runs max
        page = _get(session, h, "/irisservices/api/v1/public/protectionRuns",
                    params={"startTimeUsecs": start_usecs,
                            "endTimeUsecs":   page_end,
                            "numRuns":        PAGE_SIZE},
                    debug=debug) or []
        if not page:
            break
        all_runs.extend(page)
        if debug:
            print(f"    DEBUG v1_runs page {page_num + 1}: {len(page)} runs "
                  f"(total so far: {len(all_runs)})")
        if len(page) < PAGE_SIZE:
            break   # last page — no more data
        # Find the start time of the oldest run on this page to use as next cursor
        oldest_t = page_end
        for r in page:
            t = ((r.get("backupRun") or {}).get("stats") or {}).get("startTimeUsecs") or 0
            if t and t < oldest_t:
                oldest_t = t
        if oldest_t <= start_usecs:
            break   # we have covered the full requested window
        page_end = oldest_t - 1  # next page ends just before this oldest run

    return all_runs

def _policies(session, h, debug):
    d = _get(session, h, "/v2/data-protect/policies",
             params={"includeTenants": "true"}, debug=debug)
    return (d or {}).get("policies") or []

def _storage_domains(session, h, debug):
    d = _get(session, h, "/v2/storage-domains",
             params={"includeStats": "true"}, debug=debug)
    return (d or {}).get("storageDomains") or []

def _views(session, h, debug):
    d = _get(session, h, "/v2/file-services/views",
             params={"includeStats": "true", "maxCount": 2000}, debug=debug)
    return (d or {}).get("views") or []

def _vaults(session, h, debug):
    d = _get(session, h, "/irisservices/api/v1/public/vaults",
             params={"includeFortKnoxVault": "true"}, debug=debug)
    return d or []

def _fortknox_data(session, h, start_usecs, end_usecs, vault_ids, debug):
    # API expects milliseconds; also requires vaultIds filter to return results
    params = {"startTimeMsecs": start_usecs // 1000,
              "endTimeMsecs":   end_usecs   // 1000}
    if vault_ids:
        params["vaultIds"] = vault_ids  # requests repeats list values correctly
    elif not vault_ids:
        return []   # no vaults — nothing to query
    d = _get(session, h,
             "/irisservices/api/v1/public/reports/dataTransferToVaults",
             params=params, debug=debug)
    return (d or {}).get("dataTransferSummary") or []


def collect_cluster(session, api_key, cluster, args):
    """Collect all health-check data for one cluster. Returns a dict."""
    name = cluster.get("name", "Unknown")
    cid  = cluster.get("clusterId")
    h    = make_helios_headers(api_key, cluster_id=cid)
    dbg  = args.debug

    start_usecs = _days_ago_usecs(args.days)
    end_usecs   = _now_usecs()

    print(f"  [{name}] infrastructure...", end="", flush=True)
    info    = _cluster_info(session, h, dbg)
    nodes   = _nodes(session, h, dbg)
    print(" alerts...", end="", flush=True)
    alerts  = _alerts(session, h, args.days, dbg)
    print(" groups...", end="", flush=True)
    groups  = _protection_groups(session, h, dbg)
    print(" policies...", end="", flush=True)
    policies = _policies(session, h, dbg)
    print(" storage...", end="", flush=True)
    domains = _storage_domains(session, h, dbg)
    views   = _views(session, h, dbg)
    print(" vaults...", end="", flush=True)
    vaults    = _vaults(session, h, dbg)
    vault_ids = [v["id"] for v in vaults if "id" in v]
    fk_data   = _fortknox_data(session, h, start_usecs, end_usecs, vault_ids, dbg)
    print(" run-summary...", end="", flush=True)
    v1_runs   = _v1_protection_runs(session, h, start_usecs, end_usecs, dbg)

    group_runs = {}
    if not args.quick:
        print(f" runs ({len(groups)} groups)...", end="", flush=True)
        for g in groups:
            gid = g.get("id")
            if gid:
                runs = _group_runs(session, h, gid, start_usecs, end_usecs, dbg)
                if runs:
                    group_runs[gid] = runs

    print(" done.")

    # Build policy ID → name lookup used throughout the report
    policy_map = {p.get("id"): p.get("name", "") for p in policies if p.get("id")}

    cd = {
        "name":        name,
        "id":          cid,
        "info":        info,
        "nodes":       nodes,
        "alerts":      alerts,
        "groups":      groups,
        "group_runs":  group_runs,
        "policies":    policies,
        "policy_map":  policy_map,
        "domains":     domains,
        "views":       views,
        "vaults":      vaults,
        "fk_data":     fk_data,
        "v1_runs":     v1_runs,
        "start_usecs": start_usecs,
        "end_usecs":   end_usecs,
    }
    cd["scores"]          = _compute_scores(cd, args.quick)
    cd["recommendations"] = _build_recommendations(cd)
    return cd


# ─────────────────────────────────────────────────────────────────────────────
# SCORING ENGINE
# ─────────────────────────────────────────────────────────────────────────────

def _score_protection(cd, quick):
    """0-100 based on backup success rate."""
    groups = cd["groups"]
    if not groups:
        return 100
    if quick:
        total   = len(groups)
        success = sum(
            1 for g in groups
            if (g.get("lastRun") or {}).get("localBackupInfo", {}).get("status", "")
               in ("kSuccess", "Succeeded", "kWarning")
        )
    else:
        all_runs = [r for runs in cd["group_runs"].values() for r in runs]
        if not all_runs:
            return 100
        total   = len(all_runs)
        success = sum(
            1 for r in all_runs
            if (r.get("localBackupInfo") or {}).get("status", "")
               in ("kSuccess", "Succeeded", "kWarning")
        )
    pct = (success / total * 100) if total else 100
    if pct >= SUCCESS_WARN_PCT: return 100
    if pct >= SUCCESS_CRIT_PCT: return int(70 + (pct - SUCCESS_CRIT_PCT) * 3)
    return max(0, int(pct))


def _score_sla(cd, quick):
    """0-100 based on SLA pass rate."""
    groups = cd["groups"]
    if not groups:
        return 100
    if quick:
        total = len(groups)
        viols = sum(1 for g in groups
                    if (g.get("lastRun") or {}).get("isSlaViolated", False))
    else:
        all_runs = [r for runs in cd["group_runs"].values() for r in runs]
        if not all_runs:
            return 100
        total = len(all_runs)
        viols = sum(1 for r in all_runs if r.get("isSlaViolated", False))
    pct_ok = ((total - viols) / total * 100) if total else 100
    if pct_ok >= 99: return 100
    if pct_ok >= 95: return 80
    if pct_ok >= 90: return 60
    if pct_ok >= 80: return 40
    return 20


def _score_infra(cd):
    """0-100 based on node health."""
    nodes = cd["nodes"]
    if not nodes:
        return 70
    healthy = sum(
        1 for n in nodes
        if ((n.get("nodeStatus") or {}).get("overallStatus", "") == "kNormal"
            or n.get("removalState") in (None, "kDontRemove"))
    )
    return int(healthy / len(nodes) * 100)


def _score_capacity(cd):
    """0-100 based on cluster capacity used %."""
    info   = cd["info"]
    usage  = ((info.get("stats") or {}).get("usagePerfStats") or {})
    usable = usage.get("physicalCapacityBytes") or 0
    used   = usage.get("totalPhysicalUsageBytes") or 0
    if not usable:
        return 70
    pct = used / usable * 100
    if pct <= CAPACITY_WARN_PCT:                return 100
    if pct <= CAPACITY_CRIT_PCT:                return int(100 - (pct - CAPACITY_WARN_PCT) * 2)
    return max(0, int(40 - (pct - CAPACITY_CRIT_PCT) * 2))


def _has_fortknox(vaults, policies):
    """Return True if FortKnox / RPaaS immutability is configured on this cluster."""
    fk_types = {"krpaas", "kfortknox", "kfort_knox", "krpaasarchival"}
    for v in vaults:
        vt = (v.get("vaultType") or "").lower()
        vn = (v.get("name") or "").lower()
        if any(t in vt for t in fk_types) or "fortknox" in vn:
            return True
    for p in policies:
        for t in ((p.get("remoteTargetPolicy") or {}).get("archivalTargets") or []):
            cfg = (t.get("archivalTargetConfig") or {})
            tt  = (cfg.get("targetType") or t.get("targetType") or "").lower()
            if any(x in tt for x in fk_types) or "fortknox" in tt:
                return True
    return False


def _fk_status(cd):
    """
    Return one of three states for FortKnox immutability:
      "active"  — FK configured AND a successful FK archival run exists within
                  the collection window (data is flowing)
      "idle"    — FK configured in vault/policy but no successful FK archival
                  found in run history (configured but not sending data)
      "none"    — FK not configured at all

    Detection strategy
    ──────────────────
    We CANNOT rely on lastRun.archivalInfo alone: that field is only populated
    when the most recent primary run included archival. Groups that run daily
    local backups but archive on a different schedule (e.g. weekly) will show
    empty archivalInfo in lastRun even when FK is working correctly.

    Primary source: group_runs (full run history, available in non-quick mode).
    Every run in the window is scanned for a successful archival result whose
    target type or name indicates FortKnox/RPaaS.

    Fallback (--quick mode): lastRun.archivalInfo on each group (less reliable
    — may miss groups whose last run was local-only).
    """
    vaults     = cd["vaults"]
    policies   = cd["policies"]
    groups     = cd["groups"]
    group_runs = cd.get("group_runs") or {}

    if not _has_fortknox(vaults, policies):
        return "none"

    fk_types = {"krpaas", "kfortknox", "kfort_knox", "krpaasarchival"}
    success_states = {"kSuccess", "Succeeded", "kSuccessWithWarning",
                      "SucceededWithWarning", "kWarning"}

    # Build set of FK vault names for name-based matching in archival results
    fk_vault_names = set()
    for v in vaults:
        vt = (v.get("vaultType") or "").lower()
        vn = (v.get("name") or "").lower()
        if any(t in vt for t in fk_types) or "fortknox" in vn:
            fk_vault_names.add(vn)
    for p in policies:
        for t in ((p.get("remoteTargetPolicy") or {}).get("archivalTargets") or []):
            cfg  = t.get("archivalTargetConfig") or {}
            name = (cfg.get("name") or cfg.get("vaultName")
                    or t.get("targetName") or "").lower()
            tt   = (cfg.get("targetType") or t.get("targetType") or "").lower()
            if any(x in tt for x in fk_types) or "fortknox" in tt:
                if name:
                    fk_vault_names.add(name)

    def _is_fk_success(run):
        """Return True if this run dict has a successful archival to an FK target."""
        arch_info = (run.get("archivalInfo") or {})
        for r in (arch_info.get("archivalTargetResults") or []):
            if r.get("status") not in success_states:
                continue
            # Match by archival target type
            ttype = (r.get("archivalTargetType")
                     or r.get("targetType") or "").lower()
            if any(x in ttype for x in fk_types) or "fortknox" in ttype:
                return True
            # Match by target name against known FK vault names
            tname = (r.get("targetName") or "").lower()
            if tname and (tname in fk_vault_names or "fortknox" in tname):
                return True
        return False

    # Primary: scan full run history (non-quick mode)
    for runs in group_runs.values():
        for run in runs:
            if _is_fk_success(run):
                return "active"

    # Fallback: check lastRun on each group (quick mode)
    for g in groups:
        if _is_fk_success(g.get("lastRun") or {}):
            return "active"

    return "idle"


def _score_security(cd):
    """0-100: 20 pts each for encryption, FIPS, vault, replication, immutability."""
    info     = cd["info"]
    vaults   = cd["vaults"]
    policies = cd["policies"]
    score    = 0
    # Encryption
    if info.get("encryptionEnabled") or \
       (info.get("encryptionConfig") or {}).get("encryptionEnabled"):
        score += 20
    # FIPS
    if info.get("fipsCompliant") or info.get("fipsEnabled"):
        score += 20
    # Vault configured
    if vaults:
        score += 20
    # Replication
    if any((p.get("remoteTargetPolicy") or {}).get("replicationTargets")
           for p in policies):
        score += 20
    fk = _fk_status({"vaults": vaults, "policies": policies, "groups": cd["groups"]})
    if fk == "active":
        score += 20
    elif fk == "idle":
        score += 10   # Configured but not sending — partial credit
    return score


def _score_alerts(cd):
    """0-100: deducted for open critical and warning alerts."""
    alerts = cd["alerts"]
    crits  = sum(1 for a in alerts if a.get("severity") == "kCritical")
    warns  = sum(1 for a in alerts if a.get("severity") == "kWarning")
    return max(0, 100 - crits * 10 - warns * 2)


def _compute_scores(cd, quick):
    s = {
        "protection": _score_protection(cd, quick),
        "sla":        _score_sla(cd, quick),
        "infra":      _score_infra(cd),
        "capacity":   _score_capacity(cd),
        "security":   _score_security(cd),
        "alerts":     _score_alerts(cd),
    }
    s["overall"] = round(sum(s[k] * _W[k] for k in _W), 1)
    s["grade"]   = grade(s["overall"])
    return s


# ─────────────────────────────────────────────────────────────────────────────
# RECOMMENDATIONS ENGINE
# ─────────────────────────────────────────────────────────────────────────────

_PRIORITY_ORDER = {"CRITICAL": 0, "HIGH": 1, "MEDIUM": 2, "LOW": 3}

def _r(priority, category, finding, action, impact):
    return {"priority": priority, "category": category,
            "finding": finding, "action": action, "impact": impact}


def _build_recommendations(cd):
    recs     = []
    info     = cd["info"]
    nodes    = cd["nodes"]
    alerts   = cd["alerts"]
    groups   = cd["groups"]
    policies = cd["policies"]
    views    = cd["views"]
    vaults   = cd["vaults"]
    usage    = ((info.get("stats") or {}).get("usagePerfStats") or {})

    # ── Infrastructure ────────────────────────────────────────────────────────
    version = info.get("clusterSoftwareVersion", "")
    if version and version < VERSION_WARN_PREFIX:
        recs.append(_r("HIGH", "Infrastructure",
            f"Software version {version} may be out of date",
            "Review EOL schedule and plan upgrade path",
            "Older versions miss security patches and new features"))

    unhealthy = [n for n in nodes
                 if not ((n.get("nodeStatus") or {}).get("overallStatus") == "kNormal"
                         or n.get("removalState") in (None, "kDontRemove"))]
    if unhealthy:
        recs.append(_r("CRITICAL", "Infrastructure",
            f"{len(unhealthy)} node(s) in unhealthy state",
            "Check hardware health, disk status, and cluster logs immediately",
            "Unhealthy nodes risk data availability and backup failures"))

    # ── Capacity ──────────────────────────────────────────────────────────────
    usable = usage.get("physicalCapacityBytes") or 0
    used   = usage.get("totalPhysicalUsageBytes") or 0
    if usable:
        cap_pct = used / usable * 100
        if cap_pct >= CAPACITY_CRIT_PCT:
            recs.append(_r("CRITICAL", "Capacity",
                f"Cluster at {cap_pct:.1f}% capacity — critically high",
                "Immediately expand capacity or archive/delete cold data",
                "Backup jobs fail when capacity is exhausted"))
        elif cap_pct >= CAPACITY_WARN_PCT:
            recs.append(_r("HIGH", "Capacity",
                f"Cluster at {cap_pct:.1f}% capacity — approaching limit",
                "Plan capacity expansion or data archival within 30-60 days",
                "Continued growth risks backup failures within weeks"))

    # ── Security ─────────────────────────────────────────────────────────────
    enc = info.get("encryptionEnabled") or \
          (info.get("encryptionConfig") or {}).get("encryptionEnabled")
    if not enc:
        recs.append(_r("HIGH", "Security",
            "Encryption at rest is not enabled",
            "Enable data-at-rest encryption from cluster security settings",
            "Physical media removal exposes customer data"))

    if not (info.get("fipsCompliant") or info.get("fipsEnabled")):
        recs.append(_r("MEDIUM", "Security",
            "FIPS compliance mode is not enabled",
            "Enable FIPS mode if required by security or compliance policy",
            "May not satisfy regulated-industry compliance requirements"))

    if not vaults:
        recs.append(_r("MEDIUM", "Security",
            "No vault (archival target) is configured",
            "Configure at least one external vault for off-site retention",
            "No off-site copy of data — single point of failure"))

    fk_st = _fk_status(cd)
    if fk_st == "none":
        recs.append(_r("MEDIUM", "Security",
            "No immutable/FortKnox archival target detected",
            "Configure FortKnox or WORM archival to protect against ransomware",
            "Without immutability, backups can be deleted or encrypted by attackers"))
    elif fk_st == "idle":
        recs.append(_r("HIGH", "Security",
            "FortKnox configured but no recent successful archival detected",
            "Verify protection group policies and schedules targeting FortKnox; "
            "check for failed archival runs and resolve blocking issues",
            "Immutable copy is configured but data is not flowing — "
            "ransomware protection gap despite vault being present"))

    has_repl = any((p.get("remoteTargetPolicy") or {}).get("replicationTargets")
                   for p in policies)
    if not has_repl:
        recs.append(_r("MEDIUM", "Resilience",
            "No policies have a replication target configured",
            "Add replication to a secondary cluster for site-level resilience",
            "Single-site data is vulnerable to site-wide outage"))

    # ── Alerts ────────────────────────────────────────────────────────────────
    crits = [a for a in alerts if a.get("severity") == "kCritical"]
    if crits:
        old = [a for a in crits
               if (_now_usecs() - (a.get("firstOccurrenceTime") or _now_usecs()))
                  > ALERT_AGE_CRIT_DAYS * 86400 * 1_000_000]
        sev = "CRITICAL" if old else "HIGH"
        msg = f"{len(crits)} open critical alert(s)"
        if old:
            msg += f", {len(old)} older than {ALERT_AGE_CRIT_DAYS} days"
        recs.append(_r(sev, "Availability",
            msg,
            "Review and resolve open critical alerts; investigate root causes",
            "Unresolved critical alerts indicate active system issues"))

    # ── Protection ───────────────────────────────────────────────────────────
    paused = [g for g in groups if g.get("isPaused")]
    if paused:
        recs.append(_r("HIGH", "Protection",
            f"{len(paused)} protection group(s) are paused",
            f"Review and resume: {', '.join(g['name'] for g in paused[:3])}",
            "Paused groups do not create backups — data is unprotected"))

    rpo_gaps = []
    for g in groups:
        lr      = g.get("lastRun") or {}
        lb      = lr.get("localBackupInfo") or {}
        start_u = lb.get("startTimeUsecs") or 0
        if not start_u:
            continue
        hrs = (_now_usecs() - start_u) / 3_600_000_000
        if hrs > RPO_CRIT_HOURS:
            rpo_gaps.append((g.get("name", "?"), hrs))
    if rpo_gaps:
        examples = ", ".join(f"{n} ({h:.0f}h)" for n, h in rpo_gaps[:3])
        recs.append(_r("HIGH", "Protection",
            f"{len(rpo_gaps)} group(s) with >{RPO_CRIT_HOURS}h RPO gap",
            f"Investigate: {examples}" + (" and more" if len(rpo_gaps) > 3 else ""),
            "Objects may be unrecoverable to a recent point in time"))

    # ── Policies ─────────────────────────────────────────────────────────────
    no_arch = [p for p in policies
               if not (p.get("remoteTargetPolicy") or {}).get("archivalTargets")]
    if policies and len(no_arch) > len(policies) * 0.5:
        recs.append(_r("MEDIUM", "Governance",
            f"{len(no_arch)}/{len(policies)} policies lack an archival target",
            "Add archival targets for long-term retention and ransomware recovery",
            "Without archival, data is unrecoverable after local retention expires"))

    # ── Views ────────────────────────────────────────────────────────────────
    no_quota = [v for v in views
                if not (v.get("logicalQuota") or
                        (v.get("storagePolicy") or {}).get("storageQuotaPolicy"))]
    if no_quota:
        recs.append(_r("LOW", "Governance",
            f"{len(no_quota)} NAS view(s) have no storage quota",
            "Apply quotas to all views to prevent uncontrolled capacity growth",
            "Unquoted views can silently consume all available capacity"))

    near_quota = []
    for v in views:
        lq  = v.get("logicalQuota") or \
              (v.get("storagePolicy") or {}).get("storageQuotaPolicy") or {}
        q   = lq.get("hardLimitBytes") or 0
        u   = ((v.get("stats") or {}).get("dataUsageStats") or {}) \
                .get("totalLogicalUsageBytes") or 0
        if q and u and (u / q * 100) >= VIEW_QUOTA_CRIT_PCT:
            near_quota.append(v.get("name", "?"))
    if near_quota:
        recs.append(_r("HIGH", "Capacity",
            f"{len(near_quota)} view(s) at ≥{VIEW_QUOTA_CRIT_PCT}% of quota",
            f"Expand quota or clean up: {', '.join(near_quota[:3])}",
            "Views at quota block all new writes to that share"))

    recs.sort(key=lambda x: _PRIORITY_ORDER.get(x["priority"], 99))
    return recs


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL SHEETS
# ─────────────────────────────────────────────────────────────────────────────

def _success_stats(cd):
    """Return (succ_pct, sla_pct) from run history or last-run fallback."""
    all_runs = [r for runs in cd["group_runs"].values() for r in runs]
    groups   = cd["groups"]
    if all_runs:
        total = len(all_runs)
        succ  = sum(1 for r in all_runs
                    if (r.get("localBackupInfo") or {}).get("status", "")
                    in ("kSuccess", "Succeeded", "kWarning"))
        viols = sum(1 for r in all_runs if r.get("isSlaViolated", False))
        return (
            round(succ / total * 100, 1) if total else 100,
            round((total - viols) / total * 100, 1) if total else 100,
        )
    # quick mode
    if not groups:
        return (100, 100)
    total = len(groups)
    succ  = sum(1 for g in groups
                if (g.get("lastRun") or {}).get("localBackupInfo", {})
                   .get("status", "") in ("kSuccess", "Succeeded", "kWarning"))
    viols = sum(1 for g in groups
                if (g.get("lastRun") or {}).get("isSlaViolated", False))
    return (
        round(succ / total * 100, 1),
        round((total - viols) / total * 100, 1),
    )


def _cap_stats(cd):
    """
    Return (usable_bytes, used_bytes, cap_pct_or_NA) for a cluster.
    Primary source: usagePerfStats from /v1/public/cluster?fetchStats=true.
    Falls back to summing storageConsumedBytes across storage domains when the
    cluster-level stats field comes back null (happens on some versions/configs).
    Returns cap_pct as a float, or the string "N/A" when capacity is unknown.
    """
    info   = cd["info"]
    usage  = ((info.get("stats") or {}).get("usagePerfStats") or {})
    usable = usage.get("physicalCapacityBytes") or 0
    used   = usage.get("totalPhysicalUsageBytes") or 0

    if not usable:
        # Fallback: sum physical used across domains (no usable total available)
        domain_used = sum(
            (d.get("stats") or {}).get("storageConsumedBytes") or 0
            for d in cd.get("domains", [])
        )
        return 0, domain_used, "N/A"

    pct = round(used / usable * 100, 1)
    return usable, used, pct


def _sheet_summary(wb, all_data):
    ws = wb.create_sheet("Executive Summary")
    ws.freeze_panes = "A3"
    _title(ws, "Cohesity Health Check — Executive Summary", "M")

    cols = ["Cluster", "Software Version", "Node Count",
            "Health Score", "Grade",
            "30d Success %", "SLA Pass %", "Capacity Used %",
            "Open Criticals", "Open Warnings",
            "Security Score /100", "Recommendations", "Top Finding"]
    _hdr(ws, 2, cols)

    for cd in all_data:
        info   = cd["info"]
        scores = cd["scores"]
        alerts = cd["alerts"]
        _, _, cap_pct = _cap_stats(cd)
        crits  = sum(1 for a in alerts if a.get("severity") == "kCritical")
        warns  = sum(1 for a in alerts if a.get("severity") == "kWarning")
        succ_pct, sla_pct = _success_stats(cd)
        top = cd["recommendations"][0]["finding"] if cd["recommendations"] else "None identified"

        row = [cd["name"],
               info.get("clusterSoftwareVersion", ""), len(cd["nodes"]),
               scores["overall"], scores["grade"],
               succ_pct, sla_pct, cap_pct,
               crits, warns,
               scores["security"], len(cd["recommendations"]), top]
        r = ws.max_row + 1
        for c, val in enumerate(row, 1):
            ws.cell(row=r, column=c, value=val).font = _font()

        # RAG colour on score + grade
        for col in (4, 5):
            ws.cell(row=r, column=col).fill = _rag_fill(scores["overall"])
            ws.cell(row=r, column=col).font = _font(bold=True)

        # Red crits
        if crits:
            ws.cell(row=r, column=9).fill = _fill(RED)
            ws.cell(row=r, column=9).font = _font(bold=True, color=WHITE)

    auto_fit_columns(ws)
    ws.column_dimensions[ws.cell(row=2, column=13).column_letter].width = 55


def _sheet_infrastructure(wb, all_data):
    ws = wb.create_sheet("Infrastructure")
    ws.freeze_panes = "A3"
    _title(ws, "Infrastructure — Cluster & Node Health", "N")

    cols = ["Cluster", "Software Version", "Cluster Type", "Cluster ID",
            "Node Count", "Healthy Nodes", "Disk Count",
            "Domain Name", "DNS Servers", "NTP Servers", "Timezone",
            "Encryption", "FIPS", "Helios Status"]
    _hdr(ws, 2, cols)

    for cd in all_data:
        info  = cd["info"]
        nodes = cd["nodes"]
        healthy = sum(
            1 for n in nodes
            if ((n.get("nodeStatus") or {}).get("overallStatus") == "kNormal"
                or n.get("removalState") in (None, "kDontRemove"))
        )
        # Disk count — v1 /public/nodes returns diskCount as an integer per node
        disks = sum(n.get("diskCount") or 0 for n in nodes)

        enc  = bool(info.get("encryptionEnabled") or
                    (info.get("encryptionConfig") or {}).get("encryptionEnabled"))
        fips = bool(info.get("fipsCompliant") or info.get("fipsEnabled"))
        dns  = ", ".join(info.get("dnsServerIps") or [])
        # NTP: v1 API returns ntpSettings.ntpServers or a flat ntpServers list
        ntp_raw = ((info.get("ntpSettings") or {}).get("ntpServers")
                   or info.get("ntpSettings", {}).get("ntpServersList")
                   or info.get("ntpServers")
                   or info.get("ntpServerIps")
                   or [])
        if isinstance(ntp_raw, list):
            ntp = ", ".join(str(s) for s in ntp_raw) if ntp_raw else "Not configured"
        else:
            ntp = str(ntp_raw) if ntp_raw else "Not configured"
        domain = (info.get("domainNames") or [""])[0]

        row = [cd["name"], info.get("clusterSoftwareVersion", ""),
               info.get("clusterType", ""), str(info.get("id", "")),
               len(nodes), healthy, disks if disks else "",
               domain, dns, ntp, info.get("timezone", ""),
               "Yes" if enc else "No",
               "Yes" if fips else "No",
               "Connected"]
        r = ws.max_row + 1
        for c, val in enumerate(row, 1):
            ws.cell(row=r, column=c, value=val).font = _font()

        # Colour encryption / FIPS
        enc_cell = ws.cell(row=r, column=12)
        enc_cell.fill = _fill(LT_GREEN) if enc else _fill(RED)
        enc_cell.font = _font(bold=True, color=WHITE if not enc else "000000")
        fips_cell = ws.cell(row=r, column=13)
        fips_cell.fill = _fill(LT_GREEN) if fips else _fill(YELLOW)
        fips_cell.font = _font(bold=True)

        # Node health
        if healthy < len(nodes):
            ws.cell(row=r, column=6).fill = _fill(ORANGE)
            ws.cell(row=r, column=6).font = _font(bold=True)

    auto_fit_columns(ws)


def _sheet_protection(wb, all_data):
    ws = wb.create_sheet("Protection Health")
    ws.freeze_panes = "A3"
    _title(ws, "Protection Health — Group Status & Run Metrics", "S")

    cols = ["Cluster", "Group Name", "Policy",
            "Active", "Paused",
            "Last Run Type", "Last Run Status", "SLA Violated",
            "Last Run Start", "Last Run End", "Duration",
            "Objects OK", "Objects Failed",
            "Logical (GB)", "Physical (GB)",
            "Replication Status", "Archival Status",
            "RPO (hrs)", "Flag"]
    _hdr(ws, 2, cols)

    for cd in all_data:
        pm = cd.get("policy_map", {})
        for g in cd["groups"]:
            lr  = g.get("lastRun") or {}
            lb  = lr.get("localBackupInfo") or {}

            # runType and isSlaViolated sit on localBackupInfo (confirmed from runs API)
            run_type    = lb.get("runType") or lr.get("runType", "")
            sla_violated = lb.get("isSlaViolated") or lr.get("isSlaViolated") or False
            status      = lb.get("status", "")

            start_u = lb.get("startTimeUsecs") or 0
            end_u   = lb.get("endTimeUsecs") or 0
            dur = ""
            if start_u and end_u:
                s = int((end_u - start_u) / 1_000_000)
                h, m = divmod(s // 60, 60)
                dur = f"{h}h {m:02d}m"
            rpo_hrs = round((_now_usecs() - start_u) / 3_600_000_000, 1) \
                      if start_u else ""

            # Object counts are directly on localBackupInfo (not nested in stats)
            objs_ok   = lb.get("successfulObjectsCount", "")
            objs_fail = lb.get("failedObjectsCount", "")
            # Sizes are under localSnapshotStats sub-dict
            snap = lb.get("localSnapshotStats") or {}
            logical_gb  = round(bytes_to_gb(snap.get("logicalSizeBytes") or 0), 2)
            physical_gb = round(bytes_to_gb(snap.get("bytesWritten") or 0), 2)

            # Replication / archival status from lastRun (same level as localBackupInfo)
            repl_st = arch_st = ""
            for ri in ((lr.get("replicationInfo") or {})
                       .get("replicationTargetResults") or []):
                repl_st = ri.get("status", ""); break
            for ai in ((lr.get("archivalInfo") or {})
                       .get("archivalTargetResults") or []):
                arch_st = ai.get("status", ""); break

            flag = "OK"
            if g.get("isPaused"):
                flag = "PAUSED"
            elif status in ("kFailure", "Failed"):
                flag = "FAILED"
            elif sla_violated:
                flag = "SLA VIOLATED"
            elif isinstance(rpo_hrs, float) and rpo_hrs > RPO_CRIT_HOURS:
                flag = f"RPO GAP {rpo_hrs:.0f}h"
            elif isinstance(rpo_hrs, float) and rpo_hrs > RPO_WARN_HOURS:
                flag = f"RPO WARN {rpo_hrs:.0f}h"

            # Policy name: prefer policyName on group; fall back to policy_map lookup
            policy_name = (g.get("policyName")
                           or pm.get(g.get("policyId"), "")
                           or g.get("policyId", ""))

            row = [cd["name"], g.get("name", ""),
                   policy_name,
                   "Yes" if g.get("isActive", True) else "No",
                   "Yes" if g.get("isPaused", False) else "No",
                   run_type, status,
                   "Yes" if sla_violated else "No",
                   usecs_to_datetime(start_u), usecs_to_datetime(end_u), dur,
                   objs_ok, objs_fail,
                   logical_gb, physical_gb,
                   repl_st, arch_st, rpo_hrs, flag]
            rn = ws.max_row + 1
            for c, val in enumerate(row, 1):
                ws.cell(row=rn, column=c, value=val).font = _font()

            fc = ws.cell(row=rn, column=19)
            if flag == "FAILED":
                fc.fill = _fill(RED);   fc.font = _font(bold=True, color=WHITE)
            elif flag in ("PAUSED",) or "RPO GAP" in flag:
                fc.fill = _fill(ORANGE); fc.font = _font(bold=True)
            elif flag == "SLA VIOLATED" or "RPO WARN" in flag:
                fc.fill = _fill(YELLOW); fc.font = _font(bold=True)
            else:
                fc.fill = _fill(LT_GREEN)

    auto_fit_columns(ws)


def _sheet_storage(wb, all_data):
    ws = wb.create_sheet("Storage & Capacity")
    ws.freeze_panes = "A3"
    _title(ws, "Storage & Capacity — Utilisation & Data Reduction", "L")

    cols = ["Cluster", "Storage Domain", "Usable (TB)", "Used (TB)",
            "Free (TB)", "Used %", "Logical Data (TB)",
            "Data Reduction Ratio", "Dedup Ratio", "Compression Ratio",
            "Savings (TB)", "Status"]
    _hdr(ws, 2, cols)

    for cd in all_data:
        info   = cd["info"]
        stats  = (info.get("stats") or {})
        usage  = (stats.get("usagePerfStats") or {})
        data   = (stats.get("dataUsageStats") or {})

        usable, used, cap_pct = _cap_stats(cd)
        logical     = usage.get("dataInBytes") or 0
        physical    = usage.get("dataInBytesAfterReduction") or 0
        dedup_after = data.get("dataInBytesAfterDedup") or 0
        dr_ratio    = round((stats.get("dataReductionRatio") or 0.0), 2)
        dedup_ratio = round(logical / dedup_after, 2) if dedup_after else 0
        comp_ratio  = round(dedup_after / physical, 2) if physical else 0
        free        = usable - used
        savings     = round(bytes_to_tb(logical - physical), 2) \
                      if logical > physical else 0

        flag = ""
        if isinstance(cap_pct, float):
            flag = ("CRITICAL" if cap_pct >= CAPACITY_CRIT_PCT
                    else "WARN" if cap_pct >= CAPACITY_WARN_PCT else "OK")
        elif cap_pct == "N/A":
            flag = "N/A"

        row = [cd["name"], "(Cluster Total)",
               round(bytes_to_tb(usable), 2) or "",
               round(bytes_to_tb(used), 2) or "",
               round(bytes_to_tb(free), 2) or "",
               cap_pct,
               round(bytes_to_tb(logical), 2),
               dr_ratio, dedup_ratio, comp_ratio, savings, flag]
        rn = ws.max_row + 1
        for c, val in enumerate(row, 1):
            cell = ws.cell(row=rn, column=c, value=val)
            cell.font = _font(bold=(c == 2))

        sc = ws.cell(row=rn, column=12)
        if flag == "CRITICAL":
            sc.fill = _fill(RED);    sc.font = _font(bold=True, color=WHITE)
        elif flag == "WARN":
            sc.fill = _fill(ORANGE); sc.font = _font(bold=True)
        elif flag == "OK":
            sc.fill = _fill(LT_GREEN)

        # Per-domain rows
        for d in cd["domains"]:
            ds = d.get("stats") or {}
            ld = ds.get("dataInBytes") or 0
            da = ds.get("dataInBytesAfterDedup") or 0
            ph = ds.get("storageConsumedBytes") or 0
            drow = [cd["name"], d.get("name", ""),
                    "", "", "", "",
                    round(bytes_to_tb(ld), 2),
                    round(ld / ph, 2) if ph else 0,
                    round(ld / da, 2) if da else 0,
                    round(da / ph, 2) if ph else 0,
                    round(bytes_to_tb(ld - ph), 2) if ld > ph else 0, ""]
            rn2 = ws.max_row + 1
            for c, val in enumerate(drow, 1):
                ws.cell(row=rn2, column=c, value=val).font = _font()

    auto_fit_columns(ws)


def _ret_str(r):
    if not r: return ""
    d = r.get("duration", "")
    u = r.get("unit", "")
    return f"{d} {u}".strip() if d else u

def _sched_str(s):
    """Convert a v2 schedule dict to a human-readable string."""
    if not s: return ""
    unit = s.get("unit", "")
    if unit == "Days":
        freq = (s.get("daySchedule") or {}).get("frequency", 1)
        return f"Every {freq} day{'s' if freq != 1 else ''}"
    if unit == "Weeks":
        days = (s.get("weekSchedule") or {}).get("dayOfWeek") or []
        return f"Weekly ({', '.join(days[:2])}{'…' if len(days) > 2 else ''})" if days else "Weekly"
    if unit == "Hours":
        freq = (s.get("hourSchedule") or {}).get("frequency", 1)
        return f"Every {freq}h"
    if unit == "Minutes":
        freq = (s.get("minuteSchedule") or {}).get("frequency", 1)
        return f"Every {freq}m"
    if unit in ("Months", "Monthly"): return "Monthly"
    if unit in ("Years", "Yearly"):   return "Yearly"
    # fallback: flat frequency dict
    for key in ("minutes", "hours", "days", "weeks", "months", "years"):
        if isinstance(s, dict) and key in s:
            return f"Every {s[key]} {key}"
    return unit or ""

def _sheet_policies(wb, all_data):
    ws = wb.create_sheet("Policy Audit")
    ws.freeze_panes = "A3"
    _title(ws, "Policy Audit — Retention, Replication & Archival", "O")

    cols = ["Cluster", "Policy Name", "Type",
            "Full Schedule", "Incremental Schedule",
            "Local Retention", "Log Retention",
            "Has Replication", "Replication Target", "Replication Retention",
            "Has Archival", "Archival Target", "Archival Retention",
            "Groups Using", "Gaps"]
    _hdr(ws, 2, cols)

    for cd in all_data:
        # Count protection groups per policy ID so column N is populated
        policy_group_count: dict = {}
        for g in cd["groups"]:
            pid = g.get("policyId")
            if pid:
                policy_group_count[pid] = policy_group_count.get(pid, 0) + 1

        for p in cd["policies"]:
            bp  = p.get("backupPolicy") or {}
            rtp = p.get("remoteTargetPolicy") or {}
            rep_tgts  = rtp.get("replicationTargets") or []
            arch_tgts = rtp.get("archivalTargets") or []

            # v2 API replication target name: targetName or remoteTargetConfig.clusterName/name
            def _rep_name(t):
                return (t.get("targetName")
                        or (t.get("remoteTargetConfig") or {}).get("clusterName")
                        or (t.get("remoteTargetConfig") or {}).get("name")
                        or "")
            # v2 API archival target name: targetName or archivalTargetConfig.name/vaultName
            def _arch_name(t):
                return (t.get("targetName")
                        or (t.get("archivalTargetConfig") or {}).get("name")
                        or (t.get("archivalTargetConfig") or {}).get("vaultName")
                        or "")

            rep_names  = ", ".join(filter(None, (_rep_name(t)  for t in rep_tgts)))
            arch_names = ", ".join(filter(None, (_arch_name(t) for t in arch_tgts)))

            # v2 schedule: backupPolicy.regular.full.schedule / regular.incremental.schedule
            regular     = bp.get("regular") or {}
            full_sched  = (regular.get("full") or bp.get("full") or {}).get("schedule") or {}
            incr_sched  = (regular.get("incremental") or bp.get("incremental") or {}).get("schedule") or {}

            local_ret = regular.get("retention") or bp.get("retention") or {}
            log_ret   = (bp.get("log") or {}).get("retention") or {}
            rep_ret   = (rep_tgts[0].get("retention") or {}) if rep_tgts else {}
            arch_ret  = (arch_tgts[0].get("retention") or {}) if arch_tgts else {}

            gaps = []
            if not rep_tgts:  gaps.append("No replication")
            if not arch_tgts: gaps.append("No archival")

            row = [cd["name"], p.get("name", ""),
                   p.get("type") or p.get("policyCategory") or "",
                   _sched_str(full_sched), _sched_str(incr_sched),
                   _ret_str(local_ret), _ret_str(log_ret),
                   "Yes" if rep_tgts else "No", rep_names, _ret_str(rep_ret),
                   "Yes" if arch_tgts else "No", arch_names, _ret_str(arch_ret),
                   policy_group_count.get(p.get("id"), 0),
                   "; ".join(gaps) if gaps else "OK"]
            rn = ws.max_row + 1
            for c, val in enumerate(row, 1):
                ws.cell(row=rn, column=c, value=val).font = _font()

            gc = ws.cell(row=rn, column=15)
            if gaps:
                gc.fill = _fill(YELLOW)
                gc.font = _font(bold=True)

    auto_fit_columns(ws)


def _sheet_policy_groups(wb, all_data):
    """Sheet: one row per protection group showing which policy governs it."""
    ws = wb.create_sheet("Policy → Groups")
    ws.freeze_panes = "A3"
    _title(ws, "Policy → Groups — Protection Groups Associated with Each Policy", "G")

    cols = ["Cluster", "Policy Name", "Group Name", "Environment",
            "Active", "Paused", "Last Run Status"]
    _hdr(ws, 2, cols)

    for cd in all_data:
        pm = cd.get("policy_map", {})
        # Sort by policy name then group name for easy reading
        sorted_groups = sorted(
            cd["groups"],
            key=lambda g: (
                g.get("policyName") or pm.get(g.get("policyId"), "") or "",
                g.get("name", "")
            )
        )
        for g in sorted_groups:
            policy_name = (g.get("policyName")
                           or pm.get(g.get("policyId"), "")
                           or g.get("policyId", ""))
            lb     = (g.get("lastRun") or {}).get("localBackupInfo") or {}
            status = lb.get("status", "")
            paused = g.get("isPaused", False)

            row = [cd["name"], policy_name, g.get("name", ""),
                   g.get("environment", ""),
                   "Yes" if g.get("isActive", True) else "No",
                   "Yes" if paused else "No",
                   status]
            rn = ws.max_row + 1
            for c, val in enumerate(row, 1):
                ws.cell(row=rn, column=c, value=val).font = _font()

            # Highlight paused groups
            if paused:
                ws.cell(row=rn, column=6).fill = _fill(YELLOW)
                ws.cell(row=rn, column=6).font = _font(bold=True)
            # Highlight failed status
            if status in ("kFailure", "Failed"):
                ws.cell(row=rn, column=7).fill = _fill(RED)
                ws.cell(row=rn, column=7).font = _font(bold=True, color=WHITE)

    auto_fit_columns(ws)


def _sheet_alerts(wb, all_data):
    ws = wb.create_sheet("Alerts")
    ws.freeze_panes = "A3"
    _title(ws, "Open Alerts — Severity, Age & Description", "J")

    cols = ["Cluster", "Severity", "Category", "Alert Code",
            "Alert Name", "Description",
            "First Occurred", "Age (Days)", "Last Updated",
            "Acknowledged"]
    _hdr(ws, 2, cols)

    sev_order = {"kCritical": 0, "kWarning": 1, "kInfo": 2}

    for cd in all_data:
        sorted_a = sorted(cd["alerts"],
                          key=lambda a: (sev_order.get(a.get("severity"), 9),
                                         a.get("firstTimestampUsecs") or 0))
        for a in sorted_a:
            # v1 alerts API uses firstTimestampUsecs (not firstOccurrenceTime)
            first_u = a.get("firstTimestampUsecs") or 0
            last_u  = a.get("latestTimestampUsecs") or 0
            age = round((_now_usecs() - first_u) / (86400 * 1_000_000), 1) \
                  if first_u else ""
            doc  = a.get("alertDocument") or {}
            desc = doc.get("alertDescription") or ""

            row = [cd["name"], a.get("severity", ""),
                   a.get("alertCategory", ""), a.get("alertCode", ""),
                   doc.get("alertName", ""),
                   desc[:120] + ("…" if len(desc) > 120 else ""),
                   usecs_to_datetime(first_u), age, usecs_to_datetime(last_u),
                   "Yes" if a.get("acknowledgeTimestampUsecs") else "No"]
            rn = ws.max_row + 1
            for c, val in enumerate(row, 1):
                ws.cell(row=rn, column=c, value=val).font = _font()

            sc = ws.cell(row=rn, column=2)
            if a.get("severity") == "kCritical":
                sc.fill = _fill(RED);    sc.font = _font(bold=True, color=WHITE)
            elif a.get("severity") == "kWarning":
                sc.fill = _fill(ORANGE); sc.font = _font(bold=True)

    auto_fit_columns(ws)
    ws.column_dimensions["F"].width = 70


def _sheet_security(wb, all_data):
    ws = wb.create_sheet("Security")
    ws.freeze_panes = "A3"
    _title(ws, "Security Posture — Checklist vs Best Practices", "K")

    cols = ["Cluster", "Encryption at Rest", "FIPS Mode",
            "Vault Configured", "FortKnox / Immutable",
            "Replication", "Archival",
            "Min Local Retention", "Min Archival Retention",
            "Security Score /100", "Gaps"]
    _hdr(ws, 2, cols)

    _RET_MULT = {"Days": 1, "Weeks": 7, "Months": 30, "Years": 365}

    def _ret_to_days(r):
        d = (r or {}).get("duration") or 0
        u = (r or {}).get("unit", "Days")
        return d * _RET_MULT.get(u, 1) if d else 0

    def _min_local_days(policies):
        vals = []
        for p in policies:
            bp  = p.get("backupPolicy") or {}
            reg = bp.get("regular") or {}
            r   = reg.get("retention") or bp.get("retention") or {}
            v   = _ret_to_days(r)
            if v: vals.append(v)
        return f"{min(vals)} days" if vals else "N/A"

    def _min_arch_days(policies):
        vals = []
        for p in policies:
            for t in ((p.get("remoteTargetPolicy") or {}).get("archivalTargets") or []):
                v = _ret_to_days(t.get("retention") or {})
                if v: vals.append(v)
        return f"{min(vals)} days" if vals else "N/A"

    for cd in all_data:
        info     = cd["info"]
        vaults   = cd["vaults"]
        policies = cd["policies"]
        scores   = cd["scores"]

        enc  = bool(info.get("encryptionEnabled") or
                    (info.get("encryptionConfig") or {}).get("encryptionEnabled"))
        fips = bool(info.get("fipsCompliant") or info.get("fipsEnabled"))
        has_vault = bool(vaults)
        fk_st    = _fk_status(cd)
        has_repl = any((p.get("remoteTargetPolicy") or {}).get("replicationTargets")
                       for p in policies)
        has_arch = any((p.get("remoteTargetPolicy") or {}).get("archivalTargets")
                       for p in policies)

        gaps = []
        if not enc:       gaps.append("No encryption")
        if not fips:      gaps.append("FIPS off")
        if not has_vault: gaps.append("No vault")
        if fk_st == "none": gaps.append("No immutability")
        if fk_st == "idle": gaps.append("FK configured — not sending data")
        if not has_repl:  gaps.append("No replication")
        if not has_arch:  gaps.append("No archival")

        fk_label = {"active": "Active", "idle": "Configured — Idle", "none": "No"}[fk_st]
        yn = lambda v: "Yes" if v else "No"
        row = [cd["name"], yn(enc), yn(fips), yn(has_vault), fk_label,
               yn(has_repl), yn(has_arch),
               _min_local_days(policies), _min_arch_days(policies),
               scores["security"],
               "; ".join(gaps) if gaps else "OK"]
        rn = ws.max_row + 1
        for c, val in enumerate(row, 1):
            ws.cell(row=rn, column=c, value=val).font = _font()

        for col, flag in [(2, enc), (3, fips), (4, has_vault), (6, has_repl), (7, has_arch)]:
            cell = ws.cell(row=rn, column=col)
            cell.fill = _fill(LT_GREEN) if flag else _fill(RED)
            cell.font = _font(bold=True, color=WHITE if not flag else "000000")

        # FK column 5: three-state colouring
        fk_cell = ws.cell(row=rn, column=5)
        if fk_st == "active":
            fk_cell.fill = _fill(LT_GREEN)
            fk_cell.font = _font(bold=True)
        elif fk_st == "idle":
            fk_cell.fill = _fill(ORANGE)
            fk_cell.font = _font(bold=True)
        else:
            fk_cell.fill = _fill(RED)
            fk_cell.font = _font(bold=True, color=WHITE)

        sc = ws.cell(row=rn, column=10)
        sc.fill = _rag_fill(scores["security"])
        sc.font = _font(bold=True)

    auto_fit_columns(ws)
    ws.column_dimensions["K"].width = 55


def _sheet_replication(wb, all_data):
    ws = wb.create_sheet("Replication & Archive")
    ws.freeze_panes = "A3"
    _title(ws, "Replication & Archival Targets", "K")

    cols = ["Cluster", "Target Type", "Target Name", "Vault Type",
            "Protection Groups",
            "Vault Storage (TB)",
            "Logical Data Transferred (TB)\n[numLogicalBytesTransferred — cumulative]",
            "Physical Data Transferred (TB)\n[numPhysicalBytesTransferred — cumulative]",
            "Status", "Notes",
            "Protection Group Names"]
    _hdr(ws, 2, cols)

    for cd in all_data:
        # ── Vault ID → name / type maps ───────────────────────────────────
        vault_name_by_id   = {}   # vault id → name
        vault_type_by_name = {}   # name     → vaultType
        vault_id_by_name   = {}   # name     → id
        for v in cd["vaults"]:
            vn = v.get("name") or v.get("vaultName") or ""
            vi = v.get("id")
            vt = v.get("vaultType", "")
            if vi: vault_name_by_id[vi] = vn or f"Vault-{vi}"
            if vn:
                vault_type_by_name[vn] = vt
                if vi: vault_id_by_name[vn] = vi

        # ── Target name extractors ────────────────────────────────────────
        def _rep_name(t):
            cfg = t.get("remoteTargetConfig") or {}
            return (t.get("targetName")
                    or cfg.get("clusterName")
                    or cfg.get("name") or "")

        def _arch_name(t):
            cfg  = t.get("archivalTargetConfig") or {}
            name = (t.get("targetName")
                    or cfg.get("name")
                    or cfg.get("vaultName") or "")
            if not name:
                vid = cfg.get("vaultId") or cfg.get("id")
                if vid:
                    name = vault_name_by_id.get(vid, f"Vault-{vid}")
            return name

        # ── Policy chain: policy id → target names ───────────────────────
        policy_rep_targets  = {}   # policy id → set of replication cluster names
        policy_arch_targets = {}   # policy id → set of vault names
        arch_target_type    = {}   # vault name → targetType

        for p in cd["policies"]:
            pid = p.get("id")
            if not pid:
                continue
            rtp = p.get("remoteTargetPolicy") or {}
            for t in (rtp.get("replicationTargets") or []):
                n = _rep_name(t)
                if n:
                    policy_rep_targets.setdefault(pid, set()).add(n)
            for t in (rtp.get("archivalTargets") or []):
                n = _arch_name(t)
                if n:
                    policy_arch_targets.setdefault(pid, set()).add(n)
                    if n not in arch_target_type:
                        cfg = t.get("archivalTargetConfig") or {}
                        arch_target_type[n] = (t.get("targetType")
                                               or cfg.get("targetType") or "")

        # ── Group → target mapping (policy chain) ────────────────────────
        rep_groups  = {}   # replication cluster name → [group names]
        arch_groups = {}   # vault name               → [group names]

        for g in cd["groups"]:
            pid   = g.get("policyId")
            gname = g.get("name", "")
            if not pid or not gname:
                continue
            for tname in (policy_rep_targets.get(pid) or set()):
                rep_groups.setdefault(tname, []).append(gname)
            for tname in (policy_arch_targets.get(pid) or set()):
                arch_groups.setdefault(tname, []).append(gname)

        # ── Supplement arch_groups from fk_data (direct transfer records) ─
        # fk_data lists protectionJobName per vault — reliable even when the
        # policy chain fails to extract vault names (e.g. ID-only references).
        for fk in (cd["fk_data"] or []):
            fk_vn = fk.get("vaultName", "")
            if not fk_vn:
                continue
            for job in (fk.get("dataTransferPerProtectionJob") or []):
                jname = job.get("protectionJobName", "")
                if jname and jname not in (arch_groups.get(fk_vn) or []):
                    arch_groups.setdefault(fk_vn, []).append(jname)

        # All unique target names — union of policy chain + fk_data + vault list
        all_rep_targets  = set().union(*policy_rep_targets.values()) \
                           if policy_rep_targets else set()
        fk_vault_names   = {fk.get("vaultName") for fk in (cd["fk_data"] or [])
                            if fk.get("vaultName")}
        vault_list_names = {v.get("name") or v.get("vaultName") or ""
                            for v in cd["vaults"]} - {""}
        all_arch_targets = (
            (set().union(*policy_arch_targets.values()) if policy_arch_targets else set())
            | fk_vault_names
            | vault_list_names
        )

        # ── FK transfer stats ─────────────────────────────────────────────
        fk_by_name = {}
        fk_by_id   = {}
        for fk in (cd["fk_data"] or []):
            fk_vn  = fk.get("vaultName", "")
            fk_vid = fk.get("vaultId")
            jobs   = fk.get("dataTransferPerProtectionJob") or []
            data   = {
                "consumed": sum((j.get("storageConsumed") or 0) for j in jobs),
                "logical":  sum((j.get("numLogicalBytesTransferred") or 0) for j in jobs),
                "physical": sum((j.get("numPhysicalBytesTransferred") or 0) for j in jobs),
            }
            if fk_vn:  fk_by_name[fk_vn] = data
            if fk_vid: fk_by_id[fk_vid]   = data

        def _fk_stats(vname):
            d = fk_by_name.get(vname)
            if not d:
                vid = vault_id_by_name.get(vname)
                if vid: d = fk_by_id.get(vid)
            return d or {}

        fk_types = {"rpaas", "fortknox", "kfortknox", "krpaas"}

        def _is_fk(vname, vtype):
            return (any(x in vtype.lower() for x in fk_types)
                    or "fortknox" in vname.lower())

        # ── Replication rows ──────────────────────────────────────────────
        for tname in sorted(all_rep_targets):
            groups = sorted(rep_groups.get(tname) or [])
            gcnt   = len(groups)
            status = "Configured" if gcnt else "No groups assigned"
            rn = ws.max_row + 1
            row = [cd["name"], "Replication", tname, "", gcnt,
                   "", "", "", status, "",
                   ", ".join(groups)]
            for c, val in enumerate(row, 1):
                ws.cell(row=rn, column=c, value=val).font = _font()

        # ── Archival / vault rows ─────────────────────────────────────────
        shown = set()
        for tname in sorted(all_arch_targets):
            groups = sorted(arch_groups.get(tname) or [])
            gcnt   = len(groups)
            vtype  = arch_target_type.get(tname) or vault_type_by_name.get(tname, "")
            fkd    = _fk_stats(tname)
            note   = "FortKnox/RPaaS" if _is_fk(tname, vtype) else ""
            status = "Configured" if gcnt else "No groups assigned"
            rn     = ws.max_row + 1
            row    = [cd["name"], "Archival/Vault", tname, vtype, gcnt,
                      round(bytes_to_tb(fkd.get("consumed") or 0), 3),
                      round(bytes_to_tb(fkd.get("logical")  or 0), 3),
                      round(bytes_to_tb(fkd.get("physical") or 0), 3),
                      status, note,
                      ", ".join(groups)]
            for c, val in enumerate(row, 1):
                ws.cell(row=rn, column=c, value=val).font = _font()
            shown.add(tname)

        # Vaults that exist in the API but are not referenced by any policy
        for v in cd["vaults"]:
            vname = v.get("name") or v.get("vaultName") or ""
            if not vname or vname in shown:
                continue
            vtype = v.get("vaultType", "")
            fkd   = _fk_stats(vname)
            note  = "FortKnox/RPaaS" if _is_fk(vname, vtype) else ""
            rn    = ws.max_row + 1
            row   = [cd["name"], "Archival/Vault", vname, vtype, 0,
                     round(bytes_to_tb(fkd.get("consumed") or 0), 3),
                     round(bytes_to_tb(fkd.get("logical")  or 0), 3),
                     round(bytes_to_tb(fkd.get("physical") or 0), 3),
                     "Vault exists — not referenced by any policy", note, ""]
            for c, val in enumerate(row, 1):
                ws.cell(row=rn, column=c, value=val).font = _font()

    auto_fit_columns(ws)
    ws.column_dimensions["K"].width = 60


def _sheet_views(wb, all_data):
    ws = wb.create_sheet("Data Services")
    ws.freeze_panes = "A3"
    _title(ws, "NAS Views — Quota Utilisation & Protocols", "L")

    cols = ["Cluster", "View Name", "Storage Domain", "Protocols",
            "Quota (GB)", "Used (GB)", "Usage %",
            "Case Insensitive", "Created",
            "SMB Mount Paths", "NFS Mount Paths", "Status"]
    _hdr(ws, 2, cols)

    for cd in all_data:
        for v in cd["views"]:
            lq  = v.get("logicalQuota") or \
                  (v.get("storagePolicy") or {}).get("storageQuotaPolicy") or {}
            quota_b = lq.get("hardLimitBytes") or 0
            vstats  = ((v.get("stats") or {}).get("dataUsageStats") or {})
            used_b  = vstats.get("totalLogicalUsageBytes") or 0
            quota_gb = round(bytes_to_gb(quota_b), 2) if quota_b else ""
            used_gb  = round(bytes_to_gb(used_b), 2)
            pct      = round(used_b / quota_b * 100, 1) if quota_b else ""

            protos = []
            if v.get("enableSmbViewDiscovery") or v.get("smbMountPaths"): protos.append("SMB")
            if v.get("nfsMountPaths") or v.get("enableNfsViewDiscovery"):  protos.append("NFS")
            if v.get("s3AccessEnabled"):                                   protos.append("S3")

            smb = ", ".join(v.get("smbMountPaths") or [])
            nfs = ", ".join(v.get("nfsMountPaths") or [])
            created_ms = v.get("createTimeMsecs") or 0
            created = (datetime.datetime.fromtimestamp(created_ms / 1000,
                                                        tz=datetime.timezone.utc)
                       .strftime("%Y-%m-%d") if created_ms else "")

            flag = "OK"
            if not quota_b:
                flag = "NO QUOTA"
            elif isinstance(pct, float) and pct >= VIEW_QUOTA_CRIT_PCT:
                flag = f"CRITICAL {pct:.0f}%"
            elif isinstance(pct, float) and pct >= VIEW_QUOTA_WARN_PCT:
                flag = f"WARN {pct:.0f}%"

            row = [cd["name"], v.get("name", ""), v.get("storageDomainName", ""),
                   "/".join(protos), quota_gb, used_gb, pct,
                   "Yes" if v.get("caseInsensitiveNamesEnabled") else "No",
                   created, smb, nfs, flag]
            rn = ws.max_row + 1
            for c, val in enumerate(row, 1):
                ws.cell(row=rn, column=c, value=val).font = _font()

            fc = ws.cell(row=rn, column=12)
            if "CRITICAL" in flag:
                fc.fill = _fill(RED);    fc.font = _font(bold=True, color=WHITE)
            elif "WARN" in flag:
                fc.fill = _fill(ORANGE); fc.font = _font(bold=True)
            elif "NO QUOTA" in flag:
                fc.fill = _fill(YELLOW); fc.font = _font(bold=True)
            else:
                fc.fill = _fill(LT_GREEN)

    auto_fit_columns(ws)


def _sheet_coverage(wb, all_data):
    ws = wb.create_sheet("Coverage Gaps")
    ws.freeze_panes = "A3"
    _title(ws, "Coverage Gaps — Groups with No Recent Successful Backup", "I")

    cols = ["Cluster", "Group Name", "Policy",
            "Active", "Paused",
            "Last Successful Run", "Hours Since Success",
            "Last Run Status", "Gap Severity"]
    _hdr(ws, 2, cols)

    for cd in all_data:
        pm = cd.get("policy_map", {})
        for g in cd["groups"]:
            lr  = g.get("lastRun") or {}
            lb  = lr.get("localBackupInfo") or {}
            st  = lb.get("status", "")
            sla_viol = lb.get("isSlaViolated") or lr.get("isSlaViolated") or False
            su  = lb.get("startTimeUsecs") or 0
            hrs = round((_now_usecs() - su) / 3_600_000_000, 1) if su else 9999

            is_gap = (g.get("isPaused")
                      or st in ("kFailure", "Failed")
                      or hrs > RPO_WARN_HOURS
                      or su == 0)
            if not is_gap:
                continue

            if g.get("isPaused"):
                sev = "PAUSED"
            elif st in ("kFailure", "Failed"):
                sev = "FAILED"
            elif hrs > RPO_CRIT_HOURS:
                sev = "CRITICAL RPO GAP"
            elif hrs > RPO_WARN_HOURS:
                sev = "RPO WARNING"
            else:
                sev = "UNKNOWN"

            policy_name = (g.get("policyName")
                           or pm.get(g.get("policyId"), "")
                           or g.get("policyId", ""))

            row = [cd["name"], g.get("name", ""),
                   policy_name,
                   "Yes" if g.get("isActive", True) else "No",
                   "Yes" if g.get("isPaused") else "No",
                   usecs_to_datetime(su) if su else "Never",
                   hrs if hrs < 9999 else "",
                   st, sev]
            rn = ws.max_row + 1
            for c, val in enumerate(row, 1):
                ws.cell(row=rn, column=c, value=val).font = _font()

            fc = ws.cell(row=rn, column=9)
            if sev in ("FAILED", "CRITICAL RPO GAP"):
                fc.fill = _fill(RED);    fc.font = _font(bold=True, color=WHITE)
            elif sev in ("PAUSED", "RPO WARNING"):
                fc.fill = _fill(ORANGE); fc.font = _font(bold=True)
            else:
                fc.fill = _fill(YELLOW); fc.font = _font(bold=True)

    auto_fit_columns(ws)


def _sheet_trends(wb, all_data):
    ws = wb.create_sheet("Trends (30d)")
    ws.freeze_panes = "A3"
    _title(ws, "30-Day Trends — Daily Backup Success Rate & Volume", "J")

    has_data = any(cd.get("v1_runs") or cd["group_runs"] for cd in all_data)
    if not has_data:
        ws["A3"] = "No run history available."
        ws["A3"].font = _font(bold=False, color="595959")
        return

    cols = ["Cluster", "Date", "Total Runs", "Successful", "Failed",
            "Warning", "Canceled", "Success % (incl. Warnings)", "SLA Violations", "Logical (GB)"]
    _hdr(ws, 2, cols)

    for cd in all_data:
        daily = {}

        # ── Primary: v1 protectionRuns — mirrors the Helios reporting page source ──
        # Single call per cluster; backupRun.slaViolated and
        # backupRun.stats.totalLogicalBackupSizeBytes are accurate.
        for run in (cd.get("v1_runs") or []):
            br    = run.get("backupRun") or {}
            stats = br.get("stats") or {}
            su    = stats.get("startTimeUsecs") or 0
            if not su:
                continue
            day = datetime.datetime.fromtimestamp(su / 1_000_000,
                                                   tz=datetime.timezone.utc) \
                          .strftime("%Y-%m-%d")
            if day not in daily:
                daily[day] = dict(total=0, succ=0, fail=0,
                                  warn=0, cancel=0, viols=0, logical=0)
            d  = daily[day]
            st = br.get("status", "")
            d["total"] += 1
            if st == "kSuccess":                d["succ"]   += 1
            elif st == "kFailure":              d["fail"]   += 1
            elif st == "kWarning":              d["warn"]   += 1
            elif "Cancel" in st:                d["cancel"] += 1
            if br.get("slaViolated"):           d["viols"]  += 1
            d["logical"] += stats.get("totalLogicalBackupSizeBytes") or 0

        # ── Fallback: v2 per-group runs (used when v1 unavailable) ──────────────
        if not daily:
            for runs in cd["group_runs"].values():
                for run in runs:
                    lb = run.get("localBackupInfo") or {}
                    su = lb.get("startTimeUsecs") or 0
                    if not su:
                        continue
                    day = datetime.datetime.fromtimestamp(su / 1_000_000,
                                                           tz=datetime.timezone.utc) \
                                  .strftime("%Y-%m-%d")
                    if day not in daily:
                        daily[day] = dict(total=0, succ=0, fail=0,
                                          warn=0, cancel=0, viols=0, logical=0)
                    d  = daily[day]
                    st = lb.get("status", "")
                    d["total"] += 1
                    if st in ("kSuccess", "Succeeded"): d["succ"]   += 1
                    elif st in ("kFailure", "Failed"):  d["fail"]   += 1
                    elif st == "kWarning":              d["warn"]   += 1
                    elif "Cancel" in st:                d["cancel"] += 1
                    if lb.get("isSlaViolated"):         d["viols"]  += 1
                    d["logical"] += (lb.get("localSnapshotStats") or {}).get("logicalSizeBytes") or 0

        if not daily:
            continue

        trend_start = ws.max_row + 1
        for day in sorted(daily):
            d   = daily[day]
            pct = round((d["succ"] + d["warn"]) / d["total"] * 100, 1) if d["total"] else 0
            rn  = ws.max_row + 1
            row = [cd["name"], day, d["total"], d["succ"], d["fail"],
                   d["warn"], d["cancel"], pct, d["viols"],
                   round(bytes_to_gb(d["logical"]), 1)]
            for c, val in enumerate(row, 1):
                ws.cell(row=rn, column=c, value=val).font = _font()
            sc = ws.cell(row=rn, column=8)
            if pct < SUCCESS_CRIT_PCT:
                sc.fill = _fill(RED);    sc.font = _font(bold=True, color=WHITE)
            elif pct < SUCCESS_WARN_PCT:
                sc.fill = _fill(ORANGE); sc.font = _font(bold=True)

        trend_end = ws.max_row
        n_days    = trend_end - trend_start + 1
        if n_days > 1:
            chart = LineChart()
            chart.title  = f"{cd['name']} — Daily Success Rate (incl. Warnings)"
            chart.style  = 10
            chart.height = 14
            chart.width  = 28

            # ── Y axis: Success %, fixed 0–100, tick marks visible ───────────
            chart.y_axis.title         = "Success %"
            chart.y_axis.numFmt        = "0.0"
            chart.y_axis.majorTickMark = "out"
            chart.y_axis.tickLblPos    = "nextTo"
            chart.y_axis.delete        = False
            chart.y_axis.scaling.min   = 0
            chart.y_axis.scaling.max   = 100

            # ── X axis: date strings, tick marks visible ──────────────────────
            chart.x_axis.title         = "Date"
            chart.x_axis.majorTickMark = "out"
            chart.x_axis.tickLblPos    = "low"
            chart.x_axis.delete        = False
            # Reduce label density so dates don't overlap on long ranges
            if n_days > 60:
                chart.x_axis.tickLblSkip = 7
            elif n_days > 14:
                chart.x_axis.tickLblSkip = 2

            data_ref = Reference(ws, min_col=8, min_row=trend_start,
                                 max_row=trend_end)
            chart.add_data(data_ref)

            # ── StrRef forces <c:strRef> so Excel renders date strings ────────
            try:
                from openpyxl.chart.data_source import DataSource, StrRef
                from openpyxl.utils import get_column_letter as _gcl
                cat_formula = (f"'{ws.title}'!$B${trend_start}:$B${trend_end}")
                chart.series[0].cat = DataSource(strRef=StrRef(f=cat_formula))
            except ImportError:
                dates_ref = Reference(ws, min_col=2, min_row=trend_start,
                                      max_row=trend_end)
                chart.set_categories(dates_ref)

            # ── Series styling: Cohesity green line + circle markers ──────────
            s = chart.series[0]
            s.title = SeriesLabel(v="Success %")
            s.graphicalProperties.line.solidFill = COH_GREEN
            s.graphicalProperties.line.width     = 18000   # 2 pt in EMUs
            s.marker.symbol  = "circle"
            s.marker.size    = 5
            s.marker.graphicalProperties.solidFill   = COH_GREEN
            s.marker.graphicalProperties.line.solidFill = COH_GREEN

            ws.add_chart(chart, f"B{trend_end + 2}")

    auto_fit_columns(ws)


def _sheet_recommendations(wb, all_data):
    ws = wb.create_sheet("Recommendations")
    ws.freeze_panes = "A3"
    _title(ws, "Prioritised Recommendations", "G")

    cols = ["#", "Priority", "Cluster", "Category",
            "Finding", "Recommended Action", "Business Impact"]
    _hdr(ws, 2, cols)

    all_recs = sorted(
        [(cd["name"], rec) for cd in all_data for rec in cd["recommendations"]],
        key=lambda x: _PRIORITY_ORDER.get(x[1]["priority"], 99)
    )
    for i, (cname, rec) in enumerate(all_recs, 1):
        row = [i, rec["priority"], cname, rec["category"],
               rec["finding"], rec["action"], rec["impact"]]
        rn = ws.max_row + 1
        for c, val in enumerate(row, 1):
            ws.cell(row=rn, column=c, value=val).font = _font()

        pc = ws.cell(row=rn, column=2)
        if rec["priority"] == "CRITICAL":
            pc.fill = _fill(RED);        pc.font = _font(bold=True, color=WHITE)
        elif rec["priority"] == "HIGH":
            pc.fill = _fill(ORANGE);     pc.font = _font(bold=True)
        elif rec["priority"] == "MEDIUM":
            pc.fill = _fill(YELLOW);     pc.font = _font(bold=True)
        else:
            pc.fill = _fill(LIGHT_GRAY); pc.font = _font(bold=True)

    auto_fit_columns(ws)
    for col_letter in ("E", "F", "G"):
        ws.column_dimensions[col_letter].width = 55


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL ORCHESTRATOR
# ─────────────────────────────────────────────────────────────────────────────

def write_excel(all_data, args):
    if not EXCEL_OK:
        print("  WARN: openpyxl not installed — skipping Excel output.")
        print("        Install: pip install openpyxl")
        return None

    wb = Workbook()
    wb.remove(wb.active)

    print("  Writing Excel sheets...")
    _sheet_summary(wb, all_data)
    _sheet_infrastructure(wb, all_data)
    _sheet_protection(wb, all_data)
    _sheet_storage(wb, all_data)
    _sheet_policies(wb, all_data)
    _sheet_policy_groups(wb, all_data)
    _sheet_alerts(wb, all_data)
    _sheet_security(wb, all_data)
    _sheet_replication(wb, all_data)
    _sheet_views(wb, all_data)
    _sheet_coverage(wb, all_data)
    _sheet_trends(wb, all_data)
    _sheet_recommendations(wb, all_data)

    out = f"{args.output}.xlsx"
    wb.save(out)
    print(f"  Excel  → {out}")
    return out


# ─────────────────────────────────────────────────────────────────────────────
# WORD DOCUMENT
# ─────────────────────────────────────────────────────────────────────────────

def write_word(all_data, args):
    if not DOCX_OK:
        print("  WARN: python-docx not installed — skipping Word output.")
        print("        Install: pip install python-docx")
        return None

    doc      = Document()
    customer = args.customer or "Customer"
    today    = datetime.date.today().strftime("%B %d, %Y")

    # ── base font ──────────────────────────────────────────────────────────
    for style_name in ("Normal", "Heading 1", "Heading 2"):
        try:
            doc.styles[style_name].font.name = "Calibri"
        except Exception:
            pass

    def _h1(text):
        p = doc.add_heading(text, level=1)
        for run in p.runs:
            run.font.color.rgb = RGBColor(0x00, 0xB3, 0x88)
        return p

    def _h2(text):
        p = doc.add_heading(text, level=2)
        for run in p.runs:
            run.font.color.rgb = RGBColor(0x00, 0xB3, 0x88)
        return p

    def _cell_shade(cell, hex_color):
        """Apply background colour to a docx table cell."""
        tc   = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd  = OxmlElement("w:shd")
        shd.set(qn("w:val"),   "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"),  hex_color)
        tcPr.append(shd)

    def _tbl_header(table, cols, bg=COH_GREEN):
        row = table.rows[0].cells
        for i, col in enumerate(cols):
            row[i].text = col
            _cell_shade(row[i], bg)
            for para in row[i].paragraphs:
                for run in para.runs:
                    run.font.bold  = True
                    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                    run.font.size  = Pt(9)

    # ── Cover page ─────────────────────────────────────────────────────────
    for _ in range(6):
        doc.add_paragraph("")

    cp = doc.add_paragraph()
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = cp.add_run("Cohesity Health Check Report")
    run.font.size  = Pt(28)
    run.font.bold  = True
    run.font.color.rgb = RGBColor(0x00, 0xB3, 0x88)

    doc.add_paragraph("")
    cp2 = doc.add_paragraph()
    cp2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = cp2.add_run(f"Prepared for: {customer}\n{today}")
    r2.font.size = Pt(16)

    doc.add_paragraph("")
    cp3 = doc.add_paragraph()
    cp3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r3 = cp3.add_run("CONFIDENTIAL — For authorised recipients only")
    r3.font.size   = Pt(10)
    r3.font.italic = True
    r3.font.color.rgb = RGBColor(0x59, 0x59, 0x59)

    doc.add_page_break()

    # ── 1. Executive Summary ───────────────────────────────────────────────
    _h1("1. Executive Summary")

    n_clusters = len(all_data)
    avg_score  = round(sum(cd["scores"]["overall"] for cd in all_data) / n_clusters, 1) \
                 if n_clusters else 0
    all_recs   = [(cd["name"], r) for cd in all_data for r in cd["recommendations"]]
    n_crit     = sum(1 for _, r in all_recs if r["priority"] == "CRITICAL")
    n_high     = sum(1 for _, r in all_recs if r["priority"] == "HIGH")

    summary = (
        f"This report covers a Cohesity environment comprising {n_clusters} "
        f"cluster(s) managed through Helios. The overall environment health score "
        f"is {avg_score}/100 ({grade(avg_score)}). "
    )
    if n_crit:
        summary += (f"There are {n_crit} critical finding(s) requiring immediate "
                    f"attention. ")
    if n_high:
        summary += f"{n_high} high-priority item(s) require action within 30 days. "
    if not n_crit and not n_high:
        summary += "No critical or high-priority gaps were identified. "
    summary += ("Detailed findings and prioritised recommendations are included "
                "in the sections below, with full data in the accompanying Excel workbook.")
    doc.add_paragraph(summary)

    t1 = doc.add_table(rows=1, cols=5)
    t1.style = "Table Grid"
    _tbl_header(t1, ["Cluster", "Health Score", "Grade", "Open Criticals", "Top Finding"])
    for cd in all_data:
        row = t1.add_row().cells
        crits = sum(1 for a in cd["alerts"] if a.get("severity") == "kCritical")
        top   = cd["recommendations"][0]["finding"] if cd["recommendations"] else "None"
        row[0].text = cd["name"]
        row[1].text = str(cd["scores"]["overall"])
        row[2].text = cd["scores"]["grade"]
        row[3].text = str(crits)
        row[4].text = top
        for para in row[1].paragraphs:
            for run in para.runs:
                run.font.bold = True
    doc.add_page_break()

    # ── 2. Environment Overview ────────────────────────────────────────────
    _h1("2. Environment Overview")
    doc.add_paragraph(
        f"The environment spans {n_clusters} cluster(s). "
        "The table below summarises key infrastructure attributes for each cluster."
    )
    t2 = doc.add_table(rows=1, cols=7)
    t2.style = "Table Grid"
    _tbl_header(t2, ["Cluster", "Version", "Nodes", "Healthy", "Encryption", "FIPS", "Status"])
    for cd in all_data:
        info  = cd["info"]
        nodes = cd["nodes"]
        healthy = sum(1 for n in nodes
                      if ((n.get("nodeStatus") or {}).get("overallStatus") == "kNormal"
                          or n.get("removalState") in (None, "kDontRemove")))
        enc  = bool(info.get("encryptionEnabled") or
                    (info.get("encryptionConfig") or {}).get("encryptionEnabled"))
        fips = bool(info.get("fipsCompliant") or info.get("fipsEnabled"))
        status = "OK" if healthy == len(nodes) else f"WARN: {len(nodes)-healthy} node(s)"
        row = t2.add_row().cells
        for i, val in enumerate([cd["name"], info.get("clusterSoftwareVersion", ""),
                                  len(nodes), healthy,
                                  "Yes" if enc else "No",
                                  "Yes" if fips else "No", status]):
            row[i].text = str(val)
    doc.add_page_break()

    # ── 3. Protection & Recovery ───────────────────────────────────────────
    _h1("3. Protection & Recovery Health")
    total_g   = sum(len(cd["groups"]) for cd in all_data)
    paused_g  = sum(sum(1 for g in cd["groups"] if g.get("isPaused"))
                    for cd in all_data)
    failed_g  = sum(sum(1 for g in cd["groups"]
                        if (g.get("lastRun") or {}).get("localBackupInfo", {})
                           .get("status", "") in ("kFailure", "Failed"))
                    for cd in all_data)
    doc.add_paragraph(
        f"The environment contains {total_g} active protection group(s) across all clusters. "
        f"{paused_g} group(s) are currently paused and {failed_g} have a failed last run. "
        "The table below summarises protection health per cluster. "
        "Refer to the Protection Health and Coverage Gaps tabs in the Excel workbook "
        "for per-group detail."
    )
    t3 = doc.add_table(rows=1, cols=6)
    t3.style = "Table Grid"
    _tbl_header(t3, ["Cluster", "Groups", "Paused", "Failed Last Run",
                      "30d Success %", "SLA Score"])
    for cd in all_data:
        succ_pct, sla_pct = _success_stats(cd)
        paused = sum(1 for g in cd["groups"] if g.get("isPaused"))
        failed = sum(1 for g in cd["groups"]
                     if (g.get("lastRun") or {}).get("localBackupInfo", {})
                        .get("status", "") in ("kFailure", "Failed"))
        row = t3.add_row().cells
        for i, val in enumerate([cd["name"], len(cd["groups"]),
                                  paused, failed, f"{succ_pct}%",
                                  f"{cd['scores']['sla']}/100"]):
            row[i].text = str(val)
    doc.add_page_break()

    # ── 4. Storage & Capacity ──────────────────────────────────────────────
    _h1("4. Storage & Capacity")
    doc.add_paragraph(
        "Cohesity applies inline deduplication and compression to maximise storage "
        "efficiency. The data reduction ratio shows how much raw data has been saved "
        "compared to the logical data ingested."
    )
    t4 = doc.add_table(rows=1, cols=7)
    t4.style = "Table Grid"
    _tbl_header(t4, ["Cluster", "Usable (TB)", "Used (TB)", "Free (TB)",
                      "Used %", "Data Reduction", "Status"])
    for cd in all_data:
        info   = cd["info"]
        usable, used, cap_pct = _cap_stats(cd)
        free   = usable - used
        dr     = round((info.get("stats") or {}).get("dataReductionRatio") or 0, 2)
        pct_str = f"{cap_pct}%" if isinstance(cap_pct, float) else "N/A"
        status = ("N/A" if cap_pct == "N/A"
                  else "CRITICAL" if cap_pct >= CAPACITY_CRIT_PCT
                  else "WARNING" if cap_pct >= CAPACITY_WARN_PCT else "OK")
        row = t4.add_row().cells
        for i, val in enumerate([cd["name"],
                                  round(bytes_to_tb(usable), 2) or "N/A",
                                  round(bytes_to_tb(used), 2) or "N/A",
                                  round(bytes_to_tb(free), 2) or "N/A",
                                  pct_str, f"{dr}x", status]):
            row[i].text = str(val)
    doc.add_page_break()

    # ── 5. Security ────────────────────────────────────────────────────────
    _h1("5. Security Posture")
    doc.add_paragraph(
        "The following checklist evaluates each cluster against Cohesity security "
        "best practices. A fully hardened deployment enables encryption, FIPS, "
        "at least one vault for archival, replication to a secondary cluster, "
        "and FortKnox or WORM archival for immutable backups."
    )
    t5 = doc.add_table(rows=1, cols=7)
    t5.style = "Table Grid"
    _tbl_header(t5, ["Cluster", "Encryption", "FIPS", "Vault",
                      "Immutability", "Replication", "Score /100"])
    for cd in all_data:
        info     = cd["info"]
        vaults   = cd["vaults"]
        policies = cd["policies"]
        enc  = bool(info.get("encryptionEnabled") or
                    (info.get("encryptionConfig") or {}).get("encryptionEnabled"))
        fips = bool(info.get("fipsCompliant") or info.get("fipsEnabled"))
        has_vault = bool(vaults)
        fk_v  = [v for v in vaults if v.get("vaultType") in ("kFortKnox", "kS3Compatible")]
        immut = bool(fk_v) or any(
            any((t.get("archivalTargetSettings") or {}).get("targetType") == "kFortKnox"
                for t in ((p.get("remoteTargetPolicy") or {}).get("archivalTargets") or []))
            for p in policies)
        repl = any((p.get("remoteTargetPolicy") or {}).get("replicationTargets")
                   for p in policies)
        row = t5.add_row().cells
        yn  = lambda v: "Yes" if v else "No"
        for i, val in enumerate([cd["name"], yn(enc), yn(fips),
                                  yn(has_vault), yn(immut), yn(repl),
                                  f"{cd['scores']['security']}/100"]):
            row[i].text = str(val)
    doc.add_page_break()

    # ── 6. Recommendations ─────────────────────────────────────────────────
    _h1("6. Prioritised Recommendations")
    doc.add_paragraph(
        "The recommendations below are ranked by priority. Critical items should "
        "be addressed immediately; High-priority items within 30 days; Medium within "
        "90 days. Refer to the Recommendations tab in the Excel workbook for the "
        "full prioritised list."
    )
    sorted_recs = sorted(all_recs, key=lambda x: _PRIORITY_ORDER.get(x[1]["priority"], 99))
    if sorted_recs:
        t6 = doc.add_table(rows=1, cols=5)
        t6.style = "Table Grid"
        _tbl_header(t6, ["Priority", "Cluster", "Category", "Finding", "Action"])
        for cname, rec in sorted_recs:
            row = t6.add_row().cells
            for i, val in enumerate([rec["priority"], cname, rec["category"],
                                      rec["finding"], rec["action"]]):
                row[i].text = str(val)
                if i == 0:
                    for para in row[i].paragraphs:
                        for run in para.runs:
                            run.font.bold = True
    else:
        doc.add_paragraph("No recommendations generated — environment appears healthy.")

    doc.add_page_break()

    # ── Appendix ───────────────────────────────────────────────────────────
    _h1("Appendix — Methodology & Thresholds")
    _h2("Scoring Model")
    doc.add_paragraph(
        "Health scores (0-100) are computed from five weighted dimensions: "
        "Protection success rate (25%), SLA compliance (20%), "
        "Infrastructure health (15%), Storage capacity (15%), "
        "Security posture (15%), Alert health (10%)."
    )
    _h2("Thresholds")
    notes = [
        f"Capacity: Warning ≥{CAPACITY_WARN_PCT}%, Critical ≥{CAPACITY_CRIT_PCT}%",
        f"Backup success rate: Warning <{SUCCESS_WARN_PCT}%, Critical <{SUCCESS_CRIT_PCT}% "
        f"— runs with status 'Warning' (kWarning) are counted as successful in the "
        f"Success % calculation. They complete with data but may have minor issues "
        f"(e.g. some objects skipped) and are shown separately in the Warning column "
        f"for visibility.",
        f"RPO: Warning >{RPO_WARN_HOURS}h, Critical >{RPO_CRIT_HOURS}h since last success",
        (f"Security: 20 pts each for encryption, FIPS, vault, "
         f"replication, FortKnox/immutability"),
        f"Alert penalty: −10 pts per open critical, −2 pts per open warning",
        f"View quota: Warning ≥{VIEW_QUOTA_WARN_PCT}%, Critical ≥{VIEW_QUOTA_CRIT_PCT}%",
        f"Lookback period used for this report: {args.days} days",
        f"Report generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} (local time)",
    ]
    for note in notes:
        p = doc.add_paragraph(note, style="List Bullet")
        for run in p.runs:
            run.font.size = Pt(9)

    out = f"{args.output}.docx"
    doc.save(out)
    print(f"  Word   → {out}")
    return out


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description=f"Cohesity Health Check Report v{__version__}",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("--apikey",    help="Helios API key")
    parser.add_argument("--cluster",   help="Limit to one cluster (partial name match)")
    parser.add_argument("--days",      type=int, default=30,
                        help="Lookback window in days (default 30)")
    parser.add_argument("--customer",  default="",
                        help="Customer name for report cover page")
    parser.add_argument("--output",    default="cohesity_health_check",
                        help="Output base path without extension "
                             "(default: cohesity_health_check)")
    parser.add_argument("--quick",     action="store_true",
                        help="Skip per-group run history; use last-run only (faster)")
    parser.add_argument("--excel-only",action="store_true",
                        help="Generate Excel only, skip Word document")
    parser.add_argument("--word-only", action="store_true",
                        help="Generate Word only, skip Excel")
    parser.add_argument("--debug",     action="store_true",
                        help="Print HTTP status for each API call")
    args = parser.parse_args()

    if not EXCEL_OK and not args.word_only:
        print("ERROR: openpyxl is required.  pip install openpyxl")
        sys.exit(1)
    if not DOCX_OK and args.word_only:
        print("ERROR: python-docx is required.  pip install python-docx")
        sys.exit(1)

    # Build output filename: include customer name (if given) + local timestamp
    if args.output == "cohesity_health_check":
        ts    = datetime.datetime.now().strftime("%Y%m%d_%H%M")
        parts = ["cohesity_health_check", f"v{__version__}"]
        if args.customer:
            safe = "".join(c if c.isalnum() or c in "-_" else "_"
                           for c in args.customer).strip("_")
            parts.append(safe)
        parts.append(ts)
        args.output = "_".join(parts)

    print(f"\nCohesity Health Check Report  v{__version__}")
    print(f"  Customer : {args.customer or '(not specified)'}")
    print(f"  Lookback : {args.days} days")
    print(f"  Mode     : {'Quick (last-run only)' if args.quick else 'Full (run history)'}")
    print()

    api_key  = get_api_key(args.apikey)
    clusters = get_helios_clusters(api_key)
    if not clusters:
        print("ERROR: No clusters returned from Helios.")
        sys.exit(1)

    if args.cluster:
        clusters = [c for c in clusters
                    if args.cluster.lower() in c.get("name", "").lower()]
        if not clusters:
            print(f"ERROR: No cluster matching '{args.cluster}'")
            sys.exit(1)

    print(f"Clusters : {', '.join(c.get('name','?') for c in clusters)}\n")

    session  = requests.Session()
    all_data = []
    for cluster in clusters:
        try:
            cd = collect_cluster(session, api_key, cluster, args)
            s  = cd["scores"]
            print(f"  [{cd['name']}] score={s['overall']}/100 ({s['grade']}) "
                  f"| {len(cd['recommendations'])} recommendation(s)")
            all_data.append(cd)
        except Exception as exc:
            import traceback
            print(f"  ERROR collecting [{cluster.get('name','?')}]: {exc}")
            if args.debug:
                traceback.print_exc()
            continue

    if not all_data:
        print("ERROR: No data collected from any cluster.")
        sys.exit(1)

    print()
    if not args.word_only:
        write_excel(all_data, args)
    if not args.excel_only:
        write_word(all_data, args)

    print(f"\nComplete.  Output: {args.output}.xlsx / {args.output}.docx")


if __name__ == "__main__":
    main()
