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
  5  Policy Audit        — retention, replication, archival
  6  Alerts              — open critical/warning with age
  7  Security            — encryption, FIPS, immutability, vault
  8  Replication/Archive — targets, last transfer, FortKnox
  9  Data Services       — NAS views, quota utilisation
  10 Coverage Gaps       — groups with no recent successful backup
  11 Trends (30d)        — daily success rate + storage growth charts
  12 Recommendations     — prioritised action list

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
  1.0 (2026-04-08) — feat: initial release. 12-sheet Excel + Word document.
                     Health scoring, recommendations engine, trend charts.
"""

__version__ = "1.0"

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
COH_GREEN   = "00B388"
DARK_GREEN  = "007A5E"
LT_GREEN    = "C6EFCE"
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

def _fortknox_data(session, h, start_usecs, end_usecs, debug):
    d = _get(session, h,
             "/irisservices/api/v1/public/reports/dataTransferToVaults",
             params={"startTimeUsecs": start_usecs, "endTimeUsecs": end_usecs},
             debug=debug)
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
    vaults  = _vaults(session, h, dbg)
    fk_data = _fortknox_data(session, h, start_usecs, end_usecs, dbg)

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

    cd = {
        "name":        name,
        "id":          cid,
        "info":        info,
        "nodes":       nodes,
        "alerts":      alerts,
        "groups":      groups,
        "group_runs":  group_runs,
        "policies":    policies,
        "domains":     domains,
        "views":       views,
        "vaults":      vaults,
        "fk_data":     fk_data,
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
               in ("kSuccess", "Succeeded")
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
    # FortKnox / immutability
    fk = [v for v in vaults if v.get("vaultType") in ("kFortKnox", "kS3Compatible")]
    immut = any(
        any((t.get("archivalTargetSettings") or {}).get("targetType") == "kFortKnox"
            for t in ((p.get("remoteTargetPolicy") or {}).get("archivalTargets") or []))
        for p in policies
    )
    if fk or immut:
        score += 20
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
    version = info.get("softwareVersion", "")
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

    has_fk = any(v.get("vaultType") in ("kFortKnox", "kS3Compatible")
                 for v in vaults)
    has_immut = any(
        any((t.get("archivalTargetSettings") or {}).get("targetType") == "kFortKnox"
            for t in ((p.get("remoteTargetPolicy") or {}).get("archivalTargets") or []))
        for p in policies
    )
    if not (has_fk or has_immut):
        recs.append(_r("MEDIUM", "Security",
            "No immutable/FortKnox archival target detected",
            "Configure FortKnox or WORM archival to protect against ransomware",
            "Without immutability, backups can be deleted or encrypted by attackers"))

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
                   .get("status", "") in ("kSuccess", "Succeeded"))
    viols = sum(1 for g in groups
                if (g.get("lastRun") or {}).get("isSlaViolated", False))
    return (
        round(succ / total * 100, 1),
        round((total - viols) / total * 100, 1),
    )


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
        usage  = ((info.get("stats") or {}).get("usagePerfStats") or {})
        usable = usage.get("physicalCapacityBytes") or 0
        used   = usage.get("totalPhysicalUsageBytes") or 0
        cap_pct = round(used / usable * 100, 1) if usable else ""
        crits  = sum(1 for a in alerts if a.get("severity") == "kCritical")
        warns  = sum(1 for a in alerts if a.get("severity") == "kWarning")
        succ_pct, sla_pct = _success_stats(cd)
        top = cd["recommendations"][0]["finding"] if cd["recommendations"] else "None identified"

        row = [cd["name"],
               info.get("softwareVersion", ""), len(cd["nodes"]),
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
        disks = sum(len(n.get("diskVec") or n.get("disks") or []) for n in nodes)
        enc  = bool(info.get("encryptionEnabled") or
                    (info.get("encryptionConfig") or {}).get("encryptionEnabled"))
        fips = bool(info.get("fipsCompliant") or info.get("fipsEnabled"))
        dns  = ", ".join(info.get("dnsServerIps") or [])
        ntp  = ", ".join(
            (info.get("ntpSettings") or {}).get("ntpServers", []) or
            info.get("ntpServers", []) or []
        )
        domain = (info.get("domainNames") or [""])[0]

        row = [cd["name"], info.get("softwareVersion", ""),
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
        for g in cd["groups"]:
            lr  = g.get("lastRun") or {}
            lb  = lr.get("localBackupInfo") or {}
            st  = (lr.get("stats") or lb.get("stats") or {})
            objs = lb.get("stats") or {}

            start_u = lb.get("startTimeUsecs") or 0
            end_u   = lb.get("endTimeUsecs") or 0
            dur = ""
            if start_u and end_u:
                s = int((end_u - start_u) / 1_000_000)
                h, m = divmod(s // 60, 60)
                dur = f"{h}h {m:02d}m"
            rpo_hrs = round((_now_usecs() - start_u) / 3_600_000_000, 1) \
                      if start_u else ""

            repl_st = arch_st = ""
            for ri in ((lr.get("replicationInfo") or {})
                       .get("replicationTargetResults") or []):
                repl_st = ri.get("status", ""); break
            for ai in ((lr.get("archivalInfo") or {})
                       .get("archivalTargetResults") or []):
                arch_st = ai.get("status", ""); break

            status = lb.get("status", "")
            flag   = "OK"
            if g.get("isPaused"):
                flag = "PAUSED"
            elif status in ("kFailure", "Failed"):
                flag = "FAILED"
            elif lr.get("isSlaViolated"):
                flag = "SLA VIOLATED"
            elif isinstance(rpo_hrs, float) and rpo_hrs > RPO_CRIT_HOURS:
                flag = f"RPO GAP {rpo_hrs:.0f}h"
            elif isinstance(rpo_hrs, float) and rpo_hrs > RPO_WARN_HOURS:
                flag = f"RPO WARN {rpo_hrs:.0f}h"

            row = [cd["name"], g.get("name", ""),
                   g.get("policyName") or "",
                   "Yes" if g.get("isActive", True) else "No",
                   "Yes" if g.get("isPaused", False) else "No",
                   lr.get("runType", ""), status,
                   "Yes" if lr.get("isSlaViolated") else "No",
                   usecs_to_datetime(start_u), usecs_to_datetime(end_u), dur,
                   objs.get("totalObjectsProtected", ""),
                   objs.get("totalObjectsFailed", ""),
                   round(bytes_to_gb(objs.get("logicalSizeBytes") or 0), 2),
                   round(bytes_to_gb(objs.get("bytesWritten") or 0), 2),
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

        usable      = usage.get("physicalCapacityBytes") or 0
        used        = usage.get("totalPhysicalUsageBytes") or 0
        logical     = usage.get("dataInBytes") or 0
        physical    = usage.get("dataInBytesAfterReduction") or 0
        dedup_after = data.get("dataInBytesAfterDedup") or 0
        dr_ratio    = round((stats.get("dataReductionRatio") or 0.0), 2)
        dedup_ratio = round(logical / dedup_after, 2) if dedup_after else 0
        comp_ratio  = round(dedup_after / physical, 2) if physical else 0
        free        = usable - used
        cap_pct     = round(used / usable * 100, 1) if usable else ""
        savings     = round(bytes_to_tb(logical - physical), 2) \
                      if logical > physical else 0

        flag = ""
        if isinstance(cap_pct, float):
            flag = ("CRITICAL" if cap_pct >= CAPACITY_CRIT_PCT
                    else "WARN" if cap_pct >= CAPACITY_WARN_PCT else "OK")

        row = [cd["name"], "(Cluster Total)",
               round(bytes_to_tb(usable), 2), round(bytes_to_tb(used), 2),
               round(bytes_to_tb(free), 2), cap_pct,
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
    if not s: return ""
    for key in ("minutes", "hours", "days", "weeks", "months", "years"):
        if isinstance(s, dict) and key in s:
            return f"Every {s[key]} {key}"
    return ""

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
        for p in cd["policies"]:
            bp  = p.get("backupPolicy") or {}
            rtp = p.get("remoteTargetPolicy") or {}
            rep_tgts  = rtp.get("replicationTargets") or []
            arch_tgts = rtp.get("archivalTargets") or []

            rep_names  = ", ".join(
                (t.get("replicationTargetSettings") or {}).get("clusterName", "")
                or t.get("targetName", "") for t in rep_tgts)
            arch_names = ", ".join(
                (t.get("archivalTargetSettings") or {}).get("targetName", "")
                or t.get("targetName", "") for t in arch_tgts)

            full_sched = (bp.get("full") or bp.get("fullBackups") or {})
            reg_sched  = (bp.get("regular") or bp.get("incrementalBackups") or {})
            full_freq  = (full_sched.get("schedule") or full_sched.get("frequency") or {})
            reg_freq   = (reg_sched.get("schedule") or reg_sched.get("frequency") or {})

            local_ret = (bp.get("regular") or {}).get("retention") or bp.get("retention") or {}
            log_ret   = (bp.get("log") or {}).get("retention") or {}
            rep_ret   = (rep_tgts[0].get("retention") or {}) if rep_tgts else {}
            arch_ret  = (arch_tgts[0].get("retention") or {}) if arch_tgts else {}

            gaps = []
            if not rep_tgts:  gaps.append("No replication")
            if not arch_tgts: gaps.append("No archival")

            row = [cd["name"], p.get("name", ""),
                   p.get("type") or p.get("policyCategory") or "",
                   _sched_str(full_freq), _sched_str(reg_freq),
                   _ret_str(local_ret), _ret_str(log_ret),
                   "Yes" if rep_tgts else "No", rep_names, _ret_str(rep_ret),
                   "Yes" if arch_tgts else "No", arch_names, _ret_str(arch_ret),
                   p.get("numProtectionGroups", ""),
                   "; ".join(gaps) if gaps else "OK"]
            rn = ws.max_row + 1
            for c, val in enumerate(row, 1):
                ws.cell(row=rn, column=c, value=val).font = _font()

            gc = ws.cell(row=rn, column=15)
            if gaps:
                gc.fill = _fill(YELLOW)
                gc.font = _font(bold=True)

    auto_fit_columns(ws)


def _sheet_alerts(wb, all_data):
    ws = wb.create_sheet("Alerts")
    ws.freeze_panes = "A3"
    _title(ws, "Open Alerts — Severity, Age & Description", "L")

    cols = ["Cluster", "Severity", "Category", "Alert Code",
            "Alert Name", "Description",
            "First Occurred", "Age (Days)", "Last Updated",
            "Acknowledged", "Node ID", "Entity"]
    _hdr(ws, 2, cols)

    sev_order = {"kCritical": 0, "kWarning": 1, "kInfo": 2}

    for cd in all_data:
        sorted_a = sorted(cd["alerts"],
                          key=lambda a: (sev_order.get(a.get("severity"), 9),
                                         a.get("firstOccurrenceTime") or 0))
        for a in sorted_a:
            first_u = a.get("firstOccurrenceTime") or 0
            last_u  = a.get("latestTimestampUsecs") or 0
            age = round((_now_usecs() - first_u) / (86400 * 1_000_000), 1) \
                  if first_u else ""
            doc   = a.get("alertDocument") or {}
            props = {p["key"]: p["value"]
                     for p in (a.get("propertyList") or []) if "key" in p}
            desc  = doc.get("alertDescription") or ""

            row = [cd["name"], a.get("severity", ""),
                   a.get("alertCategory", ""), a.get("alertCode", ""),
                   doc.get("alertName", ""),
                   desc[:120] + ("…" if len(desc) > 120 else ""),
                   usecs_to_datetime(first_u), age, usecs_to_datetime(last_u),
                   "Yes" if a.get("acknowledgeTimestampUsecs") else "No",
                   props.get("nodeId", ""),
                   props.get("entityName") or props.get("jobName", "")]
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

    def _min_ret_days(policies, key="replicationTargets"):
        days_list = []
        for p in policies:
            for t in ((p.get("remoteTargetPolicy") or {}).get(key) or []):
                r = t.get("retention") or {}
                d = r.get("duration") or 0
                u = r.get("unit", "Days")
                m = {"Days": 1, "Weeks": 7, "Months": 30, "Years": 365}
                days_list.append(d * m.get(u, 1))
        return min(days_list) if days_list else ""

    def _min_local_days(policies):
        days_list = []
        for p in policies:
            bp = p.get("backupPolicy") or {}
            r  = (bp.get("regular") or {}).get("retention") or bp.get("retention") or {}
            d  = r.get("duration") or 0
            u  = r.get("unit", "Days")
            m  = {"Days": 1, "Weeks": 7, "Months": 30, "Years": 365}
            if d:
                days_list.append(d * m.get(u, 1))
        return min(days_list) if days_list else ""

    for cd in all_data:
        info     = cd["info"]
        vaults   = cd["vaults"]
        policies = cd["policies"]
        scores   = cd["scores"]

        enc  = bool(info.get("encryptionEnabled") or
                    (info.get("encryptionConfig") or {}).get("encryptionEnabled"))
        fips = bool(info.get("fipsCompliant") or info.get("fipsEnabled"))
        has_vault = bool(vaults)
        fk_v = [v for v in vaults if v.get("vaultType") in ("kFortKnox", "kS3Compatible")]
        has_immut = bool(fk_v) or any(
            any((t.get("archivalTargetSettings") or {}).get("targetType") == "kFortKnox"
                for t in ((p.get("remoteTargetPolicy") or {}).get("archivalTargets") or []))
            for p in policies)
        has_repl = any((p.get("remoteTargetPolicy") or {}).get("replicationTargets")
                       for p in policies)
        has_arch = any((p.get("remoteTargetPolicy") or {}).get("archivalTargets")
                       for p in policies)

        gaps = []
        if not enc:      gaps.append("No encryption")
        if not fips:     gaps.append("FIPS off")
        if not has_vault: gaps.append("No vault")
        if not has_immut: gaps.append("No immutability")
        if not has_repl: gaps.append("No replication")
        if not has_arch: gaps.append("No archival")

        yn = lambda v: "Yes" if v else "No"
        row = [cd["name"], yn(enc), yn(fips), yn(has_vault), yn(has_immut),
               yn(has_repl), yn(has_arch),
               _min_local_days(policies), _min_ret_days(policies, "archivalTargets"),
               scores["security"],
               "; ".join(gaps) if gaps else "OK"]
        rn = ws.max_row + 1
        for c, val in enumerate(row, 1):
            ws.cell(row=rn, column=c, value=val).font = _font()

        for col, flag in [(2, enc), (3, fips), (4, has_vault),
                          (5, has_immut), (6, has_repl), (7, has_arch)]:
            cell = ws.cell(row=rn, column=col)
            cell.fill = _fill(LT_GREEN) if flag else _fill(RED)
            cell.font = _font(bold=True, color=WHITE if not flag else "000000")

        sc = ws.cell(row=rn, column=10)
        sc.fill = _rag_fill(scores["security"])
        sc.font = _font(bold=True)

    auto_fit_columns(ws)
    ws.column_dimensions["K"].width = 55


def _sheet_replication(wb, all_data):
    ws = wb.create_sheet("Replication & Archive")
    ws.freeze_panes = "A3"
    _title(ws, "Replication & Archival Targets", "J")

    cols = ["Cluster", "Target Type", "Target Name", "Vault Type",
            "Policies Using",
            "FK Storage Consumed (TB)", "FK Logical (TB)", "FK Physical (TB)",
            "Status", "Notes"]
    _hdr(ws, 2, cols)

    for cd in all_data:
        # Tally policy usage of each target
        rep_use  = {}
        arch_use = {}
        for p in cd["policies"]:
            for t in ((p.get("remoteTargetPolicy") or {}).get("replicationTargets") or []):
                n = (t.get("replicationTargetSettings") or {}).get("clusterName", "") \
                    or t.get("targetName", "")
                rep_use[n] = rep_use.get(n, 0) + 1
            for t in ((p.get("remoteTargetPolicy") or {}).get("archivalTargets") or []):
                n = (t.get("archivalTargetSettings") or {}).get("targetName", "") \
                    or t.get("targetName", "")
                arch_use[n] = arch_use.get(n, 0) + 1

        for name, cnt in rep_use.items():
            rn = ws.max_row + 1
            for c, val in enumerate(
                [cd["name"], "Replication", name, "", cnt, "", "", "", "Configured", ""], 1
            ):
                ws.cell(row=rn, column=c, value=val).font = _font()

        # FortKnox / archival vault stats
        fk_by_vault = {}
        for fk in (cd["fk_data"] or []):
            vn = (fk.get("dataTransferPerProtectionJob") or [{}])[0].get("vaultName", "") \
                 or fk.get("vaultName", "")
            consumed = sum((j.get("storageConsumedBytes") or 0)
                           for j in (fk.get("dataTransferPerProtectionJob") or []))
            logical  = sum((j.get("logicalTransferredBytes") or 0)
                           for j in (fk.get("dataTransferPerProtectionJob") or []))
            physical = sum((j.get("physicalTransferredBytes") or 0)
                           for j in (fk.get("dataTransferPerProtectionJob") or []))
            fk_by_vault[vn] = {"consumed": consumed, "logical": logical, "physical": physical}

        for v in cd["vaults"]:
            vn  = v.get("name", "")
            vt  = v.get("vaultType", "")
            fkd = fk_by_vault.get(vn, {})
            rn  = ws.max_row + 1
            row = [cd["name"], "Archival/Vault", vn, vt,
                   arch_use.get(vn, 0),
                   round(bytes_to_tb(fkd.get("consumed") or 0), 3),
                   round(bytes_to_tb(fkd.get("logical") or 0), 3),
                   round(bytes_to_tb(fkd.get("physical") or 0), 3),
                   "Configured",
                   "FortKnox" if vt == "kFortKnox" else ""]
            for c, val in enumerate(row, 1):
                ws.cell(row=rn, column=c, value=val).font = _font()

    auto_fit_columns(ws)


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
        for g in cd["groups"]:
            lr  = g.get("lastRun") or {}
            lb  = lr.get("localBackupInfo") or {}
            st  = lb.get("status", "")
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

            row = [cd["name"], g.get("name", ""),
                   g.get("policyName") or "",
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

    has_runs = any(cd["group_runs"] for cd in all_data)
    if not has_runs:
        ws["A3"] = ("No run history available. Re-run without --quick to populate "
                    "trend data and charts.")
        ws["A3"].font = _font(bold=False, color="595959")
        return

    cols = ["Cluster", "Date", "Total Runs", "Successful", "Failed",
            "Warning", "Canceled", "Success %", "SLA Violations", "Logical (GB)"]
    _hdr(ws, 2, cols)

    for cd in all_data:
        if not cd["group_runs"]:
            continue
        daily = {}
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
                d   = daily[day]
                st  = lb.get("status", "")
                d["total"] += 1
                if st in ("kSuccess", "Succeeded"): d["succ"]   += 1
                elif st in ("kFailure", "Failed"):  d["fail"]   += 1
                elif st == "kWarning":              d["warn"]   += 1
                elif "anceл" in st or "Cancel" in st: d["cancel"] += 1
                if run.get("isSlaViolated"):        d["viols"]  += 1
                d["logical"] += (lb.get("stats") or {}).get("logicalSizeBytes") or 0

        trend_start = ws.max_row + 1
        for day in sorted(daily):
            d   = daily[day]
            pct = round(d["succ"] / d["total"] * 100, 1) if d["total"] else 0
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
        if trend_end - trend_start > 0:
            chart = LineChart()
            chart.title   = f"{cd['name']} — Daily Success Rate"
            chart.style   = 10
            chart.y_axis.title = "Success %"
            chart.x_axis.title = "Date"
            chart.height  = 12
            chart.width   = 25
            data_ref = Reference(ws, min_col=8, min_row=trend_start,
                                 max_row=trend_end)
            chart.add_data(data_ref)
            dates_ref = Reference(ws, min_col=2, min_row=trend_start,
                                  max_row=trend_end)
            chart.set_categories(dates_ref)
            chart.series[0].title = None
            chart.series[0].graphicalProperties.line.solidFill = COH_GREEN
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
        for i, val in enumerate([cd["name"], info.get("softwareVersion", ""),
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
        info  = cd["info"]
        usage = ((info.get("stats") or {}).get("usagePerfStats") or {})
        usable = usage.get("physicalCapacityBytes") or 0
        used   = usage.get("totalPhysicalUsageBytes") or 0
        free   = usable - used
        pct    = round(used / usable * 100, 1) if usable else 0
        dr     = round((info.get("stats") or {}).get("dataReductionRatio") or 0, 2)
        status = ("CRITICAL" if pct >= CAPACITY_CRIT_PCT
                  else "WARNING" if pct >= CAPACITY_WARN_PCT else "OK")
        row = t4.add_row().cells
        for i, val in enumerate([cd["name"],
                                  round(bytes_to_tb(usable), 2),
                                  round(bytes_to_tb(used), 2),
                                  round(bytes_to_tb(free), 2),
                                  f"{pct}%", f"{dr}x", status]):
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
        f"Backup success rate: Warning <{SUCCESS_WARN_PCT}%, Critical <{SUCCESS_CRIT_PCT}%",
        f"RPO: Warning >{RPO_WARN_HOURS}h, Critical >{RPO_CRIT_HOURS}h since last success",
        (f"Security: 20 pts each for encryption, FIPS, vault, "
         f"replication, FortKnox/immutability"),
        f"Alert penalty: −10 pts per open critical, −2 pts per open warning",
        f"View quota: Warning ≥{VIEW_QUOTA_WARN_PCT}%, Critical ≥{VIEW_QUOTA_CRIT_PCT}%",
        f"Lookback period used for this report: {args.days} days",
        f"Report generated: {datetime.datetime.now(datetime.timezone.utc).strftime('%Y-%m-%d %H:%M:%S')} UTC",
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
