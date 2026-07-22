"""
Cohesity MCP Server v2 — cluster-direct
Personal SE demo tool: monitor -> report -> act (with confirm pattern)

Transports:
  stdio (default, for Claude Desktop):    python cohesity_mcp_v2.py
  streamable HTTP (for Copilot/VS Code):  MCP_TRANSPORT=http python cohesity_mcp_v2.py

Env vars:
  COHESITY_CLUSTER      e.g. mycluster.lab.local (no scheme)
  COHESITY_API_KEY      cluster API key (preferred), OR
  COHESITY_USERNAME / COHESITY_PASSWORD / COHESITY_DOMAIN (session-token fallback)
  COHESITY_VERIFY_SSL   "true" | "false" (default false — lab clusters are self-signed)
  MCP_TRANSPORT         "stdio" (default) | "http"
  MCP_HTTP_PORT         default 8720

Requires: pip install "mcp[cli]" httpx
"""

import os
import time
import httpx
from typing import Any, Optional
from mcp.server.fastmcp import FastMCP

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

CLUSTER = os.environ.get("COHESITY_CLUSTER", "")
API_KEY = os.environ.get("COHESITY_API_KEY", "")
USERNAME = os.environ.get("COHESITY_USERNAME", "")
PASSWORD = os.environ.get("COHESITY_PASSWORD", "")
DOMAIN = os.environ.get("COHESITY_DOMAIN", "LOCAL")
VERIFY_SSL = os.environ.get("COHESITY_VERIFY_SSL", "false").lower() == "true"

BASE_V1 = f"https://{CLUSTER}/irisservices/api/v1"
BASE_V2 = f"https://{CLUSTER}/v2"

mcp = FastMCP("cohesity")

_session_token: Optional[str] = None
_session_expiry: float = 0.0


def _headers() -> dict:
    """API key auth if available, else session token (fetched lazily)."""
    if API_KEY:
        return {"apiKey": API_KEY, "Accept": "application/json"}
    global _session_token, _session_expiry
    if not _session_token or time.time() > _session_expiry:
        resp = httpx.post(
            f"{BASE_V1}/public/accessTokens",
            json={"username": USERNAME, "password": PASSWORD, "domain": DOMAIN},
            verify=VERIFY_SSL,
            timeout=30,
        )
        resp.raise_for_status()
        tok = resp.json()
        _session_token = f"{tok['tokenType']} {tok['accessToken']}"
        _session_expiry = time.time() + 23 * 3600  # tokens last 24h; refresh at 23
    return {"Authorization": _session_token, "Accept": "application/json"}


def _get(url: str, params: dict | None = None) -> Any:
    resp = httpx.get(url, headers=_headers(), params=params or {},
                     verify=VERIFY_SSL, timeout=60)
    resp.raise_for_status()
    return resp.json()


def _post(url: str, body: dict | None = None) -> Any:
    resp = httpx.post(url, headers=_headers(), json=body or {},
                      verify=VERIFY_SSL, timeout=120)
    resp.raise_for_status()
    return resp.json() if resp.text else {"status": resp.status_code}


def _put(url: str, body: dict | None = None) -> Any:
    resp = httpx.put(url, headers=_headers(), json=body or {},
                     verify=VERIFY_SSL, timeout=120)
    resp.raise_for_status()
    return resp.json() if resp.text else {"status": resp.status_code}


# ---------------------------------------------------------------------------
# Tier 1 — Monitor
# ---------------------------------------------------------------------------

@mcp.tool()
def get_cluster_health() -> dict:
    """Cluster identity, version, node count, and overall health/status."""
    info = _get(f"{BASE_V1}/public/cluster")
    return {
        "name": info.get("name"),
        "clusterSoftwareVersion": info.get("clusterSoftwareVersion"),
        "nodeCount": info.get("nodeCount"),
        "healthy": info.get("clusterAuditLogConfig") is not None,  # VERIFY: use basicClusterInfo/health endpoint on your version
        "raw": {k: info.get(k) for k in
                ("id", "clusterType", "usedMetadataSpacePct", "availableMetadataSpace")},
    }


@mcp.tool()
def get_alerts(max_alerts: int = 25, severity: str = "") -> list[dict]:
    """Open alerts on the cluster. severity: kCritical | kWarning | kInfo (blank = all)."""
    params: dict = {"maxAlerts": max_alerts, "alertStateList": "kOpen"}
    if severity:
        params["alertSeverityList"] = severity
    alerts = _get(f"{BASE_V1}/public/alerts", params)
    return [
        {
            "id": a.get("id"),
            "severity": a.get("severity"),
            "name": a.get("alertDocument", {}).get("alertName"),
            "description": a.get("alertDocument", {}).get("alertDescription"),
            "firstSeen": a.get("firstTimestampUsecs"),
            "latest": a.get("latestTimestampUsecs"),
        }
        for a in alerts
    ]


@mcp.tool()
def get_storage_stats() -> dict:
    """Capacity, usage, and data-reduction ratios for the cluster."""
    # VERIFY: on 7.x the v2 stats endpoint may differ; check /docs swagger on the cluster
    stats = _get(f"{BASE_V1}/public/stats/storage")
    return stats


# ---------------------------------------------------------------------------
# Tier 2 — Report
# ---------------------------------------------------------------------------

@mcp.tool()
def get_protection_groups(only_active: bool = True) -> list[dict]:
    """List protection groups with id, name, environment, and paused state."""
    params = {"isDeleted": "false"} if only_active else {}
    data = _get(f"{BASE_V2}/data-protect/protection-groups", params)
    groups = data.get("protectionGroups", [])
    return [
        {
            "id": g.get("id"),
            "name": g.get("name"),
            "environment": g.get("environment"),
            "isPaused": g.get("isPaused"),
            "isActive": g.get("isActive"),
            "lastRunStatus": (g.get("lastRun", {}) or {})
                .get("localBackupInfo", {}).get("status"),
        }
        for g in groups
    ]


@mcp.tool()
def get_protection_runs(group_id: str, num_runs: int = 10,
                        only_failed: bool = False) -> list[dict]:
    """Recent runs for a protection group. Set only_failed=True to filter failures."""
    params = {"numRuns": num_runs, "includeObjectDetails": "false"}
    data = _get(f"{BASE_V2}/data-protect/protection-groups/{group_id}/runs", params)
    runs = data.get("runs", [])
    out = []
    for r in runs:
        local = r.get("localBackupInfo", {}) or {}
        status = local.get("status")
        if only_failed and status not in ("Failed", "Warning", "Canceled"):
            continue
        out.append({
            "runId": r.get("id"),
            "status": status,
            "startTimeUsecs": local.get("startTimeUsecs"),
            "endTimeUsecs": local.get("endTimeUsecs"),
            "messages": local.get("messages"),
            "successfulObjects": local.get("successfulObjectsCount"),
            "failedObjects": local.get("failedObjectsCount"),
        })
    return out


@mcp.tool()
def get_recovery_points(object_name: str) -> list[dict]:
    """Snapshots/recovery points for a named object (VM, database, NAS share)."""
    # Search for the object, then list its snapshots
    search = _get(f"{BASE_V2}/data-protect/search/objects",
                  {"searchString": object_name, "count": 5})
    objects = search.get("objects", [])
    if not objects:
        return [{"error": f"No protected object matching '{object_name}'"}]
    obj = objects[0]
    obj_id = obj.get("id")
    snaps = _get(f"{BASE_V2}/data-protect/objects/{obj_id}/snapshots")
    return [
        {
            "object": obj.get("name"),
            "snapshotId": s.get("id"),
            "runType": s.get("runType"),
            "protectionGroup": s.get("protectionGroupName"),
            "timestampUsecs": s.get("snapshotTimestampUsecs"),
            "expiryUsecs": s.get("expiryTimeUsecs"),
        }
        for s in snaps.get("snapshots", [])
    ]


@mcp.tool()
def healthcheck_summary() -> dict:
    """One-shot demo summary: health + open criticals + failed runs across all groups."""
    health = get_cluster_health()
    criticals = get_alerts(max_alerts=10, severity="kCritical")
    groups = get_protection_groups()
    failing = [g for g in groups
               if g.get("lastRunStatus") in ("Failed", "Warning")]
    return {
        "cluster": health.get("name"),
        "version": health.get("clusterSoftwareVersion"),
        "openCriticalAlerts": len(criticals),
        "criticalAlerts": criticals,
        "protectionGroups": len(groups),
        "groupsWithFailingLastRun": failing,
    }


# ---------------------------------------------------------------------------
# Tier 3 — Act (confirm pattern: confirm=False proposes, confirm=True executes)
# ---------------------------------------------------------------------------

@mcp.tool()
def pause_protection_group(group_id: str, confirm: bool = False) -> dict:
    """Pause future runs of a protection group. Call with confirm=False first to
    get a proposal; call again with confirm=True to execute. Reversible via
    resume_protection_group."""
    group = _get(f"{BASE_V2}/data-protect/protection-groups/{group_id}")
    if not confirm:
        return {
            "proposed_action": "PAUSE protection group",
            "group": group.get("name"),
            "currently_paused": group.get("isPaused"),
            "impact": "Future scheduled runs will not start. Existing snapshots unaffected.",
            "to_execute": "Call again with confirm=true",
        }
    body = {"action": "kPause", "ids": [group_id]}
    result = _post(f"{BASE_V2}/data-protect/protection-groups/states", body)
    return {"executed": "pause", "group": group.get("name"), "result": result}


@mcp.tool()
def resume_protection_group(group_id: str, confirm: bool = False) -> dict:
    """Resume a paused protection group. Confirm pattern applies."""
    group = _get(f"{BASE_V2}/data-protect/protection-groups/{group_id}")
    if not confirm:
        return {
            "proposed_action": "RESUME protection group",
            "group": group.get("name"),
            "currently_paused": group.get("isPaused"),
            "to_execute": "Call again with confirm=true",
        }
    body = {"action": "kResume", "ids": [group_id]}
    result = _post(f"{BASE_V2}/data-protect/protection-groups/states", body)
    return {"executed": "resume", "group": group.get("name"), "result": result}


@mcp.tool()
def retry_failed_run(group_id: str, run_id: str = "",
                     confirm: bool = False) -> dict:
    """Re-run a protection group (on-demand run), typically after a failure.
    Confirm pattern: proposes first, executes with confirm=True."""
    group = _get(f"{BASE_V2}/data-protect/protection-groups/{group_id}")
    if not confirm:
        last = get_protection_runs(group_id, num_runs=1)
        return {
            "proposed_action": "TRIGGER on-demand run",
            "group": group.get("name"),
            "last_run": last[0] if last else None,
            "impact": "Starts a new backup run now. Consumes cluster resources during run.",
            "to_execute": "Call again with confirm=true",
        }
    body = {"runType": "kRegular"}
    result = _post(f"{BASE_V2}/data-protect/protection-groups/{group_id}/runs",
                   body)
    return {"executed": "on-demand run", "group": group.get("name"),
            "result": result}


@mcp.tool()
def acknowledge_alert(alert_id: str, confirm: bool = False) -> dict:
    """Mark an alert as resolved/acknowledged. Confirm pattern applies."""
    if not confirm:
        return {
            "proposed_action": "RESOLVE alert",
            "alertId": alert_id,
            "impact": "Alert moves out of the open queue. Underlying issue is NOT fixed.",
            "to_execute": "Call again with confirm=true",
        }
    # VERIFY: resolution endpoint/body varies by version; check swagger
    body = {"alertIdList": [alert_id],
            "resolutionDetails": {
                "resolutionSummary": "Resolved via Cohesity MCP demo",
                "resolutionDetails": "Acknowledged by operator through LLM assistant"}}
    result = _post(f"{BASE_V1}/public/alertResolutions", body)
    return {"executed": "resolve alert", "alertId": alert_id, "result": result}


# ---------------------------------------------------------------------------
# Entrypoint
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    if not CLUSTER:
        raise SystemExit("Set COHESITY_CLUSTER (and COHESITY_API_KEY or username/password)")
    transport = os.environ.get("MCP_TRANSPORT", "stdio")
    if transport == "http":
        mcp.settings.port = int(os.environ.get("MCP_HTTP_PORT", "8720"))
        mcp.run(transport="streamable-http")
    else:
        mcp.run()
