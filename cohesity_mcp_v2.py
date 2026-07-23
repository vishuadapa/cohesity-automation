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


def _delete(url: str, params: dict | None = None) -> Any:
    resp = httpx.delete(url, headers=_headers(), params=params or {},
                        verify=VERIFY_SSL, timeout=120)
    resp.raise_for_status()
    return resp.json() if resp.text else {"status": resp.status_code}


def _patch(url: str, body: dict | None = None) -> Any:
    resp = httpx.patch(url, headers=_headers(), json=body or {},
                       verify=VERIFY_SSL, timeout=120)
    resp.raise_for_status()
    return resp.json() if resp.text else {"status": resp.status_code}


# ---------------------------------------------------------------------------
# Risk taxonomy, RBAC, and dynamic API helpers
# ---------------------------------------------------------------------------

# Risk levels (integer, higher = more dangerous)
_RISK_READ        = 0   # GET — safe, no side effects
_RISK_REVERSIBLE  = 1   # POST/PUT that can be undone (pause, resume, trigger run)
_RISK_CAUTION     = 2   # POST/PUT with lasting side-effects (create objects, modify config)
_RISK_DESTRUCTIVE = 3   # DELETE, or operations that permanently remove/alter data
_RISK_BLOCKED     = 4   # Hardcoded floor — refused regardless of role

_RISK_LABEL = {
    _RISK_READ:        "READ (safe)",
    _RISK_REVERSIBLE:  "REVERSIBLE",
    _RISK_CAUTION:     "CAUTION",
    _RISK_DESTRUCTIVE: "DESTRUCTIVE",
    _RISK_BLOCKED:     "BLOCKED",
}

# Path substrings that hard-block the request regardless of role
_ALWAYS_BLOCKED_SUBSTRINGS = [
    "factory-reset", "destroy-cluster", "destroycluster",
    "wipe-data", "erase-data", "unregister-cluster",
]

# Path substrings that escalate a POST/PUT to CAUTION or DESTRUCTIVE
_ESCALATE_TO_DESTRUCTIVE = ["delete", "destroy", "purge", "retire", "expir"]
_ESCALATE_TO_CAUTION     = ["create", "config", "setting", "gflag", "register", "modify", "update"]

# Cohesity role → maximum risk level allowed for that role
# Two naming conventions exist across Cohesity builds:
#   k-prefixed   — older / internal API format (kAdmin, kOperator, …)
#   COHESITY_-prefixed — format returned by /public/sessionUser on 7.x clusters
_ROLE_CEILING: dict[str, int] = {
    # k-prefixed variants
    "kViewer":              _RISK_READ,
    "kHeliosData":          _RISK_READ,
    "kRestoreOperator":     _RISK_REVERSIBLE,
    "kOperator":            _RISK_REVERSIBLE,
    "kDataSecurity":        _RISK_CAUTION,
    "kAdmin":               _RISK_DESTRUCTIVE,
    "kSuperAdmin":          _RISK_DESTRUCTIVE,
    # COHESITY_-prefixed variants (7.x sessionUser API)
    "COHESITY_VIEWER":          _RISK_READ,
    "COHESITY_OPERATOR":        _RISK_REVERSIBLE,
    "COHESITY_DATA_SECURITY":   _RISK_CAUTION,
    "COHESITY_ADMIN":           _RISK_DESTRUCTIVE,
    "COHESITY_SUPER_ADMIN":     _RISK_DESTRUCTIVE,
}

_user_roles_cache: list[str] | None = None
_swagger_cache: dict | None = None


def _classify_risk(method: str, path: str) -> int:
    m = method.upper()
    p = path.lower()

    if any(s in p for s in _ALWAYS_BLOCKED_SUBSTRINGS):
        return _RISK_BLOCKED

    if m == "GET":
        return _RISK_READ
    if m == "DELETE":
        base = _RISK_DESTRUCTIVE
    else:
        base = _RISK_REVERSIBLE

    if any(s in p for s in _ESCALATE_TO_DESTRUCTIVE):
        base = max(base, _RISK_DESTRUCTIVE)
    elif any(s in p for s in _ESCALATE_TO_CAUTION):
        base = max(base, _RISK_CAUTION)

    return base


def _fetch_user_roles() -> list[str]:
    global _user_roles_cache
    if _user_roles_cache is not None:
        return _user_roles_cache
    try:
        user = _get(f"{BASE_V1}/public/sessionUser")
        roles = user.get("roles", [])
    except Exception:
        roles = []
    if not roles and USERNAME:
        try:
            users = _get(f"{BASE_V1}/public/users", {"usernames": USERNAME})
            roles = (users[0].get("roles", []) if isinstance(users, list) and users else [])
        except Exception:
            pass
    _user_roles_cache = roles
    return roles


def _rbac_gate(risk: int) -> tuple[bool, str]:
    """Return (allowed, reason). Allowed=False means the MCP refuses to execute."""
    if risk == _RISK_BLOCKED:
        return False, (
            "REFUSED: This operation matches the hardcoded block list "
            "(factory-reset, cluster destroy, data-wipe, etc.). "
            "It will not be executed by this MCP server under any circumstances."
        )
    roles = _fetch_user_roles()
    if roles:
        ceiling = max((_ROLE_CEILING.get(r, _RISK_READ) for r in roles), default=_RISK_READ)
        role_str = ", ".join(roles)
    else:
        # Can't determine roles (common with API-key auth on some builds).
        # Default to REVERSIBLE as a conservative fallback.
        ceiling = _RISK_REVERSIBLE
        role_str = "unknown — defaulting to REVERSIBLE ceiling"

    if risk > ceiling:
        return False, (
            f"REFUSED: RBAC check failed. "
            f"Your roles ({role_str}) permit up to '{_RISK_LABEL[ceiling]}' operations. "
            f"This request is classified as '{_RISK_LABEL[risk]}'. "
            f"Elevate your Cohesity role or ask an admin to perform this action."
        )
    return True, (
        f"RBAC OK — roles: {role_str} | "
        f"ceiling: {_RISK_LABEL[ceiling]} | "
        f"this request: {_RISK_LABEL[risk]}"
    )


# ---------------------------------------------------------------------------
# Tier 1 — Monitor
# ---------------------------------------------------------------------------

@mcp.tool()
def get_cluster_health() -> dict:
    """Cluster identity, version, node count, and overall health/status."""
    info = _get(f"{BASE_V1}/public/cluster")
    status = _get(f"{BASE_V1}/nexus/cluster/status")
    stopped = (status.get("bulletinState") or {}).get("stoppedServices", [])
    return {
        "name": info.get("name"),
        "clusterSoftwareVersion": info.get("clusterSoftwareVersion"),
        "nodeCount": info.get("nodeCount"),
        "healthy": status.get("isServiceStateSynced") is True and not stopped,
        "healingStatus": status.get("healingStatus"),
        "isServiceStateSynced": status.get("isServiceStateSynced"),
        "stoppedServices": stopped,
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
        if only_failed and status not in ("Failed", "SucceededWithWarning", "Canceled"):
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
               if g.get("lastRunStatus") in ("Failed", "SucceededWithWarning")]
    return {
        "cluster": health.get("name"),
        "version": health.get("clusterSoftwareVersion"),
        "openCriticalAlerts": len(criticals),
        "criticalAlerts": criticals,
        "protectionGroups": len(groups),
        "groupsWithFailingLastRun": failing,
    }


# ---------------------------------------------------------------------------
# Tier 2.5 — Dynamic API discovery and identity
# ---------------------------------------------------------------------------

@mcp.tool()
def get_whoami() -> dict:
    """Return the current authenticated user's Cohesity identity, RBAC roles, and
    the effective risk ceiling those roles allow when using call_api.
    Call this first when you are unsure what operations the current user can perform."""
    global _user_roles_cache
    _user_roles_cache = None  # force a fresh fetch
    try:
        user = _get(f"{BASE_V1}/public/sessionUser")
    except Exception:
        user = {}
    roles = user.get("roles", [])
    _user_roles_cache = roles

    if roles:
        ceiling = max((_ROLE_CEILING.get(r, _RISK_READ) for r in roles), default=_RISK_READ)
    else:
        ceiling = _RISK_REVERSIBLE

    return {
        "username":       user.get("username", USERNAME or "(API key auth — username unavailable)"),
        "domain":         user.get("domain"),
        "description":    user.get("description"),
        "roles":          roles,
        "effectiveCeiling": _RISK_LABEL[ceiling],
        "ceilingNote": (
            "Roles and ceiling are fetched from /public/sessionUser. "
            "With API key auth on some builds, roles may be empty — "
            "ceiling then defaults to REVERSIBLE as a conservative fallback."
        ),
    }


@mcp.tool()
def search_cluster_api(query: str) -> list[dict]:
    """Search the cluster's live Swagger/OpenAPI spec for endpoints matching a
    natural-language or keyword query (e.g. 'delete snapshot', 'list users',
    'storage domain'). Returns up to 20 matching endpoints with their HTTP method,
    path, summary, pre-classified risk level, and parameter names.

    Use this to discover the right endpoint before calling call_api.
    Results include the risk label so you can advise the user before executing."""
    global _swagger_cache
    if _swagger_cache is None:
        for url in [
            f"https://{CLUSTER}/docs/api/swagger.json",
            f"https://{CLUSTER}/irisservices/docs/swagger.json",
            f"{BASE_V2}/swagger.json",
            f"{BASE_V1}/swagger.json",
        ]:
            try:
                resp = httpx.get(url, headers=_headers(), verify=VERIFY_SSL, timeout=30)
                if resp.status_code == 200:
                    _swagger_cache = resp.json()
                    break
            except Exception:
                continue

    if not _swagger_cache:
        return [{
            "error": (
                "Could not fetch the cluster Swagger spec. "
                "Tried /docs/api/swagger.json, /irisservices/docs/swagger.json, "
                "/v2/swagger.json, and /irisservices/api/v1/swagger.json. "
                "You can still call call_api with a known path."
            )
        }]

    words = query.lower().split()
    results = []
    for path, methods in (_swagger_cache.get("paths") or {}).items():
        for method, spec in methods.items():
            if method.upper() not in ("GET", "POST", "PUT", "PATCH", "DELETE"):
                continue
            summary     = spec.get("summary", "")
            description = spec.get("description", "")
            tags        = " ".join(spec.get("tags", []))
            text        = f"{path} {summary} {description} {tags}".lower()
            if not any(w in text for w in words):
                continue
            risk = _classify_risk(method, path)
            results.append({
                "method":     method.upper(),
                "path":       path,
                "summary":    summary,
                "risk":       _RISK_LABEL[risk],
                "parameters": [
                    {
                        "name":     p.get("name"),
                        "in":       p.get("in"),
                        "required": p.get("required", False),
                    }
                    for p in (spec.get("parameters") or [])
                ][:8],
            })

    results.sort(key=lambda r: (
        not any(w in r["path"].lower() for w in words),
        not any(w in r["summary"].lower() for w in words),
    ))
    return results[:20]




# ---------------------------------------------------------------------------
# Tier 2.6 — Community script discovery (Brian Seltzer's repo)
# ---------------------------------------------------------------------------

@mcp.tool()
def search_community_scripts(query: str, max_results: int = 3) -> list[dict]:
    """Search Brian Seltzer's Cohesity community scripts repo for a use case
    not natively supported by this MCP.

    Fetches matching Python scripts from github.com/bseltz-cohesity/scripts,
    extracts their Cohesity REST API call patterns, and returns them so the
    caller can execute the pattern via call_api.

    If ANTHROPIC_API_KEY is set in the environment, Claude is used to
    intelligently extract the endpoint and request body from each script.
    Without it, a regex fallback extracts raw api() call signatures.

    Args:
      query        Natural-language description of the use case, e.g.
                   "expire old snapshots", "create protection policy",
                   "clone VM to different vCenter", "list storage domains".
      max_results  Number of scripts to inspect (default 3, max 5).

    Workflow:
      1. Call this tool with a description of what you want to do.
      2. Review the returned api_pattern for the best-matching script.
      3. Fill in cluster-specific IDs (from prior get_* calls).
      4. Execute via call_api(method, path, body, confirm=False/True).
    """
    import re
    import json as _json

    GITHUB_SEARCH = "https://api.github.com/search/code"
    RAW_BASE      = "https://raw.githubusercontent.com/bseltz-cohesity/scripts/master"
    ANTHROPIC_URL = "https://api.anthropic.com/v1/messages"
    ANTHROPIC_KEY = os.environ.get("ANTHROPIC_API_KEY", "")

    max_results = min(max_results, 5)

    # ── 1. GitHub code search ────────────────────────────────────────────────
    try:
        gh_resp = httpx.get(
            GITHUB_SEARCH,
            params={
                "q":        query + " repo:bseltz-cohesity/scripts language:Python",
                "per_page": max_results * 3,
            },
            headers={"Accept": "application/vnd.github+json"},
            timeout=30,
            verify=True,
        )
        gh_resp.raise_for_status()
    except Exception as exc:
        return [{"error": "GitHub search failed: " + str(exc)}]

    skip  = {"pyhesity.py", "cohesity_api.py", "__init__.py"}
    items = [i for i in gh_resp.json().get("items", [])
             if i["name"].endswith(".py") and i["name"] not in skip][:max_results]

    if not items:
        return [{"error": "No matching scripts found for: " + query}]

    results = []
    for item in items:
        path   = item["path"]
        name   = item["name"]
        gh_url = "https://github.com/bseltz-cohesity/scripts/blob/master/" + path

        # ── 2. Fetch raw script ──────────────────────────────────────────────
        try:
            raw_resp = httpx.get(RAW_BASE + "/" + path, timeout=30, verify=True)
            raw_resp.raise_for_status()
            script = raw_resp.text
        except Exception as exc:
            results.append({"name": name, "github_url": gh_url,
                            "error": "Could not fetch: " + str(exc)})
            continue

        # ── 3. Extract API pattern ───────────────────────────────────────────
        if ANTHROPIC_KEY:
            prompt_lines = [
                "You are analysing a Cohesity community Python script to extract",
                "its primary REST API call pattern for the use case: " + query,
                "",
                "Script (may be truncated):",
                script[:7000],
                "",
                "Respond ONLY with a JSON object, no markdown fences:",
                "{",
                '  "summary": "one-sentence description",',
                '  "method": "POST",',
                '  "path": "/v2/...",',
                '  "body_template": {},',
                '  "required_inputs": ["ids or names the caller must supply"],',
                '  "notes": "caveats, ordering requirements, etc."',
                "}",
            ]
            extraction_prompt = "\n".join(prompt_lines)
            try:
                ai_resp = httpx.post(
                    ANTHROPIC_URL,
                    headers={
                        "x-api-key":         ANTHROPIC_KEY,
                        "anthropic-version": "2023-06-01",
                        "content-type":      "application/json",
                    },
                    json={
                        "model":      "claude-sonnet-4-6",
                        "max_tokens": 1000,
                        "messages":   [{"role": "user", "content": extraction_prompt}],
                    },
                    timeout=45,
                    verify=True,
                )
                raw_text = ai_resp.json()["content"][0]["text"]
                raw_text = re.sub(r"```[a-z]*", "", raw_text).strip("`\n ")
                pattern  = _json.loads(raw_text)
                pattern["extraction"] = "claude"
            except Exception as exc:
                pattern = {"extraction": "claude_failed", "error": str(exc)}
        else:
            # Regex fallback — extract pyhesity api() call signatures
            v2_pat  = re.compile(
                r"api\(['\"](\w+)['\"],\s*['\"]([^'\"]+)['\"][^)]*v\s*=\s*2",
                re.IGNORECASE,
            )
            v1_pat  = re.compile(
                r"api\(['\"](\w+)['\"],\s*['\"]([^'\"]+)['\"]",
                re.IGNORECASE,
            )
            v2_hits = [(m.upper(), p) for m, p in v2_pat.findall(script)]
            v1_hits = [(m.upper(), p) for m, p in v1_pat.findall(script)]
            seen: set = set()
            calls = []
            for method, path_str in (v2_hits + v1_hits):
                key = (method, path_str)
                if key not in seen:
                    seen.add(key)
                    calls.append({
                        "method":      method,
                        "path":        path_str,
                        "api_version": "v2" if (method, path_str) in v2_hits else "v1",
                    })
            pattern = {
                "extraction": "regex",
                "note": "Set ANTHROPIC_API_KEY for intelligent body extraction.",
                "api_calls":  calls[:8],
            }

        results.append({
            "name":        name,
            "github_url":  gh_url,
            "api_pattern": pattern,
        })

    return results


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
    body = {"alertIdList": [alert_id],
            "resolutionDetails": {
                "resolutionSummary": "Resolved via Cohesity MCP demo",
                "resolutionDetails": "Acknowledged by operator through LLM assistant"}}
    result = _post(f"{BASE_V1}/public/alertResolutions", body)
    return {"executed": "resolve alert", "alertId": alert_id, "result": result}


# ---------------------------------------------------------------------------
# Tier 3b — Instant Access (VMware VM clone from backup, with rename support)
# ---------------------------------------------------------------------------

@mcp.tool()
def instant_access_vm(
    vm_name:          str,
    prefix:           str  = "",
    suffix:           str  = "",
    power_on:         bool = True,
    detach_network:   bool = True,
    vcenter_id:       int  = 0,
    target_source_id: int  = 0,
    confirm:          bool = False,
) -> dict:
    """Instant Access — boot a single VMware VM directly from a Cohesity backup snapshot.

    Scoped to ONE VM via its snapshot ID (not the whole protection group run).
    The VM runs live off Cohesity storage — no full restore needed.

    Args:
      vm_name          Exact name of the source VM (case-sensitive).
      prefix           String to prepend to the recovered VM name (e.g. 'vishu-').
      suffix           String to append to the recovered VM name (e.g. '-test').
      power_on         Power the VM on after mount (default True).
      detach_network   Disconnect all NICs (default True — recommended for DR/test).
      vcenter_id       Cohesity source ID of the vCenter to search. 0 = auto-detect.
      target_source_id Cohesity source ID of the TARGET vCenter for recovery.
                       Set to a different vCenter to avoid naming conflicts with
                       live source VMs. 0 = recover to original vCenter (same source).
      confirm          False = propose only; True = execute.

    Recovered VM name: {prefix}{original_name}{suffix}
    """
    # ── 1. Find the protected object ─────────────────────────────────────────
    search = _get(f"{BASE_V2}/data-protect/search/objects",
                  {"searchString": vm_name, "environments": "kVMware", "count": 10})
    objects = search.get("objects", [])
    # Exact-name match first, fall back to first result
    obj = next((o for o in objects if o.get("name") == vm_name), None)
    if obj is None and objects:
        obj = objects[0]
    if obj is None:
        return {"error": f"No protected VMware object matching '{vm_name}'."}

    obj_id   = obj.get("id") or (
        (obj.get("objectProtectionInfos") or [{}])[0].get("objectId")
    )
    src_id   = (obj.get("objectProtectionInfos") or [{}])[0].get("sourceId") or vcenter_id
    src_name = (obj.get("sourceInfo") or {}).get("name", "unknown vCenter")

    # ── 2. Find the latest local snapshot ────────────────────────────────────
    snaps_data = _get(f"{BASE_V2}/data-protect/objects/{obj_id}/snapshots")
    snaps = [s for s in snaps_data.get("snapshots", [])
             if s.get("snapshotTargetType") == "Local"]
    if not snaps:
        return {"error": f"No local snapshots found for '{vm_name}'."}
    snap = max(snaps, key=lambda s: s.get("snapshotTimestampUsecs", 0))

    # Decode job/run info from snapshot fields
    pg_id    = snap.get("protectionGroupId", "")          # "clId:incId:jobId"
    run_id   = snap.get("protectionGroupRunId", "")        # "jobId:instanceId"
    ts_usecs = snap.get("snapshotTimestampUsecs", 0)

    parts      = pg_id.split(":")
    job_id     = int(parts[2]) if len(parts) == 3 else 0
    cluster_id = int(parts[0]) if len(parts) == 3 else 0
    incarn_id  = int(parts[1]) if len(parts) == 3 else 0

    new_name = f"{prefix}{vm_name}{suffix}"
    snap_id  = snap.get("id", "")

    # Determine if we're recovering to a different vCenter
    recover_to_new = target_source_id != 0 and target_source_id != src_id
    target_name    = "original vCenter (same source)" if not recover_to_new else f"source ID {target_source_id}"

    if not confirm:
        return {
            "proposed_action":     "INSTANT ACCESS — single VM (v2 snapshot-scoped)",
            "source_vm":           vm_name,
            "recovered_vm_name":   new_name,
            "snapshot_id":         snap_id,
            "snapshot_time_usecs": ts_usecs,
            "protection_group":    snap.get("protectionGroupName"),
            "source_vcenter":      src_name,
            "target":              target_name,
            "power_on":            power_on,
            "detach_network":      detach_network,
            "scope_note": (
                "Recovery is scoped to THIS VM only via its snapshot ID. "
                "No other VMs in the protection group will be touched."
            ),
            "impact": (
                "Boots the VM directly off Cohesity storage. Cohesity mounts its "
                "NFS view to the ESXi host automatically. Teardown removes the clone."
            ),
            "to_execute": "Call again with confirm=True to launch.",
        }

    # ── 4. RBAC gate — check before touching the cluster ────────────────────
    allowed, rbac_reason = _rbac_gate(_RISK_REVERSIBLE)
    if not allowed:
        return {
            "status":  "REFUSED",
            "rbac":    rbac_reason,
            "action":  "instant_access_vm",
            "vm_name": vm_name,
        }

    # ── 5. POST via v2 recoveries — snapshotId scopes to exactly ONE VM ──────
    # Key insight from Brian Seltzer's recoverVMs.py:
    #   - renameRecoveredVmsParams lives at vmwareTargetParams level (NOT inside originalSourceConfig)
    #   - powerOnVms lives at vmwareTargetParams level (NOT inside originalSourceConfig)
    #   - network disable on same-source recovery uses disableNetwork:true (NOT detachNetwork:true)
    if recover_to_new:
        # Recover to a different vCenter — look up its resource pool
        tgt_tree = _get(f"{BASE_V1}/public/protectionSources",
                        {"id": target_source_id, "environments": "kVMware",
                         "includeHosts": True})
        tgt_rpool = 0
        def _find_rpool(node: Any) -> None:
            nonlocal tgt_rpool
            if not isinstance(node, dict):
                return
            ps  = node.get("protectionSource", {})
            vmw = ps.get("vmWareProtectionSource", {})
            if vmw.get("type") == "kResourcePool" and not tgt_rpool:
                tgt_rpool = ps.get("id", 0)
            for child in node.get("nodes", []):
                _find_rpool(child)
        for root in (tgt_tree if isinstance(tgt_tree, list) else [tgt_tree]):
            _find_rpool(root)

        new_src_cfg: dict = {
            "sourceType": "kVMware",
            "vmwareTargetConfig": {
                "sourceId":  target_source_id,
                "networkConfig": {
                    "detachNetwork":  detach_network,
                    "disableNetwork": False,
                },
            },
        }
        if tgt_rpool:
            new_src_cfg["vmwareTargetConfig"]["resourcePoolId"] = tgt_rpool

        recovery_target: dict = {
            "recoverToNewSource": True,
            "newSourceConfig":    new_src_cfg,
        }
    else:
        # Recover to same vCenter
        recovery_target = {
            "recoverToNewSource": False,
            "originalSourceConfig": {
                "networkConfig": {
                    "detachNetwork":  False,
                    "disableNetwork": detach_network,  # disableNetwork, not detachNetwork
                },
            },
        }

    vmware_target_params: dict = {
        "recoveryTargetConfig": recovery_target,
        "powerOnVms":           power_on,       # top-level, NOT inside originalSourceConfig
        "continueOnError":      False,
        "recoveryProcessType":  "InstantRecovery",
    }

    # renameRecoveredVmsParams is at vmwareTargetParams level — NOT inside recoveryTargetConfig
    if prefix or suffix:
        rename: dict = {}
        if prefix:
            rename["prefix"] = prefix
        if suffix:
            rename["suffix"] = suffix
        vmware_target_params["renameRecoveredVmsParams"] = rename

    body: dict = {
        "name":                f"InstantAccess-{new_name}",
        "snapshotEnvironment": "kVMware",
        "vmwareParams": {
            "objects": [{"snapshotId": snap_id}],
            "recoveryAction": "RecoverVMs",
            "recoverVmParams": {
                "targetEnvironment": "kVMware",
                "vmwareTargetParams": vmware_target_params,
            },
        },
    }

    try:
        result = _post(f"{BASE_V2}/data-protect/recoveries", body)
    except httpx.HTTPStatusError as exc:
        return {
            "status":      "ERROR",
            "http_status": exc.response.status_code,
            "detail":      exc.response.text[:1000],
            "sent_body":   body,
        }

    recovery_id = result.get("id", "")
    return {
        "status":            "LAUNCHED",
        "recovered_vm_name": new_name,
        "source_vm":         vm_name,
        "recovery_id":       recovery_id,
        "target":            target_name,
        "power_on":          power_on,
        "detach_network":    detach_network,
        "result":            result,
    }


# ---------------------------------------------------------------------------
# Tier 4 — Generic executor (research → RBAC → risk → confirm → execute)
# ---------------------------------------------------------------------------

@mcp.tool()
def call_api(
    method:  str,
    path:    str,
    params:  dict | None = None,
    body:    dict | None = None,
    confirm: bool = False,
) -> dict:
    """Generic Cohesity API executor gated by RBAC and risk classification.

    Workflow:
      1. Call search_cluster_api to find the right endpoint.
      2. Call get_whoami to understand what the current user is allowed to do.
      3. Call call_api with confirm=False — the tool will classify the risk,
         run the RBAC check, and return a proposal. If the operation is too
         intrusive or destructive for the user's role, it REFUSES here and stops.
      4. If the proposal looks correct, call again with confirm=True to execute.

    Args:
      method   HTTP verb: GET | POST | PUT | PATCH | DELETE
      path     Full API path, e.g. /v2/data-protect/protection-groups
               or /irisservices/api/v1/public/cluster
      params   Query-string parameters (GET) or additional filters
      body     JSON request body (POST / PUT / PATCH)
      confirm  False = propose only (safe); True = execute (requires prior review)

    The tool REFUSES (returns status=REFUSED) and stops without calling the
    cluster API if:
      • The path matches the hardcoded block list (factory-reset, wipe, etc.)
      • The operation's risk level exceeds the current user's RBAC ceiling
    """
    m    = method.upper()
    risk = _classify_risk(m, path)
    allowed, reason = _rbac_gate(risk)

    # Build the full URL — accept paths with or without the scheme/host
    if path.startswith("https://") or path.startswith("http://"):
        url = path
    else:
        p = path if path.startswith("/") else f"/{path}"
        url = f"https://{CLUSTER}{p}"

    proposal = {
        "method":  m,
        "url":     url,
        "params":  params or {},
        "body":    body   or {},
        "risk":    _RISK_LABEL[risk],
        "rbac":    reason,
        "allowed": allowed,
    }

    if not allowed:
        proposal["status"] = "REFUSED"
        return proposal

    if not confirm:
        proposal["status"] = "PROPOSED"
        proposal["to_execute"] = (
            "Review the risk and rbac fields above, then call again with confirm=True to execute."
        )
        return proposal

    # Execute
    try:
        if m == "GET":
            result = _get(url, params)
        elif m == "POST":
            result = _post(url, body)
        elif m == "PUT":
            result = _put(url, body)
        elif m == "PATCH":
            result = _patch(url, body)
        elif m == "DELETE":
            result = _delete(url, params)
        else:
            return {"status": "REFUSED", "message": f"Unsupported HTTP method: {m}"}
    except httpx.HTTPStatusError as exc:
        return {
            "status":      "ERROR",
            "http_status": exc.response.status_code,
            "detail":      exc.response.text[:1000],
        }

    return {
        "status": "EXECUTED",
        "method": m,
        "url":    url,
        "risk":   _RISK_LABEL[risk],
        "rbac":   reason,
        "result": result,
    }


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
