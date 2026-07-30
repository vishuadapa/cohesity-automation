"""
Cohesity MCP Server v5 — multi-cluster profiles + Helios, secure storage
Personal SE demo tool: monitor -> report -> act (with confirm pattern)

v5 = v4 + named profiles + Helios support.

  * Profiles: store credentials for any number of clusters in the OS
    lockbox (macOS Keychain / Windows Credential Manager / Linux Secret
    Service), each under a name like 'prod', 'dr', 'lab'. Claude Desktop
    selects one per server entry with the non-secret COHESITY_PROFILE env
    var — so you can expose several clusters side by side, still with
    zero secrets in the config file.

  * Helios: a profile can be Helios-type (Helios API key). Two extra MCP
    tools appear: list_clusters (show clusters connected to Helios) and
    select_cluster (pick the one to operate on). All other tools then work
    against the selected cluster through Helios passthrough
    (accessClusterId header).

One-time setup per profile (secrets never echoed):
    python3 cohesity_mcp_v5.py setup

Other commands:
    python3 cohesity_mcp_v5.py list            # list stored profiles
    python3 cohesity_mcp_v5.py use <name>      # set the default profile
    python3 cohesity_mcp_v5.py show [name]     # view config (key masked)
    python3 cohesity_mcp_v5.py test [name]     # connect and verify auth
    python3 cohesity_mcp_v5.py clear [name]    # remove one profile ('--all' = everything)
    python3 cohesity_mcp_v5.py                 # run the MCP server (default)

Profile selection when the server runs:
    COHESITY_PROFILE env var  ->  stored default  ->  'default'

v4 compatibility: a v4 lockbox (un-namespaced keys) is read automatically
as the profile named 'default'. v2-style env vars (COHESITY_CLUSTER,
COHESITY_API_KEY, ...) still override everything.

Transports:
  stdio (default, for Claude Desktop):    python cohesity_mcp_v5.py
  streamable HTTP (for Copilot/VS Code):  MCP_TRANSPORT=http python cohesity_mcp_v5.py

Requires: pip install "mcp[cli]" httpx keyring
"""

import os
import sys
import time
import getpass
import httpx
from typing import Any, Optional

try:
    from mcp.server.fastmcp import FastMCP        # MCP SDK 1.x
except ImportError:
    from mcp.server.mcpserver import MCPServer as FastMCP  # MCP SDK 2.x renamed it

__version__ = "5.0"

# ---------------------------------------------------------------------------
# Secure credential storage (OS lockbox via `keyring`) — profile-aware
# ---------------------------------------------------------------------------

SERVICE_NAME = "cohesity-mcp"
# Per-profile entries are stored as "<profile>:<key>". Two housekeeping
# entries track the profile index and the default: "_profiles", "_default".
_PROFILE_KEYS = ("type", "cluster", "api_key", "username", "password",
                 "domain", "verify_ssl")
_LEGACY_KEYS = ("cluster", "api_key", "username", "password", "domain",
                "verify_ssl")   # v4 layout, un-namespaced

HELIOS_DEFAULT_ADDRESS = "helios.cohesity.com"


def _keyring():
    """Return the keyring module, or None if it isn't installed."""
    try:
        import keyring
        return keyring
    except ImportError:
        return None


def _lockbox_get(key: str) -> str:
    kr = _keyring()
    if kr is None:
        return ""
    try:
        return kr.get_password(SERVICE_NAME, key) or ""
    except Exception:
        # A locked/unavailable backend must never crash the server —
        # env vars remain a valid fallback.
        return ""


def _lockbox_set(key: str, value: str) -> None:
    import keyring
    keyring.set_password(SERVICE_NAME, key, value)


def _lockbox_delete(key: str) -> None:
    import keyring
    import keyring.errors
    try:
        keyring.delete_password(SERVICE_NAME, key)
    except keyring.errors.PasswordDeleteError:
        pass  # not stored — nothing to delete


def _profile_names() -> list[str]:
    names = [n for n in _lockbox_get("_profiles").split(",") if n]
    # A v4-era lockbox (un-namespaced keys, no index) shows up as 'default'
    if not names and _lockbox_get("cluster"):
        names = ["default"]
    return names


def _save_profile_names(names: list[str]) -> None:
    _lockbox_set("_profiles", ",".join(sorted(set(names))))


def _default_profile() -> str:
    stored = _lockbox_get("_default")
    if stored:
        return stored
    names = _profile_names()
    return names[0] if names else "default"


def _profile_get(profile: str, key: str) -> str:
    val = _lockbox_get(f"{profile}:{key}")
    if not val and profile == "default" and key in _LEGACY_KEYS:
        val = _lockbox_get(key)          # fall back to the v4 layout
    return val


def _load_profile(profile: str) -> dict:
    """Resolve a profile's settings. Env vars win (v2/v4 compatibility)."""
    def resolve(env_name: str, key: str, default: str = "") -> str:
        return os.environ.get(env_name) or _profile_get(profile, key) or default

    ptype = _profile_get(profile, "type") or "cluster"
    return {
        "profile":    profile,
        "type":       ptype,
        "cluster":    resolve("COHESITY_CLUSTER", "cluster"),
        "api_key":    resolve("COHESITY_API_KEY", "api_key"),
        "username":   resolve("COHESITY_USERNAME", "username"),
        "password":   resolve("COHESITY_PASSWORD", "password"),
        "domain":     resolve("COHESITY_DOMAIN", "domain", "LOCAL"),
        "verify_ssl": resolve("COHESITY_VERIFY_SSL", "verify_ssl",
                              "true" if ptype == "helios" else "false"),
    }


# ---------------------------------------------------------------------------
# Config — the active profile, chosen once at startup
# ---------------------------------------------------------------------------

ACTIVE_PROFILE = os.environ.get("COHESITY_PROFILE") or _default_profile()
CFG = _load_profile(ACTIVE_PROFILE)

CLUSTER    = CFG["cluster"]
API_KEY    = CFG["api_key"]
USERNAME   = CFG["username"]
PASSWORD   = CFG["password"]
DOMAIN     = CFG["domain"]
VERIFY_SSL = CFG["verify_ssl"].lower() == "true"
IS_HELIOS  = CFG["type"] == "helios"

BASE_V1 = f"https://{CLUSTER}/irisservices/api/v1"
BASE_V2 = f"https://{CLUSTER}/v2"

mcp = FastMCP("cohesity")

_session_token: Optional[str] = None
_session_expiry: float = 0.0

# Helios: the cluster all passthrough calls are routed to (accessClusterId)
_access_cluster_id: Optional[int] = None
_access_cluster_name: str = ""
_helios_clusters: list[dict] | None = None


def _base_headers() -> dict:
    """Auth headers only — no Helios cluster routing."""
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


def _headers() -> dict:
    """Headers for cluster-scoped calls. In Helios mode a target cluster
    must have been chosen with select_cluster first."""
    h = _base_headers()
    if IS_HELIOS:
        if _access_cluster_id is None:
            raise RuntimeError(
                "Helios mode: no target cluster selected. "
                "Call list_clusters to see what is connected, then "
                "select_cluster('<name or id>') before using other tools."
            )
        h["accessClusterId"] = str(_access_cluster_id)
    return h


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
# Helios — connected-cluster discovery and selection
# ---------------------------------------------------------------------------

def _fetch_helios_clusters() -> list[dict]:
    """Clusters registered to this Helios account, via whichever endpoint
    the account/build supports."""
    global _helios_clusters
    if _helios_clusters is not None:
        return _helios_clusters
    errors = []
    for url, extract in [
        (f"{BASE_V1}/public/mcm/clusters/connectionStatus",
         lambda d: [{"clusterId": c.get("clusterId"),
                     "name": c.get("clusterName") or c.get("name"),
                     "connected": c.get("connectedToCluster")}
                    for c in (d if isinstance(d, list) else [])]),
        (f"{BASE_V2}/mcm/cluster-mgmt/info",
         lambda d: [{"clusterId": c.get("clusterId") or c.get("id"),
                     "name": c.get("clusterName") or c.get("name"),
                     "connected": c.get("connectedToCluster", True)}
                    for c in (d.get("cohesityClusters") or d.get("clusters") or [])]),
    ]:
        try:
            resp = httpx.get(url, headers=_base_headers(),
                             verify=VERIFY_SSL, timeout=60)
            resp.raise_for_status()
            clusters = [c for c in extract(resp.json()) if c.get("clusterId")]
            if clusters:
                _helios_clusters = clusters
                return clusters
        except Exception as exc:
            errors.append(f"{url}: {exc}")
    raise RuntimeError("Could not list Helios clusters. Tried:\n" + "\n".join(errors))


@mcp.tool()
def list_clusters() -> list[dict] | dict:
    """Helios mode only: list all Cohesity clusters connected to this Helios
    account, with their id, name, and connection state. Call this first, then
    select_cluster to choose which cluster the other tools operate on."""
    if not IS_HELIOS:
        return {"note": (
            f"This profile ('{ACTIVE_PROFILE}') connects directly to a single "
            f"cluster ({CLUSTER}) — there is nothing to list or select. "
            "list_clusters/select_cluster only apply to Helios-type profiles."
        )}
    clusters = _fetch_helios_clusters()
    return [
        {**c, "selected": c["clusterId"] == _access_cluster_id}
        for c in clusters
    ]


@mcp.tool()
def select_cluster(cluster: str) -> dict:
    """Helios mode only: choose the target cluster (by name or numeric id)
    for all subsequent tool calls. Returns the selection so you can confirm."""
    global _access_cluster_id, _access_cluster_name
    if not IS_HELIOS:
        return {"note": (
            f"This profile ('{ACTIVE_PROFILE}') is cluster-direct ({CLUSTER}); "
            "selection is not needed."
        )}
    wanted = cluster.strip().lower()
    clusters = _fetch_helios_clusters()
    match = next((c for c in clusters
                  if str(c["clusterId"]) == wanted
                  or (c.get("name") or "").lower() == wanted), None)
    if match is None:
        return {
            "error": f"No connected cluster matching '{cluster}'.",
            "available": [{"clusterId": c["clusterId"], "name": c["name"]}
                          for c in clusters],
        }
    if not match.get("connected", True):
        return {
            "error": f"Cluster '{match['name']}' is registered but not "
                     "currently connected to Helios.",
        }
    _access_cluster_id = match["clusterId"]
    _access_cluster_name = match.get("name") or str(match["clusterId"])
    return {
        "selected": _access_cluster_name,
        "clusterId": _access_cluster_id,
        "note": "All tools now operate on this cluster via Helios passthrough.",
    }


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
_ROLE_CEILING: dict[str, int] = {
    "kViewer":           _RISK_READ,
    "kHeliosData":       _RISK_READ,
    "kRestoreOperator":  _RISK_REVERSIBLE,
    "kOperator":         _RISK_REVERSIBLE,
    "kDataSecurity":     _RISK_CAUTION,
    "kAdmin":            _RISK_DESTRUCTIVE,
    "kSuperAdmin":       _RISK_DESTRUCTIVE,
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
    # nexus status is cluster-local; it may be unavailable through Helios
    # passthrough — degrade gracefully rather than fail the whole tool.
    try:
        status = _get(f"{BASE_V1}/nexus/cluster/status")
    except Exception:
        status = {}
    stopped = (status.get("bulletinState") or {}).get("stoppedServices", [])
    return {
        "name": info.get("name"),
        "clusterSoftwareVersion": info.get("clusterSoftwareVersion"),
        "nodeCount": info.get("nodeCount"),
        "healthy": status.get("isServiceStateSynced") is True and not stopped,
        "healingStatus": status.get("healingStatus"),
        "isServiceStateSynced": status.get("isServiceStateSynced"),
        "stoppedServices": stopped,
        "statusNote": (None if status else
                       "Service-level status unavailable via this connection "
                       "(normal for Helios passthrough); 'healthy' may read false."),
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
        "profile":        ACTIVE_PROFILE,
        "connection":     (f"Helios ({CLUSTER}) -> "
                           f"{_access_cluster_name or '(no cluster selected)'}"
                           if IS_HELIOS else f"direct ({CLUSTER})"),
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
    if IS_HELIOS:
        proposal["target_cluster"] = _access_cluster_name or "(none selected)"

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
# Setup / management CLI (setup | list | use | show | test | clear)
# ---------------------------------------------------------------------------

_USAGE = f"""Cohesity MCP Server v{__version__} — profiles + Helios, secure credential storage

Usage:
  python3 cohesity_mcp_v5.py setup         Add or update a profile (interactive).
                                           Stores credentials in the OS lockbox —
                                           macOS Keychain / Windows Credential Manager.
  python3 cohesity_mcp_v5.py list          List stored profiles.
  python3 cohesity_mcp_v5.py use <name>    Set the default profile.
  python3 cohesity_mcp_v5.py show [name]   Show a profile's config (secrets masked).
  python3 cohesity_mcp_v5.py test [name]   Connect and verify authentication.
  python3 cohesity_mcp_v5.py clear <name>  Remove one profile ('--all' = everything).
  python3 cohesity_mcp_v5.py               Run the MCP server (used by Claude Desktop).

The running server picks its profile from the COHESITY_PROFILE env var,
falling back to the stored default. Add one Claude Desktop entry per cluster:

     "cohesity-prod": {{ "command": "python3", "args": ["...cohesity_mcp_v5.py"],
                        "env": {{ "COHESITY_PROFILE": "prod" }} }}
"""


def _require_keyring():
    kr = _keyring()
    if kr is None:
        raise SystemExit(
            "The 'keyring' package is not installed.\n"
            "Install it with:  pip install keyring\n"
            "(or reinstall everything:  pip install \"mcp[cli]\" httpx keyring)"
        )
    return kr


def _backend_name() -> str:
    """Human-readable name of the lockbox in use on this machine."""
    kr = _keyring()
    if kr is None:
        return "(keyring not installed)"
    try:
        backend = kr.get_keyring()
        name = type(backend).__module__ + "." + type(backend).__name__
        friendly = {
            "keyring.backends.macOS.Keyring": "macOS Keychain",
            "keyring.backends.Windows.WinVaultKeyring": "Windows Credential Manager",
            "keyring.backends.SecretService.Keyring": "Linux Secret Service (GNOME Keyring)",
            "keyring.backends.kwallet.DBusKeyring": "KDE KWallet",
        }
        return friendly.get(name, name)
    except Exception:
        return "(unknown backend)"


def _mask(secret: str) -> str:
    if not secret:
        return "(not set)"
    if len(secret) <= 4:
        return "****"
    return "****" + secret[-4:]


def _valid_profile_name(name: str) -> bool:
    return bool(name) and ":" not in name and "," not in name and " " not in name


def _cmd_setup() -> None:
    _require_keyring()
    print(f"Cohesity MCP v{__version__} — profile setup")
    print(f"Credentials will be stored in: {_backend_name()}")
    existing = _profile_names()
    if existing:
        print(f"Existing profiles: {', '.join(existing)}")
    print("Secrets are hidden while you type (paste works too).\n")

    name = input("Profile name [default]: ").strip() or "default"
    if not _valid_profile_name(name):
        raise SystemExit("Profile names cannot contain spaces, ':' or ','. Nothing was saved.")

    kind = input("Connection type — [1] single cluster (direct)  [2] Helios [1]: ").strip()
    is_helios = kind == "2"

    if is_helios:
        cluster = input(f"Helios address [{HELIOS_DEFAULT_ADDRESS}]: ").strip() \
                  or HELIOS_DEFAULT_ADDRESS
        api_key = getpass.getpass("Helios API key (required): ").strip()
        if not api_key:
            raise SystemExit("Helios requires an API key (username/password is not "
                             "supported). Nothing was saved.")
        username = password = ""
        domain = "LOCAL"
        default_verify = "Y/n"
    else:
        cluster = input("Cluster address (e.g. mycluster.lab.local — no https://): ").strip()
        if not cluster:
            raise SystemExit("Cluster address is required. Nothing was saved.")
        api_key = getpass.getpass(
            "API key (recommended — press Enter to use username/password instead): "
        ).strip()
        username = password = ""
        domain = "LOCAL"
        if not api_key:
            username = input("Username: ").strip()
            password = getpass.getpass("Password: ")
            if not username or not password:
                raise SystemExit("Username and password are both required. Nothing was saved.")
            domain = input("Domain [LOCAL]: ").strip() or "LOCAL"
        default_verify = "y/N"

    verify = input(f"Verify SSL certificates? (n for lab/self-signed) [{default_verify}]: "
                   ).strip().lower()
    if not verify:
        verify_ssl = "true" if is_helios else "false"
    else:
        verify_ssl = "true" if verify in ("y", "yes") else "false"

    _lockbox_set(f"{name}:type", "helios" if is_helios else "cluster")
    _lockbox_set(f"{name}:cluster", cluster)
    _lockbox_set(f"{name}:verify_ssl", verify_ssl)
    if api_key:
        _lockbox_set(f"{name}:api_key", api_key)
        for key in ("username", "password", "domain"):
            _lockbox_delete(f"{name}:{key}")
        auth_desc = f"API key {_mask(api_key)}"
    else:
        _lockbox_set(f"{name}:username", username)
        _lockbox_set(f"{name}:password", password)
        _lockbox_set(f"{name}:domain", domain)
        _lockbox_delete(f"{name}:api_key")
        auth_desc = f"username '{username}' (domain {domain})"

    names = _profile_names()
    if name not in names:
        names.append(name)
    _save_profile_names(names)

    if len(names) == 1:
        _lockbox_set("_default", name)
        made_default = True
    else:
        current_default = _default_profile()
        ans = input(f"Make '{name}' the default profile? "
                    f"(current: {current_default}) [y/N]: ").strip().lower()
        made_default = ans in ("y", "yes")
        if made_default:
            _lockbox_set("_default", name)

    kind_desc = "Helios" if is_helios else "cluster-direct"
    print(f"\nSaved profile '{name}' ({kind_desc}). "
          f"Address: {cluster} | Auth: {auth_desc} | Verify SSL: {verify_ssl}"
          + (" | DEFAULT" if made_default else ""))
    print(f"Stored in {_backend_name()} under service '{SERVICE_NAME}'.")
    print("\nNext steps:")
    print(f"  1. Verify the connection:   python3 cohesity_mcp_v5.py test {name}")
    print("  2. Add an entry to claude_desktop_config.json — no secrets needed.")
    print("     One entry per profile lets Claude see several clusters at once:")
    print('''
     {
       "mcpServers": {
         "cohesity-''' + name + '''": {
           "command": "python3",
           "args": ["''' + os.path.abspath(__file__) + '''"],
           "env": { "COHESITY_PROFILE": "''' + name + '''" }
         }
       }
     }
''')
    print("  3. Fully quit and relaunch Claude Desktop.")
    if is_helios:
        print("\n  Helios tip: in Claude, start with \"list my Helios clusters\" "
              "then \"select cluster <name>\".")


def _cmd_list() -> None:
    _require_keyring()
    names = _profile_names()
    if not names:
        print("No profiles stored. Run:  python3 cohesity_mcp_v5.py setup")
        return
    default = _default_profile()
    print(f"Lockbox backend : {_backend_name()}")
    print(f"Profiles ({len(names)}):")
    for n in sorted(names):
        ptype = _profile_get(n, "type") or "cluster"
        addr = _profile_get(n, "cluster") or "(no address)"
        mark = "  *default*" if n == default else ""
        print(f"  {n:<16} {ptype:<8} {addr}{mark}")


def _cmd_use(name: str) -> None:
    _require_keyring()
    if name not in _profile_names():
        raise SystemExit(f"No profile named '{name}'. "
                         f"Stored profiles: {', '.join(_profile_names()) or '(none)'}")
    _lockbox_set("_default", name)
    print(f"Default profile is now '{name}'.")


def _cmd_show(name: str = "") -> None:
    _require_keyring()
    profile = name or _default_profile()
    if profile not in _profile_names():
        raise SystemExit(f"No profile named '{profile}'. "
                         f"Stored profiles: {', '.join(_profile_names()) or '(none)'}")
    ptype = _profile_get(profile, "type") or "cluster"
    print(f"Lockbox backend : {_backend_name()}")
    print(f"Service name    : {SERVICE_NAME}")
    print(f"Profile         : {profile}"
          + ("  (default)" if profile == _default_profile() else ""))
    print(f"Type            : {ptype}")
    print(f"Address         : {_profile_get(profile, 'cluster') or '(not set)'}")
    api_key = _profile_get(profile, "api_key")
    if api_key:
        print(f"Auth method     : API key {_mask(api_key)}")
    elif _profile_get(profile, "username"):
        pw_state = "(set, hidden)" if _profile_get(profile, "password") else "(not set)"
        print(f"Auth method     : username '{_profile_get(profile, 'username')}' "
              f"(domain {_profile_get(profile, 'domain') or 'LOCAL'}), password {pw_state}")
    else:
        print("Auth method     : (not set — run: python3 cohesity_mcp_v5.py setup)")
    print(f"Verify SSL      : {_profile_get(profile, 'verify_ssl') or ('true' if ptype == 'helios' else 'false')}")
    env_overrides = [v for v in ("COHESITY_CLUSTER", "COHESITY_API_KEY", "COHESITY_USERNAME",
                                 "COHESITY_PASSWORD", "COHESITY_DOMAIN", "COHESITY_VERIFY_SSL")
                     if os.environ.get(v)]
    if env_overrides:
        print(f"NOTE: env vars currently override the lockbox: {', '.join(env_overrides)}")


def _cmd_clear(name: str = "") -> None:
    _require_keyring()
    if not name:
        raise SystemExit(
            "Say which profile to remove:  python3 cohesity_mcp_v5.py clear <name>\n"
            "Or remove everything:         python3 cohesity_mcp_v5.py clear --all"
        )
    if name == "--all":
        for n in _profile_names():
            for key in _PROFILE_KEYS:
                _lockbox_delete(f"{n}:{key}")
        for key in _LEGACY_KEYS:            # v4-era un-namespaced entries
            _lockbox_delete(key)
        _lockbox_delete("_profiles")
        _lockbox_delete("_default")
        print(f"All Cohesity profiles removed from {_backend_name()} "
              f"(service '{SERVICE_NAME}').")
        return
    if name not in _profile_names():
        raise SystemExit(f"No profile named '{name}'. "
                         f"Stored profiles: {', '.join(_profile_names()) or '(none)'}")
    for key in _PROFILE_KEYS:
        _lockbox_delete(f"{name}:{key}")
    if name == "default":
        for key in _LEGACY_KEYS:
            _lockbox_delete(key)
    remaining = [n for n in _profile_names() if n != name]
    _save_profile_names(remaining)
    if _lockbox_get("_default") == name:
        if remaining:
            _lockbox_set("_default", remaining[0])
            print(f"Profile '{name}' removed. Default is now '{remaining[0]}'.")
        else:
            _lockbox_delete("_default")
            print(f"Profile '{name}' removed. No profiles left.")
    else:
        print(f"Profile '{name}' removed.")


def _cmd_test(name: str = "") -> None:
    profile = name or ACTIVE_PROFILE
    cfg = _load_profile(profile)
    if not cfg["cluster"]:
        raise SystemExit(
            f"Profile '{profile}' has no address configured. "
            "Run:  python3 cohesity_mcp_v5.py setup"
        )
    verify = cfg["verify_ssl"].lower() == "true"
    headers = {"apiKey": cfg["api_key"], "Accept": "application/json"}
    base_v1 = f"https://{cfg['cluster']}/irisservices/api/v1"

    if cfg["type"] == "helios":
        print(f"Connecting to Helios at {cfg['cluster']} ...")
        url = f"{base_v1}/public/mcm/clusters/connectionStatus"
    else:
        print(f"Connecting to {cfg['cluster']} ...")
        if not cfg["api_key"]:
            resp = httpx.post(f"{base_v1}/public/accessTokens",
                              json={"username": cfg["username"],
                                    "password": cfg["password"],
                                    "domain": cfg["domain"]},
                              verify=verify, timeout=30)
            if resp.status_code in (401, 403):
                raise SystemExit("Authentication FAILED — username/password rejected. "
                                 "Re-run: python3 cohesity_mcp_v5.py setup")
            resp.raise_for_status()
            tok = resp.json()
            headers = {"Authorization": f"{tok['tokenType']} {tok['accessToken']}",
                       "Accept": "application/json"}
        url = f"{base_v1}/public/cluster"

    try:
        resp = httpx.get(url, headers=headers, verify=verify, timeout=60)
        resp.raise_for_status()
        info = resp.json()
    except httpx.HTTPStatusError as exc:
        if exc.response.status_code in (401, 403):
            raise SystemExit(
                f"Authentication FAILED (HTTP {exc.response.status_code}). "
                "The stored API key or credentials were rejected — "
                "re-run: python3 cohesity_mcp_v5.py setup"
            )
        raise SystemExit(f"Server returned HTTP {exc.response.status_code}: "
                         f"{exc.response.text[:300]}")
    except httpx.RequestError as exc:
        raise SystemExit(
            f"Could not reach {cfg['cluster']}: {exc}\n"
            "Check the address, your network/VPN, and firewall rules."
        )

    print("SUCCESS — authenticated.")
    if cfg["type"] == "helios":
        clusters = info if isinstance(info, list) else []
        connected = [c for c in clusters if c.get("connectedToCluster")]
        print(f"  Helios account with {len(clusters)} registered cluster(s), "
              f"{len(connected)} connected:")
        for c in clusters:
            state = "connected" if c.get("connectedToCluster") else "NOT connected"
            print(f"    - {c.get('clusterName') or c.get('name')} "
                  f"(id {c.get('clusterId')}) — {state}")
    else:
        print(f"  Cluster : {info.get('name')}")
        print(f"  Version : {info.get('clusterSoftwareVersion')}")
        print(f"  Nodes   : {info.get('nodeCount')}")


# ---------------------------------------------------------------------------
# Entrypoint
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    cmd = sys.argv[1].lower() if len(sys.argv) > 1 else ""
    arg = sys.argv[2] if len(sys.argv) > 2 else ""
    if cmd == "setup":
        _cmd_setup()
    elif cmd == "list":
        _cmd_list()
    elif cmd == "use":
        if not arg:
            raise SystemExit("Usage: python3 cohesity_mcp_v5.py use <profile>")
        _cmd_use(arg)
    elif cmd == "show":
        _cmd_show(arg)
    elif cmd == "test":
        _cmd_test(arg)
    elif cmd in ("clear", "reset"):
        _cmd_clear(arg)
    elif cmd in ("help", "-h", "--help"):
        print(_USAGE)
    elif cmd:
        raise SystemExit(f"Unknown command: {cmd}\n\n{_USAGE}")
    else:
        if not CLUSTER:
            raise SystemExit(
                f"Profile '{ACTIVE_PROFILE}' has no address configured.\n"
                "Run the one-time setup first:  python3 cohesity_mcp_v5.py setup\n"
                "(or set COHESITY_CLUSTER and COHESITY_API_KEY env vars as in v2)"
            )
        transport = os.environ.get("MCP_TRANSPORT", "stdio")
        if transport == "http":
            port = int(os.environ.get("MCP_HTTP_PORT", "8720"))
            if hasattr(mcp, "settings"):          # MCP SDK 1.x
                mcp.settings.port = port
                mcp.run(transport="streamable-http")
            else:                                 # MCP SDK 2.x
                mcp.run(transport="streamable-http", port=port)
        else:
            mcp.run()
