"""
Cohesity MCP Server v3
Multi-cluster | AD + Helios auth | OS keyring | RBAC

Setup (one-time per machine):
    python cohesity_auth_setup.py

Config (no secrets, safe to distribute):
    ~/.cohesity-mcp/config.json        ← default location
    COHESITY_CONFIG=/custom/path.json  ← override via env var
    ./cohesity_clusters.json           ← fallback (same dir as script)

Transports:
    stdio (Claude Desktop):    python cohesity_mcp_v3.py
    HTTP  (VS Code / Copilot): MCP_TRANSPORT=http python cohesity_mcp_v3.py

Env vars:
    COHESITY_CONFIG   path to cluster config JSON (optional)
    MCP_TRANSPORT     "stdio" (default) | "http"
    MCP_HTTP_PORT     HTTP port (default 8720)

Requires: pip install "mcp[cli]" httpx keyring pyotp

Multi-cluster workflow:
    1. list_clusters()          — see all configured clusters + auth status
    2. switch_cluster("prod")   — set active cluster (all tools use this)
    3. get_whoami()             — confirm your identity and RBAC ceiling

Helios workflow:
    1. switch_cluster("helios")            — activate Helios connection
    2. helios_list_clusters()             — list all MCM-connected clusters
    3. helios_select_cluster(cluster_id)  — route calls to a specific cluster
    4. Any tool (get_alerts, etc.)        — now executes via Helios MCM proxy

SAFETY RULE (absolute, never overridden):
    Action tools (pause/resume/retry/acknowledge) and call_api with
    confirm=True must never be called unless the user has explicitly
    approved the proposed action in that specific conversation turn.
"""

import json
import os
import time
from pathlib import Path
from typing import Any, Optional

import httpx
import keyring
from mcp.server.fastmcp import FastMCP

try:
    import pyotp as _pyotp
    _PYOTP_OK = True
except ImportError:
    _PYOTP_OK = False

# ── Config loading ───────────────────────────────────────────────────────────

_CONFIG_SEARCH = [
    Path(os.environ.get("COHESITY_CONFIG", "__none__")),
    Path.home() / ".cohesity-mcp" / "config.json",
    Path(__file__).parent / "cohesity_clusters.json",
]


def _load_config() -> dict:
    for p in _CONFIG_SEARCH:
        if p.exists():
            return json.loads(p.read_text())
    raise SystemExit(
        "No cluster config found.\n"
        "Run:  python cohesity_auth_setup.py\n"
        "Or set COHESITY_CONFIG=/path/to/config.json"
    )


_cfg      = _load_config()
_clusters = {c["alias"]: c for c in _cfg.get("clusters", [])}

if not _clusters:
    raise SystemExit("Config defines no clusters. Edit the config file and restart.")

_default_alias: str = _cfg.get("default", next(iter(_clusters)))

# ── Per-cluster session state (in-memory only) ───────────────────────────────

KEYRING_SVC = "cohesity-mcp"


class _Sess:
    __slots__ = ("cfg", "token", "expiry", "mcm_cluster_id", "_roles")

    def __init__(self, cfg: dict) -> None:
        self.cfg              = cfg
        self.token: str       = ""
        self.expiry: float    = 0.0
        self.mcm_cluster_id: Optional[int] = None
        self._roles: Optional[list[str]]   = None


_sessions:     dict[str, _Sess] = {a: _Sess(c) for a, c in _clusters.items()}
_active_alias: str = _default_alias


# ── Session helpers ──────────────────────────────────────────────────────────

def _sess(alias: str = "") -> _Sess:
    a = alias or _active_alias
    if a not in _sessions:
        raise ValueError(f"Unknown cluster alias '{a}'. Known: {sorted(_sessions)}")
    return _sessions[a]


def _is_helios(alias: str = "") -> bool:
    return _sess(alias).cfg.get("type") == "helios"


def _host(alias: str = "") -> str:
    s = _sess(alias)
    return "helios.cohesity.com" if _is_helios(alias) else s.cfg["host"]


def _base_v1(alias: str = "") -> str:
    return f"https://{_host(alias)}/irisservices/api/v1"


def _base_v2(alias: str = "") -> str:
    return f"https://{_host(alias)}/v2"


def _ssl(alias: str = "") -> bool:
    return _sess(alias).cfg.get("verify_ssl", False)


# ── Authentication ───────────────────────────────────────────────────────────

def _is_mfa_challenge(body: dict) -> bool:
    code = body.get("errorCode", "").upper()
    msg  = body.get("message", "").lower()
    return (
        "MFA" in code or "OTP" in code or "MULTIFACTOR" in code
        or "two-factor" in msg or "otp" in msg or "one-time" in msg
        # Cohesity 7.x: KValidationError + "mandatory parameters" = missing otpCode
        or (code == "KVALIDATIONERROR" and "mandatory" in msg)
    )


def _get_token(alias: str = "") -> str:
    """Retrieve (or refresh) a session token for a direct-auth cluster.
    Handles TOTP MFA automatically when a secret is stored in the keyring."""
    a   = alias or _active_alias
    s   = _sessions[a]
    if s.token and time.time() < s.expiry:
        return s.token

    cfg  = s.cfg
    host = cfg["host"]
    user = cfg.get("username", "")
    dom  = cfg.get("domain", "LOCAL")
    pw   = keyring.get_password(KEYRING_SVC, f"{user}@{host}")

    if not pw:
        raise RuntimeError(
            f"No credentials found for cluster '{a}' ({user}@{host}).\n"
            f"Run  python cohesity_auth_setup.py  to set them up."
        )

    url    = f"https://{host}/irisservices/api/v1/public/accessTokens"
    verify = cfg.get("verify_ssl", False)

    # ── First attempt: password only ────────────────────────────────────
    try:
        resp = httpx.post(
            url,
            json={"username": user, "password": pw, "domain": dom},
            verify=verify, timeout=30,
        )
        resp.raise_for_status()

    except httpx.HTTPStatusError as exc:
        # ── MFA challenge: try TOTP auto-generation ──────────────────
        try:
            body = exc.response.json()
        except Exception:
            body = {}

        if _is_mfa_challenge(body):
            totp_secret = keyring.get_password(KEYRING_SVC, f"totp-secret@{host}")
            if totp_secret and _PYOTP_OK:
                otp_code = _pyotp.TOTP(totp_secret).now()
                resp = httpx.post(
                    url,
                    json={"username": user, "password": pw, "domain": dom,
                          "otpCode": otp_code, "otpType": "Totp"},
                    verify=verify, timeout=30,
                )
                resp.raise_for_status()
            elif totp_secret and not _PYOTP_OK:
                raise RuntimeError(
                    "MFA required and a TOTP secret is stored, but pyotp is not installed.\n"
                    "Run:  pip install pyotp"
                )
            else:
                raise RuntimeError(
                    f"MFA is required for cluster '{a}' but no TOTP secret is configured.\n"
                    "Run  python cohesity_auth_setup.py  and store your TOTP secret,\n"
                    "or use a service account or API key that does not require MFA."
                )
        else:
            raise   # re-raise non-MFA errors unchanged

    tok      = resp.json()
    s.token  = f"{tok['tokenType']} {tok['accessToken']}"
    s.expiry = time.time() + 23 * 3600
    s._roles = None
    return s.token


def _get_helios_key(alias: str = "") -> str:
    """Retrieve the Helios API key from keyring."""
    a    = alias or _active_alias
    host = _sessions[a].cfg.get("host", "helios.cohesity.com")
    key  = keyring.get_password(KEYRING_SVC, f"helios-apikey@{host}")
    if not key:
        raise RuntimeError(
            f"No Helios API key found for '{a}'.\n"
            f"Run  python cohesity_auth_setup.py  to configure."
        )
    return key


def _headers(alias: str = "") -> dict:
    """Build the correct auth headers for the active (or named) cluster."""
    a = alias or _active_alias
    s = _sessions[a]

    if _is_helios(a):
        h: dict = {
            "apiKey": _get_helios_key(a),
            "Accept": "application/json",
        }
        if s.mcm_cluster_id is not None:
            h["clusterIdentifier"] = str(s.mcm_cluster_id)
        return h

    return {
        "Authorization": _get_token(a),
        "Accept":        "application/json",
    }


# ── HTTP helpers ─────────────────────────────────────────────────────────────

def _get(url: str, params: dict | None = None, alias: str = "") -> Any:
    r = httpx.get(url, headers=_headers(alias), params=params or {},
                  verify=_ssl(alias), timeout=60)
    r.raise_for_status()
    return r.json()


def _post(url: str, body: dict | None = None, alias: str = "") -> Any:
    r = httpx.post(url, headers=_headers(alias), json=body or {},
                   verify=_ssl(alias), timeout=120)
    r.raise_for_status()
    return r.json() if r.text else {"status": r.status_code}


def _put(url: str, body: dict | None = None, alias: str = "") -> Any:
    r = httpx.put(url, headers=_headers(alias), json=body or {},
                  verify=_ssl(alias), timeout=120)
    r.raise_for_status()
    return r.json() if r.text else {"status": r.status_code}


def _patch(url: str, body: dict | None = None, alias: str = "") -> Any:
    r = httpx.patch(url, headers=_headers(alias), json=body or {},
                    verify=_ssl(alias), timeout=120)
    r.raise_for_status()
    return r.json() if r.text else {"status": r.status_code}


def _delete(url: str, params: dict | None = None, alias: str = "") -> Any:
    r = httpx.delete(url, headers=_headers(alias), params=params or {},
                     verify=_ssl(alias), timeout=120)
    r.raise_for_status()
    return r.json() if r.text else {"status": r.status_code}


# ── Risk taxonomy and RBAC ───────────────────────────────────────────────────

_RISK_READ        = 0   # GET — no side effects
_RISK_REVERSIBLE  = 1   # Pause, resume, trigger run — undoable
_RISK_CAUTION     = 2   # Create / config changes — lasting but not destructive
_RISK_DESTRUCTIVE = 3   # Delete, purge, permanently alter data
_RISK_BLOCKED     = 4   # Hardcoded floor — refused regardless of role

_RISK_LABEL = {
    0: "READ (safe)",
    1: "REVERSIBLE",
    2: "CAUTION",
    3: "DESTRUCTIVE",
    4: "BLOCKED",
}

_ALWAYS_BLOCKED = [
    "factory-reset", "destroy-cluster", "destroycluster",
    "wipe-data", "erase-data", "unregister-cluster",
]
_ESCALATE_TO_DESTRUCTIVE = ["delete", "destroy", "purge", "retire", "expir"]
_ESCALATE_TO_CAUTION     = ["create", "config", "setting", "gflag",
                             "register", "modify", "update"]

_ROLE_CEILING: dict[str, int] = {
    "kViewer":          _RISK_READ,
    "kHeliosData":      _RISK_READ,
    "kRestoreOperator": _RISK_REVERSIBLE,
    "kOperator":        _RISK_REVERSIBLE,
    "kDataSecurity":    _RISK_CAUTION,
    "kAdmin":           _RISK_DESTRUCTIVE,
    "kSuperAdmin":      _RISK_DESTRUCTIVE,
}

_swagger_cache: dict[str, Optional[dict]] = {}   # keyed by alias


def _classify_risk(method: str, path: str) -> int:
    m, p = method.upper(), path.lower()
    if any(s in p for s in _ALWAYS_BLOCKED):
        return _RISK_BLOCKED
    if m == "GET":
        return _RISK_READ
    base = _RISK_DESTRUCTIVE if m == "DELETE" else _RISK_REVERSIBLE
    if any(s in p for s in _ESCALATE_TO_DESTRUCTIVE):
        base = max(base, _RISK_DESTRUCTIVE)
    elif any(s in p for s in _ESCALATE_TO_CAUTION):
        base = max(base, _RISK_CAUTION)
    return base


def _fetch_roles(alias: str = "") -> list[str]:
    """Fetch RBAC roles for the current user, with in-session caching."""
    a = alias or _active_alias
    s = _sessions[a]
    if s._roles is not None:
        return s._roles

    roles: list[str] = []
    try:
        user  = _get(f"{_base_v1(a)}/public/sessionUser", alias=a)
        roles = user.get("roles", [])
    except Exception:
        pass

    if not roles and not _is_helios(a):
        uname = s.cfg.get("username", "")
        if uname:
            try:
                users = _get(f"{_base_v1(a)}/public/users",
                             {"usernames": uname}, alias=a)
                roles = (users[0].get("roles", [])
                         if isinstance(users, list) and users else [])
            except Exception:
                pass

    s._roles = roles
    return roles


def _rbac_gate(risk: int, alias: str = "") -> tuple[bool, str]:
    """Return (allowed, reason). Refused = hard stop, no cluster call made."""
    if risk == _RISK_BLOCKED:
        return False, (
            "REFUSED: matches hardcoded block list "
            "(factory-reset, destroy-cluster, wipe-data, etc.). "
            "This will never be executed by this MCP server."
        )

    roles = _fetch_roles(alias)
    if roles:
        ceiling  = max((_ROLE_CEILING.get(r, _RISK_READ) for r in roles), default=_RISK_READ)
        role_str = ", ".join(roles)
    else:
        ceiling  = _RISK_REVERSIBLE
        role_str = "unknown — defaulting to REVERSIBLE ceiling"

    if risk > ceiling:
        return False, (
            f"REFUSED: RBAC check failed. "
            f"Your roles ({role_str}) permit up to '{_RISK_LABEL[ceiling]}'. "
            f"This request is classified as '{_RISK_LABEL[risk]}'. "
            f"Ask an admin to perform this action or elevate your Cohesity role."
        )

    return True, (
        f"RBAC OK — roles: {role_str} | "
        f"ceiling: {_RISK_LABEL[ceiling]} | "
        f"this request: {_RISK_LABEL[risk]}"
    )


# ── MCP server ───────────────────────────────────────────────────────────────

mcp = FastMCP("cohesity")


# ─── Cluster management ─────────────────────────────────────────────────────

@mcp.tool()
def list_clusters() -> list[dict]:
    """List all configured clusters with alias, type, host, active status, and
    whether credentials are present in the OS keyring. Call this first to see
    what's available and which cluster is currently active."""
    out = []
    for alias, s in _sessions.items():
        cfg   = s.cfg
        ctype = cfg.get("type", "direct")
        row: dict = {
            "alias":  alias,
            "type":   ctype,
            "host":   cfg["host"],
            "active": alias == _active_alias,
        }
        if ctype == "helios":
            h    = cfg.get("host", "helios.cohesity.com")
            row["auth_configured"] = bool(
                keyring.get_password(KEYRING_SVC, f"helios-apikey@{h}")
            )
            row["mcm_cluster_id"] = s.mcm_cluster_id
        else:
            user = cfg.get("username", "")
            row["username"]        = user
            row["domain"]          = cfg.get("domain", "LOCAL")
            row["auth_configured"] = bool(
                keyring.get_password(KEYRING_SVC, f"{user}@{cfg['host']}")
            )
            row["token_cached"]    = bool(s.token and time.time() < s.expiry)
        out.append(row)
    return out


@mcp.tool()
def switch_cluster(alias: str) -> dict:
    """Switch the active cluster. All subsequent tool calls target this cluster
    until switched again. Use list_clusters() to see available aliases."""
    global _active_alias
    if alias not in _sessions:
        return {"error": f"Unknown alias '{alias}'. Known: {sorted(_sessions)}"}
    _active_alias = alias
    s             = _sessions[alias]
    cfg           = s.cfg
    return {
        "active":       alias,
        "host":         cfg["host"],
        "type":         cfg.get("type", "direct"),
        "username":     cfg.get("username", "(helios — API key auth)"),
        "domain":       cfg.get("domain", ""),
        "note":         "All tools now target this cluster.",
    }


@mcp.tool()
def helios_list_clusters() -> list[dict]:
    """List all clusters connected to Helios (MCM view). The active cluster must
    be type=helios — use switch_cluster('helios') first if not already set.
    Returns cluster_id, name, version, health, and connection status.
    Pass cluster_id to helios_select_cluster() to route subsequent calls."""
    if not _is_helios():
        return [{"error": (
            f"Active cluster '{_active_alias}' is type=direct, not type=helios. "
            f"Call switch_cluster('helios') first."
        )}]
    data = _get(f"{_base_v1()}/mcm/clusters/info")
    items = data if isinstance(data, list) else data.get("clusterList", [])
    return [
        {
            "cluster_id":    c.get("clusterId"),
            "name":          c.get("name"),
            "version":       c.get("clusterSoftwareVersion"),
            "health_status": c.get("healthStatus"),
            "connected":     c.get("connectedToHelios"),
        }
        for c in items
    ]


@mcp.tool()
def helios_select_cluster(cluster_id: int) -> dict:
    """Route Helios API calls to a specific cluster via MCM proxy.
    cluster_id comes from helios_list_clusters(). Pass 0 to clear routing
    (subsequent calls go to Helios-level endpoints only)."""
    if not _is_helios():
        return {"error": (
            f"Active cluster '{_active_alias}' is not type=helios. "
            f"Call switch_cluster('helios') first."
        )}
    s = _sessions[_active_alias]
    if cluster_id == 0:
        s.mcm_cluster_id = None
        return {"mcm_target": None, "note": "MCM routing cleared — calls go to Helios-level endpoints."}
    s.mcm_cluster_id = cluster_id
    return {
        "mcm_target": cluster_id,
        "note": (
            f"All subsequent API calls route to cluster {cluster_id} via "
            f"helios.cohesity.com MCM (clusterIdentifier header). "
            f"Call get_cluster_health() to verify connectivity."
        ),
    }


# ─── Monitor ────────────────────────────────────────────────────────────────

@mcp.tool()
def get_cluster_health() -> dict:
    """Cluster identity, version, node count, and overall health.
    Works for both direct clusters and Helios MCM-routed clusters."""
    info    = _get(f"{_base_v1()}/public/cluster")
    status  = _get(f"{_base_v1()}/nexus/cluster/status")
    stopped = (status.get("bulletinState") or {}).get("stoppedServices", [])
    return {
        "via_cluster":            _active_alias,
        "name":                   info.get("name"),
        "clusterSoftwareVersion": info.get("clusterSoftwareVersion"),
        "nodeCount":              info.get("nodeCount"),
        "healthy":                status.get("isServiceStateSynced") is True and not stopped,
        "healingStatus":          status.get("healingStatus"),
        "isServiceStateSynced":   status.get("isServiceStateSynced"),
        "stoppedServices":        stopped,
    }


@mcp.tool()
def get_alerts(max_alerts: int = 25, severity: str = "") -> list[dict]:
    """Open alerts on the active cluster.
    severity: kCritical | kWarning | kInfo  (blank = all)"""
    params: dict = {"maxAlerts": max_alerts, "alertStateList": "kOpen"}
    if severity:
        params["alertSeverityList"] = severity
    alerts = _get(f"{_base_v1()}/public/alerts", params)
    return [
        {
            "id":          a.get("id"),
            "severity":    a.get("severity"),
            "name":        a.get("alertDocument", {}).get("alertName"),
            "description": a.get("alertDocument", {}).get("alertDescription"),
            "firstSeen":   a.get("firstTimestampUsecs"),
            "latest":      a.get("latestTimestampUsecs"),
        }
        for a in alerts
    ]


@mcp.tool()
def get_storage_stats() -> dict:
    """Capacity, usage, and data-reduction ratios for the active cluster."""
    return _get(f"{_base_v1()}/public/stats/storage")


# ─── Report ─────────────────────────────────────────────────────────────────

@mcp.tool()
def get_protection_groups(only_active: bool = True) -> list[dict]:
    """List protection groups with id, name, environment, and last-run status."""
    params = {"isDeleted": "false"} if only_active else {}
    data   = _get(f"{_base_v2()}/data-protect/protection-groups", params)
    return [
        {
            "id":            g.get("id"),
            "name":          g.get("name"),
            "environment":   g.get("environment"),
            "isPaused":      g.get("isPaused"),
            "isActive":      g.get("isActive"),
            "lastRunStatus": (g.get("lastRun") or {})
                             .get("localBackupInfo", {}).get("status"),
        }
        for g in data.get("protectionGroups", [])
    ]


@mcp.tool()
def get_protection_runs(group_id: str, num_runs: int = 10,
                        only_failed: bool = False) -> list[dict]:
    """Recent runs for a protection group. only_failed=True filters to failures only."""
    params = {"numRuns": num_runs, "includeObjectDetails": "false"}
    data   = _get(f"{_base_v2()}/data-protect/protection-groups/{group_id}/runs", params)
    out    = []
    for r in data.get("runs", []):
        local  = r.get("localBackupInfo") or {}
        status = local.get("status")
        if only_failed and status not in ("Failed", "SucceededWithWarning", "Canceled"):
            continue
        out.append({
            "runId":             r.get("id"),
            "status":            status,
            "startTimeUsecs":    local.get("startTimeUsecs"),
            "endTimeUsecs":      local.get("endTimeUsecs"),
            "messages":          local.get("messages"),
            "successfulObjects": local.get("successfulObjectsCount"),
            "failedObjects":     local.get("failedObjectsCount"),
        })
    return out


@mcp.tool()
def get_recovery_points(object_name: str) -> list[dict]:
    """Snapshots/recovery points for a named object (VM, database, NAS share)."""
    search  = _get(f"{_base_v2()}/data-protect/search/objects",
                   {"searchString": object_name, "count": 5})
    objects = search.get("objects", [])
    if not objects:
        return [{"error": f"No protected object matching '{object_name}'"}]
    obj    = objects[0]
    obj_id = obj.get("id")
    snaps  = _get(f"{_base_v2()}/data-protect/objects/{obj_id}/snapshots")
    return [
        {
            "object":          obj.get("name"),
            "snapshotId":      s.get("id"),
            "runType":         s.get("runType"),
            "protectionGroup": s.get("protectionGroupName"),
            "timestampUsecs":  s.get("snapshotTimestampUsecs"),
            "expiryUsecs":     s.get("expiryTimeUsecs"),
        }
        for s in snaps.get("snapshots", [])
    ]


@mcp.tool()
def healthcheck_summary() -> dict:
    """One-shot summary: health + open critical alerts + failing protection groups."""
    health    = get_cluster_health()
    criticals = get_alerts(max_alerts=10, severity="kCritical")
    groups    = get_protection_groups()
    failing   = [g for g in groups
                 if g.get("lastRunStatus") in ("Failed", "SucceededWithWarning")]
    return {
        "via_cluster":              _active_alias,
        "clusterName":              health.get("name"),
        "version":                  health.get("clusterSoftwareVersion"),
        "openCriticalAlerts":       len(criticals),
        "criticalAlerts":           criticals,
        "protectionGroupsTotal":    len(groups),
        "groupsWithFailingLastRun": failing,
    }


# ─── Identity + API discovery ────────────────────────────────────────────────

@mcp.tool()
def get_whoami() -> dict:
    """Return the current authenticated user's identity, RBAC roles, and effective
    risk ceiling. For Helios, shows the Helios account and the active MCM target.
    Call this first when unsure what operations the current user can perform."""
    a = _active_alias
    s = _sessions[a]
    s._roles = None   # force fresh role fetch

    if _is_helios(a):
        try:
            info = _get(f"{_base_v1(a)}/mcm/userInfo")
        except Exception:
            info = {}
        roles = info.get("roles", [])
        s._roles = roles
        ceiling = (
            max((_ROLE_CEILING.get(r, _RISK_READ) for r in roles), default=_RISK_READ)
            if roles else _RISK_REVERSIBLE
        )
        return {
            "via_cluster":      a,
            "type":             "helios",
            "email":            info.get("email"),
            "username":         info.get("username"),
            "roles":            roles,
            "effectiveCeiling": _RISK_LABEL[ceiling],
            "mcm_cluster_id":   s.mcm_cluster_id,
            "note": (
                "Helios roles are account-level. When MCM routing targets a cluster, "
                "the cluster's role mapping for your AD user may further restrict access."
            ),
        }

    try:
        user = _get(f"{_base_v1(a)}/public/sessionUser")
    except Exception:
        user = {}
    roles   = user.get("roles", [])
    s._roles = roles
    ceiling = (
        max((_ROLE_CEILING.get(r, _RISK_READ) for r in roles), default=_RISK_READ)
        if roles else _RISK_REVERSIBLE
    )
    return {
        "via_cluster":      a,
        "type":             "direct",
        "username":         user.get("username", s.cfg.get("username", "")),
        "domain":           user.get("domain", s.cfg.get("domain", "")),
        "roles":            roles,
        "effectiveCeiling": _RISK_LABEL[ceiling],
    }


@mcp.tool()
def search_cluster_api(query: str) -> list[dict]:
    """Search the cluster's live Swagger/OpenAPI spec for endpoints matching a
    natural-language or keyword query (e.g. 'delete snapshot', 'storage domain').
    Returns up to 20 matches with method, path, risk classification, and parameters.
    Use this to discover the right endpoint before calling call_api."""
    a = _active_alias
    if a not in _swagger_cache:
        _swagger_cache[a] = None
        h = _host(a)
        for url in [
            f"https://{h}/docs/api/swagger.json",
            f"https://{h}/irisservices/docs/swagger.json",
            f"{_base_v2(a)}/swagger.json",
            f"{_base_v1(a)}/swagger.json",
        ]:
            try:
                r = httpx.get(url, headers=_headers(a), verify=_ssl(a), timeout=30)
                if r.status_code == 200:
                    _swagger_cache[a] = r.json()
                    break
            except Exception:
                continue

    if not _swagger_cache.get(a):
        return [{"error": (
            "Could not fetch cluster Swagger spec. "
            "You can still call call_api with a known path."
        )}]

    words   = query.lower().split()
    results = []
    for path, methods in (_swagger_cache[a].get("paths") or {}).items():
        for method, spec in methods.items():
            if method.upper() not in ("GET", "POST", "PUT", "PATCH", "DELETE"):
                continue
            text = (
                f"{path} {spec.get('summary','')} "
                f"{spec.get('description','')} "
                f"{' '.join(spec.get('tags',[]))}"
            ).lower()
            if not any(w in text for w in words):
                continue
            risk = _classify_risk(method, path)
            results.append({
                "method":     method.upper(),
                "path":       path,
                "summary":    spec.get("summary", ""),
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


# ─── Act (confirm pattern) ───────────────────────────────────────────────────

@mcp.tool()
def pause_protection_group(group_id: str, confirm: bool = False) -> dict:
    """Pause future runs of a protection group.
    SAFETY: Call with confirm=False first to review; confirm=True only after
    explicit user approval in this conversation turn."""
    group = _get(f"{_base_v2()}/data-protect/protection-groups/{group_id}")
    if not confirm:
        return {
            "proposed_action":  "PAUSE protection group",
            "group":            group.get("name"),
            "currently_paused": group.get("isPaused"),
            "impact":           "Future scheduled runs will not start. Existing snapshots are unaffected.",
            "to_execute":       "Call again with confirm=True after user approval.",
        }
    result = _post(
        f"{_base_v2()}/data-protect/protection-groups/states",
        {"action": "kPause", "ids": [group_id]},
    )
    return {"executed": "pause", "group": group.get("name"), "result": result}


@mcp.tool()
def resume_protection_group(group_id: str, confirm: bool = False) -> dict:
    """Resume a paused protection group. Confirm pattern applies."""
    group = _get(f"{_base_v2()}/data-protect/protection-groups/{group_id}")
    if not confirm:
        return {
            "proposed_action":  "RESUME protection group",
            "group":            group.get("name"),
            "currently_paused": group.get("isPaused"),
            "to_execute":       "Call again with confirm=True after user approval.",
        }
    result = _post(
        f"{_base_v2()}/data-protect/protection-groups/states",
        {"action": "kResume", "ids": [group_id]},
    )
    return {"executed": "resume", "group": group.get("name"), "result": result}


@mcp.tool()
def retry_failed_run(group_id: str, confirm: bool = False) -> dict:
    """Trigger an on-demand backup run for a protection group (typically after failure).
    Confirm pattern: propose first, execute only with confirm=True and user approval."""
    group = _get(f"{_base_v2()}/data-protect/protection-groups/{group_id}")
    if not confirm:
        last = get_protection_runs(group_id, num_runs=1)
        return {
            "proposed_action": "TRIGGER on-demand backup run",
            "group":           group.get("name"),
            "last_run":        last[0] if last else None,
            "impact":          "Starts a new backup run immediately. Consumes cluster resources.",
            "to_execute":      "Call again with confirm=True after user approval.",
        }
    result = _post(
        f"{_base_v2()}/data-protect/protection-groups/{group_id}/runs",
        {"runType": "kRegular"},
    )
    return {"executed": "on-demand run", "group": group.get("name"), "result": result}


@mcp.tool()
def acknowledge_alert(alert_id: str, confirm: bool = False) -> dict:
    """Mark an alert as resolved/acknowledged. Confirm pattern applies.
    Note: acknowledging an alert does NOT fix the underlying issue."""
    if not confirm:
        return {
            "proposed_action": "RESOLVE alert",
            "alertId":         alert_id,
            "impact":          "Alert moves out of the open queue. Underlying issue is NOT fixed by this action.",
            "to_execute":      "Call again with confirm=True after user approval.",
        }
    result = _post(
        f"{_base_v1()}/public/alertResolutions",
        {
            "alertIdList": [alert_id],
            "resolutionDetails": {
                "resolutionSummary": "Resolved via Cohesity MCP",
                "resolutionDetails": "Acknowledged by operator through AI assistant",
            },
        },
    )
    return {"executed": "resolve alert", "alertId": alert_id, "result": result}


# ─── Generic executor ─────────────────────────────────────────────────────────

@mcp.tool()
def call_api(
    method:  str,
    path:    str,
    params:  dict | None = None,
    body:    dict | None = None,
    confirm: bool = False,
) -> dict:
    """Generic Cohesity API executor gated by RBAC and risk classification.

    Recommended workflow:
      1. search_cluster_api() — find the right endpoint and its risk level.
      2. get_whoami()         — confirm what the current user is allowed to do.
      3. call_api(..., confirm=False) — classifies risk, checks RBAC, returns proposal.
         If the operation exceeds the user's RBAC ceiling it REFUSES here.
      4. call_api(..., confirm=True) — executes ONLY after explicit user approval.

    SAFETY RULE: Never call with confirm=True unless the user has explicitly
    reviewed and approved the proposal in this conversation turn.

    The tool REFUSES (status=REFUSED, no cluster call made) if:
      • The path matches the hardcoded block list (factory-reset, wipe, etc.)
      • The operation's risk level exceeds the current user's RBAC ceiling

    Args:
      method   GET | POST | PUT | PATCH | DELETE
      path     Full API path, e.g. /v2/data-protect/protection-groups
      params   Query-string parameters
      body     JSON request body (POST / PUT / PATCH)
      confirm  False = propose only (safe); True = execute
    """
    m    = method.upper()
    risk = _classify_risk(m, path)
    allowed, reason = _rbac_gate(risk)

    h = _host()
    if path.startswith("https://") or path.startswith("http://"):
        url = path
    else:
        p   = path if path.startswith("/") else f"/{path}"
        url = f"https://{h}{p}"

    proposal: dict = {
        "via_cluster": _active_alias,
        "method":      m,
        "url":         url,
        "params":      params or {},
        "body":        body   or {},
        "risk":        _RISK_LABEL[risk],
        "rbac":        reason,
        "allowed":     allowed,
    }

    if not allowed:
        proposal["status"] = "REFUSED"
        return proposal

    if not confirm:
        proposal["status"]     = "PROPOSED"
        proposal["to_execute"] = (
            "Review the risk and rbac fields above. "
            "Call again with confirm=True only after the user explicitly approves."
        )
        return proposal

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
            return {"status": "REFUSED", "message": f"Unsupported method: {m}"}
    except httpx.HTTPStatusError as exc:
        return {
            "status":      "ERROR",
            "http_status": exc.response.status_code,
            "detail":      exc.response.text[:1000],
        }

    return {
        "status":      "EXECUTED",
        "via_cluster": _active_alias,
        "method":      m,
        "url":         url,
        "risk":        _RISK_LABEL[risk],
        "rbac":        reason,
        "result":      result,
    }


# ── Entrypoint ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    transport = os.environ.get("MCP_TRANSPORT", "stdio")
    if transport == "http":
        mcp.settings.port = int(os.environ.get("MCP_HTTP_PORT", "8720"))
        mcp.run(transport="streamable-http")
    else:
        mcp.run()
