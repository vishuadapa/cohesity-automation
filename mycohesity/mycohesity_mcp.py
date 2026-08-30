"""
MyCohesity MCP Server — my.cohesity.com from Claude Desktop

Search the MyCohesity knowledge base (Cohesity DataProtect and Veritas NetBackup)
and pull asset / licensing entitlement information, using the signed-in user's own
portal session.

Authentication
  my.cohesity.com is a Salesforce Experience Cloud portal. There is no password
  grant to call:
    * External users (customers, partners) have MFA enforced.
    * Internal Cohesity users never sign in to my.cohesity.com directly — they
      authenticate to the Cohesity Salesforce org and launch MyCohesity from the
      app launcher (the icon stack, top left).
  So the first connection opens a real Chromium window and hands the keyboard to
  the user for the whole flow. The resulting browser session is then encrypted
  (Fernet, key in the OS keychain) into ~/.mycohesity/session.enc and reused for
  every subsequent request until the portal rejects it.

Transports:
  stdio (default, for Claude Desktop):    python mycohesity_mcp.py
  streamable HTTP (for Copilot/VS Code):  MCP_TRANSPORT=http python mycohesity_mcp.py

Env vars:
  MYCOHESITY_PORTAL_URL      default https://my.cohesity.com
  MYCOHESITY_SALESFORCE_URL  default https://cohesity.lightning.force.com
  MYCOHESITY_HOME            default ~/.mycohesity   (session + config directory)
  MYCOHESITY_LOGIN_TIMEOUT   default 600 seconds     (time allowed for SSO + MFA)
  MYCOHESITY_HTTP_TIMEOUT    default 60 seconds
  MCP_TRANSPORT              "stdio" (default) | "http"
  MCP_HTTP_PORT              default 8721

Requires: pip install "mcp[cli]" playwright cryptography httpx keyring
          (works with both mcp 1.x and 2.x)
          playwright install chromium

Version: 1.0
"""

from __future__ import annotations

import json
import os
import sys
from concurrent.futures import ThreadPoolExecutor
from typing import Any, Callable

try:                                            # mcp 1.x
    from mcp.server.fastmcp import FastMCP as MCPServer
except ModuleNotFoundError:                     # mcp 2.x renamed FastMCP -> MCPServer
    from mcp.server.mcpserver import MCPServer

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import mycohesity_auth as auth
from mycohesity_auth import MyCohesityClient, SessionExpired

__version__ = "1.0"

mcp = MCPServer("mycohesity")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

# Playwright's sync API refuses to run on a thread that owns a live asyncio loop,
# and the MCP server may call sync tools straight from the event loop thread. Bouncing
# every browser call through a plain worker thread keeps it valid either way.
_POOL = ThreadPoolExecutor(max_workers=1, thread_name_prefix="mycohesity-browser")


def _in_thread(fn: Callable, *args, **kwargs) -> Any:
    return _POOL.submit(fn, *args, **kwargs).result()


def _error(detail: str, hint: str = "") -> dict:
    out = {"status": "ERROR", "detail": detail}
    if hint:
        out["hint"] = hint
    return out


def _dig(obj: Any, path: str) -> Any:
    """Walk a dotted path such as 'result.records' or 'data.0.items'."""
    if not path:
        return obj
    current = obj
    for part in path.split("."):
        if current is None:
            return None
        if part.isdigit() and isinstance(current, list):
            index = int(part)
            current = current[index] if index < len(current) else None
        elif isinstance(current, dict):
            current = current.get(part)
        else:
            return None
    return current


# Query terms appended so a search is scoped to the right product family.
_PRODUCT_TERMS = {
    "dataprotect": "DataProtect",
    "dp": "DataProtect",
    "cohesity": "DataProtect",
    "netbackup": "NetBackup",
    "nbu": "NetBackup",
    "veritas": "NetBackup",
    "all": "",
    "both": "",
    "": "",
}

# A link is treated as a knowledge-base result if its URL looks like an article.
_ARTICLE_MARKERS = ("/s/article/", "/article/", "/knowledge", "/s/kb", "articleid", "/s/topic/")

# Chrome, footer and utility links that pollute scraped Lightning pages.
_NAV_NOISE = {
    "home", "search", "login", "log in", "logout", "log out", "sign in", "sign out",
    "contact us", "privacy policy", "terms of use", "cookie preferences", "skip to main content",
    "support", "downloads", "documentation", "my cases", "back", "next", "previous",
}


def _run_endpoint(name: str, wait_ms: int = 4000, **fmt) -> dict:
    """
    Execute a configured endpoint in whichever mode it declares.

      mode="api"   → authenticated JSON request over the stored cookies (fast)
      mode="page"  → headless render of the Lightning SPA, then DOM extraction

    Returns {"mode": ..., "url": ..., "data"/"page": ...} or raises SessionExpired.
    """
    endpoints = auth.load_endpoints()
    config = endpoints.get(name)
    if not isinstance(config, dict) or not config.get("url"):
        raise ValueError(
            f"Endpoint '{name}' is not configured. Set it with mycohesity_set_endpoint, "
            "or run mycohesity_discover_endpoints to find the right URL for your tenant."
        )

    url = auth.resolve_url(config["url"], endpoints, **fmt)
    mode = (config.get("mode") or "page").lower()

    if mode == "api":
        params = dict(config.get("params") or {})
        for key, value in fmt.items():
            params = {k: (str(v).replace("{" + key + "}", str(value)) if isinstance(v, str) else v)
                      for k, v in params.items()}
        with MyCohesityClient() as client:
            response = client.request(config.get("method", "GET"), url, params=params)
            try:
                payload = response.json()
            except ValueError:
                raise ValueError(
                    f"{url} returned {response.headers.get('content-type', 'an unknown type')} "
                    "rather than JSON. Set this endpoint's mode to 'page', or point it at the "
                    "JSON call recorded by mycohesity_discover_endpoints."
                )
        return {"mode": "api", "url": url,
                "data": _dig(payload, config.get("result_path") or "")}

    page = _in_thread(auth.render_page, url,
                      config.get("wait_selector") or "", wait_ms)
    return {"mode": "page", "url": url, "page": page}


def _article_links(links: list[dict], limit: int) -> list[dict]:
    out, seen = [], set()
    for link in links:
        url = (link.get("url") or "")
        title = (link.get("title") or "").strip()
        if not url or not title:
            continue
        if title.lower() in _NAV_NOISE or len(title) < 8:
            continue
        if not any(marker in url.lower() for marker in _ARTICLE_MARKERS):
            continue
        if url in seen:
            continue
        seen.add(url)
        out.append({"title": title, "url": url})
        if len(out) >= limit:
            break
    return out


def _tables_to_records(tables: list[list[list[str]]], limit: int) -> list[dict]:
    """Turn the widest scraped table into a list of header-keyed records."""
    if not tables:
        return []
    table = max(tables, key=lambda t: (len(t[0]) if t else 0, len(t)))
    header = [h or f"column_{i + 1}" for i, h in enumerate(table[0])]
    records = []
    for row in table[1:]:
        if not any(cell for cell in row):
            continue
        record = {header[i] if i < len(header) else f"column_{i + 1}": cell
                  for i, cell in enumerate(row)}
        records.append(record)
        if len(records) >= limit:
            break
    return records


# ---------------------------------------------------------------------------
# Authentication tools
# ---------------------------------------------------------------------------

@mcp.tool()
def mycohesity_login(mode: str = "external", timeout_seconds: int = 0) -> dict:
    """
    Sign in to MyCohesity. Opens a real browser window so you can complete SSO and MFA
    yourself, then stores the resulting session encrypted on this machine and reuses it.

    mode="external": customers and partners — starts at my.cohesity.com; MFA prompts
      appear inline.
    mode="internal": Cohesity employees — starts at the Cohesity Salesforce instance.
      Sign in there, then open the app launcher (icon stack, top left) and launch
      MyCohesity. Do not sign in to my.cohesity.com directly.

    A green bar appears at the bottom of the browser window; click it once the portal
    has loaded to capture the session. Run this once — later calls reuse the session
    until the portal rejects it.
    """
    timeout = timeout_seconds if timeout_seconds > 0 else auth.LOGIN_TIMEOUT
    try:
        return _in_thread(auth.interactive_login, mode, timeout,
                          auth.PORTAL_URL, auth.SALESFORCE_URL)
    except Exception as exc:
        return _error(f"Login failed: {exc}")


@mcp.tool()
def mycohesity_session_status(probe: bool = True) -> dict:
    """
    Report whether a stored MyCohesity session exists and, by default, whether it still
    works. Never returns cookie values. Set probe=false to skip the live network check.
    """
    summary = auth.session_summary()
    if not summary.get("authenticated") or not probe:
        return summary
    try:
        with MyCohesityClient() as client:
            summary["probe"] = client.probe()
    except SessionExpired as exc:
        summary["probe"] = {"valid": False, "detail": str(exc)}
    except Exception as exc:
        summary["probe"] = {"valid": False, "detail": f"Probe failed: {exc}"}
    if not summary["probe"].get("valid"):
        summary["hint"] = "Run mycohesity_login to re-authenticate."
    return summary


@mcp.tool()
def mycohesity_logout() -> dict:
    """
    Delete the stored MyCohesity session from this machine. The next request will
    require mycohesity_login again. Does not sign you out of the portal itself.
    """
    removed = auth.clear_session()
    return {
        "status": "OK",
        "detail": ("Stored session deleted." if removed else "No stored session to delete."),
        "storage_file": str(auth.SESSION_FILE),
    }


# ---------------------------------------------------------------------------
# Knowledge base
# ---------------------------------------------------------------------------

@mcp.tool()
def search_knowledge_base(query: str, product: str = "all", limit: int = 15) -> dict:
    """
    Search the MyCohesity knowledge base.

    query:   what to search for, e.g. "backup job fails with error 84"
    product: "dataprotect" (Cohesity DataProtect), "netbackup" (Veritas NetBackup),
             or "all" (default) to search both.
    limit:   maximum articles to return (default 15).

    Returns matching article titles and URLs. Use get_knowledge_base_article to read
    the full text of any result.
    """
    if not (query or "").strip():
        return _error("query is required.")

    key = (product or "all").strip().lower()
    if key not in _PRODUCT_TERMS:
        return _error(f"Unknown product '{product}'.",
                      "Use 'dataprotect', 'netbackup', or 'all'.")

    term = _PRODUCT_TERMS[key]
    effective = f"{query.strip()} {term}".strip() if term else query.strip()

    try:
        result = _run_endpoint("kb_search", wait_ms=5000,
                               query=effective, q=effective)
    except SessionExpired as exc:
        return _error(str(exc), "Run mycohesity_login.")
    except Exception as exc:
        return _error(f"Knowledge base search failed: {exc}")

    if result["mode"] == "api":
        records = result.get("data")
        return {
            "status": "OK", "query": effective, "product": key,
            "source": result["url"], "results": records if records is not None else [],
        }

    page = result["page"]
    articles = _article_links(page.get("links") or [], limit)
    payload = {
        "status": "OK",
        "query": effective,
        "product": key,
        "source": result["url"],
        "result_count": len(articles),
        "results": articles,
    }
    if not articles:
        payload["page_text_excerpt"] = (page.get("text") or "")[:2000]
        payload["hint"] = (
            "No article links were found on the rendered search page. The portal layout may "
            "differ for your tenant — run mycohesity_discover_endpoints while performing a "
            "search in the browser, then wire the real search call in with mycohesity_set_endpoint."
        )
    return payload


@mcp.tool()
def get_knowledge_base_article(article: str) -> dict:
    """
    Retrieve the full text of a MyCohesity knowledge base article.

    article: either a full article URL returned by search_knowledge_base, or a bare
             article id / slug.
    """
    reference = (article or "").strip()
    if not reference:
        return _error("article is required — pass a URL or an article id.")

    try:
        if reference.startswith(("http://", "https://")):
            page = _in_thread(auth.render_page, reference, "", 4000)
            source = reference
        else:
            result = _run_endpoint("kb_article", id=reference, query=reference)
            if result["mode"] == "api":
                return {"status": "OK", "article": reference,
                        "source": result["url"], "content": result.get("data")}
            page, source = result["page"], result["url"]
    except SessionExpired as exc:
        return _error(str(exc), "Run mycohesity_login.")
    except Exception as exc:
        return _error(f"Could not retrieve article: {exc}")

    return {
        "status": "OK",
        "source": source,
        "title": page.get("title"),
        "content": page.get("text"),
        "related_links": _article_links(page.get("links") or [], 20),
    }


# ---------------------------------------------------------------------------
# Assets and entitlements
# ---------------------------------------------------------------------------

@mcp.tool()
def list_assets(search: str = "", limit: int = 100) -> dict:
    """
    List the assets registered to your MyCohesity account — appliances, serial numbers,
    software instances and their support status.

    search: optional case-insensitive filter applied to every field of each record.
    limit:  maximum records to return (default 100).
    """
    return _records_tool("assets", search, limit, "assets")


@mcp.tool()
def get_entitlements(search: str = "", limit: int = 100) -> dict:
    """
    Retrieve licensing and support entitlement information for your MyCohesity account —
    licensed capacity, entitlement terms, start and end dates, and support levels.

    search: optional case-insensitive filter applied to every field of each record.
    limit:  maximum records to return (default 100).
    """
    return _records_tool("entitlements", search, limit, "entitlements")


def _records_tool(endpoint: str, search: str, limit: int, label: str) -> dict:
    try:
        result = _run_endpoint(endpoint, wait_ms=6000)
    except SessionExpired as exc:
        return _error(str(exc), "Run mycohesity_login.")
    except Exception as exc:
        return _error(f"Could not retrieve {label}: {exc}")

    if result["mode"] == "api":
        data = result.get("data")
        records = data if isinstance(data, list) else ([data] if data else [])
    else:
        records = _tables_to_records(result["page"].get("tables") or [], limit)

    needle = (search or "").strip().lower()
    if needle:
        records = [r for r in records
                   if needle in json.dumps(r, default=str).lower()]

    payload = {
        "status": "OK",
        "source": result["url"],
        "record_count": len(records[:limit]),
        label: records[:limit],
    }
    if not records:
        if result["mode"] == "page":
            payload["page_text_excerpt"] = (result["page"].get("text") or "")[:2000]
        payload["hint"] = (
            f"No {label} rows were extracted. The portal layout may differ for your tenant — "
            f"run mycohesity_discover_endpoints while opening the {label} page in the browser, "
            "then wire the real call in with mycohesity_set_endpoint."
        )
    return payload


# ---------------------------------------------------------------------------
# Escape hatches: generic fetch, discovery, configuration
# ---------------------------------------------------------------------------

@mcp.tool()
def mycohesity_fetch(path: str, method: str = "GET", params_json: str = "",
                     body_json: str = "", render: bool = False) -> dict:
    """
    Make an authenticated request to any MyCohesity URL using the stored session.
    Use this for portal pages or APIs that the dedicated tools do not cover.

    path:        absolute URL, or a path relative to the portal (e.g. "/s/mycases")
    method:      GET (default), POST, PUT, PATCH, DELETE — ignored when render=true
    params_json: optional JSON object of query-string parameters
    body_json:   optional JSON object request body
    render:      true to load the URL in a headless browser and return the rendered
                 text, links and tables. Needed for Lightning pages, which serve no
                 useful raw HTML. false (default) returns the raw HTTP response.
    """
    if not (path or "").strip():
        return _error("path is required.")

    try:
        params = json.loads(params_json) if params_json.strip() else None
        body = json.loads(body_json) if body_json.strip() else None
    except json.JSONDecodeError as exc:
        return _error(f"params_json / body_json must be valid JSON: {exc}")

    try:
        if render:
            with MyCohesityClient() as client:
                url = client.absolute(path)   # resolve against the signed-in portal
            page = _in_thread(auth.render_page, url, "", 4000)
            return {"status": "OK", "mode": "rendered", "url": page.get("url"),
                    "title": page.get("title"), "text": (page.get("text") or "")[:20000],
                    "links": (page.get("links") or [])[:100],
                    "tables": (page.get("tables") or [])[:5]}

        with MyCohesityClient() as client:
            response = client.request(method, path, params=params, json_body=body)
            content_type = response.headers.get("content-type", "")
            out: dict[str, Any] = {
                "status": "OK", "mode": "http", "http_status": response.status_code,
                "url": str(response.url), "content_type": content_type,
            }
            if "json" in content_type:
                out["data"] = response.json()
            else:
                out["text"] = response.text[:20000]
                out["hint"] = "Non-JSON response — try render=true for Lightning pages."
            return out
    except SessionExpired as exc:
        return _error(str(exc), "Run mycohesity_login.")
    except Exception as exc:
        return _error(f"Request failed: {exc}")


@mcp.tool()
def mycohesity_discover_endpoints(duration_seconds: int = 180,
                                  start_path: str = "/s/") -> dict:
    """
    Find the real API endpoints your MyCohesity tenant uses. Opens a browser on the
    portal with the stored session and records every XHR/fetch call while you click
    through knowledge base search, assets and entitlements. Click the green bar when
    you are done. Feed the URLs it reports into mycohesity_set_endpoint.

    Run this once if the built-in defaults do not return data for your tenant.
    """
    try:
        return _in_thread(auth.discover_endpoints, duration_seconds,
                          start_path, auth.PORTAL_URL)
    except Exception as exc:
        return _error(f"Discovery failed: {exc}")


@mcp.tool()
def mycohesity_set_endpoint(name: str, url: str, mode: str = "page",
                            method: str = "GET", result_path: str = "",
                            wait_selector: str = "") -> dict:
    """
    Point one of the tools at the correct URL for your tenant. Saved to
    ~/.mycohesity/endpoints.json and used from the next call onwards.

    name:         "kb_search", "kb_article", "assets" or "entitlements"
    url:          may contain {portal_url}, {salesforce_url}, {query} and {id}
                  placeholders, e.g. "{portal_url}/s/global-search/{query}"
    mode:         "page" to render the Lightning UI and scrape it, "api" for a JSON call
    method:       HTTP method for mode="api"
    result_path:  dotted path to the record list inside the JSON, e.g. "result.records"
    wait_selector: CSS selector to wait for before scraping, for mode="page"
    """
    known = ("kb_search", "kb_article", "assets", "entitlements")
    if name not in known:
        return _error(f"Unknown endpoint '{name}'.", f"Choose one of: {', '.join(known)}")
    if mode not in ("page", "api"):
        return _error("mode must be 'page' or 'api'.")
    if not (url or "").strip():
        return _error("url is required.")

    endpoints = auth.load_endpoints()
    endpoints[name] = {
        "mode": mode, "url": url.strip(), "method": method.upper(),
        "params": (endpoints.get(name) or {}).get("params", {}),
        "result_path": result_path, "wait_selector": wait_selector,
    }
    auth.save_endpoints(endpoints)
    return {"status": "OK", "endpoint": name, "config": endpoints[name],
            "saved_to": str(auth.ENDPOINTS_FILE)}


@mcp.tool()
def mycohesity_show_config() -> dict:
    """
    Show the current MyCohesity MCP configuration: portal and Salesforce URLs, the
    endpoint map in use, and where session data is stored. No secrets are returned.
    """
    return {
        "status": "OK",
        "version": __version__,
        "portal_url": auth.PORTAL_URL,
        "salesforce_url": auth.SALESFORCE_URL,
        "session_directory": str(auth.HOME_DIR),
        "session_file": str(auth.SESSION_FILE),
        "endpoints_file": str(auth.ENDPOINTS_FILE),
        "endpoints_file_exists": auth.ENDPOINTS_FILE.exists(),
        "discovery_file": str(auth.DISCOVERY_FILE),
        "key_storage": auth.key_storage_backend(),
        "login_timeout_seconds": auth.LOGIN_TIMEOUT,
        "endpoints": auth.load_endpoints(),
    }


# ---------------------------------------------------------------------------
# Entrypoint
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    transport = os.environ.get("MCP_TRANSPORT", "stdio")
    if transport == "http":
        mcp.settings.port = int(os.environ.get("MCP_HTTP_PORT", "8721"))
        mcp.run(transport="streamable-http")
    else:
        mcp.run()
