"""
mycohesity_auth.py — browser-driven authentication and encrypted session storage
for the MyCohesity (my.cohesity.com) MCP server.

Why a browser?
  my.cohesity.com is a Salesforce Experience Cloud portal. External partner and
  customer accounts have MFA enforced, and internal Cohesity staff reach the
  portal through the Salesforce app launcher rather than by signing in to
  my.cohesity.com directly. Neither flow can be reproduced with a username /
  password POST, so the first connection opens a real Chromium window, hands the
  user the keyboard for the whole SSO + MFA dance, and then captures the
  resulting browser session.

What is stored
  Playwright storage state (cookies + localStorage + the browser User-Agent),
  encrypted with Fernet (AES-128-CBC + HMAC). The encryption key lives in the OS
  keychain via `keyring`; if no keychain backend is available it falls back to a
  0600 key file inside the session directory. Nothing is ever written to disk in
  plaintext, and cookie values are never returned by any MCP tool or logged.

Session lifetime
  Salesforce does not tell us when the session dies, so validity is probed: a
  cheap authenticated GET is issued and the response inspected for a login
  redirect. A stored session is reused until that probe fails, at which point the
  caller is told to run the login tool again.

Session directory (0700), override with MYCOHESITY_HOME:
  ~/.mycohesity/
    session.enc      encrypted storage state
    key.bin          fallback encryption key (only if keyring is unavailable)
    endpoints.json   endpoint map (see endpoints.example.json)
    discovered.json  network calls captured by the discovery tool

Env overrides: MYCOHESITY_HOME, MYCOHESITY_PORTAL_URL, MYCOHESITY_SALESFORCE_URL,
               MYCOHESITY_LOGIN_TIMEOUT, MYCOHESITY_HTTP_TIMEOUT, MYCOHESITY_CHROMIUM_PATH

Requires: pip install playwright cryptography httpx keyring
          playwright install chromium

Version: 1.0
"""

from __future__ import annotations

import json
import os
import re
import stat
import time
from pathlib import Path
from typing import Any, Optional
from urllib.parse import quote, urljoin, urlparse

import httpx

__version__ = "1.0"

# ---------------------------------------------------------------------------
# Paths and configuration
# ---------------------------------------------------------------------------

HOME_DIR = Path(os.environ.get("MYCOHESITY_HOME", Path.home() / ".mycohesity"))
SESSION_FILE = HOME_DIR / "session.enc"
KEY_FILE = HOME_DIR / "key.bin"
ENDPOINTS_FILE = HOME_DIR / "endpoints.json"
DISCOVERY_FILE = HOME_DIR / "discovered.json"

PORTAL_URL = os.environ.get("MYCOHESITY_PORTAL_URL", "https://my.cohesity.com").rstrip("/")
SALESFORCE_URL = os.environ.get(
    "MYCOHESITY_SALESFORCE_URL", "https://cohesity.lightning.force.com"
).rstrip("/")
LOGIN_TIMEOUT = int(os.environ.get("MYCOHESITY_LOGIN_TIMEOUT", "600"))
HTTP_TIMEOUT = int(os.environ.get("MYCOHESITY_HTTP_TIMEOUT", "60"))

_KEYRING_SERVICE = "mycohesity-mcp"
_KEYRING_USER = "session-encryption-key"

# Identity-provider hosts that mean "still mid-authentication"
_LOGIN_HOSTS = (
    "okta", "microsoftonline", "login.windows.net", "adfs", "duosecurity",
    "onelogin", "pingidentity", "auth0",
)

# Path *segments* that mean "still on a login / MFA / consent screen". Matched on
# exact segment boundaries so an article slug like /s/article/authentication-guide
# is not mistaken for a sign-in page.
_LOGIN_SEGMENTS = {
    "login", "signin", "sign-in", "secur", "idp", "sso", "auth", "authenticate",
    "verify", "mfa", "_nc_external", "saml", "oauth",
}


# ---------------------------------------------------------------------------
# Session directory and encryption key
# ---------------------------------------------------------------------------

def ensure_home() -> Path:
    """Create the session directory with owner-only permissions."""
    HOME_DIR.mkdir(parents=True, exist_ok=True)
    try:
        HOME_DIR.chmod(stat.S_IRWXU)  # 0700
    except OSError:
        pass  # Windows / exotic filesystems — ACLs already restrict to the user
    return HOME_DIR


def _keyring_available() -> bool:
    try:
        import keyring
        backend = keyring.get_keyring()
        name = backend.__class__.__name__.lower()
        return "fail" not in name  # keyring.backends.fail.Keyring means "no backend"
    except Exception:
        return False


def _load_or_create_key() -> bytes:
    """
    Fetch the Fernet key from the OS keychain, creating one on first use.
    Falls back to a 0600 key file when no keychain backend exists (common on
    headless Linux without a Secret Service daemon).
    """
    from cryptography.fernet import Fernet

    ensure_home()

    if _keyring_available():
        import keyring
        stored = keyring.get_password(_KEYRING_SERVICE, _KEYRING_USER)
        if stored:
            return stored.encode()
        key = Fernet.generate_key()
        keyring.set_password(_KEYRING_SERVICE, _KEYRING_USER, key.decode())
        return key

    if KEY_FILE.exists():
        return KEY_FILE.read_bytes().strip()

    key = Fernet.generate_key()
    KEY_FILE.write_bytes(key)
    try:
        KEY_FILE.chmod(stat.S_IRUSR | stat.S_IWUSR)  # 0600
    except OSError:
        pass
    return key


def key_storage_backend() -> str:
    """Human-readable description of where the encryption key lives."""
    if _keyring_available():
        try:
            import keyring
            return f"OS keychain ({keyring.get_keyring().__class__.__name__})"
        except Exception:
            return "OS keychain"
    return f"key file {KEY_FILE} (0600) — no OS keychain backend available"


# ---------------------------------------------------------------------------
# Encrypted session store
# ---------------------------------------------------------------------------

def save_session(storage_state: dict, mode: str, user_agent: str = "",
                 portal_url: str = PORTAL_URL) -> None:
    """Encrypt and persist a Playwright storage state."""
    from cryptography.fernet import Fernet

    ensure_home()
    envelope = {
        "version": __version__,
        "saved_at": time.time(),
        "mode": mode,
        "portal_url": portal_url,
        "user_agent": user_agent,
        "storage_state": storage_state,
    }
    token = Fernet(_load_or_create_key()).encrypt(json.dumps(envelope).encode())
    SESSION_FILE.write_bytes(token)
    try:
        SESSION_FILE.chmod(stat.S_IRUSR | stat.S_IWUSR)  # 0600
    except OSError:
        pass


def load_session() -> Optional[dict]:
    """Decrypt and return the stored session envelope, or None."""
    if not SESSION_FILE.exists():
        return None
    from cryptography.fernet import Fernet, InvalidToken
    try:
        raw = Fernet(_load_or_create_key()).decrypt(SESSION_FILE.read_bytes())
        return json.loads(raw.decode())
    except (InvalidToken, ValueError, json.JSONDecodeError):
        # Key rotated or file corrupt — treat as "no session" rather than crash.
        return None


def clear_session() -> bool:
    """Delete the stored session. Returns True if a session was removed."""
    existed = SESSION_FILE.exists()
    if existed:
        SESSION_FILE.unlink()
    return existed


def session_summary() -> dict:
    """Non-sensitive description of the stored session. Never returns cookie values."""
    env = load_session()
    if not env:
        return {"authenticated": False, "detail": "No stored session — run mycohesity_login first."}

    state = env.get("storage_state") or {}
    cookies = state.get("cookies") or []
    age = time.time() - env.get("saved_at", 0)
    domains = sorted({(c.get("domain") or "").lstrip(".") for c in cookies if c.get("domain")})
    soonest = [c.get("expires") for c in cookies
               if isinstance(c.get("expires"), (int, float)) and c["expires"] > 0]

    return {
        "authenticated": True,
        "mode": env.get("mode"),
        "portal_url": env.get("portal_url"),
        "saved_at": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(env.get("saved_at", 0))),
        "age_hours": round(age / 3600, 2),
        "cookie_count": len(cookies),
        "cookie_domains": domains,
        "earliest_cookie_expiry": (
            time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(min(soonest))) if soonest else "session cookies only"
        ),
        "storage_file": str(SESSION_FILE),
        "encryption": "Fernet (AES-128-CBC + HMAC-SHA256)",
        "key_storage": key_storage_backend(),
        "note": "Validity is probed live — call mycohesity_session_status to confirm the session still works.",
    }


# ---------------------------------------------------------------------------
# Endpoint map
# ---------------------------------------------------------------------------

DEFAULT_ENDPOINTS: dict[str, Any] = {
    "portal_url": PORTAL_URL,
    "salesforce_url": SALESFORCE_URL,
    "kb_search": {
        "mode": "page",
        "url": "{portal_url}/s/global-search/{query}",
        "method": "GET",
        "params": {},
        "result_path": "",
        "wait_selector": "",
    },
    "kb_article": {
        "mode": "page",
        "url": "{portal_url}/s/article/{id}",
        "method": "GET",
        "params": {},
        "result_path": "",
        "wait_selector": "",
    },
    "assets": {
        "mode": "page",
        "url": "{portal_url}/s/assets",
        "method": "GET",
        "params": {},
        "result_path": "",
        "wait_selector": "",
    },
    "entitlements": {
        "mode": "page",
        "url": "{portal_url}/s/entitlements",
        "method": "GET",
        "params": {},
        "result_path": "",
        "wait_selector": "",
    },
}


def load_endpoints() -> dict:
    """Endpoint map, defaults merged with the user's ~/.mycohesity/endpoints.json."""
    endpoints = json.loads(json.dumps(DEFAULT_ENDPOINTS))  # deep copy
    if ENDPOINTS_FILE.exists():
        try:
            user = json.loads(ENDPOINTS_FILE.read_text())
        except json.JSONDecodeError:
            return endpoints
        for key, value in user.items():
            if isinstance(value, dict) and isinstance(endpoints.get(key), dict):
                endpoints[key].update(value)
            else:
                endpoints[key] = value
    return endpoints


def save_endpoints(endpoints: dict) -> None:
    ensure_home()
    ENDPOINTS_FILE.write_text(json.dumps(endpoints, indent=2))
    try:
        ENDPOINTS_FILE.chmod(stat.S_IRUSR | stat.S_IWUSR)
    except OSError:
        pass


def resolve_url(template: str, endpoints: dict, **kwargs) -> str:
    """Expand {portal_url}, {salesforce_url} and caller-supplied placeholders."""
    values = {
        "portal_url": (endpoints.get("portal_url") or PORTAL_URL).rstrip("/"),
        "salesforce_url": (endpoints.get("salesforce_url") or SALESFORCE_URL).rstrip("/"),
    }
    # Caller-supplied values land inside a URL, so they are percent-encoded.
    values.update({k: quote(str(v), safe="") for k, v in kwargs.items() if v is not None})
    out = template
    for key, value in values.items():
        out = out.replace("{" + key + "}", str(value))
    return out


# ---------------------------------------------------------------------------
# Interactive browser login
# ---------------------------------------------------------------------------

# Injected into every page/frame before site scripts run. Adds a fixed banner
# with a "Finish" button so the user can tell the MCP server that login is
# complete, which is far more reliable than guessing from URLs alone.
_BANNER_SCRIPT = r"""
(() => {
  if (window.__mycoBannerInstalled) return;
  window.__mycoBannerInstalled = true;
  window.__MYCO_LOGIN_DONE = false;

  const install = () => {
    try {
      if (window.top !== window) return;              // top frame only
      if (!document.body) return;
      if (document.getElementById('__myco_bar')) return;

      const bar = document.createElement('div');
      bar.id = '__myco_bar';
      bar.style.cssText = [
        'position:fixed', 'z-index:2147483647', 'left:0', 'right:0', 'bottom:0',
        'padding:10px 16px', 'background:#0b1e3a', 'color:#fff',
        'font:13px/1.4 -apple-system,Segoe UI,Roboto,sans-serif',
        'display:flex', 'align-items:center', 'gap:14px',
        'box-shadow:0 -2px 10px rgba(0,0,0,.35)'
      ].join(';');

      const text = document.createElement('span');
      text.style.cssText = 'flex:1';
      text.textContent = window.__MYCO_HINT ||
        'Claude MyCohesity MCP: complete sign-in (and MFA) in this window.';

      const btn = document.createElement('button');
      btn.textContent = "I'm signed in — capture session";
      btn.style.cssText = [
        'cursor:pointer', 'border:0', 'border-radius:5px', 'padding:8px 14px',
        'background:#2ecc71', 'color:#06331b', 'font-weight:700', 'font-size:13px'
      ].join(';');
      btn.onclick = () => {
        window.__MYCO_LOGIN_DONE = true;
        text.textContent = 'Capturing session… this window will close on its own.';
        btn.disabled = true;
        btn.style.background = '#888';
      };

      bar.appendChild(text);
      bar.appendChild(btn);
      document.body.appendChild(bar);
    } catch (e) { /* never break the host page */ }
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', install);
  } else {
    install();
  }
  setInterval(install, 1500);   // survives SPA re-renders (Lightning wipes the DOM)
})();
"""

_INTERNAL_HINT = (
    "Claude MyCohesity MCP — INTERNAL flow: sign in to Salesforce, then open the "
    "app launcher (icon stack, top left) and launch MyCohesity. Click the green "
    "button once the MyCohesity portal has loaded."
)
_EXTERNAL_HINT = (
    "Claude MyCohesity MCP — EXTERNAL flow: sign in to my.cohesity.com and complete "
    "MFA. Click the green button once the portal has loaded."
)


def _launch_browser(pw, headless: bool):
    """
    Launch Chromium, preferring the user's installed Chrome so corporate SSO
    extensions and certificates behave as they do in normal browsing.

    MYCOHESITY_CHROMIUM_PATH overrides everything — set it on machines where the
    browser lives somewhere Playwright does not look.
    """
    explicit = os.environ.get("MYCOHESITY_CHROMIUM_PATH", "").strip()
    if explicit:
        return pw.chromium.launch(headless=headless, executable_path=explicit)
    try:
        return pw.chromium.launch(headless=headless, channel="chrome")
    except Exception:
        return pw.chromium.launch(headless=headless)


def _looks_like_login(url: str) -> bool:
    """True when a URL is an identity-provider host or has a sign-in path segment."""
    try:
        parts = urlparse(url.lower())
    except ValueError:
        return False
    if any(host in parts.netloc for host in _LOGIN_HOSTS):
        return True
    segments = {seg for seg in parts.path.split("/") if seg}
    return bool(segments & _LOGIN_SEGMENTS)


def _on_portal(url: str, portal_url: str) -> bool:
    try:
        return urlparse(url).netloc.lower() == urlparse(portal_url).netloc.lower()
    except ValueError:
        return False


def interactive_login(mode: str = "external", timeout: int = LOGIN_TIMEOUT,
                      portal_url: str = PORTAL_URL,
                      salesforce_url: str = SALESFORCE_URL) -> dict:
    """
    Open a visible Chromium window, let the user complete SSO/MFA, then capture
    and encrypt the resulting session.

    mode="external"  → start at my.cohesity.com (MFA prompts appear inline)
    mode="internal"  → start at the Salesforce instance; the user then uses the
                       app launcher to reach MyCohesity without a second login
    """
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        return {
            "status": "ERROR",
            "detail": "Playwright is not installed. Run: pip install playwright && playwright install chromium",
        }

    mode = (mode or "external").strip().lower()
    if mode not in ("external", "internal"):
        return {"status": "ERROR", "detail": "mode must be 'external' or 'internal'"}

    start_url = salesforce_url if mode == "internal" else portal_url
    hint = _INTERNAL_HINT if mode == "internal" else _EXTERNAL_HINT

    with sync_playwright() as pw:
        try:
            browser = _launch_browser(pw, headless=False)
        except Exception as exc:
            return {
                "status": "ERROR",
                "detail": f"Could not launch a browser ({exc}). Run: playwright install chromium",
            }

        # Reuse any existing session so a partially-valid login skips straight through.
        existing = load_session()
        context_args: dict[str, Any] = {"viewport": {"width": 1440, "height": 900}}
        if existing and existing.get("storage_state"):
            context_args["storage_state"] = existing["storage_state"]

        context = browser.new_context(**context_args)
        context.add_init_script(f"window.__MYCO_HINT = {json.dumps(hint)};")
        context.add_init_script(_BANNER_SCRIPT)

        page = context.new_page()
        try:
            page.goto(start_url, wait_until="domcontentloaded", timeout=60_000)
        except Exception as exc:
            context.close()
            browser.close()
            return {"status": "ERROR", "detail": f"Could not open {start_url}: {exc}"}

        deadline = time.time() + max(60, timeout)
        stable_hits = 0
        captured = False
        final_url = start_url

        while time.time() < deadline:
            pages = [p for p in context.pages if not p.is_closed()]
            if not pages:
                break

            # 1. Explicit user confirmation via the injected banner button.
            for candidate in pages:
                try:
                    if candidate.evaluate("() => window.__MYCO_LOGIN_DONE === true"):
                        final_url = candidate.url
                        captured = True
                        break
                except Exception:
                    continue        # mid-navigation; try again next tick
            if captured:
                break

            # 2. Automatic detection: a portal page that is not a login screen,
            #    backed by a session cookie, seen on three consecutive checks.
            portal_pages = [
                p for p in pages
                if _on_portal(_safe_url(p), portal_url) and not _looks_like_login(_safe_url(p))
            ]
            if portal_pages and _has_session_cookie(context, portal_url):
                stable_hits += 1
                if stable_hits >= 3:
                    final_url = _safe_url(portal_pages[-1])
                    captured = True
                    break
            else:
                stable_hits = 0

            time.sleep(1)

        if not captured:
            context.close()
            browser.close()
            return {
                "status": "TIMEOUT",
                "detail": (
                    f"Login was not completed within {timeout}s. Nothing was saved. "
                    "Re-run mycohesity_login, or raise MYCOHESITY_LOGIN_TIMEOUT."
                ),
            }

        try:
            user_agent = page.evaluate("() => navigator.userAgent") or ""
        except Exception:
            user_agent = ""

        storage_state = context.storage_state()
        context.close()
        browser.close()

    cookies = storage_state.get("cookies") or []
    if not cookies:
        return {"status": "ERROR", "detail": "No cookies were captured — the session was not saved."}

    save_session(storage_state, mode=mode, user_agent=user_agent, portal_url=portal_url)

    return {
        "status": "OK",
        "mode": mode,
        "landed_on": final_url,
        "cookie_count": len(cookies),
        "cookie_domains": sorted({(c.get("domain") or "").lstrip(".") for c in cookies if c.get("domain")}),
        "stored_at": str(SESSION_FILE),
        "encryption": "Fernet (AES-128-CBC + HMAC-SHA256)",
        "key_storage": key_storage_backend(),
        "detail": "Session captured and encrypted. It will be reused until it stops working.",
    }


def _safe_url(page) -> str:
    try:
        return page.url or ""
    except Exception:
        return ""


def _has_session_cookie(context, portal_url: str) -> bool:
    """Salesforce sets `sid` (and friends) once authentication actually succeeds."""
    try:
        cookies = context.cookies(portal_url)
    except Exception:
        return False
    return any("sid" in (c.get("name") or "").lower() and c.get("value") for c in cookies)


# ---------------------------------------------------------------------------
# Authenticated HTTP client
# ---------------------------------------------------------------------------

class SessionExpired(RuntimeError):
    """Raised when the stored session no longer authenticates."""


class MyCohesityClient:
    """
    Thin authenticated HTTP client over the captured browser session.

    The stored User-Agent is replayed alongside the cookies: Salesforce ties a
    session to the browser fingerprint that created it, so sending httpx's own
    User-Agent is a reliable way to get the session invalidated server-side.
    """

    def __init__(self) -> None:
        envelope = load_session()
        if not envelope:
            raise SessionExpired("No stored session — run mycohesity_login first.")

        self.mode = envelope.get("mode", "external")
        self.portal_url = (envelope.get("portal_url") or PORTAL_URL).rstrip("/")
        self.user_agent = envelope.get("user_agent") or ""
        self.storage_state = envelope.get("storage_state") or {}

        headers = {
            "Accept": "application/json, text/plain, */*",
            "Accept-Language": "en-US,en;q=0.9",
        }
        if self.user_agent:
            headers["User-Agent"] = self.user_agent

        self._client = httpx.Client(
            headers=headers,
            timeout=HTTP_TIMEOUT,
            follow_redirects=True,
        )
        for cookie in self.storage_state.get("cookies") or []:
            name, value = cookie.get("name"), cookie.get("value")
            if not name or value is None:
                continue
            try:
                self._client.cookies.set(
                    name, value,
                    domain=(cookie.get("domain") or "").lstrip("."),
                    path=cookie.get("path") or "/",
                )
            except Exception:
                continue

    def close(self) -> None:
        self._client.close()

    def __enter__(self) -> "MyCohesityClient":
        return self

    def __exit__(self, *exc_info) -> None:
        self.close()

    def absolute(self, path_or_url: str) -> str:
        if path_or_url.startswith(("http://", "https://")):
            return path_or_url
        return urljoin(self.portal_url + "/", path_or_url.lstrip("/"))

    def request(self, method: str, path_or_url: str, params: dict | None = None,
                json_body: dict | None = None, headers: dict | None = None) -> httpx.Response:
        url = self.absolute(path_or_url)
        response = self._client.request(
            method.upper(), url, params=params or None,
            json=json_body if json_body else None, headers=headers or None,
        )
        if _is_login_response(response):
            raise SessionExpired(
                f"The stored session was rejected by {url} (HTTP {response.status_code}, "
                "redirected to a login screen). Run mycohesity_login again."
            )
        return response

    def probe(self) -> dict:
        """Cheap authenticated GET used to decide whether the session is still alive."""
        try:
            response = self.request("GET", self.portal_url + "/s/")
        except SessionExpired as exc:
            return {"valid": False, "detail": str(exc)}
        except httpx.HTTPError as exc:
            return {"valid": False, "detail": f"Network error contacting the portal: {exc}"}
        return {
            "valid": True,
            "http_status": response.status_code,
            "final_url": str(response.url),
        }


def _is_login_response(response: httpx.Response) -> bool:
    """True when the portal bounced us to a sign-in page instead of serving content."""
    if response.status_code in (401, 403):
        return True
    if _looks_like_login(str(response.url)):
        return True
    content_type = response.headers.get("content-type", "")
    if "html" in content_type:
        head = response.text[:4000].lower()
        if ("<form" in head and "password" in head) or "loginform" in head:
            return True
    return False


# ---------------------------------------------------------------------------
# Headless rendering (for the Lightning SPA, which serves no useful raw HTML)
# ---------------------------------------------------------------------------

# Pulls anchors, tables and Lightning data-grids out of a rendered page.
_EXTRACT_JS = r"""
() => {
  const clean = (s) => (s || '').replace(/\s+/g, ' ').trim();

  const links = [];
  const seen = new Set();
  document.querySelectorAll('a[href]').forEach((a) => {
    const href = a.href;
    const title = clean(a.innerText || a.getAttribute('title'));
    if (!href || !title || title.length < 3) return;
    if (seen.has(href)) return;
    seen.add(href);
    links.push({ title: title.slice(0, 300), url: href });
  });

  const tables = [];
  document.querySelectorAll('table, [role="grid"]').forEach((tbl) => {
    const rows = [];
    tbl.querySelectorAll('tr, [role="row"]').forEach((tr) => {
      const cells = [];
      tr.querySelectorAll('th, td, [role="gridcell"], [role="columnheader"]').forEach((td) => {
        cells.push(clean(td.innerText).slice(0, 400));
      });
      if (cells.some((c) => c)) rows.push(cells);
    });
    if (rows.length > 1) tables.push(rows.slice(0, 500));
  });

  const h1 = document.querySelector('h1');
  return {
    url: location.href,
    has_password_field: !!document.querySelector('input[type="password"]'),
    title: document.title || clean(h1 ? h1.innerText : '').slice(0, 200),
    text: clean(document.body ? document.body.innerText : '').slice(0, 60000),
    links: links.slice(0, 400),
    tables: tables.slice(0, 25),
  };
}
"""


def render_page(url: str, wait_selector: str = "", wait_ms: int = 4000,
                timeout_ms: int = 60_000) -> dict:
    """
    Load `url` in headless Chromium using the stored session and return its
    rendered title, text, links and tables. Raises SessionExpired if the portal
    redirects to a login screen.
    """
    envelope = load_session()
    if not envelope:
        raise SessionExpired("No stored session — run mycohesity_login first.")

    try:
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise RuntimeError(
            "Playwright is not installed. Run: pip install playwright && playwright install chromium"
        ) from exc

    with sync_playwright() as pw:
        browser = _launch_browser(pw, headless=True)

        context_args: dict[str, Any] = {
            "storage_state": envelope.get("storage_state") or {},
            "viewport": {"width": 1600, "height": 1200},
        }
        if envelope.get("user_agent"):
            context_args["user_agent"] = envelope["user_agent"]

        context = browser.new_context(**context_args)
        page = context.new_page()
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=timeout_ms)
            if wait_selector:
                try:
                    page.wait_for_selector(wait_selector, timeout=timeout_ms)
                except Exception:
                    pass                       # selector never appeared — return what we have
            else:
                try:
                    page.wait_for_load_state("networkidle", timeout=15_000)
                except Exception:
                    pass                       # Lightning polls forever; the sleep below covers it
            page.wait_for_timeout(max(0, wait_ms))
            result = page.evaluate(_EXTRACT_JS)
        finally:
            context.close()
            browser.close()

    landed = result.get("url") or ""
    if _looks_like_login(landed):
        raise SessionExpired(
            f"Rendering {url} landed on a login page ({landed}). Run mycohesity_login again."
        )
    # Some tenants render a sign-in screen without changing the URL.
    if result.get("has_password_field") and len(result.get("text") or "") < 2000:
        raise SessionExpired(
            f"Rendering {url} returned a sign-in screen. Run mycohesity_login again."
        )
    return result


# ---------------------------------------------------------------------------
# Endpoint discovery
# ---------------------------------------------------------------------------

_DISCOVERY_IGNORE = re.compile(
    r"\.(png|jpe?g|gif|svg|woff2?|ttf|eot|ico|css|map)(\?|$)", re.I
)


def discover_endpoints(duration: int = 120, start_path: str = "/s/",
                       portal_url: str = PORTAL_URL) -> dict:
    """
    Open a visible browser on the portal and record every XHR/fetch call the
    Lightning app makes while the user clicks around. Use this to find the real
    knowledge-base / assets / entitlements endpoints for this tenant, then wire
    them in with mycohesity_set_endpoint.
    """
    envelope = load_session()
    if not envelope:
        return {"status": "ERROR", "detail": "No stored session — run mycohesity_login first."}

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        return {"status": "ERROR", "detail": "Playwright is not installed."}

    calls: list[dict] = []
    seen: set[tuple[str, str]] = set()

    def on_response(response) -> None:
        try:
            request = response.request
            if request.resource_type not in ("xhr", "fetch"):
                return
            url = request.url
            if _DISCOVERY_IGNORE.search(url):
                return
            key = (request.method, url.split("?")[0])
            if key in seen:
                return
            seen.add(key)
            entry = {
                "method": request.method,
                "url": url,
                "status": response.status,
                "content_type": response.headers.get("content-type", ""),
            }
            if "json" in entry["content_type"]:
                try:
                    entry["response_preview"] = response.text()[:1500]
                except Exception:
                    entry["response_preview"] = "<body unavailable>"
            calls.append(entry)
        except Exception:
            pass

    hint = (
        "Claude MyCohesity MCP — DISCOVERY: browse to knowledge-base search, assets and "
        "entitlements. Every API call is being recorded. Click the green button when done."
    )

    with sync_playwright() as pw:
        browser = _launch_browser(pw, headless=False)

        context_args: dict[str, Any] = {
            "storage_state": envelope.get("storage_state") or {},
            "viewport": {"width": 1440, "height": 900},
        }
        if envelope.get("user_agent"):
            context_args["user_agent"] = envelope["user_agent"]

        context = browser.new_context(**context_args)
        context.add_init_script(f"window.__MYCO_HINT = {json.dumps(hint)};")
        context.add_init_script(_BANNER_SCRIPT)
        context.on("response", on_response)

        page = context.new_page()
        target = start_path if start_path.startswith("http") else portal_url.rstrip("/") + "/" + start_path.lstrip("/")
        try:
            page.goto(target, wait_until="domcontentloaded", timeout=60_000)
        except Exception as exc:
            context.close()
            browser.close()
            return {"status": "ERROR", "detail": f"Could not open {target}: {exc}"}

        deadline = time.time() + max(30, duration)
        while time.time() < deadline:
            done = False
            for candidate in [p for p in context.pages if not p.is_closed()]:
                try:
                    if candidate.evaluate("() => window.__MYCO_LOGIN_DONE === true"):
                        done = True
                        break
                except Exception:
                    continue
            if done:
                break
            time.sleep(1)

        # Refresh the stored session — browsing usually extends it.
        try:
            save_session(context.storage_state(), mode=envelope.get("mode", "external"),
                         user_agent=envelope.get("user_agent", ""), portal_url=portal_url)
        except Exception:
            pass

        context.close()
        browser.close()

    ensure_home()
    DISCOVERY_FILE.write_text(json.dumps(calls, indent=2))
    try:
        DISCOVERY_FILE.chmod(stat.S_IRUSR | stat.S_IWUSR)
    except OSError:
        pass

    return {
        "status": "OK",
        "captured": len(calls),
        "saved_to": str(DISCOVERY_FILE),
        "calls": calls[:60],
        "detail": "Pick the calls that returned the data you want, then wire them in with mycohesity_set_endpoint.",
    }
