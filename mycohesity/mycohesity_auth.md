# mycohesity_auth.py

Browser-driven authentication and encrypted session storage for the MyCohesity MCP server. Imported by [`mycohesity_mcp.py`](mycohesity_mcp.md); it holds everything that touches credentials, the browser, or the disk.

Required packages:
```bash
pip install playwright cryptography httpx keyring
playwright install chromium
```

**Version:** 1.0

---

## Why a browser

my.cohesity.com is a Salesforce Experience Cloud portal, and neither supported sign-in path can be reproduced with an HTTP POST:

- **External users** (customers, partners) have MFA enforced — a script cannot supply the one-time code.
- **Internal Cohesity users** do not sign in to my.cohesity.com at all. They authenticate to the Cohesity Salesforce org and reach MyCohesity through the app launcher.

So the first connection opens a visible Chromium window, hands the user the keyboard for the whole SSO + MFA flow, and captures the resulting browser session. Everything after that is ordinary authenticated HTTP.

---

## What is stored

[Playwright storage state](https://playwright.dev/python/docs/auth) — cookies, localStorage and the browser User-Agent — inside an envelope recording the mode, portal URL and capture time.

| Path | Contents | Mode |
|---|---|---|
| `~/.mycohesity/` | Session directory | `0700` |
| `~/.mycohesity/session.enc` | Encrypted session envelope | `0600` |
| `~/.mycohesity/key.bin` | Encryption key — **only if no OS keychain exists** | `0600` |
| `~/.mycohesity/endpoints.json` | Endpoint map | `0600` |
| `~/.mycohesity/discovered.json` | Captured XHR calls | `0600` |

Override the directory with `MYCOHESITY_HOME`.

### Encryption

- **Fernet** — AES-128-CBC with an HMAC-SHA256 authentication tag.
- The key lives in the **OS keychain** via `keyring` (service `mycohesity-mcp`) — macOS Keychain, Windows Credential Manager, Linux Secret Service.
- Where no keychain backend is available, the key falls back to a `0600` file in the session directory. `key_storage_backend()` reports which is in use, and it is surfaced by `mycohesity_show_config`.
- A session file that cannot be decrypted — rotated key, corrupt file — is treated as *no session* rather than an error, so a rotated keychain entry does not wedge the server.
- **No cookie value is ever returned or logged.** `session_summary()` reports counts, domains and expiry dates only.

The stored session is a live credential for the user's MyCohesity account. It should be treated like a saved password.

---

## Session lifetime

Salesforce does not publish a session expiry, so validity is **probed rather than assumed**. Every request checks whether the portal answered with content or bounced to a sign-in page; if it bounced, `SessionExpired` is raised naming `mycohesity_login`. A stored session is reused indefinitely until that happens.

The captured User-Agent is replayed on every subsequent request, because Salesforce ties a session to the browser fingerprint that created it — sending httpx's own User-Agent is a reliable way to get the session invalidated server-side.

---

## API

### Session store

| Function | Description |
|---|---|
| `ensure_home()` | Creates the session directory with `0700` permissions |
| `save_session(storage_state, mode, user_agent, portal_url)` | Encrypts and writes the session envelope at `0600` |
| `load_session()` | Decrypts and returns the envelope, or `None` |
| `clear_session()` | Deletes the stored session; returns whether one existed |
| `session_summary()` | Non-sensitive description of the stored session |
| `key_storage_backend()` | Where the encryption key lives, as human-readable text |

### Endpoint map

| Function | Description |
|---|---|
| `load_endpoints()` | `DEFAULT_ENDPOINTS` merged with the user's `endpoints.json` |
| `save_endpoints(endpoints)` | Writes the map at `0600` |
| `resolve_url(template, endpoints, **kwargs)` | Expands `{portal_url}`, `{salesforce_url}` and caller placeholders. Caller values are percent-encoded, since they land inside a URL. |

### Authentication and retrieval

| Function | Description |
|---|---|
| `interactive_login(mode, timeout, portal_url, salesforce_url)` | Opens a visible browser, waits for sign-in, captures and encrypts the session |
| `MyCohesityClient` | Authenticated `httpx` client over the stored session — `request()`, `probe()`, `absolute()`, context-manager support |
| `render_page(url, wait_selector, wait_ms, timeout_ms)` | Loads a URL in headless Chromium with the stored session; returns title, text, links and tables |
| `discover_endpoints(duration, start_path, portal_url)` | Records every XHR/fetch call while the user browses; writes `discovered.json` |
| `SessionExpired` | Raised when the stored session no longer authenticates |

---

## How sign-in completion is detected

Two independent signals, either of which captures the session:

1. **The user says so.** An init script injects a fixed banner into the bottom of every top-level page carrying the mode-specific instructions and a green **"I'm signed in — capture session"** button. It reinstalls itself every 1.5 s, so it survives the DOM wipes Lightning performs on navigation, and it is added via `add_init_script` so page CSP cannot block it.

2. **Automatic detection.** The browser has settled on a portal page that is not a sign-in screen, and a Salesforce `sid` cookie is present, confirmed over three consecutive one-second checks. The consecutive-check requirement prevents capturing mid-redirect.

All tabs in the context are polled, because the internal app launcher opens MyCohesity in a new tab.

If neither signal arrives before the timeout, **nothing is saved** and the caller is told to try again.

A login screen is recognised by identity-provider hostname (`okta`, `microsoftonline`, `adfs`, `duosecurity`, `onelogin`, `pingidentity`, `auth0`) or by exact URL path *segment* (`login`, `signin`, `secur`, `idp`, `sso`, `auth`, `mfa`, `saml`, `oauth`, …). Matching on whole segments rather than substrings means an article slug such as `/s/article/authentication-guide` is not mistaken for a sign-in page, while a same-origin bounce to `/s/login/` — what Experience Cloud actually does when a session dies — is caught.

An existing session is loaded into the login browser before it starts, so a partially-valid session skips straight through instead of forcing a full re-authentication.

---

## Browser launch

`_launch_browser()` resolves the browser in this order:

1. `MYCOHESITY_CHROMIUM_PATH`, if set — for machines where Playwright's copy will not run
2. The user's installed Chrome (`channel="chrome"`), so corporate SSO extensions and certificates behave as they do in normal browsing
3. Playwright's bundled Chromium

`interactive_login` and `discover_endpoints` launch headed and need a visible desktop. `render_page` launches headless.

---

## Environment variables

| Variable | Default | Description |
|---|---|---|
| `MYCOHESITY_HOME` | `~/.mycohesity` | Session and configuration directory |
| `MYCOHESITY_PORTAL_URL` | `https://my.cohesity.com` | Portal to authenticate against |
| `MYCOHESITY_SALESFORCE_URL` | `https://cohesity.lightning.force.com` | Salesforce org for the internal flow |
| `MYCOHESITY_LOGIN_TIMEOUT` | `600` | Seconds allowed to complete SSO and MFA |
| `MYCOHESITY_HTTP_TIMEOUT` | `60` | Per-request timeout |
| `MYCOHESITY_CHROMIUM_PATH` | *(unset)* | Explicit browser executable |

---

## Usage outside the MCP server

```python
from mycohesity_auth import MyCohesityClient, render_page, SessionExpired

try:
    with MyCohesityClient() as client:
        response = client.request("GET", "/s/assets")
        print(response.status_code)

    page = render_page("https://my.cohesity.com/s/entitlements")
    print(page["tables"])
except SessionExpired as exc:
    print(exc)   # run the login tool, or call interactive_login() directly
```
