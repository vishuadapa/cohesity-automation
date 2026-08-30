# mycohesity_mcp.py

MCP server exposing the MyCohesity support portal (my.cohesity.com) to Claude Desktop. Searches the knowledge base for both **Cohesity DataProtect** and **Veritas NetBackup**, and retrieves **asset** and **licensing entitlement** information, using the signed-in user's own portal session.

Read-only: every tool searches or retrieves. Nothing in the portal is modified.

For step-by-step setup aimed at non-developers, see [README.md](README.md). This file is the reference.

Required packages:
```bash
pip install "mcp[cli]" playwright cryptography httpx keyring
playwright install chromium
```

**Version:** 1.0 — works with both `mcp` 1.x (`FastMCP`) and 2.x (`MCPServer`).

---

## Why authentication works the way it does

my.cohesity.com is a Salesforce Experience Cloud portal. There is no password grant to call:

- **External users** (customers, partners) have MFA enforced.
- **Internal Cohesity users** never sign in to my.cohesity.com directly — they authenticate to the Cohesity Salesforce org and launch MyCohesity from the app launcher (the icon stack, top left).

So `mycohesity_login` opens a real Chromium window and hands the keyboard to the user for the whole flow. A banner injected at the bottom of every page carries a **"I'm signed in — capture session"** button; clicking it is the signal to capture. Sessions are also captured automatically once the browser settles on a non-login portal page holding a Salesforce `sid` cookie, confirmed over three consecutive one-second checks.

The captured [Playwright storage state](https://playwright.dev/python/docs/auth) — cookies, localStorage and the browser User-Agent — is encrypted into `~/.mycohesity/session.enc` and reused for every later call. The User-Agent is replayed with each request because Salesforce ties a session to the browser fingerprint that created it.

Session validity is **probed, not assumed**: Salesforce does not publish an expiry, so each request checks whether the portal bounced it to a sign-in page and raises `SessionExpired` if so.

See [mycohesity_auth.md](mycohesity_auth.md) for the storage and encryption details.

---

## Running

```bash
# stdio — the transport Claude Desktop uses
python3 mycohesity_mcp.py

# streamable HTTP — for Copilot / VS Code
MCP_TRANSPORT=http MCP_HTTP_PORT=8721 python3 mycohesity_mcp.py
```

The server needs a **visible desktop** for `mycohesity_login` and `mycohesity_discover_endpoints`. Every other tool runs headless.

---

## Tools

### Authentication

| Tool | Arguments | Description |
|---|---|---|
| `mycohesity_login` | `mode` (`external` \| `internal`), `timeout_seconds` | Opens a browser for interactive sign-in, then stores the session encrypted. `external` starts at the portal; `internal` starts at Salesforce, from where the user launches MyCohesity via the app launcher. |
| `mycohesity_session_status` | `probe` (default `true`) | Reports whether a session is stored and, unless `probe=false`, whether it still works. Returns cookie counts, domains and expiry dates — **never cookie values**. |
| `mycohesity_logout` | — | Deletes the stored session from this machine. Does not sign the user out of the portal itself. |

### Knowledge base

| Tool | Arguments | Description |
|---|---|---|
| `search_knowledge_base` | `query`, `product` (`dataprotect` \| `netbackup` \| `all`, default `all`), `limit` (default 15) | Searches the knowledge base. `product` appends the matching product term to scope the search. Returns article titles and URLs. |
| `get_knowledge_base_article` | `article` | Retrieves an article's full text. Accepts either a URL from a search result or a bare article id / slug. Also returns related links. |

`product` accepts aliases: `dp` and `cohesity` map to DataProtect; `nbu` and `veritas` map to NetBackup; `both` maps to all.

### Assets and entitlements

| Tool | Arguments | Description |
|---|---|---|
| `list_assets` | `search`, `limit` (default 100) | Assets registered to the account — appliances, serial numbers, software instances, support status. |
| `get_entitlements` | `search`, `limit` (default 100) | Licensing and support entitlements — licensed capacity, terms, start and end dates, support levels. |

`search` is a case-insensitive substring filter applied across every field of each record, so `search="netbackup"` narrows a mixed estate to the NetBackup rows.

### Configuration and escape hatches

| Tool | Arguments | Description |
|---|---|---|
| `mycohesity_fetch` | `path`, `method`, `params_json`, `body_json`, `render` | Authenticated request to any portal URL. `render=true` loads it in a headless browser and returns rendered text, links and tables; `render=false` (default) returns the raw HTTP response. Relative paths resolve against the signed-in portal. |
| `mycohesity_discover_endpoints` | `duration_seconds` (default 180), `start_path` | Opens a browser with the stored session and records every XHR/fetch call while the user clicks through the portal. Writes `~/.mycohesity/discovered.json`. |
| `mycohesity_set_endpoint` | `name`, `url`, `mode`, `method`, `result_path`, `wait_selector` | Points `kb_search`, `kb_article`, `assets` or `entitlements` at the correct URL for this tenant. Persists to `~/.mycohesity/endpoints.json`. |
| `mycohesity_show_config` | — | Current portal and Salesforce URLs, the endpoint map, and storage locations. No secrets. |

---

## Page mode and API mode

Each of the four data endpoints runs in one of two modes.

| | `page` (default) | `api` |
|---|---|---|
| How it works | Renders the Lightning SPA in headless Chromium, then extracts links, tables and `role="grid"` data-grids from the DOM | Issues a JSON request over the stored session cookies |
| Speed | Seconds — a real browser starts up | Fast |
| When to use | Works out of the box; the portal serves no useful raw HTML, so this is the reliable default | Once `mycohesity_discover_endpoints` has revealed the tenant's real JSON endpoint |

Switching a tool to `api` mode is the single biggest speed-up available. Discover the endpoint, then:

```
Set the assets endpoint to https://my.cohesity.com/<discovered path> in api mode with result path "records"
```

### Endpoint map

Stored at `~/.mycohesity/endpoints.json`; see [endpoints.example.json](endpoints.example.json). Defaults are merged with the file, so partial overrides are fine.

| Field | Applies to | Meaning |
|---|---|---|
| `mode` | both | `page` or `api` |
| `url` | both | May contain `{portal_url}`, `{salesforce_url}`, `{query}` and `{id}`. `{query}` and `{id}` are percent-encoded on substitution. |
| `method` | `api` | HTTP method |
| `result_path` | `api` | Dotted path to the record list inside the response, e.g. `result.records` or `data.0.items` |
| `wait_selector` | `page` | CSS selector to wait for before scraping, for pages that populate late |

Defaults:

| Endpoint | Default URL | Mode |
|---|---|---|
| `kb_search` | `{portal_url}/s/global-search/{query}` | `page` |
| `kb_article` | `{portal_url}/s/article/{id}` | `page` |
| `assets` | `{portal_url}/s/assets` | `page` |
| `entitlements` | `{portal_url}/s/entitlements` | `page` |

Portal layouts vary per tenant. When a tool extracts no rows it returns a `hint` pointing at `mycohesity_discover_endpoints` rather than silently reporting "nothing found", along with a `page_text_excerpt` showing what it actually loaded.

---

## Environment variables

| Variable | Default | Description |
|---|---|---|
| `MYCOHESITY_PORTAL_URL` | `https://my.cohesity.com` | Portal to authenticate against |
| `MYCOHESITY_SALESFORCE_URL` | `https://cohesity.lightning.force.com` | Salesforce org for the internal flow |
| `MYCOHESITY_HOME` | `~/.mycohesity` | Session and configuration directory |
| `MYCOHESITY_LOGIN_TIMEOUT` | `600` | Seconds allowed to complete SSO and MFA |
| `MYCOHESITY_HTTP_TIMEOUT` | `60` | Per-request timeout |
| `MYCOHESITY_CHROMIUM_PATH` | *(unset)* | Explicit browser executable, for machines where Playwright's copy will not run |
| `MCP_TRANSPORT` | `stdio` | `stdio` or `http` |
| `MCP_HTTP_PORT` | `8721` | Port for `http` transport |

---

## Return shape

Every tool returns a dict. Failures return `{"status": "ERROR", "detail": ..., "hint": ...}` rather than raising, so Claude can explain what went wrong and what to do next.

All data tools **fail closed**: with no stored session, or with an expired one, they return an error naming `mycohesity_login` instead of an empty result that could be mistaken for "no matches".

```python
{
  "status": "OK",
  "query": "backup fails NetBackup",
  "product": "netbackup",
  "source": "https://my.cohesity.com/s/global-search/backup%20fails%20NetBackup",
  "result_count": 2,
  "results": [
    {"title": "NetBackup job fails with status 84", "url": "https://my.cohesity.com/s/article/..."}
  ]
}
```

---

## Implementation notes

- **Threading.** Playwright's sync API refuses to run on a thread owning a live asyncio loop, and the MCP server may call sync tools straight from the event loop thread. Every browser call is dispatched through a dedicated single-worker `ThreadPoolExecutor` so it is valid either way. A browser call blocks the server for its duration — expected for a single-user desktop tool, and the reason `mycohesity_login` is a one-time cost.
- **Login detection** matches identity-provider hostnames and exact path *segments* (`login`, `secur`, `sso`, `mfa`, …), so an article slug such as `/s/article/authentication-guide` is not mistaken for a sign-in page, while a same-origin bounce to `/s/login/` is caught.
- **Table extraction** picks the widest table on the page and keys rows by its header row, tolerating ragged rows and blank spacer rows. `role="grid"` Lightning data-grids are read the same way as `<table>`.
- **Search result filtering** keeps links whose URL looks like an article, drops portal chrome (Home, Contact Us, Log In and similar), and de-duplicates by URL — Lightning renders the same result more than once.

---

## Disclaimer

Personal project, not an official Cohesity tool. Not supported by Cohesity. Test before relying on it.
