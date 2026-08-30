# MyCohesity MCP — Setup Guide

**Search the MyCohesity knowledge base and check your licensing entitlements from Claude Desktop.**

This guide walks you through every step, even if you have never heard of MCP before. By the end you will be able to open Claude Desktop and type things like:

> *"Search MyCohesity for NetBackup status 84 errors"*
> *"What DataProtect articles cover upgrade readiness?"*
> *"Show me my registered assets and when their support expires"*
> *"How much licensed capacity do we have left, and when does the entitlement end?"*

---

## What Is MCP and Why Does It Matter?

MCP stands for **Model Context Protocol**. Think of it as a power cable that connects Claude to external systems — in this case the MyCohesity support portal. Without MCP, Claude can only answer from what it already knows. With MCP, Claude can search the real knowledge base and read your real entitlement data.

The files in this folder are that power cable — a small Python program that sits between Claude and my.cohesity.com.

---

## How Sign-In Works (Read This First)

MyCohesity is a **Salesforce Experience Cloud** portal. There is no API key you can paste into a config file, and there is no username/password you can hand to a script:

- **External users** (customers and partners) have **MFA enforced**. A script cannot type your one-time code for you.
- **Internal Cohesity users** never sign in to my.cohesity.com directly at all. They authenticate to the Cohesity **Salesforce** org, then launch MyCohesity from the **app launcher** — the icon stack in the top-left corner.

So this MCP does the only thing that actually works: **the first time you use it, it opens a real browser window and hands you the keyboard.** You sign in exactly as you normally would — SSO, MFA, app launcher, all of it. When the portal has loaded, you click a green button and the MCP captures the browser session.

That session is then **encrypted and stored on your own machine** and reused for every later question. You will not be asked to sign in again until the portal itself expires the session — typically after a period of inactivity, or when your organisation's session policy says so.

```
First run                                    Every run after that
─────────                                    ────────────────────
Claude → browser opens                       Claude → reads the stored session
       → you sign in + MFA                          → queries the portal
       → you click "I'm signed in"                  → answers you
       → session encrypted to disk
```

---

## Before You Start — What You Need

| Requirement | How to check |
|---|---|
| **Python 3.10 or newer** | Open Terminal / Command Prompt and run: `python3 --version` |
| **Claude Desktop app** | Download from [claude.ai/download](https://claude.ai/download) |
| **A MyCohesity account** | Either an external portal account, or Cohesity Salesforce access |
| **A desktop session you can see** | The sign-in browser window has to appear on a real screen |

> **Windows users:** all commands below work the same in PowerShell. Replace `python3` with `python` if needed.

---

## Step 1 — Download This Folder

If you are not a developer, the easiest way to get the files is:

1. Go to the GitHub repository page.
2. Click the green **Code** button → **Download ZIP**.
3. Unzip it somewhere easy to find, like Desktop or Documents.

The file you need is:

```
cohesity-automation/mycohesity/mycohesity_mcp.py
```

Note the **full path** — you will need it in Step 4. For example:

- Mac/Linux: `/Users/yourname/Downloads/cohesity-automation/mycohesity/mycohesity_mcp.py`
- Windows: `C:\Users\yourname\Downloads\cohesity-automation\mycohesity\mycohesity_mcp.py`

---

## Step 2 — Install the Required Python Packages

Open **Terminal** (Mac/Linux) or **Command Prompt / PowerShell** (Windows) and run these two commands:

```bash
pip install "mcp[cli]" playwright cryptography httpx keyring
playwright install chromium
```

The second command downloads a private copy of the Chromium browser that the sign-in window uses. Together they take a minute or two. You only need to do this once.

> **Why a separate browser?** So the MCP can watch for the moment sign-in completes and capture the session, without touching your everyday browser profile or its cookies.

---

## Step 3 — Find Your Claude Desktop Config File

Claude Desktop uses a configuration file to know which MCP servers to load at startup. It is called `claude_desktop_config.json`.

| Operating System | Location |
|---|---|
| **Mac** | `~/Library/Application Support/Claude/claude_desktop_config.json` |
| **Windows** | `%APPDATA%\Claude\claude_desktop_config.json` |

**Tip — open the config directly from Claude Desktop:**

1. Open Claude Desktop.
2. Click the **Claude** menu (Mac) or the hamburger menu (Windows).
3. Click **Settings → Developer → Edit Config**.

If the file does not exist yet, create it — Claude Desktop picks it up on the next launch.

---

## Step 4 — Add the MyCohesity MCP Server to the Config

Open `claude_desktop_config.json` in any text editor and paste the block below. If the file already has content, add the `mycohesity` entry inside the existing `mcpServers` object.

### Mac / Linux

```json
{
  "mcpServers": {
    "mycohesity": {
      "command": "python3",
      "args": ["/full/path/to/mycohesity_mcp.py"]
    }
  }
}
```

### Windows

```json
{
  "mcpServers": {
    "mycohesity": {
      "command": "python",
      "args": ["C:\\full\\path\\to\\mycohesity_mcp.py"]
    }
  }
}
```

**There are no credentials in this file.** Nothing to paste, nothing to leak. Sign-in happens in the browser in Step 6.

If you already have the cluster MCP configured, the two live side by side:

```json
{
  "mcpServers": {
    "cohesity":   { "command": "python3", "args": ["/path/to/mcp/cohesity_mcp_v2.py"],
                    "env": { "COHESITY_CLUSTER": "...", "COHESITY_API_KEY": "..." } },
    "mycohesity": { "command": "python3", "args": ["/path/to/mycohesity/mycohesity_mcp.py"] }
  }
}
```

---

## Step 5 — Restart Claude Desktop

Quit Claude Desktop **completely** (Mac: `Cmd+Q` — closing the window is not enough) and reopen it. You should see a tools or plug icon in the message box. Click it and you should see the MyCohesity tools listed.

---

## Step 6 — Sign In (Once)

In Claude Desktop, type one of these:

**If you are a customer or partner (external user):**

> Log in to MyCohesity

**If you work at Cohesity (internal user):**

> Log in to MyCohesity using the internal Salesforce flow

Claude will ask your permission to run the tool. Approve it, and a browser window opens.

### What to do in that window

| You are | What you will see | What to do |
|---|---|---|
| **External** | The my.cohesity.com sign-in page | Sign in and complete MFA |
| **Internal** | The Cohesity Salesforce sign-in page | Sign in, then click the **app launcher** (the icon stack, top left) and choose **MyCohesity**. Do **not** try to sign in to my.cohesity.com directly. |

A **dark blue bar** sits at the bottom of the window the whole time, reminding you which flow you are in. Once the MyCohesity portal has finished loading, click the green **"I'm signed in — capture session"** button. The window closes on its own and Claude confirms the session was stored.

> Take as long as you need — you have 10 minutes by default. Nothing is saved unless you finish.

You should not have to do this again for a long time.

---

## Step 7 — Ask Questions

```
Search MyCohesity for "backup fails with status 84" in NetBackup
```
```
Find DataProtect articles about upgrade readiness and summarise the top three
```
```
List my registered assets — which ones have support ending in the next 6 months?
```
```
What are my current licensing entitlements? Show licensed capacity and end dates.
```
```
Read the full article at <url> and tell me the workaround
```

Claude picks the right tool automatically. To check on the connection at any time:

```
What's my MyCohesity session status?
```

---

## What Gets Stored, and Where

Everything lives in a single folder on your own machine, created with owner-only permissions (`0700`):

| Path | What it is |
|---|---|
| `~/.mycohesity/session.enc` | Your portal session — **encrypted**, file mode `0600` |
| `~/.mycohesity/key.bin` | Encryption key — **only created if your OS has no keychain** |
| `~/.mycohesity/endpoints.json` | Optional URL overrides for your tenant |
| `~/.mycohesity/discovered.json` | Output of the endpoint discovery tool |

**How the session is protected**

- Encrypted with **Fernet** (AES-128-CBC with an HMAC-SHA256 authentication tag).
- The encryption key is held in your **OS keychain** — macOS Keychain, Windows Credential Manager, or Linux Secret Service. If no keychain backend exists, the key falls back to a `0600` file in the same folder.
- **No tool ever returns a cookie value.** `mycohesity_session_status` reports counts, domains and expiry dates only.
- Nothing is sent anywhere except my.cohesity.com. There is no telemetry and no third-party service.

To wipe the stored session at any time, ask Claude to **log out of MyCohesity**, or just delete the folder.

> **Note:** the stored session is a live credential for your MyCohesity account. Treat `~/.mycohesity/` like you would a saved password — do not copy it to another machine or into a shared folder.

---

## If Results Come Back Empty

The tools ship with sensible default URLs, but Salesforce portals are configured per tenant — your assets page may not live at the same address as someone else's. If a tool returns no rows, it tells you so and suggests the fix. Here it is in full:

1. Ask Claude:
   > Discover the MyCohesity endpoints — I'll click through the portal
2. A browser opens with your stored session. **Click through the pages you care about** — run a knowledge base search, open your assets page, open your entitlements page. Every API call the portal makes is recorded.
3. Click the green button when you are done.
4. Claude shows you the calls it captured. Ask it to save the right ones, for example:
   > Set the assets endpoint to that first JSON call, in api mode, with result path "records"

From then on the tools use your tenant's real endpoints. The settings persist in `~/.mycohesity/endpoints.json`.

You can also copy [`endpoints.example.json`](endpoints.example.json) to `~/.mycohesity/endpoints.json` and edit it by hand.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| **No MyCohesity tools in Claude** | Fully quit and reopen Claude Desktop. Check the path in the config is correct and absolute. |
| **"Could not launch a browser"** | Run `playwright install chromium`. On a locked-down machine, set `MYCOHESITY_CHROMIUM_PATH` to your Chrome executable. |
| **Login window never appears** | The MCP server needs a visible desktop. It cannot sign in over SSH or on a headless server. |
| **"Login was not completed within 600s"** | Nothing was saved — just ask Claude to log in again. To allow longer, set `MYCOHESITY_LOGIN_TIMEOUT`. |
| **"Run mycohesity_login again"** | Your portal session expired. Ask Claude to log in to MyCohesity again. |
| **Internal login lands on a MyCohesity sign-in page** | You started the wrong flow. Ask for the **internal** flow, sign in to Salesforce, and use the app launcher. |
| **Searches return nothing** | See *If Results Come Back Empty* above. |
| **Wrong Salesforce org** | Set `MYCOHESITY_SALESFORCE_URL` in the config `env` block. |
| **Everything is slow** | Page mode drives a real browser. If discovery finds a JSON endpoint, switching that tool to `api` mode makes it far faster. |

To see exactly what the server is configured with:

> Show me the MyCohesity MCP configuration

---

## Optional Settings

Add any of these to an `"env"` block in `claude_desktop_config.json`. All are optional.

| Variable | Default | What it does |
|---|---|---|
| `MYCOHESITY_PORTAL_URL` | `https://my.cohesity.com` | The portal to sign in to |
| `MYCOHESITY_SALESFORCE_URL` | `https://cohesity.lightning.force.com` | Salesforce org for the internal flow |
| `MYCOHESITY_HOME` | `~/.mycohesity` | Where the session and settings are stored |
| `MYCOHESITY_LOGIN_TIMEOUT` | `600` | Seconds allowed to finish signing in |
| `MYCOHESITY_HTTP_TIMEOUT` | `60` | Seconds before a portal request gives up |
| `MYCOHESITY_CHROMIUM_PATH` | *(unset)* | Full path to a browser executable, if Playwright's copy will not run |

```json
{
  "mcpServers": {
    "mycohesity": {
      "command": "python3",
      "args": ["/full/path/to/mycohesity_mcp.py"],
      "env": {
        "MYCOHESITY_SALESFORCE_URL": "https://cohesity.my.salesforce.com",
        "MYCOHESITY_LOGIN_TIMEOUT": "900"
      }
    }
  }
}
```

---

## Files in This Folder

| File | Description |
|---|---|
| `mycohesity_mcp.py` | The MCP server — the 11 tools Claude calls. [Reference](mycohesity_mcp.md) |
| `mycohesity_auth.py` | Browser sign-in, encrypted session storage, page rendering. [Reference](mycohesity_auth.md) |
| `endpoints.example.json` | Template for tenant-specific URL overrides |
| `requirements.txt` | Python dependencies |
| `README.md` | This guide |

---

## Also Available: Talk to Your Cluster

This MCP covers the **support portal**. If you also want Claude to talk to a **Cohesity cluster** — health, alerts, protection groups, recovery points — see [`../mcp/README.md`](../mcp/README.md). The two run side by side.

---

## Disclaimer

Personal project, not an official Cohesity tool. Not supported by Cohesity. Read-only by design: it searches and retrieves, and never modifies anything in the portal. Test before relying on it.
