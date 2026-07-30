# Cohesity MCP Server v4 — Secure Credential Storage

**Everything v2 does, minus the plain-text API key in your config file.**

v4 stores your Cohesity API key (or username/password) in your operating
system's built-in credential lockbox — the same protected store your browsers
and password managers use — instead of in `claude_desktop_config.json`.

| Platform | Where your credentials live |
|---|---|
| **macOS** | Keychain (Keychain Access app) |
| **Windows** | Credential Manager (Control Panel → Credential Manager) |
| Linux | Secret Service (GNOME Keyring / KWallet) |

---

## Why — The Problem v4 Solves

With v2, the setup instructions told you to paste your Cohesity API key
directly into `claude_desktop_config.json`:

```json
"env": {
  "COHESITY_CLUSTER": "mycluster.lab.local",
  "COHESITY_API_KEY": "sk-abc123-your-actual-key-here"
}
```

That works, but it means your API key sits **in plain text** in a file that:

- anyone with access to your machine can open and read;
- gets swept up by backup tools, sync clients (iCloud/OneDrive), and
  "share your config" screenshots;
- is easy to accidentally paste into a chat, email, or Git commit when
  troubleshooting.

A Cohesity API key is a real credential — whoever holds it can do whatever
your cluster role allows. It deserves the same protection as a password.

**What about v3?** v3 attempted TOTP (time-based one-time passwords) combined
with keyring storage. It added a second factor to every session, but the
moving parts (secret enrollment, clock drift, re-prompting inside a
non-interactive MCP launch) made it fragile and it never worked reliably —
an MCP server is launched silently in the background by Claude Desktop, so
there is no natural moment to type a 6-digit code. v4 drops TOTP and keeps
the part that worked and matters most: **the secret never sits in a plain-text
file**. Simple, reliable, and invisible to the user after a one-time setup.

## What — What v4 Is

`cohesity_mcp_v4.py` is `cohesity_mcp_v2.py` plus a credential lockbox layer.
All 14 tools, the confirm pattern, the RBAC/risk gating, and the dynamic API
discovery are identical to v2. What changed:

1. **Secure storage** — cluster address, API key (or username/password/domain),
   and the SSL-verify preference are stored in the OS lockbox via the
   cross-platform [`keyring`](https://pypi.org/project/keyring/) library,
   under the service name `cohesity-mcp`.
2. **A tiny built-in CLI** — four commands (`setup`, `show`, `test`, `clear`)
   so you never have to touch the lockbox directly.
3. **Config resolution order** — for every setting, an environment variable
   still wins if present (full v2 backward compatibility), otherwise the
   value comes from the lockbox. Your Claude Desktop config no longer needs
   an `env` block at all.
4. **MCP SDK 1.x and 2.x compatibility** — the import works with both the
   older `mcp` package (`FastMCP`) and the 2.x rename (`MCPServer`).

Nothing about how you *talk* to your cluster changes — every prompt that
worked with v2 works with v4.

## How — Setup in Three Steps

### Step 1 — Install the required packages (once)

```bash
pip install "mcp[cli]" httpx keyring
```

(Windows: use `pip` in PowerShell; if `python3` isn't found below, use `python`.)

### Step 2 — Run the one-time setup

```bash
python3 /path/to/cohesity_mcp_v4.py setup
```

You will be asked for:

1. **Cluster address** — e.g. `mycluster.lab.local` (no `https://`)
2. **API key** — input is hidden while you type or paste.
   Press Enter instead to fall back to username/password/domain.
3. **Verify SSL?** — answer `n` for lab clusters with self-signed certificates.

The values are written straight into your OS lockbox. Verify with:

```bash
python3 /path/to/cohesity_mcp_v4.py test    # connects and authenticates
python3 /path/to/cohesity_mcp_v4.py show    # displays config, key masked
```

### Step 3 — Point Claude Desktop at the script

Edit `claude_desktop_config.json`
(Mac: `~/Library/Application Support/Claude/claude_desktop_config.json`,
Windows: `%APPDATA%\Claude\claude_desktop_config.json` — or in Claude
Desktop: **Settings → Developer → Edit Config**):

**Mac:**
```json
{
  "mcpServers": {
    "cohesity": {
      "command": "python3",
      "args": ["/full/path/to/cohesity_mcp_v4.py"]
    }
  }
}
```

**Windows:**
```json
{
  "mcpServers": {
    "cohesity": {
      "command": "python",
      "args": ["C:\\full\\path\\to\\cohesity_mcp_v4.py"]
    }
  }
}
```

Notice: **no `env` block, no secrets.** This file is now safe to back up,
sync, screenshot, and share.

Fully quit and relaunch Claude Desktop. Done.

> **macOS only, first launch:** macOS may show a dialog like
> *"python wants to access confidential information stored in your
> keychain."* Click **Always Allow** so you're not asked again.
> This is the Keychain doing its job — every program that reads the key
> must be approved by you.

### Everyday commands

| Command | What it does |
|---|---|
| `python3 cohesity_mcp_v4.py setup` | Store / replace credentials (run again any time the key rotates) |
| `python3 cohesity_mcp_v4.py show`  | Display stored config — API key masked to last 4 characters |
| `python3 cohesity_mcp_v4.py test`  | Connect to the cluster and confirm authentication works |
| `python3 cohesity_mcp_v4.py clear` | Wipe all stored Cohesity credentials from the lockbox |
| `python3 cohesity_mcp_v4.py`       | Run the MCP server (what Claude Desktop does automatically) |

## Where — Where Exactly Are My Credentials?

Everything is stored under the keyring **service name `cohesity-mcp`**, one
entry per setting: `cluster`, `api_key` (or `username` / `password` /
`domain`), and `verify_ssl`.

- **macOS** — open **Keychain Access** (Spotlight → "Keychain Access"),
  search for `cohesity-mcp`. Each entry appears as an application password
  in your *login* keychain, encrypted at rest and unlocked with your Mac
  login. You can inspect or delete entries there directly.
- **Windows** — open **Credential Manager** (Start → type "Credential
  Manager") → **Windows Credentials** → look under *Generic Credentials*
  for `cohesity-mcp`. Entries are encrypted with DPAPI, tied to your
  Windows account.
- **Linux** — stored in the desktop's Secret Service (GNOME Keyring or
  KWallet), viewable with Seahorse ("Passwords and Keys") or
  KWalletManager.

Nothing is written to any file in this repository, your home directory, or
the Claude Desktop config. Deleting is one command
(`python3 cohesity_mcp_v4.py clear`) or a right-click in the OS tool above.

### Honest security boundaries

- The lockbox protects the key **at rest** and against casual/file-level
  access. Any program running *as you* that you approve (that's the macOS
  Keychain prompt) can read it — this is the same trust model your browser's
  saved passwords use.
- Environment variables still override the lockbox. That's a feature (CI,
  quick tests, v2 configs keep working) but remember: if you put
  `COHESITY_API_KEY` back in the Claude config, you're back to plain text.
- Headless Linux servers without a desktop keyring need either a Secret
  Service daemon or env vars — v4 falls back to env vars gracefully.

## Troubleshooting

**"The 'keyring' package is not installed"**
Run `pip install keyring` with the same Python you use to run the script.

**"No cluster configured" when Claude Desktop starts the server**
The lockbox is empty for the Python that Claude Desktop launched. Two common causes:
1. You never ran `setup` — run it.
2. You have multiple Pythons and `setup` ran under a different one than the
   `command` in your Claude config. Use the same interpreter for both, e.g.
   run `/usr/bin/python3 cohesity_mcp_v4.py setup` and put `/usr/bin/python3`
   as the `command`.

**macOS keeps prompting for keychain access**
Click **Always Allow** (not "Allow") in the dialog. If you upgraded or
moved your Python, macOS treats it as a new app and asks once more.

**Key rotated / auth suddenly failing**
`python3 cohesity_mcp_v4.py test` will report an authentication failure —
just re-run `setup` with the new key.

**I want to go back to env vars**
That still works exactly as in v2 — set `COHESITY_CLUSTER` and
`COHESITY_API_KEY` in the config's `env` block; env vars take precedence
over the lockbox.

## Reference

### Configuration resolution (per setting)

```
environment variable  →  OS lockbox (service "cohesity-mcp")  →  default
```

| Setting | Env var | Lockbox key | Default |
|---|---|---|---|
| Cluster address | `COHESITY_CLUSTER` | `cluster` | — (required) |
| API key | `COHESITY_API_KEY` | `api_key` | — |
| Username | `COHESITY_USERNAME` | `username` | — |
| Password | `COHESITY_PASSWORD` | `password` | — |
| Domain | `COHESITY_DOMAIN` | `domain` | `LOCAL` |
| Verify SSL | `COHESITY_VERIFY_SSL` | `verify_ssl` | `false` |
| Transport | `MCP_TRANSPORT` | — | `stdio` |
| HTTP port | `MCP_HTTP_PORT` | — | `8720` |

### Tools

Identical to v2 — see the tool table in [`README.md`](README.md):
monitor (`get_cluster_health`, `get_alerts`, `get_storage_stats`), report
(`get_protection_groups`, `get_protection_runs`, `get_recovery_points`,
`healthcheck_summary`), identity/discovery (`get_whoami`,
`search_cluster_api`), actions with confirm pattern
(`pause_protection_group`, `resume_protection_group`, `retry_failed_run`,
`acknowledge_alert`), and the RBAC-gated generic executor (`call_api`).

### Version history

| Version | Summary |
|---|---|
| v2 | Cluster-direct MCP server; credentials via env vars in the Claude config (plain text) |
| v3 | TOTP + keyring experiment — abandoned (unreliable in non-interactive MCP launches; not in this repo) |
| **v4** | v2 + OS lockbox credential storage via `keyring`; `setup`/`show`/`test`/`clear` CLI; MCP SDK 1.x/2.x compatible |
