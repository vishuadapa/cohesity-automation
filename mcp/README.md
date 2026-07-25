# Cohesity MCP v2 — Setup Guide

**Talk to your Cohesity cluster directly from Claude — no code required.**

This guide walks you through every step, even if you have never heard of MCP before. By the end you will be able to open Claude Desktop and type things like:

> *"Show me the health of my cluster"*
> *"Which protection groups failed last night?"*
> *"Pause the group called Daily-VMs and confirm it"*

---

## What Is MCP and Why Does It Matter?

MCP stands for **Model Context Protocol**. Think of it as a power cable that connects Claude to external systems — in this case, your Cohesity cluster. Without MCP, Claude can only answer questions using what it already knows. With MCP, Claude can actually reach out to your cluster in real time, pull live data, and take actions on your behalf.

The file `cohesity_mcp_v2.py` in this folder is that power cable — a small Python program that sits between Claude and your cluster.

---

## Before You Start — What You Need

| Requirement | How to check |
|---|---|
| **Python 3.10 or newer** | Open Terminal / Command Prompt and run: `python3 --version` |
| **Claude Desktop app** | Download from [claude.ai/download](https://claude.ai/download) (free or Pro plan) |
| **Your Cohesity cluster address** | e.g. `mycluster.company.com` or an IP address |
| **A Cohesity API key or username + password** | Ask your Cohesity admin if you don't have one |

> **Windows users:** all commands below work the same in PowerShell. Replace `python3` with `python` if needed.

---

## Step 1 — Download This Folder

If you are not a developer, the easiest way to get the files is:

1. Go to the GitHub repository page.
2. Click the green **Code** button → **Download ZIP**.
3. Unzip the downloaded file somewhere easy to find, like your Desktop or Documents folder.

The file you need is:
```
cohesity-automation/mcp/cohesity_mcp_v2.py
```

Note the **full path** to this file — you will need it in Step 4. For example:
- Mac/Linux: `/Users/yourname/Downloads/cohesity-automation/mcp/cohesity_mcp_v2.py`
- Windows: `C:\Users\yourname\Downloads\cohesity-automation\mcp\cohesity_mcp_v2.py`

---

## Step 2 — Install the Required Python Package

Open **Terminal** (Mac/Linux) or **Command Prompt / PowerShell** (Windows) and run:

```bash
pip install "mcp[cli]" httpx
```

This installs two small libraries the MCP server needs. It takes about 30 seconds. You only need to do this once.

---

## Step 3 — Find Your Claude Desktop Config File

Claude Desktop uses a configuration file to know which MCP servers to load at startup. The file is called `claude_desktop_config.json`.

| Operating System | Location |
|---|---|
| **Mac** | `~/Library/Application Support/Claude/claude_desktop_config.json` |
| **Windows** | `%APPDATA%\Claude\claude_desktop_config.json` |

**Tip — open the config directly from Claude Desktop:**
1. Open Claude Desktop.
2. Click the **Claude** menu (Mac) or the hamburger menu (Windows).
3. Click **Settings → Developer → Edit Config**.

If the file does not exist yet, create it — Claude Desktop will pick it up on the next launch.

---

## Step 4 — Add the Cohesity MCP Server to the Config

Open `claude_desktop_config.json` in any text editor (Notepad, TextEdit, VS Code, etc.) and paste the block below. If the file already has content, add the Cohesity entry inside the existing `mcpServers` object.

### Mac / Linux

```json
{
  "mcpServers": {
    "cohesity": {
      "command": "python3",
      "args": ["/full/path/to/cohesity_mcp_v2.py"],
      "env": {
        "COHESITY_CLUSTER": "your-cluster-address-here",
        "COHESITY_API_KEY": "your-api-key-here"
      }
    }
  }
}
```

### Windows

```json
{
  "mcpServers": {
    "cohesity": {
      "command": "python",
      "args": ["C:\\full\\path\\to\\cohesity_mcp_v2.py"],
      "env": {
        "COHESITY_CLUSTER": "your-cluster-address-here",
        "COHESITY_API_KEY": "your-api-key-here"
      }
    }
  }
}
```

**Replace the placeholder values:**

| Placeholder | What to put there |
|---|---|
| `/full/path/to/cohesity_mcp_v2.py` | The actual path you noted in Step 1 |
| `your-cluster-address-here` | Your cluster hostname or IP, e.g. `mycluster.lab.local` — no `https://` |
| `your-api-key-here` | Your Cohesity API key |

### Using Username + Password Instead of an API Key

If you don't have an API key, you can use your Cohesity credentials instead. Replace the `env` block with:

```json
"env": {
  "COHESITY_CLUSTER": "your-cluster-address-here",
  "COHESITY_USERNAME": "your-username",
  "COHESITY_PASSWORD": "your-password",
  "COHESITY_DOMAIN": "LOCAL"
}
```

Change `"LOCAL"` to your Active Directory domain name if your account is an AD account (e.g. `"CORP"`).

---

## Step 5 — Restart Claude Desktop

Completely quit Claude Desktop (don't just close the window — quit it fully), then re-open it.

When Claude Desktop starts up it will automatically launch the Cohesity MCP server in the background. You will see a small **plug icon** (🔌) or a tool count indicator in the chat interface — that confirms the MCP is connected.

---

## Step 6 — Start Talking to Your Cluster

You are ready. Open a new conversation in Claude Desktop and try these:

### Health and Status
- *"What is the health of my Cohesity cluster?"*
- *"Are there any critical alerts open right now?"*
- *"Show me storage usage and capacity."*

### Protection Groups and Backup Jobs
- *"List all my protection groups."*
- *"Which protection groups had a failed run recently?"*
- *"Show me the last 5 runs for the group called Daily-VMs."*

### Recovery Points
- *"What snapshots are available for the VM called prod-db-01?"*

### Taking Action (Always Proposes Before Executing)
The MCP uses a **confirm pattern** for any action that changes something. Claude will always show you a proposal first — nothing executes until you say yes.

- *"Pause the protection group Daily-VMs."* → Claude proposes, you confirm.
- *"Resume the group Daily-VMs."* → Same pattern.
- *"Trigger an on-demand backup run for group ID 12345."*
- *"Acknowledge the alert with ID abc-123."*

### Advanced — Any Cohesity API Endpoint
- *"Search the Cohesity API for endpoints related to storage domains."*
- *"Who am I logged in as, and what am I allowed to do?"*
- *"Call the Cohesity API to GET /v2/data-protect/protection-groups — propose it first."*

---

## Safety Built In

This MCP server has two layers of protection so you don't accidentally break something:

1. **Confirm pattern** — Every action (pause, resume, run, resolve) asks for your approval before executing. Claude shows you exactly what it will do and waits for you to say "yes, confirm."

2. **RBAC enforcement** — The server checks your Cohesity role against what you are trying to do. A read-only viewer cannot accidentally trigger a destructive action. Destructive operations like factory reset and data wipe are blocked unconditionally.

---

## Troubleshooting

**The plug icon does not appear / Claude says it has no tools**
- Make sure you fully quit and relaunched Claude Desktop after saving the config.
- Double-check the file path in `args` — it must be the exact path to `cohesity_mcp_v2.py` with no typos.
- Run the server manually in your terminal to see any errors:
  ```bash
  COHESITY_CLUSTER=mycluster.lab.local COHESITY_API_KEY=your-key python3 /path/to/cohesity_mcp_v2.py
  ```

**SSL / certificate errors**
- Lab clusters use self-signed certificates. The server disables SSL verification by default (`COHESITY_VERIFY_SSL=false`). If your cluster uses a trusted certificate and you want to enforce SSL, add `"COHESITY_VERIFY_SSL": "true"` to the `env` block.

**"Connection refused" or "Could not reach cluster"**
- Confirm the cluster address is reachable from your machine: `ping mycluster.lab.local`
- Ensure you are on the same network or VPN as the cluster.

**"RBAC check failed" in Claude's response**
- Your Cohesity account does not have permission for that operation. Ask your Cohesity admin to elevate your role, or ask them to perform the action.

---

## Environment Variables Reference

| Variable | Required | Description |
|---|---|---|
| `COHESITY_CLUSTER` | **Yes** | Cluster hostname or IP (no scheme) |
| `COHESITY_API_KEY` | One of these two | Preferred auth method |
| `COHESITY_USERNAME` + `COHESITY_PASSWORD` | One of these two | Fallback if no API key |
| `COHESITY_DOMAIN` | No | Default: `LOCAL`. Set to your AD domain if applicable |
| `COHESITY_VERIFY_SSL` | No | `true` or `false` (default: `false`) |
| `MCP_TRANSPORT` | No | `stdio` (default, for Claude Desktop) or `http` |
| `MCP_HTTP_PORT` | No | HTTP port when using `http` transport (default: `8720`) |

---

## Available Tools at a Glance

| Tool | What It Does |
|---|---|
| `get_cluster_health` | Cluster name, version, node count, health status |
| `get_alerts` | Open alerts, filterable by severity |
| `get_storage_stats` | Capacity, usage, dedup/compression ratios |
| `get_protection_groups` | List all active protection groups and their last run status |
| `get_protection_runs` | Recent run history for a specific group |
| `get_recovery_points` | Available snapshots for a named object (VM, DB, share) |
| `healthcheck_summary` | One-shot summary: health + criticals + failing groups |
| `get_whoami` | Your identity, roles, and what you are allowed to do |
| `search_cluster_api` | Search the live Swagger spec for any API endpoint |
| `pause_protection_group` | Pause a group (confirm pattern) |
| `resume_protection_group` | Resume a paused group (confirm pattern) |
| `retry_failed_run` | Trigger an on-demand backup run (confirm pattern) |
| `acknowledge_alert` | Resolve/acknowledge an open alert (confirm pattern) |
| `call_api` | Generic executor for any Cohesity API endpoint (RBAC + confirm gated) |
