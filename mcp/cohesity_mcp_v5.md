# Cohesity MCP Server v5 — Multiple Clusters + Helios

**Everything v4 does (secure credential storage), plus any number of clusters
and Helios — still with zero secrets in your Claude Desktop config.**

| Capability | v2 | v4 | v5 |
|---|---|---|---|
| Talk to a Cohesity cluster from Claude | ✅ | ✅ | ✅ |
| API key stored securely (Keychain / Credential Manager) | — | ✅ | ✅ |
| Multiple clusters (named profiles) | — | — | ✅ |
| Helios (one login, all connected clusters) | — | — | ✅ |

---

## Why — What v5 Adds and Why

**Multiple clusters.** v4 stores exactly one cluster's credentials. If you
work with a prod cluster, a DR cluster, and a lab, you had to re-run `setup`
every time you switched — or fall back to plain-text env vars, defeating the
point of v4. v5 introduces **named profiles**: each cluster's address and key
live in the OS lockbox under a name you choose (`prod`, `dr`, `lab`). Claude
Desktop selects a profile with a single non-secret env var, so you can expose
several clusters **side by side in the same Claude conversation** — ask
"compare open alerts on prod and dr" and Claude queries both.

**Helios.** If your clusters are registered to Cohesity Helios, one Helios
API key covers all of them. A Helios-type profile connects to
`helios.cohesity.com`, and two new tools appear in Claude:
`list_clusters` (what's connected) and `select_cluster` (pick a target).
Every other tool then operates on the selected cluster through Helios
passthrough — Cohesity's `accessClusterId` routing header. Switch clusters
mid-conversation just by saying *"now switch to the DR cluster"*.

## What — What v5 Is

`cohesity_mcp_v5.py` is `cohesity_mcp_v4.py` with a profile layer and a
Helios connection type. Unchanged from v2/v4: all monitoring, reporting, and
action tools, the confirm pattern, RBAC/risk gating, and dynamic API
discovery. New:

1. **Profiles** — lockbox entries are namespaced per profile
   (`prod:api_key`, `dr:cluster`, …) under the same service name
   `cohesity-mcp`. A stored index tracks profile names and the default.
2. **Two connection types per profile**:
   - `cluster` — direct to one cluster (v2/v4 behaviour, API key or
     username/password).
   - `helios` — via Helios (API key required; username/password isn't
     supported by Helios API auth). SSL verification defaults **on** for
     Helios (it has a real certificate) and off for lab clusters.
3. **Two new MCP tools** — `list_clusters` and `select_cluster`
   (Helios profiles only; on direct profiles they politely explain they
   don't apply). Until a cluster is selected, cluster-scoped tools return a
   clear instruction to select one first, and `get_cluster_health` degrades
   gracefully where Helios passthrough doesn't expose service-level status.
4. **Profile-aware CLI** — `setup` (add/update a profile), `list`, `use`
   (set default), `show [name]`, `test [name]`, `clear <name>` /
   `clear --all`.
5. **Full backward compatibility** — a v4 lockbox is read automatically as
   the profile named `default` (no migration needed), and v2-style env vars
   still override everything.

## How — Setup

### Step 1 — Install packages (once)

```bash
pip install "mcp[cli]" httpx keyring
```

### Step 2 — Add a profile per cluster (and/or one for Helios)

```bash
python3 /path/to/cohesity_mcp_v5.py setup
```

The prompts, in order:

1. **Profile name** — e.g. `prod` (default: `default`)
2. **Connection type** — `1` single cluster (direct) or `2` Helios
3. **Address** — cluster hostname, or Enter to accept `helios.cohesity.com`
4. **API key** — hidden while you type. Direct profiles may press Enter to
   use username/password instead; Helios requires an API key.
5. **Verify SSL** — Enter accepts the sensible default (on for Helios,
   off for self-signed lab clusters)
6. **Default profile?** — the first profile becomes the default automatically

Repeat `setup` for each cluster. Then verify and review:

```bash
python3 cohesity_mcp_v5.py test prod     # connects and authenticates
python3 cohesity_mcp_v5.py list          # all profiles, default marked
```

For a Helios profile, `test` also lists every cluster registered to your
Helios account and whether each is currently connected.

### Step 3 — One Claude Desktop entry per profile

`COHESITY_PROFILE` is just a name — not a secret — so this config stays safe
to share:

```json
{
  "mcpServers": {
    "cohesity-prod": {
      "command": "python3",
      "args": ["/full/path/to/cohesity_mcp_v5.py"],
      "env": { "COHESITY_PROFILE": "prod" }
    },
    "cohesity-dr": {
      "command": "python3",
      "args": ["/full/path/to/cohesity_mcp_v5.py"],
      "env": { "COHESITY_PROFILE": "dr" }
    },
    "cohesity-helios": {
      "command": "python3",
      "args": ["/full/path/to/cohesity_mcp_v5.py"],
      "env": { "COHESITY_PROFILE": "helios" }
    }
  }
}
```

(Windows: `"command": "python"` and a `C:\\...` path. A single-profile setup
can omit the `env` block entirely — the stored default profile is used.)

Fully quit and relaunch Claude Desktop.

### Using it in Claude

Direct profiles work exactly as before. With a Helios profile:

> *"List my Helios clusters"* → Claude calls `list_clusters`
> *"Select the Prod-East cluster"* → `select_cluster`, by name or id
> *"Show its health and open critical alerts"* → all tools now route there
> *"Now switch to DR-West and compare"* → `select_cluster` again

With multiple direct profiles loaded, just name the server:
*"Using cohesity-dr, list protection groups that failed last night."*

### Everyday commands

| Command | What it does |
|---|---|
| `setup` | Add a new profile or update an existing one (re-run on key rotation) |
| `list` | All stored profiles with type, address, and the default marker |
| `use <name>` | Change which profile is the default |
| `show [name]` | One profile's config — API key masked, password hidden |
| `test [name]` | Connect and verify auth; Helios: also lists connected clusters |
| `clear <name>` | Remove one profile (`clear --all` wipes everything) |
| *(no command)* | Run the MCP server — what Claude Desktop does |

## Where — Where Everything Lives

Same lockbox as v4 — macOS **Keychain**, Windows **Credential Manager**,
Linux Secret Service — service name `cohesity-mcp`. Each profile's settings
are separate entries named `<profile>:<setting>`, e.g. `prod:api_key`,
`dr:cluster`. Two housekeeping entries, `_profiles` and `_default`, hold the
profile index and default name (no secrets in either). Inspect or delete any
of it in Keychain Access / Credential Manager, or with `clear`.

Nothing is written to any file. The security boundaries are the same as v4
(see `cohesity_mcp_v4.md`): protected at rest, same trust model as your
browser's saved passwords, env vars deliberately override for CI and
debugging.

## Troubleshooting

**"Helios mode: no target cluster selected"**
Working as designed — ask Claude to *"list clusters"* then *"select cluster
\<name\>"*. Helios needs to know which cluster to route each call to.

**A cluster shows "NOT connected" in `test` / `list_clusters`**
Helios can only reach clusters that are currently phoning home. Check the
cluster's Helios connection (cluster UI → Settings → Helios).

**`get_cluster_health` says service-level status unavailable**
Normal through Helios passthrough — the `nexus` status endpoint is
cluster-local. Identity, version, and node count still come through; use a
direct profile for deep service-level health.

**Two Pythons problem** (profile saved but server says "no address
configured"): run `setup` with the same interpreter that Claude Desktop's
`command` uses — see the v4 guide's troubleshooting section.

**Upgrading from v4**
Nothing to migrate: v5 reads a v4 lockbox as the profile `default`. Point
your Claude config at `cohesity_mcp_v5.py` and it behaves identically until
you add more profiles.

## Reference

### Configuration resolution (per setting)

```
env var  →  OS lockbox entry "<profile>:<key>"  →  (profile 'default' only) v4 legacy key  →  default value
```

Profile selection at server start: `COHESITY_PROFILE` env var → stored
default (`_default`) → `default`.

| Setting | Env var (overrides all profiles) | Lockbox key | Default |
|---|---|---|---|
| Connection type | — | `<p>:type` | `cluster` |
| Address | `COHESITY_CLUSTER` | `<p>:cluster` | — (Helios: `helios.cohesity.com` at setup) |
| API key | `COHESITY_API_KEY` | `<p>:api_key` | — |
| Username | `COHESITY_USERNAME` | `<p>:username` | — (direct only) |
| Password | `COHESITY_PASSWORD` | `<p>:password` | — (direct only) |
| Domain | `COHESITY_DOMAIN` | `<p>:domain` | `LOCAL` |
| Verify SSL | `COHESITY_VERIFY_SSL` | `<p>:verify_ssl` | `false` direct / `true` Helios |
| Profile | `COHESITY_PROFILE` | `_default` | `default` |
| Transport / HTTP port | `MCP_TRANSPORT` / `MCP_HTTP_PORT` | — | `stdio` / `8720` |

### Tools

The 14 v2/v4 tools unchanged, plus:

| Tool | Mode | What it does |
|---|---|---|
| `list_clusters` | Helios | Clusters registered to the Helios account: id, name, connected, selected |
| `select_cluster` | Helios | Choose the target cluster by name or id; refuses disconnected clusters |

### Version history

| Version | Summary |
|---|---|
| v2 | Cluster-direct; credentials as plain-text env vars in the Claude config |
| v3 | TOTP + keyring experiment — abandoned (not in this repo) |
| v4 | v2 + OS lockbox credential storage; `setup`/`show`/`test`/`clear` CLI |
| **v5** | v4 + named multi-cluster profiles (`COHESITY_PROFILE`) + Helios connection type with `list_clusters`/`select_cluster`; reads v4 lockboxes unchanged |
