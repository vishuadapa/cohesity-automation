# Cohesity Automation Scripts

Personal collection of Cohesity REST API automation scripts for reporting, operations, and customer use cases. Built and maintained by Vishu Adapa.

Targets Cohesity 7.1 and 7.3. Uses the [Cohesity REST API](https://developer.cohesity.com/) (v1 and v2 endpoints as appropriate).
Inspired by [Brian Seltzer's scripts](https://github.com/bseltz-cohesity/scripts).

---

## FortKnox Vault Report

> Per-protection-group FortKnox (RPaaS) vault reporting script. Connects to Helios, iterates every connected cluster, and produces a **multi-tab Excel workbook** with a data report, trend charts, and a built-in About tab explaining every field and API used.

### What it produces

| Output | Description |
|--------|-------------|
| **About tab** | Command run, report parameters, full field reference (name, description, source field path, API endpoint), APIs used table, and a Notes section explaining local vs. vault activity fields |
| **Report tab** | One row per protection group per vault. Source-banner row groups columns by originating API. Columns frozen at row 3. |
| **Chart tabs** | One line-chart sheet per cluster (trend mode only) showing Remote Storage Consumed (TB) over time, one series per protection group |

### Report columns

All Bytes values are grouped together, followed by TB equivalents in the same order:

| Column | Source | Notes |
|--------|--------|-------|
| Period Start / Period End | Computed | Date range covered by this row |
| Cluster | Helios MCM | Cluster name |
| Vault Name | dataTransferToVaults | FortKnox vault name |
| Vault Type | dataTransferToVaults | e.g. kFortKnox, kRPaaS |
| Protection Group | dataTransferToVaults | Protection job name |
| **Remote Storage Consumed (Bytes)** | dataTransferToVaults | Total retained bytes in the vault — cumulative, not windowed |
| **Activities: Data Read (Bytes)** | Protection Activities | Data read from source during Backup activity — sent to **local cluster** |
| **Activities: Data Written (Bytes)** | Protection Activities | Data written to local storage during Backup activity — sent to **local cluster** |
| **Activities: Logical Transferred (Bytes)** | Protection Activities | Logical bytes sent to the **remote FortKnox vault** (Vault activity) |
| **Activities: Physical Transferred (Bytes)** | Protection Activities | Physical (post-dedup/compress) bytes sent to the **remote FortKnox vault** |
| Remote Storage Consumed (TB) | Computed | ÷ 1,000,000,000,000 |
| Activities: Data Read (TB) | Computed | ÷ 1,000,000,000,000 |
| Activities: Data Written (TB) | Computed | ÷ 1,000,000,000,000 |
| Activities: Logical Transferred (TB) | Computed | ÷ 1,000,000,000,000 |
| Activities: Physical Transferred (TB) | Computed | ÷ 1,000,000,000,000 |

> **Note**: Data Read / Data Written reflect the Protection Activities report **Activity Type = Backup** (local cluster). Logical / Physical Transferred reflect **Activity Type = Vault** (remote FortKnox). This is equivalent to the **Protection Activities Report** on Helios, filtered to cloud vault (FortKnox).

### Field value types — cumulative vs. point-in-time

| Column | Value type | Explanation |
|--------|-----------|-------------|
| **Remote Storage Consumed** | **Cumulative** | Total retained bytes across all snapshots ever stored in the vault for this protection group. The `dataTransferToVaults` API ignores your `--start`/`--end`/`--days` window for this field — the value is always the current vault footprint. In trend mode every day row shows the same value. |
| **Activities: Data Read** | **Point-in-time** | Per-run stat from the runs API, filtered to runs that started within your queried window. In summary mode it sums all qualifying runs across the full range; in trend mode it shows only runs that started on that specific day. |
| **Activities: Data Written** | **Point-in-time** | Same filtering as Data Read. |
| **Activities: Logical Transferred** | **Point-in-time** | Same filtering — only vault archival runs that started within your window. |
| **Activities: Physical Transferred** | **Point-in-time** | Same filtering as Logical Transferred. |

**Practical implications**

- Do not use **Remote Storage Consumed** to measure ingest activity or growth within a period — it is always the current total and will be the same value regardless of the date range you query.
- Use **Logical / Physical Transferred** to understand how much data was sent to the vault during a specific window.
- Use **Data Read / Data Written** to understand local backup workload (how much data was read from source and written to the local cluster) during a specific window.
- In **trend mode**, plotting Logical/Physical Transferred per day gives a true daily vault ingest view. Plotting Remote Storage Consumed per day shows the running vault footprint (all values will be identical within a single report run).

### APIs used

| API | Endpoint | Purpose |
|-----|----------|---------|
| Helios MCM | `GET /mcm/clusters/connectionStatus` | List all connected clusters |
| Cluster v1 | `GET /public/cluster` | Detect cluster timezone |
| Cluster v1 | `GET /public/vaults?includeFortKnoxVault=true` | Fetch FortKnox vault IDs and names |
| Cluster v1 | `GET /public/reports/dataTransferToVaults` | Storage consumed per protection group (cumulative) |
| Cluster v2 | `GET /v2/data-protect/protection-groups` | List protection groups to resolve group IDs |
| Cluster v2 | `GET /v2/data-protect/protection-groups/{id}/runs` | Per-run backup and archival stats for Activities columns |

All cluster API calls are routed via the Helios proxy using the `accessClusterId` header — a single Helios API key is all that is required.

### Prerequisites

```bash
pip install requests openpyxl keyring
```

| Package | Required? | Purpose |
|---------|-----------|---------|
| `requests` | **Yes** | HTTP client |
| `openpyxl` | **Yes** | Excel workbook output |
| `keyring` | Optional | Store API key in OS keychain between runs |

### Usage

```bash
# First run — prompts for API key, saves to keychain
python3 reporting/fortknox_vault_report.py

# Summary report — last 30 days, all clusters
python3 reporting/fortknox_vault_report.py --days 30

# Trend mode — one row per day per group, with charts
python3 reporting/fortknox_vault_report.py --mode trend --days 30

# Specific date range, single cluster
python3 reporting/fortknox_vault_report.py --start 2026-03-01 --end 2026-03-31 --cluster prod-east

# Corporate proxy with custom CA cert
python3 reporting/fortknox_vault_report.py --ca-bundle /path/to/ca.pem

# Remove stored API key
python3 reporting/fortknox_vault_report.py --clear-credentials
```

### All options

| Flag | Default | Description |
|------|---------|-------------|
| `--apikey KEY` | keychain / prompted | Helios API key |
| `--clear-credentials` | off | Remove stored key from OS keychain and exit |
| `--cluster NAME` | all | Filter to one cluster (partial name match) |
| `--vault NAME` | all | Filter to one vault (partial name match) |
| `--mode` | `summary` | `summary`: one row per group for the full range. `trend`: one row per group per day with charts |
| `--days N` | 7 | Last N days |
| `--start YYYY-MM-DD` | — | Start date (inclusive) |
| `--end YYYY-MM-DD` | today | End date (inclusive) |
| `--start-msecs MS` | — | Exact start timestamp in milliseconds |
| `--end-msecs MS` | — | Exact end timestamp in milliseconds |
| `--output PATH` | auto-generated | Output `.xlsx` path |
| `--ca-bundle PATH` | system bundle | CA cert for corporate proxy or self-signed cluster |
| `--insecure` | off | Disable TLS verification (prints warning) |
| `--debug` | off | Print raw API response samples and request URLs |

### Output filename

Auto-generated with version and timestamp:

```
fortknox_vault_report_v4.22_20260421_143000.xlsx
```

---

## MyCohesity MCP — Support Portal from Claude Desktop

> MCP server that lets Claude search the **MyCohesity knowledge base** — Cohesity DataProtect and Veritas NetBackup — and retrieve **asset** and **licensing entitlement** information, using your own portal session.

Full setup guide: **[mycohesity/README.md](mycohesity/README.md)** · Reference: [mycohesity_mcp.md](mycohesity/mycohesity_mcp.md)

### What you can ask

```
Search MyCohesity for "backup fails with status 84" in NetBackup
Find DataProtect articles about upgrade readiness and summarise the top three
List my registered assets — which have support ending in the next 6 months?
What are my current licensing entitlements? Show licensed capacity and end dates.
```

### Sign-in

my.cohesity.com is a Salesforce Experience Cloud portal, so there is no API key and no password to put in a config file:

- **External users** (customers, partners) have **MFA enforced**.
- **Internal Cohesity users** never sign in to my.cohesity.com directly — they authenticate to the Cohesity **Salesforce** org and launch MyCohesity from the **app launcher** (the icon stack, top left).

The first time you use it, a real browser window opens and you sign in exactly as you normally would — SSO, MFA, app launcher, all of it. You click a green button when the portal has loaded, and the session is **encrypted to `~/.mycohesity/session.enc`** (Fernet, key held in your OS keychain) and reused until the portal expires it. No credentials are stored in the Claude Desktop config.

Read-only by design: it searches and retrieves, and never modifies anything in the portal.

### Tools

| Group | Tools |
|-------|-------|
| Authentication | `mycohesity_login`, `mycohesity_session_status`, `mycohesity_logout` |
| Knowledge base | `search_knowledge_base`, `get_knowledge_base_article` |
| Assets & licensing | `list_assets`, `get_entitlements` |
| Configuration | `mycohesity_fetch`, `mycohesity_discover_endpoints`, `mycohesity_set_endpoint`, `mycohesity_show_config` |

### Prerequisites

```bash
pip install "mcp[cli]" playwright cryptography httpx keyring
playwright install chromium
```

Needs a visible desktop for the one-time sign-in — it cannot authenticate over SSH or on a headless server.

---

## Repository Structure

| Folder | Purpose |
|--------|---------|
| [reporting/](reporting/) | Protection group run reports and FortKnox vault data transfer reports |
| [protection/](protection/) | Protection group management — run, pause, clone jobs |
| [recovery/](recovery/) | Restore automation — VMs, files, file/folder recovery |
| [storage/](storage/) | Capacity planning, views, storage domain reporting |
| [policies/](policies/) | Policy audit, SLA compliance, and cloning |
| [alerts/](alerts/) | Alert queries, summaries, bulk resolve, and CSV export |
| [infrastructure/](infrastructure/) | Cluster health, node status, version inventory, upgrade readiness |
| [mcp/](mcp/) | Cohesity MCP server for Claude Desktop — talk to your cluster in natural language |
| [mycohesity/](mycohesity/) | MyCohesity MCP server for Claude Desktop — search the support knowledge base and check licensing entitlements |
| [utils/](utils/) | Shared auth and helper modules |

---

## Authentication

All scripts support two modes:

**Helios (recommended)** — single API key covers all connected clusters:
```bash
python3 <script>.py --apikey <your-helios-api-key>
```

The key is stored in the OS keychain after the first use — `--apikey` is optional on subsequent runs.

**Direct cluster** — username/password against a specific cluster:
```bash
python3 <script>.py --cluster <ip-or-hostname> --username admin --domain LOCAL
```

---

## Requirements

All scripts require `requests` and `openpyxl`. Install everything:

```bash
pip3 install requests openpyxl keyring
```

| Package | Scripts | Required? |
|---------|---------|-----------|
| `requests` | All scripts | **Yes** — core HTTP client |
| `openpyxl` | Reporting, infrastructure, alerts, storage, policies | **Yes** — Excel workbook output |
| `keyring` | All scripts | Optional — OS keychain credential storage |

Both MCP servers have their own dependencies — see [mcp/README.md](mcp/README.md) and [mycohesity/README.md](mycohesity/README.md).

---

## Quick Reference

### Reporting
```bash
python3 reporting/protection_group_report.py --days 7
python3 reporting/fortknox_vault_report.py --days 30
python3 reporting/fortknox_vault_report.py --mode trend --days 30
python3 reporting/fortknox_vault_report.py --start 2026-03-01 --end 2026-03-31 --cluster prod-east
```

### Alerts
```bash
python3 alerts/alert_summary_report.py --days 1
python3 alerts/resolve_alerts.py --severity kCritical --cluster prod --yes
python3 alerts/alert_to_csv.py --days 30
```

### Infrastructure
```bash
python3 infrastructure/cluster_info_report.py
python3 infrastructure/node_status.py --unhealthy-only
python3 infrastructure/upgrade_readiness.py --min-version 7.1
```

### Policies
```bash
python3 policies/list_policies.py
python3 policies/policy_compliance_report.py --min-retention 14 --require-replication
python3 policies/clone_policy.py --source "Gold 30-Day" --name "Gold 60-Day" --cluster prod
```

### Protection
```bash
python3 protection/run_protection_group.py --group "VMware Daily" --cluster prod
python3 protection/pause_resume_groups.py --action pause --pattern "VMware" --cluster prod
python3 protection/clone_protection_group.py --source "SQL Prod" --name "SQL DR" --cluster prod
```

### Recovery
```bash
python3 recovery/list_snapshots.py --object "vm-prod-01" --cluster prod
python3 recovery/restore_vm.py --vm "vm-prod-01" --cluster prod --suffix "-restored"
python3 recovery/recover_files.py --source "linux-host" --paths /var/log --cluster prod
```

### Storage
```bash
python3 storage/capacity_report.py
python3 storage/list_views.py --protocol nfs
python3 storage/storage_domain_report.py
```

---

## API Versions

- **v1 API** — Used for infrastructure, alerts, storage, and other operations for better Helios reliability
- **v2 API** — Used for protection group reporting, policies, protection operations, and recovery

Individual `README.md` files in each folder document which API version is used by those scripts.

---

## Disclaimer

These scripts are personal tools, not official Cohesity products. Use at your own risk in production environments.
