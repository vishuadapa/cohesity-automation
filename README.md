# Cohesity Automation Scripts

Personal collection of Cohesity REST API automation scripts for reporting, operations, and customer use cases. Built and maintained by Vishu Adapa.

Targets Cohesity 7.1 and 7.3. Uses the [Cohesity v2 REST API](https://developer.cohesity.com/).
Inspired by [Brian Seltzer's scripts](https://github.com/bseltz-cohesity/scripts).

---

## Repository Structure

| Folder | Purpose |
|--------|---------|
| [reporting/](reporting/) | Health check and status reports — CSV output for Excel and customer decks |
| [protection/](protection/) | Protection group management — run, pause, clone jobs |
| [recovery/](recovery/) | Restore automation — VMs, files, databases |
| [storage/](storage/) | Capacity planning, views, storage domain reporting |
| [policies/](policies/) | Policy audit, creation, and SLA compliance |
| [alerts/](alerts/) | Alert queries, summaries, and routing |
| [infrastructure/](infrastructure/) | Cluster health, node status, version inventory |
| [utils/](utils/) | Shared auth and helper modules |

---

## Authentication

All scripts support two modes:

**Helios (recommended)** — single API key covers all connected clusters:
```bash
python3 <script>.py --apikey <your-helios-api-key>
```
Generate your API key in Helios → Settings → Access Management → API Keys.

**Direct cluster** — username/password against a specific cluster:
```bash
python3 <script>.py --cluster <ip-or-hostname> --username admin --domain LOCAL
```

---

## Requirements

```bash
pip3 install requests
```

---

## Disclaimer

These scripts are personal tools, not official Cohesity products. Use at your own risk in production environments.
