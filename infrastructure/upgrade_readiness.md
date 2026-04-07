# upgrade_readiness.py

Checks prerequisites before a major Cohesity software version upgrade.

Runs a series of readiness checks per cluster and prints a PASS/WARN/FAIL
summary to the console. Optionally saves results to a color-coded Excel workbook
for sharing with stakeholders.

Requires a **Helios API key** stored securely in the OS keychain after the first run.

Required packages:
```bash
pip install openpyxl keyring
```

**Version:** 1.0

---

## API Endpoints Used

| Purpose | Method | Endpoint |
|---------|--------|----------|
| Cluster list | `GET` | `https://helios.cohesity.com/mcm/clusters/connectionStatus` |
| Cluster detail | `GET` | `https://helios.cohesity.com/irisservices/api/v1/public/cluster` |
| Nodes | `GET` | `https://helios.cohesity.com/irisservices/api/v1/public/nodes` |
| Active runs | `GET` | `https://helios.cohesity.com/irisservices/api/v1/public/protectionRuns` |

> **Note:** All calls use the v1 public API via Helios routing. The v2 endpoints
> return 404 or 500 when proxied through Helios for on-prem clusters.
> Capacity is computed from node-level `maxPhysicalCapacityBytes` because the
> `/cluster` endpoint returns `stats: null` in v1.

---

## Usage

```bash
# Check all clusters:
python3 upgrade_readiness.py

# One cluster only:
python3 upgrade_readiness.py --cluster <name>

# Flag clusters below a version minimum:
python3 upgrade_readiness.py --min-version 7.1

# Save results to a specific Excel file:
python3 upgrade_readiness.py --output readiness.xlsx

# Print raw cluster API response keys for debugging:
python3 upgrade_readiness.py --debug

# Remove stored credentials:
python3 upgrade_readiness.py --clear-credentials
```

Output is auto-generated as `upgrade_readiness_YYYYMMDD_HHMMSS.xlsx` in the current directory.

---

## CLI Options

| Option | Description |
|--------|-------------|
| `--apikey` | Helios API key (optional — keychain is used if omitted) |
| `--cluster` | Filter to one cluster by name |
| `--min-version` | Minimum required version (e.g. `7.1`) — clusters below this FAIL |
| `--output` | Output `.xlsx` path (auto-generated if omitted) |
| `--debug` | Print raw cluster API response keys and stats value |
| `--clear-credentials` | Remove the stored API key from the OS keychain and exit |

---

## Checks Performed

| Check | PASS | WARN | FAIL |
|-------|------|------|------|
| **Node Health** | All nodes healthy | Could not retrieve nodes | Any node marked for removal or in abnormal state |
| **Disk Health** | All disks healthy with tier breakdown | Nodes with disk issues flagged | — |
| **Active Runs** | No protection runs in progress | Runs currently executing | — |
| **Software Ver.** | Version ≥ minimum (if `--min-version` set) | Version not available | Version < minimum |
| **Disk Capacity** | Raw physical capacity reported | Capacity data unavailable | — |

---

## Node Health Logic

Derived from v1 node fields (no `nodeStatus.overallStatus` in v1):
- `isMarkedForRemoval == True` → unhealthy
- `removalState != "kDontRemove"` → unhealthy
- All other nodes are healthy

---

## Excel Output

One row per (cluster, check). Color-coded:
- Green (`#C6EFCE`) — PASS
- Yellow (`#FFEB9C`) — WARN
- Red (`#FFC7CE`) — FAIL

Columns: Cluster, Check, Status, Detail
