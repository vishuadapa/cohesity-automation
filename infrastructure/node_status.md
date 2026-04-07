# node_status.py

Reports node health, disk tier status, and raw capacity for all
Helios-connected clusters (or a single named cluster).

Outputs an Excel workbook with two sheets:
- **Nodes** — one row per node with health status, role, IP, serial, disk count, and capacity
- **Disk Tiers** — one row per storage tier per node with disk count and tier capacity

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
| Nodes | `GET` | `https://helios.cohesity.com/irisservices/api/v1/public/nodes` |

> **Note:** The v1 public API is the correct path for Helios-proxied calls. Individual
> disk records are not exposed by the v1 public API — disk inventory is aggregated at the
> tier level via `diskCountByTier` and `capacityByTier` in the node response.

---

## Usage

```bash
# All clusters via Helios:
python3 node_status.py

# One cluster only:
python3 node_status.py --cluster <name>

# Only nodes in a non-healthy state:
python3 node_status.py --unhealthy-only

# Print raw node API response for debugging:
python3 node_status.py --debug

# Remove stored credentials:
python3 node_status.py --clear-credentials
```

Output is auto-generated as `node_status_YYYYMMDD_HHMMSS.xlsx` in the current directory.

---

## CLI Options

| Option | Description |
|--------|-------------|
| `--apikey` | Helios API key (optional — keychain is used if omitted) |
| `--cluster` | Filter to one cluster by name |
| `--unhealthy-only` | Only include nodes that are not in a healthy state |
| `--output` | Output `.xlsx` path (auto-generated if omitted) |
| `--debug` | Print raw node API response sample |
| `--clear-credentials` | Remove the stored API key from the OS keychain and exit |

---

## Node Health Logic

Node health is derived from two v1 fields (no `nodeStatus.overallStatus` in v1):

| Condition | Status shown |
|-----------|-------------|
| `isMarkedForRemoval == True` | `Marked for Removal` |
| `removalState == "kDontRemove"` | `Healthy` |
| Any other `removalState` | The raw state value |
| No state field | `Unknown` |

---

## Nodes Sheet Fields

| Column | Source field | Description |
|--------|-------------|-------------|
| Cluster | _(script)_ | Cluster name |
| Node ID | `id` | Cohesity node ID |
| Node IP | `ip` | Node management IP address |
| Status | _(derived)_ | Health status (see above) |
| Node Type | `nodeType` | `StorageNode`, `AllFlashNode`, etc. |
| Total Disks | `diskCount` | Total number of disks on this node |
| Capacity (TB) | `maxPhysicalCapacityBytes` | Raw physical capacity of the node in TB |
| Model | `hardwareModel` | Hardware model (e.g. `Alletra Storage Server 4120`) |
| Host Name | `hostName` | Node hostname |
| Serial | `cohesityNodeSerial` | Cohesity node serial number |
| Software Version | `nodeSoftwareVersion` | Node-level software version |

## Disk Tiers Sheet Fields

| Column | Source field | Description |
|--------|-------------|-------------|
| Cluster | _(script)_ | Cluster name |
| Node IP | `ip` | Node management IP |
| Host Name | `hostName` | Node hostname |
| Storage Tier | `diskCountByTier[].storageTier` | Tier name: `PCIeSSD`, `SATA-HDD`, `SATA-SSD`, etc. |
| Disk Count | `diskCountByTier[].diskCount` | Number of disks in this tier on this node |
| Tier Capacity (TB) | `capacityByTier[].tierMaxPhysicalCapacityBytes` | Raw capacity for this tier in TB |

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.0 | 2026-04-06 | feat: initial release |
