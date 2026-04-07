# cluster_info_report.py

Reports software version, node count, cluster ID, and key configuration across all
Helios-connected clusters (or a single named cluster).

Useful as a quick inventory snapshot before upgrades, for presales assessments, or
for multi-cluster environment audits.

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
| Node list | `GET` | `https://helios.cohesity.com/irisservices/api/v1/public/nodes` |

> **Note:** v2 `/v2/clusters` and `/v2/nodes` return 404 when routed through Helios.
> The v1 public API endpoints are the correct paths for Helios-proxied cluster calls.

---

## Usage

```bash
# All clusters via Helios:
python3 cluster_info_report.py

# One cluster only:
python3 cluster_info_report.py --cluster <name>

# Custom output path:
python3 cluster_info_report.py --output report.xlsx

# Print raw API response for debugging:
python3 cluster_info_report.py --debug

# Remove stored credentials:
python3 cluster_info_report.py --clear-credentials
```

Output is auto-generated as `cluster_info_report_YYYYMMDD_HHMMSS.xlsx` in the current directory.

---

## CLI Options

| Option | Description |
|--------|-------------|
| `--apikey` | Helios API key (optional — keychain is used if omitted) |
| `--cluster` | Filter to one cluster by name |
| `--output` | Output `.xlsx` path (auto-generated if omitted) |
| `--debug` | Print raw cluster and node API responses |
| `--clear-credentials` | Remove the stored API key from the OS keychain and exit |

---

## Report Fields

One row per cluster.

| Column | Source field | Description |
|--------|-------------|-------------|
| Cluster Name | Helios cluster list | Cluster name as registered in Helios |
| Cluster ID | `id` | Cohesity cluster ID |
| Incarnation ID | `incarnationId` | Unique instance ID (changes on re-image) |
| Software Version | `clusterSoftwareVersion` | Installed Cohesity software version |
| Cluster Type | `clusterType` | `kPhysical`, `kVirtualRobo`, etc. |
| Node Count | count of node list | Total number of nodes in the cluster |
| Healthy Nodes | _(derived)_ | Nodes not flagged as marked-for-removal or in a known-unhealthy state |
| Domain Name | `domainNames[0]` | Primary DNS domain name |
| DNS Servers | `dnsServerIps` | Comma-separated list of configured DNS servers |
| NTP Servers | `ntpSettings.ntpServers` or `ntpServers` | Comma-separated list of configured NTP servers |
| Timezone | `timezone` | Cluster timezone (e.g. `America/Chicago`) |
| Encryption Enabled | `encryptionEnabled` | Whether data-at-rest encryption is enabled |
| FIPS Enabled | `fipsEnabled` | Whether FIPS 140-2 mode is enabled |
| Helios Connected | _(derived)_ | Always `Yes` — clusters returned from Helios are connected by definition |
