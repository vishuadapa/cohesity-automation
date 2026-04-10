#!/usr/bin/env python3
# =============================================================================
# DISCLAIMER: This code is provided on a best-effort basis and is not in any
# way officially supported or sanctioned by Cohesity. The code is intentionally
# kept simple to retain value as example code. The code in this repository is
# provided as-is and the author accepts no liability for damages resulting
# from its use.
# =============================================================================
"""
health_check_report.py  v1.44

Multi-cluster Cohesity health check — 19-sheet Excel workbook + Word document + editable topology diagram.
Designed for enterprise customer business reviews (EBRs) and SE trusted-advisor
engagements.  Gathers live data from Cohesity Helios (or directly from a single
cluster) and produces:

  Excel sheets
  ────────────
  1  Executive Summary       — per-cluster health score & grade
  2  Infrastructure          — versions, nodes, disks, config
  3  Node Hardware           — per-node model, serial, EOL status
  4  Disk Health             — per-disk status, SSD wear %, encryption
  5  Protection Health       — success rates, SLA, RPO gaps, DataLock per group
  6  Storage & Capacity      — utilization, data reduction, runway
  7  Policy Audit            — retention, replication, archival, DataLock per policy
  8  Policy → Groups         — every group with the policy that governs it
  9  Alerts                  — open critical/warning with age
  10 Security                — expanded checklist: encryption, FIPS, audit log, MFA,
                               NTP auth, tunnel, SSO, cert expiry, vault, FK, DataLock
  11 Agent Health            — per-host agent version, status, upgradability
  12 Source Coverage         — registered sources with protected/unprotected counts
  13 Replication/Archive     — targets, last transfer, FortKnox
  14 FortKnox Data Transfer  — per-group transfer to external targets (30d full / 1d quick)
  15 Data Services           — NAS views, quota utilization
  16 Coverage Gaps           — groups with no recent successful backup
  17 User Security           — user accounts, MFA, locked status, roles, last login
  18 Trends (30d)            — daily success rate + storage growth charts
  19 Recommendations         — prioritized action list

  Word document
  ─────────────
  Cover, Executive Summary, Environment, Protection, Storage,
  Security, Agent & Source Coverage, User Security, Recommendations,
  Methodology appendix.

Usage
─────
  # Helios — API key (key saved to OS keychain after first prompt)
  python3 health_check_report.py --apikey <key> [options]

  # Direct cluster — username/password + optional MFA (bypasses Helios)
  python3 health_check_report.py --cluster-host 10.1.2.3 --username admin [--domain LOCAL]

  NOTE: Helios does NOT expose a scriptable username/password login endpoint.
  Its web login uses OIDC/OAuth2 which cannot be driven from a script.
  Generate a Helios API key at:
    helios.cohesity.com → Settings → Access Management → API Keys

Options
───────
  --apikey KEY        Helios API key (stored in keychain if omitted)
  --cluster-host HOST Bypass Helios; connect directly to this cluster hostname/IP
  --username USER     Cluster username (required with --cluster-host)
  --password PASS     Password (prompted securely if omitted; saved to keychain)
  --domain DOMAIN     Auth domain — LOCAL or AD name (default: LOCAL)
  --mfa-code CODE     TOTP/OTP code for MFA-enabled cluster accounts
  --cluster NAME      (Helios mode) Limit to one cluster by partial name match
  --clear-credentials Remove stored credentials from keychain and exit
  --days N            Lookback window for alerts/runs (default 30)
  --customer NAME     Customer name printed on report cover
  --output PATH       Base path for output files (no extension)
  --quick             Skip per-group run history; use last-run only (faster)
  --excel-only        Skip Word document generation
  --word-only         Skip Excel generation
  --debug             Print raw API payloads for troubleshooting

Requirements
────────────
  Required:  pip install requests openpyxl python-docx
  Optional:  pip install matplotlib keyring

  The script checks all prerequisites at startup and reports missing
  packages with exact install commands.

Version history
───────────────
  1.44 (2026-04-10) — feat: Prereq check now always reports ALL five packages
                     (requests, openpyxl, python-docx, matplotlib, keyring)
                     in a full status table — required vs optional, installed
                     vs NOT FOUND. Previously matplotlib and keyring were only
                     reported as small NOTE lines when missing; they are now
                     first-class entries in the startup check every run.
  1.43 (2026-04-10) — fix: Standalone / single-file distribution on Windows.
                     import urllib3 failed with ModuleNotFoundError when
                     requests was not installed (urllib3 is a dep of requests,
                     not a top-level package). Now catches ImportError at the
                     requests import itself and prints a clear install command.
                     urllib3 warning suppression falls back to
                     requests.packages.urllib3 when urllib3 is not importable
                     as a standalone package.
                     fix: cohesity_auth and formatters imports wrapped in
                     try/except with full inline fallbacks — script now runs
                     as a single downloaded .py file without needing the repo's
                     utils/ directory alongside it.
  1.42 (2026-04-10) — fix: Direct cluster auth — get_auth_token() now tries
                     v1 endpoint first (POST /irisservices/api/v1/public/
                     accessTokens) then falls back to v2 (POST /v2/users/
                     sessions). The v1 endpoint has the widest compatibility
                     across Cohesity versions and account types; v2 returned
                     400 "KValidationError Access denied" on some clusters.
                     fix: collect_cluster() now stores "mode" in the cd dict
                     so downstream code (quorum N/A, Word security table)
                     correctly detects direct mode.
                     cohesity_auth.py v1.5.
  1.41 (2026-04-10) — feat: Startup prerequisite check. main() now validates
                     all required packages (requests, openpyxl, python-docx)
                     before execution and prints clear error messages with
                     exact pip install commands for any missing package.
                     Optional packages (matplotlib, keyring) produce a NOTE
                     at startup but do not block. README.md fully rewritten
                     with complete prerequisite table, all CLI options
                     (including direct-cluster auth flags), and updated
                     descriptions for Word topology (native shapes) and
                     security table (full RAG coloring).
  1.40 (2026-04-10) — feat: Word doc Security Posture table (section 5) now
                     color-codes ALL columns: red for "No"/missing, amber for
                     "Idle"/"Partial", green for "Yes"/"Active"/"Full"; score
                     column uses RAG (red <50, amber 50-74, green 75+).
                     feat: Environment Topology diagram replaced with native
                     Word DrawingML shapes — editable rounded rectangles
                     (#96C73D green) for source clusters, cloud shapes for
                     FortKnox (#1A5C3A), cylinders for archival, rounded rects
                     for replication (#0062B1). Connector arrows drawn between
                     each cluster and its targets. matplotlib PNG kept as
                     fallback only. Dark blue header bar removed from PNG.
                     Cluster color updated from #00B388 (teal) to #96C73D
                     (Cohesity green) globally. draw.io fonts 11→13 (clusters),
                     10→12 (targets); FK shape changed to cloud; box sizes
                     increased for readability.
  1.39 (2026-04-10) — fix: Executive Summary "30d Success %" now matches the
                     Trends tab. _success_stats() rebuilt to use the same data
                     source priority as _sheet_trends() (v1_runs → group_runs
                     fallback) and now returns the average of per-day success %
                     values (= average of Trends col H) instead of a flat run-
                     count ratio across the entire window.
  1.38 (2026-04-10) — fix: Environment Topology table (Word) FortKnox Vaults
                     column now uses _fk_status() (active/idle/none) instead of
                     a narrow vault-type check. FK vault names collected via the
                     same comprehensive _is_fk logic used elsewhere (krpaas,
                     kfortknox, kfort_knox, krpaasarchival, "fortknox" in name).
                     Configured-but-idle vaults now show vault name +
                     "(Configured — not sending data)" in amber. Active vaults
                     show vault name(s) on green background. "No" only when FK
                     is genuinely not configured.
  1.37 (2026-04-10) — fix: Quorum detection revised — Quorum is a Helios control-
                     plane feature; any enabled quorum group now marks ALL Helios-
                     connected clusters as quorum-enabled (per-cluster ID matching
                     was unreliable when clusterIdentifiers is unpopulated).
                     fix: DataLock label "None" → "No"; removed broken policy-only
                     fallback from _cluster_datalock() — when no group data exists
                     return "No" rather than claiming full/partial coverage.
                     feat: Word doc Security Posture table expanded from 8→10
                     columns: "Indelible (FK/DL)" split into "Indelible (FK)" and
                     "DataLock (WORM)"; "Admins w/o MFA" column added with red
                     cell highlight when count > 0. Appendix scoring text updated.
  1.36 (2026-04-10) — feat: Environment Topology diagram. New _topology_data()
                     helper extracts cluster/replication/archival/FortKnox graph
                     from all_data. _generate_topology_drawio() writes an editable
                     draw.io XML file (.drawio) alongside the Excel/Word output —
                     open with diagrams.net or the draw.io desktop app. Nodes:
                     clusters (Cohesity teal), replication targets (blue cluster
                     boxes), archival vaults (amber cylinders), FortKnox/RPaaS
                     (dark-green, dashed arrows). _render_topology_png() generates
                     a matplotlib PNG preview (requires matplotlib) embedded in the
                     Word doc after the topology table; caption references the
                     .drawio file. Cohesity brand colors used throughout (#00B388
                     teal, #0062B1 replication, #E07A00 archival, #1A5C3A FK).
  1.35 (2026-04-10) — fix: Quorum detection now fetches Helios-level quorum groups
                     via GET /v2/mcm/quorum/groups instead of probing per-cluster
                     info dict (which never had quorum fields). quorum_enabled is
                     resolved once for all clusters and stored in cd.
                     feat: FortKnox implicit DataLock — policies archiving to
                     FortKnox/RPaaS now show "FortKnox (Indelible)" (green) in
                     Policy Audit DataLock column and are not flagged as "No DataLock".
                     feat: Admin MFA elevated from MEDIUM → HIGH recommendation;
                     missing admin MFA now deducts 10 pts from security score.
  1.34 (2026-04-10) — feat: replace FIPS with split Cluster/SD Encryption columns
                     across Excel Security sheet, Infrastructure sheet, and Word doc.
                     DataLock coverage now per-group (partial compliance shown as
                     "X/Y groups"). Security score redesigned: 6 dimensions totalling
                     100 pts. Add Quorum column to Security sheet (N/A for direct
                     mode). Rename "Immutability"→"Indelible" throughout. Agent
                     upgrade recommendation uses softer "MAY cause issues" language
                     with compatibility doc URL. Word doc gains Environment Topology
                     table (clusters → replication → archival → FortKnox).
  1.33 (2026-04-09) — fix: Agent Health — hostname vs IP routing by content:
                     new _split_host_ip() helper inspects each value and routes
                     it to Host or IP Address column based on whether it matches
                     an IPv4/IPv6 pattern (not field name). Fixes empty Host
                     column when hostIp contains FQDNs.
                     fix: Infrastructure Domain (col K) — domainNames[0] returns
                     the full cluster FQDN (e.g. "mycluster.example.com"); strip
                     the cluster hostname prefix so only the domain is shown.
  1.32 (2026-04-09) — feat: DataLock (WORM) column in Protection Health sheet
                     (col 18 per group, green=Compliance/yellow=Administrative/
                     red=None); new "FortKnox Data Transfer" sheet (sheet 14)
                     showing per-protection-group transfer to every external
                     target: logical TB, physical TB, storage consumed TB,
                     snapshots. Full mode covers the full lookback window
                     (args.days); quick mode fetches a separate 1-day window.
                     Sheet count: 18 → 19.
  1.31 (2026-04-09) — polish: formatting improvements — Word doc headings and
                     cover title standardized to #70AD47 (was #00B388); Excel
                     header rows gain a thin black border on all sides and row
                     height 32; auto_fit_columns() handles multiline cell values
                     (measures longest line) with tuned defaults (min 12, max 45);
                     Trends charts moved to the right of data (col L) to prevent
                     overlap, height increased to 16, legend placed at bottom,
                     chart title rendered at 14 pt bold with overlay disabled.
                     formatters.py: auto_fit_columns defaults min=12, max=45.
  1.30 (2026-04-09) — feat: DataLock (WORM) tracking across Policy Audit,
                     Security, and Recommendations. New helper functions
                     _policy_datalock_str() and _cluster_datalock() read
                     backupPolicy.regular.retention.dataLockConfig from the
                     v2 policies API (already fetched — no new API call).
                     Policy Audit gains a "DataLock (WORM)" column (col 15,
                     green=Compliance, yellow=Administrative, red=None) with
                     "No DataLock" added to the Gaps cell when absent.
                     Security sheet gains a "DataLock (WORM)" column (col 16)
                     showing cluster-level rollup: Compliance / Administrative
                     / Mixed / None, with "No DataLock" added to the Gaps
                     column when absent. Recommendations: MEDIUM if no policies
                     have DataLock enabled; LOW if Administrative mode only
                     (urging upgrade to Compliance mode).
  1.29 (2026-04-09) — fix: removed Helios username/password auth mode —
                     Helios /irisservices/api/v1/mcm/login requires an OIDC
                     token in the Authorization header, not credentials in
                     the POST body. Helios access requires an API key
                     (generate at helios.cohesity.com → Settings → Access
                     Management → API Keys). --username/--password/--mfa-code
                     are now only supported with --cluster-host (direct
                     cluster mode). cohesity_auth.py bumped to v1.4.
  1.28 (2026-04-09) — feat: direct cluster mode (--cluster-host) bypasses
                     Helios; authenticates via POST /v2/users/sessions →
                     Bearer token. Cluster name auto-detected from cluster's
                     own /public/cluster endpoint. Added --clear-credentials
                     flag and --mfa-code for cluster MFA. _get() now reads
                     base URL from session.base_url. cohesity_auth.py v1.2.
  1.27 (2026-04-09) — fix: Disk Health sheet still empty after v1.26 — GET
                     /v2/disks is not a valid Helios-proxied endpoint.
                     Rewrote _disks() to use GET /v2/disks/local?nodeId={id}
                     per node, matching the approach used by the Cohesity ELF
                     inventory tool (cohesityInv.ps1, apiType='ClusterV2',
                     path='/disks/local?nodeId={0}'). Updated field fallbacks
                     throughout to include ELF field names: ssdUsedPercentage
                     (was ssdWearLevelPct), capacityInBytes (was capacityBytes),
                     encryptionStatus (was encryptionEnabled).
  1.26 (2026-04-09) — fix: Disk Health sheet empty — switched _disks() from the
                     private /v1/disks/local?nodeId=X endpoint (always 400 via
                     Helios proxy) to the supported GET /v2/disks endpoint;
                     private endpoint kept as fallback. Added isinstance guard
                     on usageStats sub-object in _sheet_disk_health().
  1.25 (2026-04-09) — chore: standardize on US English throughout — utilization,
                     prioritized, normalized, color-coded, behavior, organization,
                     authorized, summarizes, maximize, minimize, penalized, labeled.
                     Affects docstring, sheet titles, Word doc narrative, comments.
  1.24 (2026-04-09) — fix: Source Coverage sheet empty on older clusters —
                     /v2/data-protect/sources returns 404 on pre-7.x firmware.
                     _sources() now falls back to v1
                     /protectionSources/registrationInfo (rootNodes[].stats.
                     protectedCount / unprotectedCount) and normalizes the
                     result into the same dict shape used by the sheet and
                     recommendations engine. Also fixed isinstance guards on
                     stats / objectCount sub-objects in _sheet_source_coverage.
  1.23 (2026-04-09) — fix: suppress debug body output for best-effort endpoints
                     (_disks, _sources, _tenants) that return 400/403/404 on
                     clusters where the endpoint is unavailable or restricted.
                     Added `silent` param to _get(); these three callers pass
                     silent=True so --debug output is not cluttered with
                     expected error bodies.
  1.22 (2026-04-09) — fix: replace all unsafe `(x or {}).get()` patterns with
                     isinstance(x, dict) guards; affected fields: healthStatus
                     (could be a string in older cluster versions), auditLogConfig,
                     clusterAuditConfig, encryptionConfig, ntpSettings,
                     remoteSupportConfig, mfaConfig, sourceInfo (could be string
                     in some source types), stats (source coverage). Added
                     isinstance(d, dict) guards to disk comprehensions in
                     _build_recommendations(). Fixed _ag_st() helper in Word
                     doc section to handle str/dict/None healthStatus uniformly.
  1.21 (2026-04-08) — feat: Tier 1/2/3 ELF inventory additions. New data
                     collection: _certificates(), _agents(), _disks(),
                     _sources(), _users(), _idps(), _tenants(). New Excel
                     sheets: Disk Health (per-disk status/SSD wear/encryption),
                     Agent Health (agent version/status/upgradability), Source
                     Coverage (registered sources with protected/unprotected
                     ratio), User Security (MFA/locked/roles/last-login).
                     Security sheet expanded: audit log, NTP auth, remote
                     tunnel, cluster MFA, SSO/IDP, TLS cert expiry. New
                     recommendations: cert expiry CRITICAL/HIGH, unhealthy
                     agents HIGH, failed/missing disks CRITICAL, SSD wear HIGH,
                     unprotected sources HIGH, users without MFA MEDIUM, no SSO
                     LOW, audit log not enabled MEDIUM. Sheet count: 14 → 18.
  1.20 (2026-04-08) — fix: 6.8.1 LTS had EOS November 15 2024 (OOS) — was
                     incorrectly matched by "6.8" prefix as in-support; split
                     table so only 6.8.2+ maps to "In Support — LTS Jun 2026".
                     Hardware EOL table expanded: added C2100, C2200, C2510,
                     C4500 model variants confirmed in Cohesity EOL doc (Oct 2025).
  1.19 (2026-04-08) — feat: software and hardware lifecycle tracking. Added
                     SOFTWARE_EOS and HARDWARE_EOL reference tables from
                     Cohesity Product Life Cycle Policy (Oct 2025). Infrastructure
                     sheet gets 3 new columns: SW Lifecycle Status (color-coded
                     red/orange/green), SW EOS Date, Days to EOS. New "Node Hardware"
                     sheet (sheet 3): per-node model, serial, node type, capacity,
                     storage tiers, HW EOL date, HW EOL status color-coded.
                     Recommendations engine updated: CRITICAL for out-of-support
                     software or past-EOL hardware, HIGH/MEDIUM for upcoming dates.
                     Word document gains a Software & Hardware Lifecycle section.
                     Sheet count: 13 → 14.
  1.18 (2026-04-08) — fix: Trends tab still truncated after v1.17 — previous
                     pagination used `len(page) < PAGE_SIZE` to detect last
                     page, but the actual server cap varies per cluster version
                     and is often lower than 1000. That check fired immediately
                     on page 1 for busy clusters, stopping pagination early.
                     Fix: remove the page-size check entirely; page purely by
                     timestamp cursor until the response is empty, oldest run
                     predates the window start, or no progress is made. Safety
                     limit raised to 200 pages.
  1.17 (2026-04-08) — fix: Trends tab shows fewer than 30 days on clusters with
                     many protection groups — v1 protectionRuns caps per page;
                     added pagination by advancing endTimeUsecs cursor.
  1.16 (2026-04-08) — feat: Trends chart — visible X/Y axes with tick marks,
                     Y axis fixed 0-100 with "0.0" format, circle data-point
                     markers (size 5, Cohesity green), 2pt line weight, StrRef
                     category fix so date strings render as labels in Excel,
                     auto tick-density reduction for >14 and >60 day ranges,
                     chart title updated to note warnings are included.
  1.15 (2026-04-08) — feat: include kWarning runs in Success % calculation
                     throughout — Trends "Success % (incl. Warnings)" col H,
                     _score_protection() quick mode, _success_stats() quick
                     mode. Warning runs complete successfully (data protected)
                     but may have minor issues (skipped objects). They still
                     appear in the Warning column for visibility. Word doc
                     Methodology appendix now documents this behavior.
  1.14 (2026-04-08) — fix: Trends tab inaccurate — switched primary data source
                     from per-group v2 runs to v1 /public/protectionRuns which
                     is the same source powering the Helios reporting page.
                     Single call per cluster; backupRun.slaViolated (col I) and
                     backupRun.stats.totalLogicalBackupSizeBytes (col J) are
                     accurate. v2 per-group runs retained as fallback. Works in
                     --quick mode (v1 call not gated behind the per-group loop).
  1.13 (2026-04-08) — fix: Trends tab cols I and J always blank — col I
                     (SLA Violations) used run.get("isSlaViolated") but the
                     field lives inside localBackupInfo, not the top-level run
                     object; fixed to lb.get("isSlaViolated"). Col J (Logical GB)
                     used lb.get("stats").get("logicalSizeBytes") but the correct
                     sub-object is localSnapshotStats; fixed to
                     lb.get("localSnapshotStats").get("logicalSizeBytes").
  1.12 (2026-04-08) — fix: Replication & Archive col E and col K still empty
                     after v1.11 — policy→group chain produced no rows because
                     _arch_name() returns "" when v2 policies store archival
                     targets by vault ID only (no name/targetName field), so
                     policy_arch_targets stayed empty and all_arch_targets was
                     empty. Fix: supplement arch_groups directly from
                     fk_data.dataTransferPerProtectionJob[].protectionJobName
                     per vault (reliable even without policy resolution); build
                     all_arch_targets as union of policy chain + fk_data vault
                     names + vault list names so at least one source always
                     produces rows and E/K are populated.
  1.11 (2026-04-08) — feat: Replication & Archive col E changed from policy
                     count to protection group count; new col K lists all group
                     names using each target. Status updated: "Configured" when
                     groups are assigned, "No groups assigned" when not.
                     Rewired target discovery through policy→group chain so E
                     and K reflect actual group assignment, not policy count.
  1.10 (2026-04-08) — fix: Replication & Archive col E (Policies Using) showed
                     0 because _arch_name() returned "" when the v2 policy stores
                     archival targets by vault ID only (no name/targetName field).
                     Added vault ID → name resolution via /public/vaults lookup.
                     Columns G/H renamed to include API field names and note that
                     values are cumulative totals. "Vault only — no policy" status
                     reworded to "Vault exists — not referenced by any policy".
  1.9 (2026-04-08) — fix: Replication & Archive still blank after v1.8.
                     dataTransferToVaults API requires vaultIds filter to
                     return any rows — added vault ID extraction and pass-
                     through from _vaults() to _fortknox_data(). Vault type
                     column D now read from archivalTargetConfig.targetType
                     inside the policy target (previously name-based lookup
                     against vault list failed on every mismatch).
  1.8 (2026-04-08) — fix: Replication & Archive sheet columns D-H blank.
                     Root causes: (1) _fortknox_data used startTimeUsecs/
                     endTimeUsecs — API expects startTimeMsecs/endTimeMsecs,
                     so fk_data was always empty; (2) archival rows driven by
                     cd["vaults"] list whose name may differ from policy target
                     names — flipped to drive rows from arch_use (policy targets)
                     and enrich vault type/FK stats via name-then-ID lookup;
                     (3) G/H (Logical/Physical Transferred) hardcoded "" — now
                     populated from numLogicalBytesTransferred/numPhysical.
                     Orphaned vaults (in API but not in any policy) still shown.
  1.7 (2026-04-08) — fix: _fk_status() false-negatives — lastRun.archivalInfo
                     is only populated when the most recent primary run included
                     archival; groups with daily local backups + less-frequent
                     archival showed as idle even when active. Fix: scan full
                     group_runs history (non-quick mode) for successful FK
                     archival results matched by target type AND vault name;
                     lastRun remains the fallback for --quick mode only.
  1.6 (2026-04-08) — feat: detect FortKnox configured-but-idle state. New
                     _fk_status() returns "active"/"idle"/"none" by inspecting
                     group lastRun.archivalInfo (dataTransferToVaults storageConsumed
                     is cumulative and not reliable for recency). Security sheet
                     column 5 shows "Active" (green) / "Configured — Idle" (orange) /
                     "No" (red). Idle triggers a HIGH recommendation and scores
                     10/20 points instead of the full 20.
  1.5 (2026-04-08) — fix: drop Alerts columns K (Node ID) and L (Entity) —
                     sparse propertyList fields only present on hardware alerts;
                     description + alert code already capture the context.
  1.4 (2026-04-08) — fix: Policy Audit column N (Groups Using) now computed by
                     counting groups whose policyId matches each policy — the v2
                     API does not return numProtectionGroups.
                     feat: new "Policy → Groups" sheet (sheet 6) listing every
                     protection group with the policy that governs it, sorted by
                     policy then group name, paused/failed rows highlighted.
  1.3 (2026-04-08) — fix: infrastructure disk count now reads diskCount integer
                     per node (v1 /public/nodes field) instead of looking for
                     non-existent list fields; NTP servers shows "Not configured"
                     instead of blank when the cluster has no NTP entries; added
                     ntpServersList and ntpServerIps fallback paths.
  1.2 (2026-04-08) — fix: correct field names throughout — firstTimestampUsecs,
                     clusterSoftwareVersion, localSnapshotStats, successfulObjectsCount,
                     policyId→name lookup, v2 schedule/target/retention paths,
                     broadened FortKnox detection, N/A capacity fallback.
                     feat: output filename includes customer name + local timestamp.
  1.0 (2026-04-08) — feat: initial release. 12-sheet Excel + Word document.
                     Health scoring, recommendations engine, trend charts.
"""

__version__ = "1.43"

import argparse
import datetime
import os
import re
import sys
import time

# ── Core HTTP client ──────────────────────────────────────────────────────────
# requests must be present.  urllib3 is bundled inside requests; we suppress
# the SSL-verify warning through whichever path is available.
# If requests is missing we exit here with a clear install command rather than
# propagating a cryptic "No module named 'urllib3'" traceback.
try:
    import requests
    try:
        import urllib3 as _urllib3
        _urllib3.disable_warnings(_urllib3.exceptions.InsecureRequestWarning)
    except ImportError:
        requests.packages.urllib3.disable_warnings(
            requests.packages.urllib3.exceptions.InsecureRequestWarning)
except ImportError:
    print("=" * 65)
    print("  MISSING REQUIRED PACKAGE: requests")
    print()
    print("  Install it with:")
    print("    pip install requests")
    print()
    print("  Or install all required packages at once:")
    print("    pip install requests openpyxl python-docx")
    print("=" * 65)
    sys.exit(1)

# ── Auth + formatting helpers ─────────────────────────────────────────────────
# Try the repo's utils/ directory first (normal dev/clone usage).
# If the utils modules are not found (e.g. single-file standalone download),
# fall back to inline implementations so the script works out of the box.
_utils_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "utils")
sys.path.insert(0, _utils_dir)
try:
    from cohesity_auth import (
        get_api_key, get_helios_clusters, make_helios_headers, HELIOS_HOST,
        get_cluster_password, get_auth_token, make_headers,
        clear_stored_credentials,
    )
    from formatters import (
        bytes_to_gb, bytes_to_tb, usecs_to_datetime, auto_fit_columns,
    )
except ImportError:
    # ── Bundled fallbacks — used only when running as a standalone file ───────
    import getpass as _getpass

    HELIOS_HOST     = "helios.cohesity.com"
    _KR_SVC_HELIOS  = "cohesity_helios"
    _KR_USER_HELIOS = "apikey"
    _KR_SVC_CLUSTER = "cohesity_cluster"

    def _keyring_available():
        try:
            import keyring  # noqa: F401
            return True
        except ImportError:
            return False

    def get_api_key(cli_key=None):
        if cli_key:
            return cli_key
        if _keyring_available():
            import keyring
            stored = keyring.get_password(_KR_SVC_HELIOS, _KR_USER_HELIOS)
            if stored:
                print("[*] Using Helios API key stored in system keychain.")
                return stored
            print("[*] No stored Helios API key found.")
        else:
            print("    NOTE: 'keyring' not installed — key will not be saved.")
        key = _getpass.getpass("    Enter Helios API key: ").strip()
        if not key:
            print("ERROR: API key cannot be empty.")
            sys.exit(1)
        if _keyring_available():
            import keyring
            keyring.set_password(_KR_SVC_HELIOS, _KR_USER_HELIOS, key)
            print("[*] Helios API key saved to system keychain.")
        return key

    def get_cluster_password(cluster, username, domain, cli_password=None):
        if cli_password:
            return cli_password
        kr_user = f"{cluster}:{domain}:{username}"
        if _keyring_available():
            import keyring
            stored = keyring.get_password(_KR_SVC_CLUSTER, kr_user)
            if stored:
                print(f"[*] Using stored password for {domain}\\{username}@{cluster}")
                return stored
        pwd = _getpass.getpass(f"Password for {domain}\\{username}@{cluster}: ").strip()
        if not pwd:
            print("ERROR: Password cannot be empty.")
            sys.exit(1)
        if _keyring_available():
            import keyring
            keyring.set_password(_KR_SVC_CLUSTER, kr_user, pwd)
            print(f"[*] Password saved to keychain for {domain}\\{username}@{cluster}")
        return pwd

    def clear_stored_credentials(cluster=None, username="admin", domain="LOCAL"):
        if not _keyring_available():
            print("ERROR: 'keyring' not installed. Nothing to clear.")
            sys.exit(1)
        import keyring
        if cluster:
            kr_user = f"{cluster}:{domain}:{username}"
            if keyring.get_password(_KR_SVC_CLUSTER, kr_user):
                keyring.delete_password(_KR_SVC_CLUSTER, kr_user)
                print(f"[*] Stored password for {domain}\\{username}@{cluster} removed.")
            else:
                print(f"[*] No stored password found for {domain}\\{username}@{cluster}.")
        else:
            if keyring.get_password(_KR_SVC_HELIOS, _KR_USER_HELIOS):
                keyring.delete_password(_KR_SVC_HELIOS, _KR_USER_HELIOS)
                print("[*] Stored Helios API key removed.")
            else:
                print("[*] No stored Helios API key found — nothing to clear.")

    def get_auth_token(cluster, username, password, domain, mfa_code=None):
        _hdrs = {"Content-Type": "application/json", "Accept": "application/json"}
        url_v1 = f"https://{cluster}/irisservices/api/v1/public/accessTokens"
        try:
            r = requests.post(url_v1,
                              json={"username": username, "password": password,
                                    "domain": domain},
                              headers=_hdrs, verify=False, timeout=30)
            r.raise_for_status()
            data = r.json()
            token_val = data.get("accessToken", "")
            if token_val:
                print(f"[+] Authenticated to {cluster} as {domain}\\{username} (v1)")
                return f"{data.get('tokenType', 'Bearer')} {token_val}"
        except requests.exceptions.ConnectionError:
            print(f"ERROR: Cannot connect to '{cluster}'. Check hostname/IP and network.")
            sys.exit(1)
        except requests.exceptions.HTTPError:
            pass
        url_v2 = f"https://{cluster}/v2/users/sessions"
        pl2 = {"username": username, "password": password, "domain": domain}
        if mfa_code:
            pl2["otpCode"] = mfa_code
            pl2["otpType"] = "Totp"
        try:
            r = requests.post(url_v2, json=pl2, headers=_hdrs, verify=False, timeout=30)
            r.raise_for_status()
        except requests.exceptions.ConnectionError:
            print(f"ERROR: Cannot connect to '{cluster}'. Check hostname/IP and network.")
            sys.exit(1)
        except requests.exceptions.HTTPError as e:
            body = r.text[:400]
            if r.status_code == 401 and "otp" in body.lower():
                print("ERROR: Cluster requires MFA. Re-run with --mfa-code <code>.")
                sys.exit(1)
            print(f"ERROR: Authentication failed.\n  {e}\n  {body}")
            sys.exit(1)
        data = r.json()
        token_val = data.get("accessToken", "")
        if not token_val:
            print(f"ERROR: No access token in response: {r.text[:200]}")
            sys.exit(1)
        print(f"[+] Authenticated to {cluster} as {domain}\\{username} (v2)")
        return f"{data.get('tokenType', 'Bearer')} {token_val}"

    def make_headers(token):
        return {"Authorization": token, "Accept": "application/json",
                "Content-Type": "application/json"}

    def make_helios_headers(api_key, cluster_id=None):
        h = {"apiKey": api_key, "Accept": "application/json",
             "Content-Type": "application/json"}
        if cluster_id is not None:
            h["accessClusterId"] = str(cluster_id)
        return h

    def get_helios_clusters(api_key):
        url = f"https://{HELIOS_HOST}/mcm/clusters/connectionStatus"
        try:
            r = requests.get(url, headers=make_helios_headers(api_key),
                             verify=False, timeout=30)
            r.raise_for_status()
        except requests.exceptions.ConnectionError:
            print("ERROR: Cannot connect to Helios. Check your network/VPN.")
            sys.exit(1)
        except requests.exceptions.HTTPError:
            print(f"ERROR: Helios auth failed ({r.status_code}). Check your API key.")
            sys.exit(1)
        clusters = r.json()
        if not isinstance(clusters, list):
            clusters = clusters.get("clusterList", [])
        result = [{"name": c.get("name", ""), "clusterId": c.get("clusterId")}
                  for c in clusters]
        print(f"[+] Connected to Helios — {len(result)} cluster(s)")
        return result

    # ── Formatting helpers ────────────────────────────────────────────────────
    from datetime import datetime as _dt, timezone as _tz_utc

    def usecs_to_datetime(usecs, tz=None):
        if not usecs:
            return ""
        dt = _dt.fromtimestamp(usecs / 1_000_000, tz=_tz_utc.utc)
        if tz is not None:
            dt = dt.astimezone(tz)
        return dt.strftime("%Y-%m-%d %H:%M:%S")

    def bytes_to_gb(b):
        return round(b / 1_000_000_000, 4) if b else 0.0

    def bytes_to_tb(b):
        return round(b / 1_000_000_000_000, 4) if b else 0.0

    def auto_fit_columns(ws, min_width=12, max_width=45):
        try:
            from openpyxl.cell.cell import MergedCell
            from openpyxl.utils import get_column_letter as _gcl
        except ImportError:
            return
        for col in ws.columns:
            if not col:
                continue
            first = col[0]
            col_letter = (first.column_letter if hasattr(first, "column_letter")
                          else _gcl(first.column))
            best = min_width
            for cell in col:
                if isinstance(cell, MergedCell):
                    continue
                if cell.value is not None:
                    lines = str(cell.value).split("\n")
                    best = max(best, max(len(ln) for ln in lines) + 2)
            ws.column_dimensions[col_letter].width = min(best, max_width)

try:
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.chart import LineChart, Reference
    from openpyxl.chart.series import SeriesLabel
    EXCEL_OK = True
except ImportError:
    EXCEL_OK = False

try:
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    DOCX_OK = True
except ImportError:
    DOCX_OK = False

# ── Thresholds / best-practice baselines ─────────────────────────────────────
CAPACITY_WARN_PCT    = 70       # cluster used % — warning
CAPACITY_CRIT_PCT    = 85       # cluster used % — critical
SUCCESS_WARN_PCT     = 95.0     # backup success rate below this = warning
SUCCESS_CRIT_PCT     = 90.0     # backup success rate below this = critical
RPO_WARN_HOURS       = 24       # hours since last successful run — warning
RPO_CRIT_HOURS       = 48       # hours since last successful run — critical
ALERT_AGE_CRIT_DAYS  = 7        # open critical alert older than N days = finding
VIEW_QUOTA_WARN_PCT  = 80       # view at this % of quota = warning
VIEW_QUOTA_CRIT_PCT  = 90       # view at this % of quota = critical
VERSION_WARN_PREFIX  = "7."     # clusters below this major version get a flag
SSD_WEAR_WARN_PCT    = 70       # SSD wear % — warning
SSD_WEAR_CRIT_PCT    = 80       # SSD wear % — critical
CERT_WARN_DAYS       = 90       # TLS cert expiry warning threshold (days)
CERT_CRIT_DAYS       = 30       # TLS cert expiry critical threshold (days)

# ── Software End-of-Support matrix (checked top-to-bottom, first prefix match wins) ──
# Source: Cohesity Product Life Cycle Policy (updated Oct 2025)
SOFTWARE_EOS = [
    # (version_prefix, status, eos_date_str_or_None, display_label)
    # Evaluated in order — first prefix match wins; put longer/more-specific prefixes first.
    ("7.3",   "current",     None,         "Current"),
    ("7.2",   "in_support",  "2026-06-08", "In Support — Feature Release"),
    ("7.1.2", "in_support",  "2026-06-08", "In Support — LTS"),
    ("7.1",   "eol",         None,         "Out of Support"),   # 7.1.0 and 7.1.1
    ("7.0",   "eol",         None,         "Out of Support"),
    ("6.8.2", "in_support",  "2026-06-08", "In Support — LTS"),
    ("6.8",   "eol",         None,         "Out of Support"),   # 6.8.0 / 6.8.1 EOS Nov 2024
    # anything older falls through to the default → eol
]

# ── Hardware End-of-Life matrix (case-insensitive substring match on model string) ──
# Source: Cohesity Platforms EOL Terms and Service Express EOL database
HARDWARE_EOL = [
    # (model_substring, eol_date_str)
    # Source: Cohesity Products End of Support and End of Life Dates (Oct 2025)
    # C2xx0 series — EOL Jan 31, 2023
    ("C2000", "2023-01-31"), ("C2100", "2023-01-31"),
    ("C2200", "2023-01-31"), ("C2300", "2023-01-31"),
    # C2xx5 series — EOL May 13, 2024
    ("C2500", "2024-05-13"), ("C2505", "2024-05-13"), ("C2510", "2024-05-13"),
    # C3500 — EOL Feb 18, 2025
    ("C3500", "2025-02-18"),
    # C4xx0 / CN series — EOL Feb 22, 2027
    ("C4000", "2027-02-22"), ("C4300", "2027-02-22"), ("C4500", "2027-02-22"),
    ("CN3",   "2027-02-22"), ("CN4",   "2027-02-22"),
    # C60x5 — no EOL date announced as of Oct 2025 (omitted intentionally)
]
HW_EOL_WARN_DAYS = 180    # flag hardware EOL within this many days
SW_EOS_WARN_DAYS = 90     # flag software EOS within this many days

# Scoring weights — must sum to 1.0
_W = {
    "protection": 0.25,
    "sla":        0.20,
    "infra":      0.15,
    "capacity":   0.15,
    "security":   0.15,
    "alerts":     0.10,
}

# ── Excel palette ─────────────────────────────────────────────────────────────
COH_GREEN   = "70AD47"
DARK_GREEN  = "4E7A2E"
LT_GREEN    = "E2EFDA"
RED         = "FF4C4C"
ORANGE      = "FF8C00"
YELLOW      = "FFD700"
LIGHT_GRAY  = "F2F2F2"
WHITE       = "FFFFFF"


# ── openpyxl style helpers (lazy — only called when EXCEL_OK is True) ─────────
_IP4_RE = re.compile(r'^\d{1,3}(\.\d{1,3}){3}$')
_IP6_RE = re.compile(r'^[0-9a-fA-F]{0,4}(:[0-9a-fA-F]{0,4}){2,7}$')

def _split_host_ip(raw_host, raw_ip):
    """Route values to (hostname, ip_address) by content, not field name.

    The Cohesity agents API sometimes puts FQDNs in hostIp and leaves hostName
    empty. Inspect each value: if it looks like an IPv4/IPv6 address put it in
    the ip slot; otherwise put it in the host slot.
    """
    host = ""
    ip   = ""
    for val in (raw_host, raw_ip):
        val = (val or "").strip()
        if not val:
            continue
        if _IP4_RE.match(val) or _IP6_RE.match(val):
            if not ip:   ip   = val
        else:
            if not host: host = val
    return host, ip


def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def _font(bold=False, color="000000", size=10):
    return Font(bold=bold, color=color, size=size)

def _rag_fill(score):
    """Return fill based on 0-100 score."""
    if score >= 90: return _fill(LT_GREEN)
    if score >= 75: return _fill("FFFF99")
    if score >= 60: return _fill("FFD580")
    if score >= 40: return _fill("FFAA80")
    return _fill("FF8080")

def _hdr(ws, row, cols):
    """Write a bold Cohesity-green header row with a thin black border."""
    _thin   = Side(style="thin", color="000000")
    _border = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
    for c, val in enumerate(cols, 1):
        cell = ws.cell(row=row, column=c, value=val)
        cell.font      = Font(bold=True, color=WHITE, size=10)
        cell.fill      = _fill(COH_GREEN)
        cell.alignment = Alignment(wrap_text=True, vertical="center",
                                   horizontal="center")
        cell.border    = _border
    ws.row_dimensions[row].height = 32

def _title(ws, text, span):
    ws.merge_cells(f"A1:{span}1")
    c = ws["A1"]
    c.value     = text
    c.font      = Font(bold=True, color=WHITE, size=13)
    c.fill      = _fill(COH_GREEN)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22

def _reg(ws, row, col, val):
    cell = ws.cell(row=row, column=col, value=val)
    cell.font = _font()
    return cell

# ── Grade from score ──────────────────────────────────────────────────────────
def grade(score):
    if score >= 90: return "Excellent"
    if score >= 75: return "Good"
    if score >= 60: return "Fair"
    if score >= 40: return "At Risk"
    return "Critical"


def _sw_eos(version: str):
    """Return (status, eos_date_or_None, label) for a cluster software version string.
    status: 'current' | 'in_support' | 'eol' | 'unknown'
    """
    import datetime as _dt
    if not version:
        return ("unknown", None, "Unknown")
    v = version.strip().lower()
    for prefix, status, eos_str, label in SOFTWARE_EOS:
        if v.startswith(prefix.lower()):
            eos = _dt.date.fromisoformat(eos_str) if eos_str else None
            return (status, eos, label)
    # No match — treat anything pre-6.8 as out of support
    return ("eol", None, "Out of Support")


def _hw_eol(model: str):
    """Return eol_date (datetime.date) or None if model not in EOL matrix."""
    import datetime as _dt
    if not model:
        return None
    m = model.upper()
    for substr, eol_str in HARDWARE_EOL:
        if substr.upper() in m:
            return _dt.date.fromisoformat(eol_str)
    return None


# ── Time helpers ──────────────────────────────────────────────────────────────
def _now_usecs():
    return int(time.time() * 1_000_000)

def _days_ago_usecs(n):
    return int((time.time() - n * 86400) * 1_000_000)


# ─────────────────────────────────────────────────────────────────────────────
# API HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _get(session, headers, path, params=None, debug=False, silent=False):
    """GET from Helios or a direct cluster.

    The base URL is read from session.base_url (set in main() based on auth
    mode). Defaults to Helios if not set.

    silent=True suppresses the debug body print on non-200 responses — use for
    best-effort endpoints where 400/403/404 is expected on some clusters.
    """
    base = getattr(session, "base_url", f"https://{HELIOS_HOST}")
    url  = f"{base}{path}"
    try:
        r = session.get(url, headers=headers, params=params,
                        verify=False, timeout=60)
        if debug:
            print(f"    DEBUG {path} → {r.status_code}")
        if r.status_code == 200:
            return r.json()
        if debug and not silent:
            print(f"    DEBUG body: {r.text[:300]}")
        return None
    except Exception as exc:
        print(f"    WARN: GET {path} failed: {exc}")
        return None


# ─────────────────────────────────────────────────────────────────────────────
# DATA COLLECTION  (one function per resource type)
# ─────────────────────────────────────────────────────────────────────────────

def _cluster_info(session, h, debug):
    d = _get(session, h, "/irisservices/api/v1/public/cluster",
             params={"fetchStats": "true"}, debug=debug)
    return d or {}

def _nodes(session, h, debug):
    d = _get(session, h, "/irisservices/api/v1/public/nodes", debug=debug)
    return d or []

def _alerts(session, h, days, debug):
    d = _get(session, h, "/irisservices/api/v1/public/alerts",
             params={
                 "alertStateList": "kOpen",
                 "maxAlerts": 1000,
                 "startDateUsecs": _days_ago_usecs(days),
                 "endDateUsecs":   _now_usecs(),
             }, debug=debug)
    return d or []

def _protection_groups(session, h, debug):
    d = _get(session, h, "/v2/data-protect/protection-groups",
             params={
                 "isDeleted": "false",
                 "includeTenants": "true",
                 "includeLastRunInfo": "true",
             }, debug=debug)
    return (d or {}).get("protectionGroups") or []

def _group_runs(session, h, group_id, start_usecs, end_usecs, debug):
    d = _get(session, h,
             f"/v2/data-protect/protection-groups/{group_id}/runs",
             params={
                 "startTimeUsecs": start_usecs,
                 "endTimeUsecs":   end_usecs,
                 "numRuns": 500,
                 "includeTenants": "true",
             }, debug=debug)
    return (d or {}).get("runs") or []

def _v1_protection_runs(session, h, start_usecs, end_usecs, debug):
    """Paginated fetch of ALL protection runs in the time window via v1 API.

    The API returns runs newest-first. The actual per-page cap varies by
    cluster version; we do NOT rely on comparing len(page) to a constant
    because that assumption breaks when the server cap is lower than expected.
    Instead we page purely by timestamp cursor until:
      - the response is empty, or
      - the oldest run on the page predates our window start, or
      - no timestamp progress is made (safety against infinite loop).
    Safety limit: 200 pages.
    """
    all_runs = []
    page_end  = end_usecs

    for page_num in range(200):
        page = _get(session, h, "/irisservices/api/v1/public/protectionRuns",
                    params={"startTimeUsecs": start_usecs,
                            "endTimeUsecs":   page_end,
                            "numRuns":        1000},
                    debug=debug) or []
        if not page:
            break
        all_runs.extend(page)
        if debug:
            print(f"    DEBUG v1_runs page {page_num + 1}: {len(page)} runs "
                  f"(total so far: {len(all_runs)})")
        # Find the oldest startTimeUsecs on this page to use as the next cursor
        oldest_t = None
        for r in page:
            t = ((r.get("backupRun") or {}).get("stats") or {}).get("startTimeUsecs") or 0
            if t and (oldest_t is None or t < oldest_t):
                oldest_t = t
        if not oldest_t:
            break   # no timestamps found — can't advance cursor
        if oldest_t <= start_usecs:
            break   # full window covered
        if oldest_t >= page_end:
            break   # no progress — prevent infinite loop
        page_end = oldest_t - 1  # advance cursor to just before oldest run seen

    return all_runs

def _policies(session, h, debug):
    d = _get(session, h, "/v2/data-protect/policies",
             params={"includeTenants": "true"}, debug=debug)
    return (d or {}).get("policies") or []

def _storage_domains(session, h, debug):
    d = _get(session, h, "/v2/storage-domains",
             params={"includeStats": "true"}, debug=debug)
    return (d or {}).get("storageDomains") or []

def _views(session, h, debug):
    d = _get(session, h, "/v2/file-services/views",
             params={"includeStats": "true", "maxCount": 2000}, debug=debug)
    return (d or {}).get("views") or []

def _vaults(session, h, debug):
    d = _get(session, h, "/irisservices/api/v1/public/vaults",
             params={"includeFortKnoxVault": "true"}, debug=debug)
    return d or []


def _helios_quorum_groups(session, api_key, debug):
    """Fetch Helios-level quorum groups from /v2/mcm/quorum/groups.

    This is a Helios-scoped endpoint — do NOT include accessClusterId in headers.
    Returns a list of group dicts, each with:
      isEnabled          — True/False
      clusterIdentifiers — list of {"clusterId": <int>, ...}
    """
    h = make_helios_headers(api_key)   # no cluster_id — Helios-level call
    d = _get(session, h, "/v2/mcm/quorum/groups", debug=debug, silent=True)
    if isinstance(d, dict):
        return d.get("quorumGroups") or []
    if isinstance(d, list):
        return d
    return []


def _fortknox_data(session, h, start_usecs, end_usecs, vault_ids, debug):
    # API expects milliseconds; also requires vaultIds filter to return results
    params = {"startTimeMsecs": start_usecs // 1000,
              "endTimeMsecs":   end_usecs   // 1000}
    if vault_ids:
        params["vaultIds"] = vault_ids  # requests repeats list values correctly
    elif not vault_ids:
        return []   # no vaults — nothing to query
    d = _get(session, h,
             "/irisservices/api/v1/public/reports/dataTransferToVaults",
             params=params, debug=debug)
    return (d or {}).get("dataTransferSummary") or []


def _certificates(session, h, debug):
    """Fetch web server TLS certificate(s). Returns list of cert dicts only."""
    d = _get(session, h, "/irisservices/api/v1/public/certificates/webServer",
             debug=debug)
    if isinstance(d, list):
        return [c for c in d if isinstance(c, dict)]
    if isinstance(d, dict) and d:
        return [d]
    return []


def _agents(session, h, debug):
    """Fetch Cohesity agent deployment/health status per registered host."""
    d = _get(session, h, "/irisservices/api/v1/public/reports/agents",
             debug=debug)
    if isinstance(d, list):
        return d
    return (d or {}).get("agentDeploymentStatusList") or []


def _disks(session, h, nodes, debug):
    """Fetch per-disk info using GET /v2/disks/local?nodeId={id} per node.

    This mirrors the approach used by the Cohesity ELF inventory tool
    (cohesityInv.ps1, apiType='ClusterV2', path='/disks/local?nodeId={0}').
    Each node is queried individually; results are annotated with _nodeId
    and _nodeIp for display convenience.

    Field names returned by this endpoint (ELF script reference):
      serialNumber, model, status, location, type, encryptionStatus,
      ssdUsedPercentage (raw 0-100), removalProgressList, capacityInBytes
    """
    all_disks = []
    for n in nodes:
        nid = n.get("id")
        if not nid:
            continue
        d = _get(session, h, "/v2/disks/local",
                 params={"nodeId": nid}, debug=debug, silent=True)
        if not d:
            continue
        dl = d if isinstance(d, list) else (d.get("disksList") or d.get("disks") or [])
        for disk in dl:
            if not isinstance(disk, dict):
                continue
            disk["_nodeId"] = nid
            disk["_nodeIp"] = n.get("ip", "")
        all_disks.extend(dk for dk in dl if isinstance(dk, dict))
    return all_disks


def _sources(session, h, debug):
    """Fetch registered protection sources with coverage stats.

    Tries v2 first (includeTenants).  Falls back to v1
    /protectionSources/registrationInfo on older cluster firmware where the
    v2 endpoint returns 404.  Returns a list of normalized source dicts that
    _sheet_source_coverage and _build_recommendations can consume via the
    standard protectedObjectsCount / unprotectedObjectsCount keys.
    """
    d = _get(session, h, "/v2/data-protect/sources",
             params={"includeTenants": "true"}, debug=debug, silent=True)
    if isinstance(d, list) and d:
        return d
    if isinstance(d, dict):
        v2 = d.get("sources") or []
        if v2:
            return v2

    # v1 fallback — works on all supported cluster versions
    d1 = _get(session, h,
              "/irisservices/api/v1/public/protectionSources/registrationInfo",
              debug=debug)
    rows = d1 if isinstance(d1, list) else ((d1 or {}).get("rootNodes") or [])
    out = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        ps   = row.get("rootNode") or row.get("protectionSource") or {}
        ps   = ps if isinstance(ps, dict) else {}
        st   = row.get("stats") or {}
        st   = st if isinstance(st, dict) else {}
        out.append({
            "name":                    ps.get("name", ""),
            "environment":             ps.get("environment", ""),
            "protectedObjectsCount":   st.get("protectedCount")   or st.get("protectedObjectsCount")   or 0,
            "unprotectedObjectsCount": st.get("unprotectedCount") or st.get("unprotectedObjectsCount") or 0,
        })
    return out


def _users(session, h, debug):
    """Fetch local and domain user accounts."""
    d = _get(session, h, "/irisservices/api/v1/public/users",
             params={"allUnderHierarchy": "true"}, debug=debug)
    return d or []


def _idps(session, h, debug):
    """Fetch configured SSO / IDP providers."""
    d = _get(session, h, "/irisservices/api/v1/public/idps", debug=debug)
    return d or []


def _tenants(session, h, debug):
    """Fetch tenant / organization list (multi-tenancy inventory)."""
    d = _get(session, h, "/irisservices/api/v1/public/tenants", debug=debug, silent=True)
    return d or []


def collect_cluster(session, api_key, cluster, args):
    """Collect all health-check data for one cluster. Returns a dict.

    cluster dict shapes:
      Helios mode: {"name": "...", "clusterId": 123}
      Direct mode: {"name": "...", "clusterId": None, "mode": "direct",
                    "host": "10.1.2.3", "token": "Bearer ..."}
    """
    name = cluster.get("name", "Unknown")
    cid  = cluster.get("clusterId")
    dbg  = args.debug

    if cluster.get("mode") == "direct":
        h = make_headers(cluster["token"])
    else:
        h = make_helios_headers(api_key, cluster_id=cid)

    start_usecs = _days_ago_usecs(args.days)
    end_usecs   = _now_usecs()

    print(f"  [{name}] infrastructure...", end="", flush=True)
    info    = _cluster_info(session, h, dbg)
    nodes   = _nodes(session, h, dbg)
    print(" alerts...", end="", flush=True)
    alerts  = _alerts(session, h, args.days, dbg)
    print(" groups...", end="", flush=True)
    groups  = _protection_groups(session, h, dbg)
    print(" policies...", end="", flush=True)
    policies = _policies(session, h, dbg)
    print(" storage...", end="", flush=True)
    domains = _storage_domains(session, h, dbg)
    views   = _views(session, h, dbg)
    print(" vaults...", end="", flush=True)
    vaults    = _vaults(session, h, dbg)
    vault_ids = [v["id"] for v in vaults if "id" in v]
    fk_data    = _fortknox_data(session, h, start_usecs, end_usecs, vault_ids, dbg)
    # For the FortKnox detail tab: quick mode shows last 1 day only
    fk_data_1d = (_fortknox_data(session, h, _days_ago_usecs(1), end_usecs, vault_ids, dbg)
                  if args.quick else [])
    print(" run-summary...", end="", flush=True)
    v1_runs   = _v1_protection_runs(session, h, start_usecs, end_usecs, dbg)
    print(" certs/agents...", end="", flush=True)
    certs    = _certificates(session, h, dbg)
    agents   = _agents(session, h, dbg)
    print(" disks...", end="", flush=True)
    disks    = _disks(session, h, nodes, dbg)
    print(" sources...", end="", flush=True)
    sources  = _sources(session, h, dbg)
    print(" users/idps...", end="", flush=True)
    users    = _users(session, h, dbg)
    idps     = _idps(session, h, dbg)
    tenants  = _tenants(session, h, dbg)

    group_runs = {}
    if not args.quick:
        print(f" runs ({len(groups)} groups)...", end="", flush=True)
        for g in groups:
            gid = g.get("id")
            if gid:
                runs = _group_runs(session, h, gid, start_usecs, end_usecs, dbg)
                if runs:
                    group_runs[gid] = runs

    print(" done.")

    # Build policy ID → name lookup used throughout the report
    policy_map = {p.get("id"): p.get("name", "") for p in policies if p.get("id")}

    cd = {
        "name":        name,
        "id":          cid,
        "mode":        cluster.get("mode"),       # "direct" or None (Helios)
        "info":        info,
        "nodes":       nodes,
        "alerts":      alerts,
        "groups":      groups,
        "group_runs":  group_runs,
        "policies":    policies,
        "policy_map":  policy_map,
        "domains":     domains,
        "views":       views,
        "vaults":      vaults,
        "fk_data":     fk_data,
        "v1_runs":     v1_runs,
        "certs":       certs,
        "agents":      agents,
        "disks":       disks,
        "sources":     sources,
        "users":       users,
        "idps":        idps,
        "tenants":     tenants,
        "start_usecs":    start_usecs,
        "end_usecs":      end_usecs,
        "days":           args.days,
        "quick":          args.quick,
        "fk_data_1d":     fk_data_1d,
        "quorum_enabled": cluster.get("quorum_enabled", False),
    }
    cd["scores"]          = _compute_scores(cd, args.quick)
    cd["recommendations"] = _build_recommendations(cd)
    return cd


# ─────────────────────────────────────────────────────────────────────────────
# SCORING ENGINE
# ─────────────────────────────────────────────────────────────────────────────

def _score_protection(cd, quick):
    """0-100 based on backup success rate."""
    groups = cd["groups"]
    if not groups:
        return 100
    if quick:
        total   = len(groups)
        success = sum(
            1 for g in groups
            if (g.get("lastRun") or {}).get("localBackupInfo", {}).get("status", "")
               in ("kSuccess", "Succeeded", "kWarning")
        )
    else:
        all_runs = [r for runs in cd["group_runs"].values() for r in runs]
        if not all_runs:
            return 100
        total   = len(all_runs)
        success = sum(
            1 for r in all_runs
            if (r.get("localBackupInfo") or {}).get("status", "")
               in ("kSuccess", "Succeeded", "kWarning")
        )
    pct = (success / total * 100) if total else 100
    if pct >= SUCCESS_WARN_PCT: return 100
    if pct >= SUCCESS_CRIT_PCT: return int(70 + (pct - SUCCESS_CRIT_PCT) * 3)
    return max(0, int(pct))


def _score_sla(cd, quick):
    """0-100 based on SLA pass rate."""
    groups = cd["groups"]
    if not groups:
        return 100
    if quick:
        total = len(groups)
        viols = sum(1 for g in groups
                    if (g.get("lastRun") or {}).get("isSlaViolated", False))
    else:
        all_runs = [r for runs in cd["group_runs"].values() for r in runs]
        if not all_runs:
            return 100
        total = len(all_runs)
        viols = sum(1 for r in all_runs if r.get("isSlaViolated", False))
    pct_ok = ((total - viols) / total * 100) if total else 100
    if pct_ok >= 99: return 100
    if pct_ok >= 95: return 80
    if pct_ok >= 90: return 60
    if pct_ok >= 80: return 40
    return 20


def _score_infra(cd):
    """0-100 based on node health."""
    nodes = cd["nodes"]
    if not nodes:
        return 70
    healthy = sum(
        1 for n in nodes
        if ((n.get("nodeStatus") or {}).get("overallStatus", "") == "kNormal"
            or n.get("removalState") in (None, "kDontRemove"))
    )
    return int(healthy / len(nodes) * 100)


def _score_capacity(cd):
    """0-100 based on cluster capacity used %."""
    info   = cd["info"]
    usage  = ((info.get("stats") or {}).get("usagePerfStats") or {})
    usable = usage.get("physicalCapacityBytes") or 0
    used   = usage.get("totalPhysicalUsageBytes") or 0
    if not usable:
        return 70
    pct = used / usable * 100
    if pct <= CAPACITY_WARN_PCT:                return 100
    if pct <= CAPACITY_CRIT_PCT:                return int(100 - (pct - CAPACITY_WARN_PCT) * 2)
    return max(0, int(40 - (pct - CAPACITY_CRIT_PCT) * 2))


def _has_fortknox(vaults, policies):
    """Return True if FortKnox / RPaaS immutability is configured on this cluster."""
    fk_types = {"krpaas", "kfortknox", "kfort_knox", "krpaasarchival"}
    for v in vaults:
        vt = (v.get("vaultType") or "").lower()
        vn = (v.get("name") or "").lower()
        if any(t in vt for t in fk_types) or "fortknox" in vn:
            return True
    for p in policies:
        for t in ((p.get("remoteTargetPolicy") or {}).get("archivalTargets") or []):
            cfg = (t.get("archivalTargetConfig") or {})
            tt  = (cfg.get("targetType") or t.get("targetType") or "").lower()
            if any(x in tt for x in fk_types) or "fortknox" in tt:
                return True
    return False


def _fk_status(cd):
    """
    Return one of three states for FortKnox immutability:
      "active"  — FK configured AND a successful FK archival run exists within
                  the collection window (data is flowing)
      "idle"    — FK configured in vault/policy but no successful FK archival
                  found in run history (configured but not sending data)
      "none"    — FK not configured at all

    Detection strategy
    ──────────────────
    We CANNOT rely on lastRun.archivalInfo alone: that field is only populated
    when the most recent primary run included archival. Groups that run daily
    local backups but archive on a different schedule (e.g. weekly) will show
    empty archivalInfo in lastRun even when FK is working correctly.

    Primary source: group_runs (full run history, available in non-quick mode).
    Every run in the window is scanned for a successful archival result whose
    target type or name indicates FortKnox/RPaaS.

    Fallback (--quick mode): lastRun.archivalInfo on each group (less reliable
    — may miss groups whose last run was local-only).
    """
    vaults     = cd["vaults"]
    policies   = cd["policies"]
    groups     = cd["groups"]
    group_runs = cd.get("group_runs") or {}

    if not _has_fortknox(vaults, policies):
        return "none"

    fk_types = {"krpaas", "kfortknox", "kfort_knox", "krpaasarchival"}
    success_states = {"kSuccess", "Succeeded", "kSuccessWithWarning",
                      "SucceededWithWarning", "kWarning"}

    # Build set of FK vault names for name-based matching in archival results
    fk_vault_names = set()
    for v in vaults:
        vt = (v.get("vaultType") or "").lower()
        vn = (v.get("name") or "").lower()
        if any(t in vt for t in fk_types) or "fortknox" in vn:
            fk_vault_names.add(vn)
    for p in policies:
        for t in ((p.get("remoteTargetPolicy") or {}).get("archivalTargets") or []):
            cfg  = t.get("archivalTargetConfig") or {}
            name = (cfg.get("name") or cfg.get("vaultName")
                    or t.get("targetName") or "").lower()
            tt   = (cfg.get("targetType") or t.get("targetType") or "").lower()
            if any(x in tt for x in fk_types) or "fortknox" in tt:
                if name:
                    fk_vault_names.add(name)

    def _is_fk_success(run):
        """Return True if this run dict has a successful archival to an FK target."""
        arch_info = (run.get("archivalInfo") or {})
        for r in (arch_info.get("archivalTargetResults") or []):
            if r.get("status") not in success_states:
                continue
            # Match by archival target type
            ttype = (r.get("archivalTargetType")
                     or r.get("targetType") or "").lower()
            if any(x in ttype for x in fk_types) or "fortknox" in ttype:
                return True
            # Match by target name against known FK vault names
            tname = (r.get("targetName") or "").lower()
            if tname and (tname in fk_vault_names or "fortknox" in tname):
                return True
        return False

    # Primary: scan full run history (non-quick mode)
    for runs in group_runs.values():
        for run in runs:
            if _is_fk_success(run):
                return "active"

    # Fallback: check lastRun on each group (quick mode)
    for g in groups:
        if _is_fk_success(g.get("lastRun") or {}):
            return "active"

    return "idle"


def _score_security(cd):
    """0-100 security score across 6 dimensions (no FIPS — requires cluster rebuild).

    Cluster Encryption  15 pts — cluster-wide encryption at rest enabled
    SD Encryption       15 pts — all storage domains encrypted (0 if any unencrypted)
    Vault configured    15 pts — at least one archival target exists
    Replication         15 pts — at least one replication target exists
    FortKnox/Indelible  20 pts — 20=active, 10=configured-idle, 0=absent
    DataLock (WORM)     20 pts — 20=full Compliance, 15=full any mode,
                                  8=partial ≥50%, 3=partial <50%, 0=none
    Admin MFA penalty   -10 pts — deducted if any admin account lacks MFA
    """
    vaults   = cd["vaults"]
    policies = cd["policies"]
    groups   = cd["groups"]
    score    = 0

    # Cluster encryption (15 pts)
    cluster_enc, sd_enc = _sd_enc_status(cd)
    if cluster_enc:
        score += 15

    # Storage domain encryption (15 pts)
    if sd_enc is True:
        score += 15
    elif sd_enc is None and cluster_enc:
        score += 8   # No domain data but cluster enc is on — partial credit

    # Vault (15 pts)
    if vaults:
        score += 15

    # Replication (15 pts)
    if any((p.get("remoteTargetPolicy") or {}).get("replicationTargets")
           for p in policies):
        score += 15

    # FortKnox / Indelible archival (20 pts)
    fk = _fk_status({"vaults": vaults, "policies": policies, "groups": groups})
    if fk == "active":
        score += 20
    elif fk == "idle":
        score += 10

    # DataLock coverage (20 pts)
    dl_status, _, covered, total = _cluster_datalock(policies, groups)
    if dl_status == "full_compliance":
        score += 20
    elif dl_status == "full_any":
        score += 15
    elif dl_status == "partial" and total > 0:
        pct = covered / total
        score += 8 if pct >= 0.5 else 3

    # Admin MFA penalty (-10 pts if any admin lacks MFA)
    admin_roles = {"COHESITY_ADMIN", "Admin", "kAdmin", "kSuperAdmin"}
    if any(
        isinstance(u, dict)
        and not (u.get("mfaEnabled") or u.get("isMfaEnabled"))
        and bool(set(u.get("roles") or []) & admin_roles)
        for u in (cd.get("users") or [])
    ):
        score = max(0, score - 10)

    return score


def _score_alerts(cd):
    """0-100: deducted for open critical and warning alerts."""
    alerts = cd["alerts"]
    crits  = sum(1 for a in alerts if a.get("severity") == "kCritical")
    warns  = sum(1 for a in alerts if a.get("severity") == "kWarning")
    return max(0, 100 - crits * 10 - warns * 2)


def _compute_scores(cd, quick):
    s = {
        "protection": _score_protection(cd, quick),
        "sla":        _score_sla(cd, quick),
        "infra":      _score_infra(cd),
        "capacity":   _score_capacity(cd),
        "security":   _score_security(cd),
        "alerts":     _score_alerts(cd),
    }
    s["overall"] = round(sum(s[k] * _W[k] for k in _W), 1)
    s["grade"]   = grade(s["overall"])
    return s


# ─────────────────────────────────────────────────────────────────────────────
# RECOMMENDATIONS ENGINE
# ─────────────────────────────────────────────────────────────────────────────

_PRIORITY_ORDER = {"CRITICAL": 0, "HIGH": 1, "MEDIUM": 2, "LOW": 3}

def _r(priority, category, finding, action, impact):
    return {"priority": priority, "category": category,
            "finding": finding, "action": action, "impact": impact}


def _build_recommendations(cd):
    recs     = []
    info     = cd["info"]
    nodes    = cd["nodes"]
    alerts   = cd["alerts"]
    groups   = cd["groups"]
    policies = cd["policies"]
    views    = cd["views"]
    vaults   = cd["vaults"]
    usage    = ((info.get("stats") or {}).get("usagePerfStats") or {})

    # ── Infrastructure ────────────────────────────────────────────────────────

    # ── Software lifecycle ────────────────────────────────────────────────────
    import datetime as _dt
    today = _dt.date.today()
    sw_ver = (cd["info"].get("clusterSoftwareVersion") or "").strip()
    sw_status, sw_eos_date, sw_label = _sw_eos(sw_ver)
    if sw_status == "eol":
        recs.append(_r("CRITICAL", "Software Lifecycle",
             f"Software version {sw_ver or '(unknown)'} is out of support — "
             "security patches and bug fixes are no longer available",
             "Upgrade to a supported release (7.1.2_u2 LTS, 7.2.x, or later) immediately",
             "Running unsupported software exposes the cluster to unpatched vulnerabilities"))
    elif sw_status == "in_support" and sw_eos_date:
        days_left = (sw_eos_date - today).days
        if days_left < SW_EOS_WARN_DAYS:
            recs.append(_r("HIGH", "Software Lifecycle",
                 f"Software version {sw_ver} reaches end of support in {days_left} days ({sw_eos_date})",
                 "Plan and schedule upgrade before EOS date to maintain support eligibility",
                 "Post-EOS clusters are ineligible for Cohesity technical support"))
        elif days_left < 365:
            recs.append(_r("MEDIUM", "Software Lifecycle",
                 f"Software version {sw_ver} EOS in {days_left} days ({sw_eos_date})",
                 "Begin upgrade planning — schedule upgrade within next quarter",
                 "Proactive planning avoids rushed upgrades near the EOS deadline"))

    # ── Hardware lifecycle ────────────────────────────────────────────────────
    hw_eol_past   = []
    hw_eol_soon   = []
    for n in cd["nodes"]:
        model = n.get("hardwareModel") or n.get("productModel") or ""
        eol   = _hw_eol(model)
        if not eol:
            continue
        days_left = (eol - today).days
        if days_left < 0 and model not in hw_eol_past:
            hw_eol_past.append(model)
        elif days_left <= HW_EOL_WARN_DAYS and model not in hw_eol_soon:
            hw_eol_soon.append(model)

    if hw_eol_past:
        models = ", ".join(sorted(set(hw_eol_past)))
        recs.append(_r("CRITICAL", "Hardware Lifecycle",
             f"Hardware model(s) {models} are past EOL — no longer eligible for manufacturer support",
             "Initiate hardware refresh or extended support contract immediately",
             "EOL hardware cannot be replaced under standard support; failure risk increases"))
    if hw_eol_soon:
        models = ", ".join(sorted(set(hw_eol_soon)))
        recs.append(_r("HIGH", "Hardware Lifecycle",
             f"Hardware model(s) {models} reaching EOL within {HW_EOL_WARN_DAYS} days",
             "Begin hardware refresh planning — lead times for Cohesity hardware can be 8–16 weeks",
             "Planning ahead avoids gaps in hardware support coverage"))

    unhealthy = [n for n in nodes
                 if not ((n.get("nodeStatus") or {}).get("overallStatus") == "kNormal"
                         or n.get("removalState") in (None, "kDontRemove"))]
    if unhealthy:
        recs.append(_r("CRITICAL", "Infrastructure",
            f"{len(unhealthy)} node(s) in unhealthy state",
            "Check hardware health, disk status, and cluster logs immediately",
            "Unhealthy nodes risk data availability and backup failures"))

    # ── Capacity ──────────────────────────────────────────────────────────────
    usable = usage.get("physicalCapacityBytes") or 0
    used   = usage.get("totalPhysicalUsageBytes") or 0
    if usable:
        cap_pct = used / usable * 100
        if cap_pct >= CAPACITY_CRIT_PCT:
            recs.append(_r("CRITICAL", "Capacity",
                f"Cluster at {cap_pct:.1f}% capacity — critically high",
                "Immediately expand capacity or archive/delete cold data",
                "Backup jobs fail when capacity is exhausted"))
        elif cap_pct >= CAPACITY_WARN_PCT:
            recs.append(_r("HIGH", "Capacity",
                f"Cluster at {cap_pct:.1f}% capacity — approaching limit",
                "Plan capacity expansion or data archival within 30-60 days",
                "Continued growth risks backup failures within weeks"))

    # ── Security ─────────────────────────────────────────────────────────────
    cluster_enc, sd_enc = _sd_enc_status(cd)
    if not cluster_enc:
        recs.append(_r("HIGH", "Security",
            "Cluster-wide encryption at rest is not enabled",
            "Enable cluster-wide encryption during cluster setup (requires rebuild if "
            "not done at creation time); enable storage domain encryption as an interim step",
            "Unencrypted data on physical media is exposed if drives are removed or stolen"))
    if sd_enc is False:
        recs.append(_r("HIGH", "Security",
            "One or more storage domains are not encrypted",
            "Enable encryption on all storage domains under cluster security settings",
            "Unencrypted storage domains expose data on physical media removal"))
    elif sd_enc is None and not cluster_enc:
        recs.append(_r("MEDIUM", "Security",
            "Unable to verify storage domain encryption — cluster encryption is off",
            "Review storage domain encryption settings and enable where missing",
            "Unencrypted storage domains expose data if media is removed"))

    if not vaults:
        recs.append(_r("MEDIUM", "Security",
            "No vault (archival target) is configured",
            "Configure at least one external vault for off-site retention",
            "No off-site copy of data — single point of failure"))

    fk_st = _fk_status(cd)
    if fk_st == "none":
        recs.append(_r("MEDIUM", "Security",
            "No indelible/FortKnox archival target detected",
            "Configure FortKnox or WORM archival to protect against ransomware",
            "Without an indelible copy, all backups can be deleted or encrypted by attackers"))
    elif fk_st == "idle":
        recs.append(_r("HIGH", "Security",
            "FortKnox configured but no recent successful archival detected",
            "Verify protection group policies and schedules targeting FortKnox; "
            "check for failed archival runs and resolve blocking issues",
            "Indelible copy is configured but data is not flowing — "
            "ransomware protection gap despite vault being present"))

    dl_status, dl_label, dl_covered, dl_total = _cluster_datalock(policies, groups)
    if dl_status == "none":
        recs.append(_r("MEDIUM", "Security",
            "No protection groups have DataLock (WORM) enabled",
            "Enable DataLock on protection policies — prefer Compliance mode to "
            "prevent any user or admin from deleting or shortening backup snapshots",
            "Without DataLock, local backups can be deleted or modified by ransomware "
            "or a compromised admin; FortKnox alone does not protect local snapshots"))
    elif dl_status == "partial":
        ungrouped = dl_total - dl_covered
        recs.append(_r("MEDIUM", "Security",
            f"DataLock partially configured — {dl_covered}/{dl_total} groups protected; "
            f"{ungrouped} group(s) have no DataLock on their governing policy",
            "Review the Policy Audit tab and enable DataLock on policies covering "
            "the unprotected groups; Compliance mode provides the strongest protection",
            "Partial DataLock coverage leaves unprotected groups vulnerable to "
            "admin-level or ransomware deletion of local snapshots"))
    elif dl_status == "full_any":
        recs.append(_r("LOW", "Security",
            "DataLock covers all groups but some policies use Administrative mode — "
            "Compliance mode provides stronger protection",
            "Upgrade DataLock from Administrative to Compliance mode on all policies",
            "Administrative mode DataLock can be overridden by a Cohesity admin; "
            "Compliance mode is indelible even to admins"))

    has_repl = any((p.get("remoteTargetPolicy") or {}).get("replicationTargets")
                   for p in policies)
    if not has_repl:
        recs.append(_r("MEDIUM", "Resilience",
            "No policies have a replication target configured",
            "Add replication to a secondary cluster for site-level resilience",
            "Single-site data is vulnerable to site-wide outage"))

    # ── Alerts ────────────────────────────────────────────────────────────────
    crits = [a for a in alerts if a.get("severity") == "kCritical"]
    if crits:
        old = [a for a in crits
               if (_now_usecs() - (a.get("firstOccurrenceTime") or _now_usecs()))
                  > ALERT_AGE_CRIT_DAYS * 86400 * 1_000_000]
        sev = "CRITICAL" if old else "HIGH"
        msg = f"{len(crits)} open critical alert(s)"
        if old:
            msg += f", {len(old)} older than {ALERT_AGE_CRIT_DAYS} days"
        recs.append(_r(sev, "Availability",
            msg,
            "Review and resolve open critical alerts; investigate root causes",
            "Unresolved critical alerts indicate active system issues"))

    # ── Protection ───────────────────────────────────────────────────────────
    paused = [g for g in groups if g.get("isPaused")]
    if paused:
        recs.append(_r("HIGH", "Protection",
            f"{len(paused)} protection group(s) are paused",
            f"Review and resume: {', '.join(g['name'] for g in paused[:3])}",
            "Paused groups do not create backups — data is unprotected"))

    rpo_gaps = []
    for g in groups:
        lr      = g.get("lastRun") or {}
        lb      = lr.get("localBackupInfo") or {}
        start_u = lb.get("startTimeUsecs") or 0
        if not start_u:
            continue
        hrs = (_now_usecs() - start_u) / 3_600_000_000
        if hrs > RPO_CRIT_HOURS:
            rpo_gaps.append((g.get("name", "?"), hrs))
    if rpo_gaps:
        examples = ", ".join(f"{n} ({h:.0f}h)" for n, h in rpo_gaps[:3])
        recs.append(_r("HIGH", "Protection",
            f"{len(rpo_gaps)} group(s) with >{RPO_CRIT_HOURS}h RPO gap",
            f"Investigate: {examples}" + (" and more" if len(rpo_gaps) > 3 else ""),
            "Objects may be unrecoverable to a recent point in time"))

    # ── Policies ─────────────────────────────────────────────────────────────
    no_arch = [p for p in policies
               if not (p.get("remoteTargetPolicy") or {}).get("archivalTargets")]
    if policies and len(no_arch) > len(policies) * 0.5:
        recs.append(_r("MEDIUM", "Governance",
            f"{len(no_arch)}/{len(policies)} policies lack an archival target",
            "Add archival targets for long-term retention and ransomware recovery",
            "Without archival, data is unrecoverable after local retention expires"))

    # ── Views ────────────────────────────────────────────────────────────────
    no_quota = [v for v in views
                if not (v.get("logicalQuota") or
                        (v.get("storagePolicy") or {}).get("storageQuotaPolicy"))]
    if no_quota:
        recs.append(_r("LOW", "Governance",
            f"{len(no_quota)} NAS view(s) have no storage quota",
            "Apply quotas to all views to prevent uncontrolled capacity growth",
            "Unquoted views can silently consume all available capacity"))

    near_quota = []
    for v in views:
        lq  = v.get("logicalQuota") or \
              (v.get("storagePolicy") or {}).get("storageQuotaPolicy") or {}
        q   = lq.get("hardLimitBytes") or 0
        u   = ((v.get("stats") or {}).get("dataUsageStats") or {}) \
                .get("totalLogicalUsageBytes") or 0
        if q and u and (u / q * 100) >= VIEW_QUOTA_CRIT_PCT:
            near_quota.append(v.get("name", "?"))
    if near_quota:
        recs.append(_r("HIGH", "Capacity",
            f"{len(near_quota)} view(s) at ≥{VIEW_QUOTA_CRIT_PCT}% of quota",
            f"Expand quota or clean up: {', '.join(near_quota[:3])}",
            "Views at quota block all new writes to that share"))

    # ── TLS Certificate expiry ────────────────────────────────────────────────
    for cert in (cd.get("certs") or []):
        if not isinstance(cert, dict):
            continue
        exp_usecs = (cert.get("expiryUsecs")
                     or cert.get("expiryTimeMsecs")
                     or cert.get("notAfter"))
        if not exp_usecs:
            continue
        # Handle usecs (>1e15) vs msecs (>1e12) vs seconds
        if exp_usecs > 1_000_000_000_000_000:
            exp_dt = _dt.date.fromtimestamp(exp_usecs / 1_000_000)
        elif exp_usecs > 1_000_000_000_000:
            exp_dt = _dt.date.fromtimestamp(exp_usecs / 1_000)
        else:
            exp_dt = _dt.date.fromtimestamp(exp_usecs)
        days_left = (exp_dt - today).days
        cn = (cert.get("commonName") or cert.get("subject")
              or cert.get("certName") or "cluster TLS certificate")
        if days_left < 0:
            recs.append(_r("CRITICAL", "Certificate",
                f"TLS certificate '{cn}' expired {abs(days_left)} days ago",
                "Replace the expired certificate immediately to avoid browser/API errors",
                "Expired certificates break HTTPS access and API connectivity"))
        elif days_left < CERT_CRIT_DAYS:
            recs.append(_r("CRITICAL", "Certificate",
                f"TLS certificate '{cn}' expires in {days_left} days ({exp_dt})",
                "Renew or replace the certificate before expiry to prevent service disruption",
                "Expired certificates break HTTPS access and API connectivity"))
        elif days_left < CERT_WARN_DAYS:
            recs.append(_r("HIGH", "Certificate",
                f"TLS certificate '{cn}' expires in {days_left} days ({exp_dt})",
                "Plan certificate renewal now — lead time for CA-signed certs can be 2–4 weeks",
                "Certificate expiry causes browser warnings and may break automated API calls"))

    # ── Agent health ──────────────────────────────────────────────────────────
    unhealthy_agents = []
    for agent in (cd.get("agents") or []):
        if not isinstance(agent, dict):
            continue
        _hs = agent.get("healthStatus")
        st = ((isinstance(_hs, dict) and _hs.get("status"))
              or (isinstance(_hs, str) and _hs)
              or agent.get("agentStatus") or agent.get("status") or "")
        if st.lower() in ("kunhealthy", "unhealthy", "unreachable", "kfailed", "failed"):
            _rh, _ri = _split_host_ip(
                agent.get("hostName") or agent.get("host") or "",
                agent.get("hostIp") or "")
            host = _rh or _ri or "unknown"
            unhealthy_agents.append(host)
    if unhealthy_agents:
        examples = ", ".join(unhealthy_agents[:3])
        recs.append(_r("HIGH", "Agent Health",
            f"{len(unhealthy_agents)} agent(s) unhealthy or unreachable: {examples}"
            + (" and more" if len(unhealthy_agents) > 3 else ""),
            "Check agent connectivity, firewall rules, and reinstall if necessary",
            "Unhealthy agents silently prevent backups from running on affected hosts"))

    upgradable_agents = []
    for agent in (cd.get("agents") or []):
        if not isinstance(agent, dict):
            continue
        upg = (agent.get("upgradability") or agent.get("upgradable") or "")
        if "upgradable" in upg.lower():
            _rh, _ri = _split_host_ip(
                agent.get("hostName") or agent.get("host") or "",
                agent.get("hostIp") or "")
            upgradable_agents.append(_rh or _ri or "unknown")
    if upgradable_agents:
        examples = ", ".join(upgradable_agents[:3])
        recs.append(_r("MEDIUM", "Agent Health",
            f"{len(upgradable_agents)} agent(s) are running older versions and may be "
            f"upgraded: {examples}" + (" and more" if len(upgradable_agents) > 3 else ""),
            "Review agent compatibility matrix and upgrade agents where appropriate — "
            "see https://docs.cohesity.com/7_4/Web/UserGuide/Content/ReleaseNotes/"
            "agent-compatibility.htm for supported back-versions",
            "Outdated agents MAY cause issues with newer cluster features; Cohesity "
            "supports a limited range of back-versions — check compatibility docs"))

    # ── Disk health and SSD wear ──────────────────────────────────────────────
    failed_disks = [d for d in (cd.get("disks") or [])
                    if isinstance(d, dict)
                    and (d.get("diskStatus") or d.get("status") or "").lower()
                    in ("kfailed", "failed", "kmissing", "missing")]
    if failed_disks:
        recs.append(_r("CRITICAL", "Disk Health",
            f"{len(failed_disks)} disk(s) in failed or missing state",
            "Replace failed disks immediately and verify cluster redundancy",
            "Failed disks reduce storage redundancy and risk data loss"))

    high_wear = [(d.get("_nodeIp", "?"),
                  d.get("ssdUsedPercentage") or d.get("ssdWearLevelPct")
                  or d.get("wearLevel") or d.get("lifeUsedPct") or 0)
                 for d in (cd.get("disks") or [])
                 if isinstance(d, dict)
                 and (d.get("ssdUsedPercentage") or d.get("ssdWearLevelPct")
                      or d.get("wearLevel") or d.get("lifeUsedPct") or 0)
                 >= SSD_WEAR_CRIT_PCT]
    if high_wear:
        examples = ", ".join(f"{ip} ({w}%)" for ip, w in high_wear[:3])
        recs.append(_r("HIGH", "Disk Health",
            f"{len(high_wear)} SSD(s) with ≥{SSD_WEAR_CRIT_PCT}% wear: {examples}",
            "Plan SSD replacement before wear reaches 100% to avoid sudden failure",
            "Worn SSDs have elevated failure risk and may degrade cluster performance"))

    # ── Source coverage ───────────────────────────────────────────────────────
    low_coverage = []
    for src in (cd.get("sources") or []):
        if not isinstance(src, dict):
            continue
        _si = src.get("sourceInfo")
        si   = _si if isinstance(_si, dict) else src
        name = si.get("name") or si.get("sourceName") or src.get("name") or ""
        _src_stats = src.get("stats")
        prot   = (src.get("protectedObjectsCount")
                  or src.get("numProtectedObjects")
                  or (isinstance(_src_stats, dict) and _src_stats.get("protectedObjectsCount")) or 0)
        unprot = (src.get("unprotectedObjectsCount")
                  or src.get("numUnprotectedObjects")
                  or (isinstance(_src_stats, dict) and _src_stats.get("unprotectedObjectsCount")) or 0)
        total  = prot + unprot
        if total > 0 and unprot > 0 and (unprot / total * 100) >= 25:
            low_coverage.append(f"{name} ({unprot}/{total} unprotected)")
    if low_coverage:
        examples = ", ".join(low_coverage[:3])
        recs.append(_r("HIGH", "Coverage",
            f"{len(low_coverage)} source(s) have ≥25% unprotected objects: {examples}"
            + (" and more" if len(low_coverage) > 3 else ""),
            "Review unprotected objects in each source and add them to a protection group",
            "Unprotected objects are unrecoverable in the event of loss or corruption"))

    # ── User security ─────────────────────────────────────────────────────────
    admin_roles = {"COHESITY_ADMIN", "Admin", "kAdmin", "kSuperAdmin"}
    users_no_mfa = [
        u.get("username", "?") for u in (cd.get("users") or [])
        if isinstance(u, dict)
        and not (u.get("mfaEnabled") or u.get("isMfaEnabled"))
        and bool(set(u.get("roles") or []) & admin_roles)
    ]
    if users_no_mfa:
        examples = ", ".join(users_no_mfa[:3])
        recs.append(_r("HIGH", "Security",
            f"{len(users_no_mfa)} admin user(s) without MFA: {examples}"
            + (" and more" if len(users_no_mfa) > 3 else ""),
            "Enable MFA for all administrator accounts in Helios → Access Management",
            "Admin accounts without MFA are the most common ransomware entry point; "
            "a stolen credential gives full cluster access without MFA as a second factor"))

    # ── Audit log ─────────────────────────────────────────────────────────────
    info_local = cd["info"]
    _audit_cfg     = info_local.get("auditLogConfig")
    _cluster_audit = info_local.get("clusterAuditConfig")
    audit_enabled = bool(
        info_local.get("auditLogEnabled")
        or (isinstance(_audit_cfg, dict)     and _audit_cfg.get("enabled"))
        or (isinstance(_cluster_audit, dict) and _cluster_audit.get("enabled"))
    )
    if not audit_enabled:
        recs.append(_r("MEDIUM", "Security",
            "Cluster audit logging is not enabled",
            "Enable audit logging in cluster security settings to capture all admin actions",
            "Without audit logs, unauthorized changes cannot be detected or investigated"))

    # ── SSO / IDP ─────────────────────────────────────────────────────────────
    if not cd.get("idps"):
        recs.append(_r("LOW", "Security",
            "No SSO / IDP is configured",
            "Configure SAML or LDAP-based SSO for centralized identity management",
            "Without SSO, local accounts with individual passwords increase credential risk"))

    recs.sort(key=lambda x: _PRIORITY_ORDER.get(x["priority"], 99))
    return recs


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL SHEETS
# ─────────────────────────────────────────────────────────────────────────────

def _success_stats(cd):
    """Return (succ_pct, sla_pct) from run history.

    succ_pct: average of per-day Success % (incl. Warnings) — identical to
              the average of column H in the Trends sheet, so the Executive
              Summary stays consistent with the Trends tab.
    sla_pct:  percentage of total runs without an SLA violation.

    Data source priority mirrors _sheet_trends():
      1. v1_runs  — Helios protectionRuns (preferred; same as Helios reports)
      2. group_runs — v2 per-group run history (fallback when v1 unavailable)
      3. lastRun per group — quick-mode last resort
    """
    # ── Build daily buckets (same logic as _sheet_trends) ────────────────────
    daily = {}
    for run in (cd.get("v1_runs") or []):
        br  = run.get("backupRun") or {}
        su  = (br.get("stats") or {}).get("startTimeUsecs") or 0
        if not su:
            continue
        day = datetime.datetime.fromtimestamp(
            su / 1_000_000, tz=datetime.timezone.utc).strftime("%Y-%m-%d")
        d  = daily.setdefault(day, dict(total=0, succ=0, warn=0, viols=0))
        st = br.get("status", "")
        d["total"] += 1
        if st == "kSuccess":   d["succ"]  += 1
        elif st == "kWarning": d["warn"]  += 1
        if br.get("slaViolated"): d["viols"] += 1

    if not daily:
        for runs in cd["group_runs"].values():
            for run in runs:
                lb = run.get("localBackupInfo") or {}
                su = lb.get("startTimeUsecs") or 0
                if not su:
                    continue
                day = datetime.datetime.fromtimestamp(
                    su / 1_000_000, tz=datetime.timezone.utc).strftime("%Y-%m-%d")
                d  = daily.setdefault(day, dict(total=0, succ=0, warn=0, viols=0))
                st = lb.get("status", "")
                d["total"] += 1
                if st in ("kSuccess", "Succeeded"): d["succ"]  += 1
                elif st == "kWarning":              d["warn"]  += 1
                if lb.get("isSlaViolated"): d["viols"] += 1

    # ── Average of per-day success % (= average of Trends col H) ─────────────
    if daily:
        day_pcts = [(d["succ"] + d["warn"]) / d["total"] * 100
                    for d in daily.values() if d["total"]]
        succ_pct    = round(sum(day_pcts) / len(day_pcts), 1) if day_pcts else 100
        total_runs  = sum(d["total"] for d in daily.values())
        total_viols = sum(d["viols"] for d in daily.values())
        sla_pct     = round((total_runs - total_viols) / total_runs * 100, 1) if total_runs else 100
        return succ_pct, sla_pct

    # ── quick-mode last resort: last-run per group ────────────────────────────
    groups = cd["groups"]
    if not groups:
        return (100, 100)
    total = len(groups)
    succ  = sum(1 for g in groups
                if (g.get("lastRun") or {}).get("localBackupInfo", {})
                   .get("status", "") in ("kSuccess", "Succeeded", "kWarning"))
    viols = sum(1 for g in groups
                if (g.get("lastRun") or {}).get("isSlaViolated", False))
    return (
        round(succ / total * 100, 1),
        round((total - viols) / total * 100, 1),
    )


def _cap_stats(cd):
    """
    Return (usable_bytes, used_bytes, cap_pct_or_NA) for a cluster.
    Primary source: usagePerfStats from /v1/public/cluster?fetchStats=true.
    Falls back to summing storageConsumedBytes across storage domains when the
    cluster-level stats field comes back null (happens on some versions/configs).
    Returns cap_pct as a float, or the string "N/A" when capacity is unknown.
    """
    info   = cd["info"]
    usage  = ((info.get("stats") or {}).get("usagePerfStats") or {})
    usable = usage.get("physicalCapacityBytes") or 0
    used   = usage.get("totalPhysicalUsageBytes") or 0

    if not usable:
        # Fallback: sum physical used across domains (no usable total available)
        domain_used = sum(
            (d.get("stats") or {}).get("storageConsumedBytes") or 0
            for d in cd.get("domains", [])
        )
        return 0, domain_used, "N/A"

    pct = round(used / usable * 100, 1)
    return usable, used, pct


def _sheet_summary(wb, all_data):
    ws = wb.create_sheet("Executive Summary")
    ws.freeze_panes = "A3"
    _title(ws, "Cohesity Health Check — Executive Summary", "M")

    cols = ["Cluster", "Software Version", "Node Count",
            "Health Score", "Grade",
            "30d Success %", "SLA Pass %", "Capacity Used %",
            "Open Criticals", "Open Warnings",
            "Security Score /100", "Recommendations", "Top Finding"]
    _hdr(ws, 2, cols)

    for cd in all_data:
        info   = cd["info"]
        scores = cd["scores"]
        alerts = cd["alerts"]
        _, _, cap_pct = _cap_stats(cd)
        crits  = sum(1 for a in alerts if a.get("severity") == "kCritical")
        warns  = sum(1 for a in alerts if a.get("severity") == "kWarning")
        succ_pct, sla_pct = _success_stats(cd)
        top = cd["recommendations"][0]["finding"] if cd["recommendations"] else "None identified"

        row = [cd["name"],
               info.get("clusterSoftwareVersion", ""), len(cd["nodes"]),
               scores["overall"], scores["grade"],
               succ_pct, sla_pct, cap_pct,
               crits, warns,
               scores["security"], len(cd["recommendations"]), top]
        r = ws.max_row + 1
        for c, val in enumerate(row, 1):
            ws.cell(row=r, column=c, value=val).font = _font()

        # RAG color on score + grade
        for col in (4, 5):
            ws.cell(row=r, column=col).fill = _rag_fill(scores["overall"])
            ws.cell(row=r, column=col).font = _font(bold=True)

        # Red crits
        if crits:
            ws.cell(row=r, column=9).fill = _fill(RED)
            ws.cell(row=r, column=9).font = _font(bold=True, color=WHITE)

    auto_fit_columns(ws)
    ws.column_dimensions[ws.cell(row=2, column=13).column_letter].width = 55


def _sheet_infrastructure(wb, all_data):
    ws = wb.create_sheet("Infrastructure")
    ws.freeze_panes = "A3"
    _title(ws, "Infrastructure — Cluster & Node Health", "Q")

    cols = ["Cluster", "Software Version", "SW Lifecycle Status", "SW EOS Date", "Days to EOS",
            "Cluster Type", "Cluster ID", "Node Count", "Healthy Nodes", "Disk Count",
            "Domain", "DNS Servers", "NTP Servers", "Timezone",
            "Cluster Encryption", "SD Encryption", "Helios Status"]
    _hdr(ws, 2, cols)

    for cd in all_data:
        info  = cd["info"]
        nodes = cd["nodes"]
        healthy = sum(
            1 for n in nodes
            if ((n.get("nodeStatus") or {}).get("overallStatus") == "kNormal"
                or n.get("removalState") in (None, "kDontRemove"))
        )
        # Disk count — v1 /public/nodes returns diskCount as an integer per node
        disks = sum(n.get("diskCount") or 0 for n in nodes)

        sw_ver = info.get("clusterSoftwareVersion", "")
        import datetime as _dt_mod
        sw_status, sw_eos_date, sw_label = _sw_eos(sw_ver)
        today = _dt_mod.date.today()
        days_to_eos = (sw_eos_date - today).days if sw_eos_date else ""
        eos_date_str = sw_eos_date.isoformat() if sw_eos_date else ("—" if sw_status == "current" else "Unknown")

        cluster_enc, sd_enc = _sd_enc_status(cd)
        dns  = ", ".join(info.get("dnsServerIps") or [])
        # NTP: v1 API returns ntpSettings.ntpServers or a flat ntpServers list
        _ntp_s  = info.get("ntpSettings")
        ntp_raw = ((isinstance(_ntp_s, dict) and (_ntp_s.get("ntpServers") or _ntp_s.get("ntpServersList")))
                   or info.get("ntpServers")
                   or info.get("ntpServerIps")
                   or [])
        if isinstance(ntp_raw, list):
            ntp = ", ".join(str(s) for s in ntp_raw) if ntp_raw else "Not configured"
        else:
            ntp = str(ntp_raw) if ntp_raw else "Not configured"
        domain_raw = (info.get("domainNames") or [""])[0]
        # domainNames returns the cluster FQDN (e.g. "mycluster.example.com").
        # Strip the cluster hostname prefix so only the domain portion is shown.
        _cluster_prefix = (info.get("name") or "").lower().rstrip(".")
        if _cluster_prefix and domain_raw.lower().startswith(_cluster_prefix + "."):
            domain = domain_raw[len(_cluster_prefix) + 1:]
        else:
            domain = domain_raw

        sd_enc_str = ("Yes" if sd_enc is True else "No" if sd_enc is False else "N/A")
        row = [cd["name"], sw_ver,
               sw_label, eos_date_str, days_to_eos,
               info.get("clusterType", ""), str(info.get("id", "")),
               len(nodes), healthy, disks if disks else "",
               domain, dns, ntp, info.get("timezone", ""),
               "Yes" if cluster_enc else "No",
               sd_enc_str,
               "Connected"]
        rn = ws.max_row + 1
        for c, val in enumerate(row, 1):
            ws.cell(row=rn, column=c, value=val).font = _font()

        # color SW Lifecycle Status cell (col 3)
        sw_cell = ws.cell(row=rn, column=3)
        if sw_status == "eol":
            sw_cell.fill = _fill(RED);    sw_cell.font = _font(bold=True, color=WHITE)
        elif sw_status == "in_support" and isinstance(days_to_eos, int) and days_to_eos < SW_EOS_WARN_DAYS:
            sw_cell.fill = _fill(ORANGE); sw_cell.font = _font(bold=True)
        elif sw_status == "current":
            sw_cell.fill = _fill(LT_GREEN)

        # color Days to EOS cell (col 5)
        eos_cell = ws.cell(row=rn, column=5)
        if isinstance(days_to_eos, int):
            if days_to_eos < 0:
                eos_cell.fill = _fill(RED);    eos_cell.font = _font(bold=True, color=WHITE)
            elif days_to_eos < SW_EOS_WARN_DAYS:
                eos_cell.fill = _fill(ORANGE); eos_cell.font = _font(bold=True)

        # color Cluster Encryption (col 15) and SD Encryption (col 16)
        enc_cell = ws.cell(row=rn, column=15)
        enc_cell.fill = _fill(LT_GREEN) if cluster_enc else _fill(RED)
        enc_cell.font = _font(bold=True, color=WHITE if not cluster_enc else "000000")
        sd_cell = ws.cell(row=rn, column=16)
        if sd_enc is True:
            sd_cell.fill = _fill(LT_GREEN); sd_cell.font = _font(bold=True)
        elif sd_enc is False:
            sd_cell.fill = _fill(RED);      sd_cell.font = _font(bold=True, color=WHITE)
        else:
            sd_cell.fill = _fill(YELLOW);   sd_cell.font = _font(bold=True)

        # Node health — now col 9
        if healthy < len(nodes):
            ws.cell(row=rn, column=9).fill = _fill(ORANGE)
            ws.cell(row=rn, column=9).font = _font(bold=True)

    auto_fit_columns(ws)


def _sheet_hardware(wb, all_data):
    import datetime as _dt
    ws = wb.create_sheet("Node Hardware")
    ws.freeze_panes = "A3"
    _title(ws, "Node Hardware Inventory — Models, Capacity & End-of-Life Status", "L")

    cols = ["Cluster", "Node ID", "IP Address", "Node Type", "Model",
            "Serial Number", "Node SW Version",
            "Raw Capacity (TB)", "Disk Count", "Storage Tiers",
            "HW EOL Date", "HW EOL Status"]
    _hdr(ws, 2, cols)

    today = _dt.date.today()

    for cd in all_data:
        for n in cd["nodes"]:
            model   = n.get("hardwareModel") or n.get("productModel") or ""
            serial  = n.get("cohesityNodeSerial") or ""
            ntype   = n.get("nodeType") or ""
            nsw     = n.get("nodeSoftwareVersion") or ""
            cap_tb  = round((n.get("maxPhysicalCapacityBytes") or 0) / 1e12, 2)
            disks   = n.get("diskCount") or 0

            # Storage tiers summary e.g. "SSD:4, HDD:8"
            tiers = []
            for t in (n.get("diskCountByTier") or []):
                tier_name = t.get("storageTier", "")
                tier_cnt  = t.get("diskCount", 0)
                if tier_name and tier_cnt:
                    tiers.append(f"{tier_name}:{tier_cnt}")
            tier_str = ", ".join(tiers) or str(disks)

            eol_date = _hw_eol(model)
            if eol_date:
                days_left = (eol_date - today).days
                if days_left < 0:
                    hw_status = f"PAST EOL ({abs(days_left)}d ago)"
                elif days_left <= HW_EOL_WARN_DAYS:
                    hw_status = f"EOL in {days_left}d"
                else:
                    hw_status = f"In Service ({days_left}d remaining)"
                eol_str = eol_date.isoformat()
            else:
                hw_status = "In Service" if model else "Unknown"
                eol_str   = "—"

            row = [cd["name"],
                   n.get("id", ""), n.get("ip", ""),
                   ntype, model, serial, nsw,
                   cap_tb, disks, tier_str,
                   eol_str, hw_status]
            rn = ws.max_row + 1
            for c, val in enumerate(row, 1):
                ws.cell(row=rn, column=c, value=val).font = _font()

            # color HW EOL Status cell (col 12)
            eol_cell = ws.cell(row=rn, column=12)
            if eol_date:
                days_left = (eol_date - today).days
                if days_left < 0:
                    eol_cell.fill = _fill(RED);    eol_cell.font = _font(bold=True, color=WHITE)
                elif days_left <= HW_EOL_WARN_DAYS:
                    eol_cell.fill = _fill(ORANGE); eol_cell.font = _font(bold=True)
                elif days_left <= 365:
                    eol_cell.fill = _fill(YELLOW); eol_cell.font = _font(bold=True)

    auto_fit_columns(ws)
    ws.column_dimensions["E"].width = 30  # Model


def _sheet_protection(wb, all_data):
    ws = wb.create_sheet("Protection Health")
    ws.freeze_panes = "A3"
    _title(ws, "Protection Health — Group Status & Run Metrics", "T")

    cols = ["Cluster", "Group Name", "Policy",
            "Active", "Paused",
            "Last Run Type", "Last Run Status", "SLA Violated",
            "Last Run Start", "Last Run End", "Duration",
            "Objects OK", "Objects Failed",
            "Logical (GB)", "Physical (GB)",
            "Replication Status", "Archival Status",
            "DataLock (WORM)",
            "RPO (hrs)", "Flag"]
    _hdr(ws, 2, cols)

    for cd in all_data:
        pm         = cd.get("policy_map", {})
        policy_by_id = {p.get("id"): p for p in cd["policies"] if p.get("id")}
        for g in cd["groups"]:
            lr  = g.get("lastRun") or {}
            lb  = lr.get("localBackupInfo") or {}

            # runType and isSlaViolated sit on localBackupInfo (confirmed from runs API)
            run_type    = lb.get("runType") or lr.get("runType", "")
            sla_violated = lb.get("isSlaViolated") or lr.get("isSlaViolated") or False
            status      = lb.get("status", "")

            start_u = lb.get("startTimeUsecs") or 0
            end_u   = lb.get("endTimeUsecs") or 0
            dur = ""
            if start_u and end_u:
                s = int((end_u - start_u) / 1_000_000)
                h, m = divmod(s // 60, 60)
                dur = f"{h}h {m:02d}m"
            rpo_hrs = round((_now_usecs() - start_u) / 3_600_000_000, 1) \
                      if start_u else ""

            # Object counts are directly on localBackupInfo (not nested in stats)
            objs_ok   = lb.get("successfulObjectsCount", "")
            objs_fail = lb.get("failedObjectsCount", "")
            # Sizes are under localSnapshotStats sub-dict
            snap = lb.get("localSnapshotStats") or {}
            logical_gb  = round(bytes_to_gb(snap.get("logicalSizeBytes") or 0), 2)
            physical_gb = round(bytes_to_gb(snap.get("bytesWritten") or 0), 2)

            # Replication / archival status from lastRun (same level as localBackupInfo)
            repl_st = arch_st = ""
            for ri in ((lr.get("replicationInfo") or {})
                       .get("replicationTargetResults") or []):
                repl_st = ri.get("status", ""); break
            for ai in ((lr.get("archivalInfo") or {})
                       .get("archivalTargetResults") or []):
                arch_st = ai.get("status", ""); break

            flag = "OK"
            if g.get("isPaused"):
                flag = "PAUSED"
            elif status in ("kFailure", "Failed"):
                flag = "FAILED"
            elif sla_violated:
                flag = "SLA VIOLATED"
            elif isinstance(rpo_hrs, float) and rpo_hrs > RPO_CRIT_HOURS:
                flag = f"RPO GAP {rpo_hrs:.0f}h"
            elif isinstance(rpo_hrs, float) and rpo_hrs > RPO_WARN_HOURS:
                flag = f"RPO WARN {rpo_hrs:.0f}h"

            # Policy name: prefer policyName on group; fall back to policy_map lookup
            policy_name = (g.get("policyName")
                           or pm.get(g.get("policyId"), "")
                           or g.get("policyId", ""))

            # DataLock: look up via the full policy object
            policy = policy_by_id.get(g.get("policyId"), {})
            dl_str = _policy_datalock_str(policy)

            row = [cd["name"], g.get("name", ""),
                   policy_name,
                   "Yes" if g.get("isActive", True) else "No",
                   "Yes" if g.get("isPaused", False) else "No",
                   run_type, status,
                   "Yes" if sla_violated else "No",
                   usecs_to_datetime(start_u), usecs_to_datetime(end_u), dur,
                   objs_ok, objs_fail,
                   logical_gb, physical_gb,
                   repl_st, arch_st,
                   dl_str,
                   rpo_hrs, flag]
            rn = ws.max_row + 1
            for c, val in enumerate(row, 1):
                ws.cell(row=rn, column=c, value=val).font = _font()

            # DataLock col 18 — color by mode
            dc = ws.cell(row=rn, column=18)
            dl_lower = dl_str.lower()
            if "compliance" in dl_lower:
                dc.fill = _fill(LT_GREEN)
            elif "administrative" in dl_lower:
                dc.fill = _fill(YELLOW); dc.font = _font(bold=True)
            else:
                dc.fill = _fill(RED); dc.font = _font(bold=True, color=WHITE)

            # Flag col 20
            fc = ws.cell(row=rn, column=20)
            if flag == "FAILED":
                fc.fill = _fill(RED);   fc.font = _font(bold=True, color=WHITE)
            elif flag in ("PAUSED",) or "RPO GAP" in flag:
                fc.fill = _fill(ORANGE); fc.font = _font(bold=True)
            elif flag == "SLA VIOLATED" or "RPO WARN" in flag:
                fc.fill = _fill(YELLOW); fc.font = _font(bold=True)
            else:
                fc.fill = _fill(LT_GREEN)

    auto_fit_columns(ws)


def _sheet_storage(wb, all_data):
    ws = wb.create_sheet("Storage & Capacity")
    ws.freeze_panes = "A3"
    _title(ws, "Storage & Capacity — Utilization & Data Reduction", "L")

    cols = ["Cluster", "Storage Domain", "Usable (TB)", "Used (TB)",
            "Free (TB)", "Used %", "Logical Data (TB)",
            "Data Reduction Ratio", "Dedup Ratio", "Compression Ratio",
            "Savings (TB)", "Status"]
    _hdr(ws, 2, cols)

    for cd in all_data:
        info   = cd["info"]
        stats  = (info.get("stats") or {})
        usage  = (stats.get("usagePerfStats") or {})
        data   = (stats.get("dataUsageStats") or {})

        usable, used, cap_pct = _cap_stats(cd)
        logical     = usage.get("dataInBytes") or 0
        physical    = usage.get("dataInBytesAfterReduction") or 0
        dedup_after = data.get("dataInBytesAfterDedup") or 0
        dr_ratio    = round((stats.get("dataReductionRatio") or 0.0), 2)
        dedup_ratio = round(logical / dedup_after, 2) if dedup_after else 0
        comp_ratio  = round(dedup_after / physical, 2) if physical else 0
        free        = usable - used
        savings     = round(bytes_to_tb(logical - physical), 2) \
                      if logical > physical else 0

        flag = ""
        if isinstance(cap_pct, float):
            flag = ("CRITICAL" if cap_pct >= CAPACITY_CRIT_PCT
                    else "WARN" if cap_pct >= CAPACITY_WARN_PCT else "OK")
        elif cap_pct == "N/A":
            flag = "N/A"

        row = [cd["name"], "(Cluster Total)",
               round(bytes_to_tb(usable), 2) or "",
               round(bytes_to_tb(used), 2) or "",
               round(bytes_to_tb(free), 2) or "",
               cap_pct,
               round(bytes_to_tb(logical), 2),
               dr_ratio, dedup_ratio, comp_ratio, savings, flag]
        rn = ws.max_row + 1
        for c, val in enumerate(row, 1):
            cell = ws.cell(row=rn, column=c, value=val)
            cell.font = _font(bold=(c == 2))

        sc = ws.cell(row=rn, column=12)
        if flag == "CRITICAL":
            sc.fill = _fill(RED);    sc.font = _font(bold=True, color=WHITE)
        elif flag == "WARN":
            sc.fill = _fill(ORANGE); sc.font = _font(bold=True)
        elif flag == "OK":
            sc.fill = _fill(LT_GREEN)

        # Per-domain rows
        for d in cd["domains"]:
            ds = d.get("stats") or {}
            ld = ds.get("dataInBytes") or 0
            da = ds.get("dataInBytesAfterDedup") or 0
            ph = ds.get("storageConsumedBytes") or 0
            drow = [cd["name"], d.get("name", ""),
                    "", "", "", "",
                    round(bytes_to_tb(ld), 2),
                    round(ld / ph, 2) if ph else 0,
                    round(ld / da, 2) if da else 0,
                    round(da / ph, 2) if ph else 0,
                    round(bytes_to_tb(ld - ph), 2) if ld > ph else 0, ""]
            rn2 = ws.max_row + 1
            for c, val in enumerate(drow, 1):
                ws.cell(row=rn2, column=c, value=val).font = _font()

    auto_fit_columns(ws)


def _ret_str(r):
    if not r: return ""
    d = r.get("duration", "")
    u = r.get("unit", "")
    return f"{d} {u}".strip() if d else u

def _sched_str(s):
    """Convert a v2 schedule dict to a human-readable string."""
    if not s: return ""
    unit = s.get("unit", "")
    if unit == "Days":
        freq = (s.get("daySchedule") or {}).get("frequency", 1)
        return f"Every {freq} day{'s' if freq != 1 else ''}"
    if unit == "Weeks":
        days = (s.get("weekSchedule") or {}).get("dayOfWeek") or []
        return f"Weekly ({', '.join(days[:2])}{'…' if len(days) > 2 else ''})" if days else "Weekly"
    if unit == "Hours":
        freq = (s.get("hourSchedule") or {}).get("frequency", 1)
        return f"Every {freq}h"
    if unit == "Minutes":
        freq = (s.get("minuteSchedule") or {}).get("frequency", 1)
        return f"Every {freq}m"
    if unit in ("Months", "Monthly"): return "Monthly"
    if unit in ("Years", "Yearly"):   return "Yearly"
    # fallback: flat frequency dict
    for key in ("minutes", "hours", "days", "weeks", "months", "years"):
        if isinstance(s, dict) and key in s:
            return f"Every {s[key]} {key}"
    return unit or ""

def _policy_datalock_str(p):
    """Return a display string for a single policy's DataLock / indelible config.

    Priority:
      1. Explicit local DataLock (backupPolicy.regular.retention.dataLockConfig).
         Returns e.g. "Compliance / 14 Days" or "Administrative / 30 Days".
      2. FortKnox / RPaaS archival target — indelible by design.
         Returns "FortKnox (Indelible)".
      3. Neither configured → "None".
    """
    bp  = p.get("backupPolicy") or {}
    reg = bp.get("regular") or {}
    ret = reg.get("retention") or bp.get("retention") or {}
    dlc = ret.get("dataLockConfig") if isinstance(ret, dict) else None
    if isinstance(dlc, dict) and dlc.get("mode"):
        mode = dlc["mode"]
        dur  = dlc.get("duration")
        unit = dlc.get("unit", "Days")
        return f"{mode} / {dur} {unit}" if dur else mode
    # No explicit WORM — check for FortKnox/RPaaS archival (implicit indelible)
    _fk = {"krpaas", "kfortknox", "kfort_knox", "krpaasarchival"}
    for t in ((p.get("remoteTargetPolicy") or {}).get("archivalTargets") or []):
        cfg   = t.get("archivalTargetConfig") or {}
        tt    = (cfg.get("targetType") or t.get("targetType") or "").lower()
        tname = (t.get("targetName") or cfg.get("name") or "").lower()
        if any(x in tt for x in _fk) or "fortknox" in tt or "fortknox" in tname:
            return "FortKnox (Indelible)"
    return "None"


def _sd_enc_status(cd):
    """Return (cluster_enc, sd_enc) booleans for encryption state.

    cluster_enc: True if cluster-wide encryption is enabled.
    sd_enc:      True if ALL storage domains have a non-empty encryptionType.
                 None if no domain data is available (unknown).
    """
    info = cd["info"]
    _enc_cfg = info.get("encryptionConfig")
    cluster_enc = bool(info.get("encryptionEnabled") or
                       (isinstance(_enc_cfg, dict) and _enc_cfg.get("encryptionEnabled")))
    domains = [d for d in (cd.get("domains") or []) if isinstance(d, dict)]
    if not domains:
        return cluster_enc, None
    sd_enc = all(
        bool((d.get("storagePolicy") or {}).get("encryptionType", ""))
        for d in domains
    )
    return cluster_enc, sd_enc


def _cluster_datalock(policies, groups=None):
    """Return (status, label, covered, total) for cluster-level DataLock coverage.

    Computes coverage at the *protection group* level: a group is covered if its
    governing policy has DataLock enabled (local retention or archival).

    status: "full_compliance" | "full_any" | "partial" | "none"
    label:  human-readable string, e.g. "Full Compliance" or "Partial — 8/15 groups"
    covered: number of groups with DataLock
    total:   total number of groups
    """
    # Build policy_id → datalock mode map
    policy_dl = {}   # policy id → "compliance" | "administrative" | "none"
    for p in policies:
        pid = p.get("id")
        if not pid:
            continue
        # Check local retention
        s = _policy_datalock_str(p).lower()
        mode = ("compliance" if "compliance" in s
                else "administrative" if "admin" in s else "none")
        # Check archival target retention (upgrade to compliance if found)
        if mode != "compliance":
            for t in ((p.get("remoteTargetPolicy") or {}).get("archivalTargets") or []):
                dlc = (t.get("retention") or {}).get("dataLockConfig")
                if isinstance(dlc, dict):
                    m = (dlc.get("mode") or "").lower()
                    if "compliance" in m:
                        mode = "compliance"; break
                    elif "admin" in m and mode == "none":
                        mode = "administrative"
        policy_dl[pid] = mode

    # Count group coverage
    total   = len(groups) if groups else 0
    covered_comp = 0
    covered_any  = 0
    if groups:
        for g in groups:
            pid  = g.get("policyId")
            mode = policy_dl.get(pid, "none")
            if mode == "compliance":
                covered_comp += 1
                covered_any  += 1
            elif mode == "administrative":
                covered_any += 1
    covered = covered_any

    if total == 0:
        # No active protection groups — cannot determine coverage.
        return ("none", "No", 0, 0)

    if covered == 0:
        return ("none", "No", 0, total)
    if covered == total and covered_comp == total:
        return ("full_compliance", "Full Compliance", covered, total)
    if covered == total:
        return ("full_any", f"Full ({covered_comp}/{total} Compliance)", covered, total)
    # Partial coverage
    return ("partial", f"Partial ({covered}/{total} groups)", covered, total)


def _sheet_policies(wb, all_data):
    ws = wb.create_sheet("Policy Audit")
    ws.freeze_panes = "A3"
    _title(ws, "Policy Audit — Retention, Replication & Archival", "P")

    cols = ["Cluster", "Policy Name", "Type",
            "Full Schedule", "Incremental Schedule",
            "Local Retention", "Log Retention",
            "Has Replication", "Replication Target", "Replication Retention",
            "Has Archival", "Archival Target", "Archival Retention",
            "Groups Using", "DataLock (WORM)", "Gaps"]
    _hdr(ws, 2, cols)

    for cd in all_data:
        # Count protection groups per policy ID so column N is populated
        policy_group_count: dict = {}
        for g in cd["groups"]:
            pid = g.get("policyId")
            if pid:
                policy_group_count[pid] = policy_group_count.get(pid, 0) + 1

        for p in cd["policies"]:
            bp  = p.get("backupPolicy") or {}
            rtp = p.get("remoteTargetPolicy") or {}
            rep_tgts  = rtp.get("replicationTargets") or []
            arch_tgts = rtp.get("archivalTargets") or []

            # v2 API replication target name: targetName or remoteTargetConfig.clusterName/name
            def _rep_name(t):
                return (t.get("targetName")
                        or (t.get("remoteTargetConfig") or {}).get("clusterName")
                        or (t.get("remoteTargetConfig") or {}).get("name")
                        or "")
            # v2 API archival target name: targetName or archivalTargetConfig.name/vaultName
            def _arch_name(t):
                return (t.get("targetName")
                        or (t.get("archivalTargetConfig") or {}).get("name")
                        or (t.get("archivalTargetConfig") or {}).get("vaultName")
                        or "")

            rep_names  = ", ".join(filter(None, (_rep_name(t)  for t in rep_tgts)))
            arch_names = ", ".join(filter(None, (_arch_name(t) for t in arch_tgts)))

            # v2 schedule: backupPolicy.regular.full.schedule / regular.incremental.schedule
            regular     = bp.get("regular") or {}
            full_sched  = (regular.get("full") or bp.get("full") or {}).get("schedule") or {}
            incr_sched  = (regular.get("incremental") or bp.get("incremental") or {}).get("schedule") or {}

            local_ret = regular.get("retention") or bp.get("retention") or {}
            log_ret   = (bp.get("log") or {}).get("retention") or {}
            rep_ret   = (rep_tgts[0].get("retention") or {}) if rep_tgts else {}
            arch_ret  = (arch_tgts[0].get("retention") or {}) if arch_tgts else {}

            gaps = []
            if not rep_tgts:  gaps.append("No replication")
            if not arch_tgts: gaps.append("No archival")

            dl_str = _policy_datalock_str(p)
            dl_low = dl_str.lower()
            if dl_str == "None":
                gaps.append("No DataLock")

            row = [cd["name"], p.get("name", ""),
                   p.get("type") or p.get("policyCategory") or "",
                   _sched_str(full_sched), _sched_str(incr_sched),
                   _ret_str(local_ret), _ret_str(log_ret),
                   "Yes" if rep_tgts else "No", rep_names, _ret_str(rep_ret),
                   "Yes" if arch_tgts else "No", arch_names, _ret_str(arch_ret),
                   policy_group_count.get(p.get("id"), 0),
                   dl_str,
                   "; ".join(gaps) if gaps else "OK"]
            rn = ws.max_row + 1
            for c, val in enumerate(row, 1):
                ws.cell(row=rn, column=c, value=val).font = _font()

            # DataLock col 15 — color by mode
            dl_cell = ws.cell(row=rn, column=15)
            if "compliance" in dl_low:
                dl_cell.fill = _fill(LT_GREEN); dl_cell.font = _font(bold=True)
            elif "admin" in dl_low:
                dl_cell.fill = _fill(YELLOW);   dl_cell.font = _font(bold=True)
            elif "fortknox" in dl_low:
                dl_cell.fill = _fill(LT_GREEN); dl_cell.font = _font(bold=True)
            else:
                dl_cell.fill = _fill(RED);      dl_cell.font = _font(bold=True, color=WHITE)

            # Gaps col 16
            gc = ws.cell(row=rn, column=16)
            if gaps:
                gc.fill = _fill(YELLOW)
                gc.font = _font(bold=True)

    auto_fit_columns(ws)


def _sheet_policy_groups(wb, all_data):
    """Sheet: one row per protection group showing which policy governs it."""
    ws = wb.create_sheet("Policy → Groups")
    ws.freeze_panes = "A3"
    _title(ws, "Policy → Groups — Protection Groups Associated with Each Policy", "G")

    cols = ["Cluster", "Policy Name", "Group Name", "Environment",
            "Active", "Paused", "Last Run Status"]
    _hdr(ws, 2, cols)

    for cd in all_data:
        pm = cd.get("policy_map", {})
        # Sort by policy name then group name for easy reading
        sorted_groups = sorted(
            cd["groups"],
            key=lambda g: (
                g.get("policyName") or pm.get(g.get("policyId"), "") or "",
                g.get("name", "")
            )
        )
        for g in sorted_groups:
            policy_name = (g.get("policyName")
                           or pm.get(g.get("policyId"), "")
                           or g.get("policyId", ""))
            lb     = (g.get("lastRun") or {}).get("localBackupInfo") or {}
            status = lb.get("status", "")
            paused = g.get("isPaused", False)

            row = [cd["name"], policy_name, g.get("name", ""),
                   g.get("environment", ""),
                   "Yes" if g.get("isActive", True) else "No",
                   "Yes" if paused else "No",
                   status]
            rn = ws.max_row + 1
            for c, val in enumerate(row, 1):
                ws.cell(row=rn, column=c, value=val).font = _font()

            # Highlight paused groups
            if paused:
                ws.cell(row=rn, column=6).fill = _fill(YELLOW)
                ws.cell(row=rn, column=6).font = _font(bold=True)
            # Highlight failed status
            if status in ("kFailure", "Failed"):
                ws.cell(row=rn, column=7).fill = _fill(RED)
                ws.cell(row=rn, column=7).font = _font(bold=True, color=WHITE)

    auto_fit_columns(ws)


def _sheet_alerts(wb, all_data):
    ws = wb.create_sheet("Alerts")
    ws.freeze_panes = "A3"
    _title(ws, "Open Alerts — Severity, Age & Description", "J")

    cols = ["Cluster", "Severity", "Category", "Alert Code",
            "Alert Name", "Description",
            "First Occurred", "Age (Days)", "Last Updated",
            "Acknowledged"]
    _hdr(ws, 2, cols)

    sev_order = {"kCritical": 0, "kWarning": 1, "kInfo": 2}

    for cd in all_data:
        sorted_a = sorted(cd["alerts"],
                          key=lambda a: (sev_order.get(a.get("severity"), 9),
                                         a.get("firstTimestampUsecs") or 0))
        for a in sorted_a:
            # v1 alerts API uses firstTimestampUsecs (not firstOccurrenceTime)
            first_u = a.get("firstTimestampUsecs") or 0
            last_u  = a.get("latestTimestampUsecs") or 0
            age = round((_now_usecs() - first_u) / (86400 * 1_000_000), 1) \
                  if first_u else ""
            doc  = a.get("alertDocument") or {}
            desc = doc.get("alertDescription") or ""

            row = [cd["name"], a.get("severity", ""),
                   a.get("alertCategory", ""), a.get("alertCode", ""),
                   doc.get("alertName", ""),
                   desc[:120] + ("…" if len(desc) > 120 else ""),
                   usecs_to_datetime(first_u), age, usecs_to_datetime(last_u),
                   "Yes" if a.get("acknowledgeTimestampUsecs") else "No"]
            rn = ws.max_row + 1
            for c, val in enumerate(row, 1):
                ws.cell(row=rn, column=c, value=val).font = _font()

            sc = ws.cell(row=rn, column=2)
            if a.get("severity") == "kCritical":
                sc.fill = _fill(RED);    sc.font = _font(bold=True, color=WHITE)
            elif a.get("severity") == "kWarning":
                sc.fill = _fill(ORANGE); sc.font = _font(bold=True)

    auto_fit_columns(ws)
    ws.column_dimensions["F"].width = 70


def _sheet_security(wb, all_data):
    import datetime as _dt
    ws = wb.create_sheet("Security")
    ws.freeze_panes = "A3"
    _title(ws, "Security Posture — Checklist vs Best Practices", "S")

    cols = ["Cluster", "Cluster Encryption", "SD Encryption",
            "Vault Configured", "FortKnox / Indelible",
            "Replication", "Archival",
            "Min Local Retention", "Min Archival Retention",
            "Security Score /100",
            "Audit Log", "NTP Auth", "Remote Tunnel",
            "Cluster MFA", "SSO / IDP",
            "Quorum",
            "DataLock (WORM)",
            "Soonest Cert Expiry",
            "Gaps"]
    _hdr(ws, 2, cols)

    _RET_MULT = {"Days": 1, "Weeks": 7, "Months": 30, "Years": 365}

    def _ret_to_days(r):
        d = (r or {}).get("duration") or 0
        u = (r or {}).get("unit", "Days")
        return d * _RET_MULT.get(u, 1) if d else 0

    def _min_local_days(policies):
        vals = []
        for p in policies:
            bp  = p.get("backupPolicy") or {}
            reg = bp.get("regular") or {}
            r   = reg.get("retention") or bp.get("retention") or {}
            v   = _ret_to_days(r)
            if v: vals.append(v)
        return f"{min(vals)} days" if vals else "N/A"

    def _min_arch_days(policies):
        vals = []
        for p in policies:
            for t in ((p.get("remoteTargetPolicy") or {}).get("archivalTargets") or []):
                v = _ret_to_days(t.get("retention") or {})
                if v: vals.append(v)
        return f"{min(vals)} days" if vals else "N/A"

    today = _dt.date.today()

    for cd in all_data:
        info     = cd["info"]
        vaults   = cd["vaults"]
        policies = cd["policies"]
        scores   = cd["scores"]

        cluster_enc, sd_enc = _sd_enc_status(cd)
        has_vault = bool(vaults)
        fk_st    = _fk_status(cd)
        has_repl = any((p.get("remoteTargetPolicy") or {}).get("replicationTargets")
                       for p in policies)
        has_arch = any((p.get("remoteTargetPolicy") or {}).get("archivalTargets")
                       for p in policies)

        # ── Extended security controls ─────────────────────────────────────
        _audit_cfg2     = info.get("auditLogConfig")
        _cluster_audit2 = info.get("clusterAuditConfig")
        audit_on = bool(
            info.get("auditLogEnabled")
            or (isinstance(_audit_cfg2, dict)     and _audit_cfg2.get("enabled"))
            or (isinstance(_cluster_audit2, dict) and _cluster_audit2.get("enabled"))
        )
        _ntp_s2  = info.get("ntpSettings")
        ntp_auth = bool(
            isinstance(_ntp_s2, dict)
            and (_ntp_s2.get("ntpAuthenticationEnabled") or _ntp_s2.get("authEnabled"))
        )
        _remote_cfg = info.get("remoteSupportConfig")
        tunnel_on = bool(
            info.get("enableRemoteSupport")
            or (isinstance(_remote_cfg, dict) and _remote_cfg.get("enabled"))
            or info.get("remoteSupportEnabled")
        )
        _mfa_cfg = info.get("mfaConfig")
        cluster_mfa = bool(
            info.get("clusterMfaEnabled")
            or info.get("mfaEnabled")
            or (isinstance(_mfa_cfg, dict) and _mfa_cfg.get("enabled"))
        )
        has_sso = bool(cd.get("idps"))

        # Quorum — Helios feature; check info dict for known field names
        # Quorum is a Helios-level feature (Security Center → Quorum).
        # Data comes from GET /v2/mcm/quorum/groups fetched once before cluster
        # collection and stored in cd["quorum_enabled"]. Direct mode = N/A.
        if cd.get("mode") == "direct":
            quorum_str = "N/A"
        else:
            quorum_str = "Yes" if cd.get("quorum_enabled") else "No"

        # Soonest TLS cert expiry
        cert_expiry_str = ""
        soonest_days    = None
        for cert in (cd.get("certs") or []):
            exp_raw = (cert.get("expiryUsecs")
                       or cert.get("expiryTimeMsecs")
                       or cert.get("notAfter"))
            if not exp_raw:
                continue
            if exp_raw > 1_000_000_000_000_000:
                exp_dt = _dt.date.fromtimestamp(exp_raw / 1_000_000)
            elif exp_raw > 1_000_000_000_000:
                exp_dt = _dt.date.fromtimestamp(exp_raw / 1_000)
            else:
                exp_dt = _dt.date.fromtimestamp(exp_raw)
            d = (exp_dt - today).days
            if soonest_days is None or d < soonest_days:
                soonest_days    = d
                cert_expiry_str = f"{exp_dt.isoformat()} ({d}d)"

        dl_status, dl_label, dl_covered, dl_total = _cluster_datalock(policies,
                                                                      cd["groups"])

        gaps = []
        if not cluster_enc:           gaps.append("No cluster encryption")
        if sd_enc is False:           gaps.append("SD encryption missing")
        if not has_vault:             gaps.append("No vault")
        if fk_st == "none":           gaps.append("No indelible copy")
        if fk_st == "idle":           gaps.append("FK configured — not sending data")
        if not has_repl:              gaps.append("No replication")
        if not has_arch:              gaps.append("No archival")
        if dl_status == "none":       gaps.append("No DataLock")
        if dl_status == "partial":    gaps.append(f"DataLock partial ({dl_covered}/{dl_total})")
        if not audit_on:              gaps.append("Audit log off")
        if not cluster_mfa:           gaps.append("No cluster MFA")
        if not has_sso:               gaps.append("No SSO")
        if quorum_str == "No":        gaps.append("Quorum not enabled")
        if soonest_days is not None and soonest_days < CERT_WARN_DAYS:
            gaps.append(f"Cert expires {soonest_days}d")

        fk_label = {"active": "Active", "idle": "Configured — Idle", "none": "No"}[fk_st]
        yn = lambda v: "Yes" if v else "No"
        sd_enc_str = ("Yes" if sd_enc is True else "No" if sd_enc is False else "N/A")
        row = [cd["name"], yn(cluster_enc), sd_enc_str, yn(has_vault), fk_label,
               yn(has_repl), yn(has_arch),
               _min_local_days(policies), _min_arch_days(policies),
               scores["security"],
               yn(audit_on), yn(ntp_auth), yn(tunnel_on),
               yn(cluster_mfa), yn(has_sso),
               quorum_str,
               dl_label,
               cert_expiry_str or "—",
               "; ".join(gaps) if gaps else "OK"]
        rn = ws.max_row + 1
        for c, val in enumerate(row, 1):
            ws.cell(row=rn, column=c, value=val).font = _font()

        # Cluster Encryption col 2, SD Encryption col 3
        ce_cell = ws.cell(row=rn, column=2)
        ce_cell.fill = _fill(LT_GREEN) if cluster_enc else _fill(RED)
        ce_cell.font = _font(bold=True, color=WHITE if not cluster_enc else "000000")
        sd_cell = ws.cell(row=rn, column=3)
        if sd_enc is True:
            sd_cell.fill = _fill(LT_GREEN); sd_cell.font = _font(bold=True)
        elif sd_enc is False:
            sd_cell.fill = _fill(RED);      sd_cell.font = _font(bold=True, color=WHITE)
        else:
            sd_cell.fill = _fill(YELLOW);   sd_cell.font = _font(bold=True)

        for col, flag in [(4, has_vault), (6, has_repl), (7, has_arch)]:
            cell = ws.cell(row=rn, column=col)
            cell.fill = _fill(LT_GREEN) if flag else _fill(RED)
            cell.font = _font(bold=True, color=WHITE if not flag else "000000")

        # FK column 5: three-state coloring
        fk_cell = ws.cell(row=rn, column=5)
        if fk_st == "active":
            fk_cell.fill = _fill(LT_GREEN); fk_cell.font = _font(bold=True)
        elif fk_st == "idle":
            fk_cell.fill = _fill(ORANGE);   fk_cell.font = _font(bold=True)
        else:
            fk_cell.fill = _fill(RED);      fk_cell.font = _font(bold=True, color=WHITE)

        # Score col 10
        sc = ws.cell(row=rn, column=10)
        sc.fill = _rag_fill(scores["security"])
        sc.font = _font(bold=True)

        # Extended controls: audit(11), NTP auth(12), tunnel(13), MFA(14), SSO(15)
        for col, flag, warn_if_off in [
            (11, audit_on,    True),
            (12, ntp_auth,    False),
            (13, tunnel_on,   False),   # tunnel: yellow if on (reduce attack surface)
            (14, cluster_mfa, True),
            (15, has_sso,     False),
        ]:
            cell = ws.cell(row=rn, column=col)
            if col == 13:
                cell.fill = _fill(YELLOW) if flag else _fill(LT_GREEN)
                cell.font = _font(bold=True)
            elif warn_if_off:
                cell.fill = _fill(LT_GREEN) if flag else _fill(RED)
                cell.font = _font(bold=True, color=WHITE if not flag else "000000")
            else:
                cell.fill = _fill(LT_GREEN) if flag else _fill(YELLOW)
                cell.font = _font(bold=True)

        # Quorum col 16 — N/A for direct mode
        qc = ws.cell(row=rn, column=16)
        if quorum_str == "Yes":
            qc.fill = _fill(LT_GREEN); qc.font = _font(bold=True)
        elif quorum_str == "No":
            qc.fill = _fill(YELLOW);   qc.font = _font(bold=True)
        # N/A: no fill

        # DataLock col 17 — four-state coloring
        dl_cell = ws.cell(row=rn, column=17)
        if dl_status == "full_compliance":
            dl_cell.fill = _fill(LT_GREEN); dl_cell.font = _font(bold=True)
        elif dl_status == "full_any":
            dl_cell.fill = _fill(LT_GREEN); dl_cell.font = _font(bold=True)
        elif dl_status == "partial":
            dl_cell.fill = _fill(YELLOW);   dl_cell.font = _font(bold=True)
        else:
            dl_cell.fill = _fill(RED);      dl_cell.font = _font(bold=True, color=WHITE)

        # Cert expiry col 18
        if soonest_days is not None:
            cc = ws.cell(row=rn, column=18)
            if soonest_days < CERT_CRIT_DAYS:
                cc.fill = _fill(RED);    cc.font = _font(bold=True, color=WHITE)
            elif soonest_days < CERT_WARN_DAYS:
                cc.fill = _fill(ORANGE); cc.font = _font(bold=True)

    auto_fit_columns(ws)
    ws.column_dimensions["S"].width = 60


def _sheet_replication(wb, all_data):
    ws = wb.create_sheet("Replication & Archive")
    ws.freeze_panes = "A3"
    _title(ws, "Replication & Archival Targets", "K")

    cols = ["Cluster", "Target Type", "Target Name", "Vault Type",
            "Protection Groups",
            "Vault Storage (TB)",
            "Logical Data Transferred (TB)\n[numLogicalBytesTransferred — cumulative]",
            "Physical Data Transferred (TB)\n[numPhysicalBytesTransferred — cumulative]",
            "Status", "Notes",
            "Protection Group Names"]
    _hdr(ws, 2, cols)

    for cd in all_data:
        # ── Vault ID → name / type maps ───────────────────────────────────
        vault_name_by_id   = {}   # vault id → name
        vault_type_by_name = {}   # name     → vaultType
        vault_id_by_name   = {}   # name     → id
        for v in cd["vaults"]:
            vn = v.get("name") or v.get("vaultName") or ""
            vi = v.get("id")
            vt = v.get("vaultType", "")
            if vi: vault_name_by_id[vi] = vn or f"Vault-{vi}"
            if vn:
                vault_type_by_name[vn] = vt
                if vi: vault_id_by_name[vn] = vi

        # ── Target name extractors ────────────────────────────────────────
        def _rep_name(t):
            cfg = t.get("remoteTargetConfig") or {}
            return (t.get("targetName")
                    or cfg.get("clusterName")
                    or cfg.get("name") or "")

        def _arch_name(t):
            cfg  = t.get("archivalTargetConfig") or {}
            name = (t.get("targetName")
                    or cfg.get("name")
                    or cfg.get("vaultName") or "")
            if not name:
                vid = cfg.get("vaultId") or cfg.get("id")
                if vid:
                    name = vault_name_by_id.get(vid, f"Vault-{vid}")
            return name

        # ── Policy chain: policy id → target names ───────────────────────
        policy_rep_targets  = {}   # policy id → set of replication cluster names
        policy_arch_targets = {}   # policy id → set of vault names
        arch_target_type    = {}   # vault name → targetType

        for p in cd["policies"]:
            pid = p.get("id")
            if not pid:
                continue
            rtp = p.get("remoteTargetPolicy") or {}
            for t in (rtp.get("replicationTargets") or []):
                n = _rep_name(t)
                if n:
                    policy_rep_targets.setdefault(pid, set()).add(n)
            for t in (rtp.get("archivalTargets") or []):
                n = _arch_name(t)
                if n:
                    policy_arch_targets.setdefault(pid, set()).add(n)
                    if n not in arch_target_type:
                        cfg = t.get("archivalTargetConfig") or {}
                        arch_target_type[n] = (t.get("targetType")
                                               or cfg.get("targetType") or "")

        # ── Group → target mapping (policy chain) ────────────────────────
        rep_groups  = {}   # replication cluster name → [group names]
        arch_groups = {}   # vault name               → [group names]

        for g in cd["groups"]:
            pid   = g.get("policyId")
            gname = g.get("name", "")
            if not pid or not gname:
                continue
            for tname in (policy_rep_targets.get(pid) or set()):
                rep_groups.setdefault(tname, []).append(gname)
            for tname in (policy_arch_targets.get(pid) or set()):
                arch_groups.setdefault(tname, []).append(gname)

        # ── Supplement arch_groups from fk_data (direct transfer records) ─
        # fk_data lists protectionJobName per vault — reliable even when the
        # policy chain fails to extract vault names (e.g. ID-only references).
        for fk in (cd["fk_data"] or []):
            fk_vn = fk.get("vaultName", "")
            if not fk_vn:
                continue
            for job in (fk.get("dataTransferPerProtectionJob") or []):
                jname = job.get("protectionJobName", "")
                if jname and jname not in (arch_groups.get(fk_vn) or []):
                    arch_groups.setdefault(fk_vn, []).append(jname)

        # All unique target names — union of policy chain + fk_data + vault list
        all_rep_targets  = set().union(*policy_rep_targets.values()) \
                           if policy_rep_targets else set()
        fk_vault_names   = {fk.get("vaultName") for fk in (cd["fk_data"] or [])
                            if fk.get("vaultName")}
        vault_list_names = {v.get("name") or v.get("vaultName") or ""
                            for v in cd["vaults"]} - {""}
        all_arch_targets = (
            (set().union(*policy_arch_targets.values()) if policy_arch_targets else set())
            | fk_vault_names
            | vault_list_names
        )

        # ── FK transfer stats ─────────────────────────────────────────────
        fk_by_name = {}
        fk_by_id   = {}
        for fk in (cd["fk_data"] or []):
            fk_vn  = fk.get("vaultName", "")
            fk_vid = fk.get("vaultId")
            jobs   = fk.get("dataTransferPerProtectionJob") or []
            data   = {
                "consumed": sum((j.get("storageConsumed") or 0) for j in jobs),
                "logical":  sum((j.get("numLogicalBytesTransferred") or 0) for j in jobs),
                "physical": sum((j.get("numPhysicalBytesTransferred") or 0) for j in jobs),
            }
            if fk_vn:  fk_by_name[fk_vn] = data
            if fk_vid: fk_by_id[fk_vid]   = data

        def _fk_stats(vname):
            d = fk_by_name.get(vname)
            if not d:
                vid = vault_id_by_name.get(vname)
                if vid: d = fk_by_id.get(vid)
            return d or {}

        fk_types = {"rpaas", "fortknox", "kfortknox", "krpaas"}

        def _is_fk(vname, vtype):
            return (any(x in vtype.lower() for x in fk_types)
                    or "fortknox" in vname.lower())

        # ── Replication rows ──────────────────────────────────────────────
        for tname in sorted(all_rep_targets):
            groups = sorted(rep_groups.get(tname) or [])
            gcnt   = len(groups)
            status = "Configured" if gcnt else "No groups assigned"
            rn = ws.max_row + 1
            row = [cd["name"], "Replication", tname, "", gcnt,
                   "", "", "", status, "",
                   ", ".join(groups)]
            for c, val in enumerate(row, 1):
                ws.cell(row=rn, column=c, value=val).font = _font()

        # ── Archival / vault rows ─────────────────────────────────────────
        shown = set()
        for tname in sorted(all_arch_targets):
            groups = sorted(arch_groups.get(tname) or [])
            gcnt   = len(groups)
            vtype  = arch_target_type.get(tname) or vault_type_by_name.get(tname, "")
            fkd    = _fk_stats(tname)
            note   = "FortKnox/RPaaS" if _is_fk(tname, vtype) else ""
            status = "Configured" if gcnt else "No groups assigned"
            rn     = ws.max_row + 1
            row    = [cd["name"], "Archival/Vault", tname, vtype, gcnt,
                      round(bytes_to_tb(fkd.get("consumed") or 0), 3),
                      round(bytes_to_tb(fkd.get("logical")  or 0), 3),
                      round(bytes_to_tb(fkd.get("physical") or 0), 3),
                      status, note,
                      ", ".join(groups)]
            for c, val in enumerate(row, 1):
                ws.cell(row=rn, column=c, value=val).font = _font()
            shown.add(tname)

        # Vaults that exist in the API but are not referenced by any policy
        for v in cd["vaults"]:
            vname = v.get("name") or v.get("vaultName") or ""
            if not vname or vname in shown:
                continue
            vtype = v.get("vaultType", "")
            fkd   = _fk_stats(vname)
            note  = "FortKnox/RPaaS" if _is_fk(vname, vtype) else ""
            rn    = ws.max_row + 1
            row   = [cd["name"], "Archival/Vault", vname, vtype, 0,
                     round(bytes_to_tb(fkd.get("consumed") or 0), 3),
                     round(bytes_to_tb(fkd.get("logical")  or 0), 3),
                     round(bytes_to_tb(fkd.get("physical") or 0), 3),
                     "Vault exists — not referenced by any policy", note, ""]
            for c, val in enumerate(row, 1):
                ws.cell(row=rn, column=c, value=val).font = _font()

    auto_fit_columns(ws)
    ws.column_dimensions["K"].width = 60


def _sheet_fortknox_detail(wb, all_data):
    """Per-protection-group data transfer to FortKnox / external vault targets.

    Full mode: covers the full lookback window (args.days, default 30 days).
    Quick mode: covers the last 1 day only (separate fk_data_1d fetch).
    """
    ws = wb.create_sheet("FortKnox Data Transfer")
    ws.freeze_panes = "A3"
    _title(ws,
           "FortKnox / External Target Data Transfer — Per Protection Group", "I")

    cols = ["Cluster", "Vault Name", "Vault Type",
            "Protection Group",
            "Logical Transferred (TB)", "Physical Transferred (TB)",
            "Storage Consumed (TB)", "Snapshots",
            "Period"]
    _hdr(ws, 2, cols)

    def _vtype_label(vt):
        tl = (vt or "").lower()
        if any(x in tl for x in ("fortknox", "rpaas", "krpaas", "kfortknox")):
            return "FortKnox / RPaaS"
        if any(x in tl for x in ("s3", "glacier", "azure", "gcp", "cloud")):
            return "Cloud Tier"
        if "nas" in tl:
            return "NAS Vault"
        if "tape" in tl:
            return "Tape"
        return vt or "Unknown"

    for cd in all_data:
        quick      = cd.get("quick", False)
        days       = cd.get("days", 30)
        fk_source  = cd.get("fk_data_1d") if quick else (cd.get("fk_data") or [])
        period     = "Last 1 day" if quick else f"Last {days} days"

        has_rows = False
        for fk in (fk_source or []):
            vname = fk.get("vaultName", "")
            vtype = _vtype_label(fk.get("vaultType", ""))
            for job in (fk.get("dataTransferPerProtectionJob") or []):
                jname    = job.get("protectionJobName", "")
                logical  = round(bytes_to_tb(job.get("numLogicalBytesTransferred")  or 0), 4)
                physical = round(bytes_to_tb(job.get("numPhysicalBytesTransferred") or 0), 4)
                consumed = round(bytes_to_tb(job.get("storageConsumed")             or 0), 4)
                snaps    = job.get("numSnapshots") or ""
                rn  = ws.max_row + 1
                row = [cd["name"], vname, vtype, jname,
                       logical, physical, consumed, snaps, period]
                for c, val in enumerate(row, 1):
                    ws.cell(row=rn, column=c, value=val).font = _font()
                has_rows = True

        if not has_rows:
            rn = ws.max_row + 1
            ws.cell(row=rn, column=1, value=cd["name"]).font = _font()
            ws.cell(row=rn, column=4,
                    value="No data transfer records for this cluster in the period.") \
               .font = _font(color="595959")

    auto_fit_columns(ws)
    ws.column_dimensions["B"].width = 30   # Vault Name
    ws.column_dimensions["D"].width = 35   # Protection Group


def _sheet_views(wb, all_data):
    ws = wb.create_sheet("Data Services")
    ws.freeze_panes = "A3"
    _title(ws, "NAS Views — Quota Utilization & Protocols", "L")

    cols = ["Cluster", "View Name", "Storage Domain", "Protocols",
            "Quota (GB)", "Used (GB)", "Usage %",
            "Case Insensitive", "Created",
            "SMB Mount Paths", "NFS Mount Paths", "Status"]
    _hdr(ws, 2, cols)

    for cd in all_data:
        for v in cd["views"]:
            lq  = v.get("logicalQuota") or \
                  (v.get("storagePolicy") or {}).get("storageQuotaPolicy") or {}
            quota_b = lq.get("hardLimitBytes") or 0
            vstats  = ((v.get("stats") or {}).get("dataUsageStats") or {})
            used_b  = vstats.get("totalLogicalUsageBytes") or 0
            quota_gb = round(bytes_to_gb(quota_b), 2) if quota_b else ""
            used_gb  = round(bytes_to_gb(used_b), 2)
            pct      = round(used_b / quota_b * 100, 1) if quota_b else ""

            protos = []
            if v.get("enableSmbViewDiscovery") or v.get("smbMountPaths"): protos.append("SMB")
            if v.get("nfsMountPaths") or v.get("enableNfsViewDiscovery"):  protos.append("NFS")
            if v.get("s3AccessEnabled"):                                   protos.append("S3")

            smb = ", ".join(v.get("smbMountPaths") or [])
            nfs = ", ".join(v.get("nfsMountPaths") or [])
            created_ms = v.get("createTimeMsecs") or 0
            created = (datetime.datetime.fromtimestamp(created_ms / 1000,
                                                        tz=datetime.timezone.utc)
                       .strftime("%Y-%m-%d") if created_ms else "")

            flag = "OK"
            if not quota_b:
                flag = "NO QUOTA"
            elif isinstance(pct, float) and pct >= VIEW_QUOTA_CRIT_PCT:
                flag = f"CRITICAL {pct:.0f}%"
            elif isinstance(pct, float) and pct >= VIEW_QUOTA_WARN_PCT:
                flag = f"WARN {pct:.0f}%"

            row = [cd["name"], v.get("name", ""), v.get("storageDomainName", ""),
                   "/".join(protos), quota_gb, used_gb, pct,
                   "Yes" if v.get("caseInsensitiveNamesEnabled") else "No",
                   created, smb, nfs, flag]
            rn = ws.max_row + 1
            for c, val in enumerate(row, 1):
                ws.cell(row=rn, column=c, value=val).font = _font()

            fc = ws.cell(row=rn, column=12)
            if "CRITICAL" in flag:
                fc.fill = _fill(RED);    fc.font = _font(bold=True, color=WHITE)
            elif "WARN" in flag:
                fc.fill = _fill(ORANGE); fc.font = _font(bold=True)
            elif "NO QUOTA" in flag:
                fc.fill = _fill(YELLOW); fc.font = _font(bold=True)
            else:
                fc.fill = _fill(LT_GREEN)

    auto_fit_columns(ws)


def _sheet_coverage(wb, all_data):
    ws = wb.create_sheet("Coverage Gaps")
    ws.freeze_panes = "A3"
    _title(ws, "Coverage Gaps — Groups with No Recent Successful Backup", "I")

    cols = ["Cluster", "Group Name", "Policy",
            "Active", "Paused",
            "Last Successful Run", "Hours Since Success",
            "Last Run Status", "Gap Severity"]
    _hdr(ws, 2, cols)

    for cd in all_data:
        pm = cd.get("policy_map", {})
        for g in cd["groups"]:
            lr  = g.get("lastRun") or {}
            lb  = lr.get("localBackupInfo") or {}
            st  = lb.get("status", "")
            sla_viol = lb.get("isSlaViolated") or lr.get("isSlaViolated") or False
            su  = lb.get("startTimeUsecs") or 0
            hrs = round((_now_usecs() - su) / 3_600_000_000, 1) if su else 9999

            is_gap = (g.get("isPaused")
                      or st in ("kFailure", "Failed")
                      or hrs > RPO_WARN_HOURS
                      or su == 0)
            if not is_gap:
                continue

            if g.get("isPaused"):
                sev = "PAUSED"
            elif st in ("kFailure", "Failed"):
                sev = "FAILED"
            elif hrs > RPO_CRIT_HOURS:
                sev = "CRITICAL RPO GAP"
            elif hrs > RPO_WARN_HOURS:
                sev = "RPO WARNING"
            else:
                sev = "UNKNOWN"

            policy_name = (g.get("policyName")
                           or pm.get(g.get("policyId"), "")
                           or g.get("policyId", ""))

            row = [cd["name"], g.get("name", ""),
                   policy_name,
                   "Yes" if g.get("isActive", True) else "No",
                   "Yes" if g.get("isPaused") else "No",
                   usecs_to_datetime(su) if su else "Never",
                   hrs if hrs < 9999 else "",
                   st, sev]
            rn = ws.max_row + 1
            for c, val in enumerate(row, 1):
                ws.cell(row=rn, column=c, value=val).font = _font()

            fc = ws.cell(row=rn, column=9)
            if sev in ("FAILED", "CRITICAL RPO GAP"):
                fc.fill = _fill(RED);    fc.font = _font(bold=True, color=WHITE)
            elif sev in ("PAUSED", "RPO WARNING"):
                fc.fill = _fill(ORANGE); fc.font = _font(bold=True)
            else:
                fc.fill = _fill(YELLOW); fc.font = _font(bold=True)

    auto_fit_columns(ws)


def _sheet_trends(wb, all_data):
    ws = wb.create_sheet("Trends (30d)")
    ws.freeze_panes = "A3"
    _title(ws, "30-Day Trends — Daily Backup Success Rate & Volume", "J")

    has_data = any(cd.get("v1_runs") or cd["group_runs"] for cd in all_data)
    if not has_data:
        ws["A3"] = "No run history available."
        ws["A3"].font = _font(bold=False, color="595959")
        return

    cols = ["Cluster", "Date", "Total Runs", "Successful", "Failed",
            "Warning", "Canceled", "Success % (incl. Warnings)", "SLA Violations", "Logical (GB)"]
    _hdr(ws, 2, cols)

    for cd in all_data:
        daily = {}

        # ── Primary: v1 protectionRuns — mirrors the Helios reporting page source ──
        # Single call per cluster; backupRun.slaViolated and
        # backupRun.stats.totalLogicalBackupSizeBytes are accurate.
        for run in (cd.get("v1_runs") or []):
            br    = run.get("backupRun") or {}
            stats = br.get("stats") or {}
            su    = stats.get("startTimeUsecs") or 0
            if not su:
                continue
            day = datetime.datetime.fromtimestamp(su / 1_000_000,
                                                   tz=datetime.timezone.utc) \
                          .strftime("%Y-%m-%d")
            if day not in daily:
                daily[day] = dict(total=0, succ=0, fail=0,
                                  warn=0, cancel=0, viols=0, logical=0)
            d  = daily[day]
            st = br.get("status", "")
            d["total"] += 1
            if st == "kSuccess":                d["succ"]   += 1
            elif st == "kFailure":              d["fail"]   += 1
            elif st == "kWarning":              d["warn"]   += 1
            elif "Cancel" in st:                d["cancel"] += 1
            if br.get("slaViolated"):           d["viols"]  += 1
            d["logical"] += stats.get("totalLogicalBackupSizeBytes") or 0

        # ── Fallback: v2 per-group runs (used when v1 unavailable) ──────────────
        if not daily:
            for runs in cd["group_runs"].values():
                for run in runs:
                    lb = run.get("localBackupInfo") or {}
                    su = lb.get("startTimeUsecs") or 0
                    if not su:
                        continue
                    day = datetime.datetime.fromtimestamp(su / 1_000_000,
                                                           tz=datetime.timezone.utc) \
                                  .strftime("%Y-%m-%d")
                    if day not in daily:
                        daily[day] = dict(total=0, succ=0, fail=0,
                                          warn=0, cancel=0, viols=0, logical=0)
                    d  = daily[day]
                    st = lb.get("status", "")
                    d["total"] += 1
                    if st in ("kSuccess", "Succeeded"): d["succ"]   += 1
                    elif st in ("kFailure", "Failed"):  d["fail"]   += 1
                    elif st == "kWarning":              d["warn"]   += 1
                    elif "Cancel" in st:                d["cancel"] += 1
                    if lb.get("isSlaViolated"):         d["viols"]  += 1
                    d["logical"] += (lb.get("localSnapshotStats") or {}).get("logicalSizeBytes") or 0

        if not daily:
            continue

        trend_start = ws.max_row + 1
        for day in sorted(daily):
            d   = daily[day]
            pct = round((d["succ"] + d["warn"]) / d["total"] * 100, 1) if d["total"] else 0
            rn  = ws.max_row + 1
            row = [cd["name"], day, d["total"], d["succ"], d["fail"],
                   d["warn"], d["cancel"], pct, d["viols"],
                   round(bytes_to_gb(d["logical"]), 1)]
            for c, val in enumerate(row, 1):
                ws.cell(row=rn, column=c, value=val).font = _font()
            sc = ws.cell(row=rn, column=8)
            if pct < SUCCESS_CRIT_PCT:
                sc.fill = _fill(RED);    sc.font = _font(bold=True, color=WHITE)
            elif pct < SUCCESS_WARN_PCT:
                sc.fill = _fill(ORANGE); sc.font = _font(bold=True)

        trend_end = ws.max_row
        n_days    = trend_end - trend_start + 1
        if n_days > 1:
            chart = LineChart()
            chart.style  = 10
            chart.height = 16   # taller for readability
            chart.width  = 26   # wide but fits to the right of the data cols

            # ── Chart title: larger font, no overlay ──────────────────────────
            _title_text = f"{cd['name']} — Daily Success Rate (incl. Warnings)"
            try:
                from openpyxl.chart.title import Title as _CTitle
                from openpyxl.drawing.text import (
                    RichTextBodyProperties as _BodyPr,
                    Paragraph as _Para,
                    RegularTextRun as _Run,
                    RunProperties as _RunPr,
                )
                from openpyxl.chart.text import RichText as _CRichText, Text as _CText
                _ctitle = _CTitle()
                _ctitle.overlay = False
                _rich = _CRichText()
                _rich.bodyPr = _BodyPr()
                _para = _Para()
                _run  = _Run()
                _run.t  = _title_text
                _run.rPr = _RunPr(sz=1400, b=True)   # 14 pt bold
                _para.r.append(_run)
                _rich.p.append(_para)
                _ctitle.tx = _CText(rich=_rich)
                chart.title = _ctitle
            except Exception:
                chart.title = _title_text   # fallback: plain string

            # ── Y axis: Success %, fixed 0–100, tick marks visible ───────────
            chart.y_axis.title         = "Success %"
            chart.y_axis.numFmt        = "0.0"
            chart.y_axis.majorTickMark = "out"
            chart.y_axis.tickLblPos    = "nextTo"
            chart.y_axis.delete        = False
            chart.y_axis.scaling.min   = 0
            chart.y_axis.scaling.max   = 100

            # ── X axis: date strings, tick marks visible ──────────────────────
            chart.x_axis.title         = "Date"
            chart.x_axis.majorTickMark = "out"
            chart.x_axis.tickLblPos    = "low"
            chart.x_axis.delete        = False
            # Reduce label density so dates don't overlap on long ranges
            if n_days > 60:
                chart.x_axis.tickLblSkip = 7
            elif n_days > 14:
                chart.x_axis.tickLblSkip = 2

            data_ref = Reference(ws, min_col=8, min_row=trend_start,
                                 max_row=trend_end)
            chart.add_data(data_ref)

            # ── StrRef forces <c:strRef> so Excel renders date strings ────────
            try:
                from openpyxl.chart.data_source import DataSource, StrRef
                from openpyxl.utils import get_column_letter as _gcl
                cat_formula = (f"'{ws.title}'!$B${trend_start}:$B${trend_end}")
                chart.series[0].cat = DataSource(strRef=StrRef(f=cat_formula))
            except ImportError:
                dates_ref = Reference(ws, min_col=2, min_row=trend_start,
                                      max_row=trend_end)
                chart.set_categories(dates_ref)

            # ── Series styling: Cohesity green line + circle markers ──────────
            s = chart.series[0]
            s.title = SeriesLabel(v="Success %")
            s.graphicalProperties.line.solidFill = COH_GREEN
            s.graphicalProperties.line.width     = 18000   # 2 pt in EMUs
            s.marker.symbol  = "circle"
            s.marker.size    = 5
            s.marker.graphicalProperties.solidFill   = COH_GREEN
            s.marker.graphicalProperties.line.solidFill = COH_GREEN

            # ── Legend at the bottom ──────────────────────────────────────────
            try:
                from openpyxl.chart.legend import Legend as _ChartLegend
                chart.legend = _ChartLegend()
                chart.legend.position = "b"
            except Exception:
                pass

            # Place chart to the right of the data columns (col L = 12)
            ws.add_chart(chart, f"L{trend_start}")

    auto_fit_columns(ws)


def _sheet_recommendations(wb, all_data):
    ws = wb.create_sheet("Recommendations")
    ws.freeze_panes = "A3"
    _title(ws, "Prioritized Recommendations", "G")

    cols = ["#", "Priority", "Cluster", "Category",
            "Finding", "Recommended Action", "Business Impact"]
    _hdr(ws, 2, cols)

    all_recs = sorted(
        [(cd["name"], rec) for cd in all_data for rec in cd["recommendations"]],
        key=lambda x: _PRIORITY_ORDER.get(x[1]["priority"], 99)
    )
    for i, (cname, rec) in enumerate(all_recs, 1):
        row = [i, rec["priority"], cname, rec["category"],
               rec["finding"], rec["action"], rec["impact"]]
        rn = ws.max_row + 1
        for c, val in enumerate(row, 1):
            ws.cell(row=rn, column=c, value=val).font = _font()

        pc = ws.cell(row=rn, column=2)
        if rec["priority"] == "CRITICAL":
            pc.fill = _fill(RED);        pc.font = _font(bold=True, color=WHITE)
        elif rec["priority"] == "HIGH":
            pc.fill = _fill(ORANGE);     pc.font = _font(bold=True)
        elif rec["priority"] == "MEDIUM":
            pc.fill = _fill(YELLOW);     pc.font = _font(bold=True)
        else:
            pc.fill = _fill(LIGHT_GRAY); pc.font = _font(bold=True)

    auto_fit_columns(ws)
    for col_letter in ("E", "F", "G"):
        ws.column_dimensions[col_letter].width = 55


def _sheet_disk_health(wb, all_data):
    ws = wb.create_sheet("Disk Health")
    ws.freeze_panes = "A3"
    _title(ws, "Disk Health — Per-Disk Status, SSD Wear & Encryption", "L")

    cols = ["Cluster", "Node ID", "Node IP", "Disk ID", "Disk Type",
            "Model", "Serial", "Status", "SSD Wear %",
            "Capacity (GB)", "Encrypted", "Storage Tier"]
    _hdr(ws, 2, cols)

    any_data = False
    for cd in all_data:
        disks = cd.get("disks") or []
        if not disks:
            continue
        any_data = True
        for disk in disks:
            status  = (disk.get("diskStatus") or disk.get("status") or "Unknown")
            # ssdUsedPercentage is the ELF field name (raw 0-100 integer)
            wear    = (disk.get("ssdUsedPercentage") or disk.get("ssdWearLevelPct")
                       or disk.get("wearLevel") or disk.get("lifeUsedPct"))
            _us     = disk.get("usageStats")
            # capacityInBytes is the ELF field name
            cap_b   = ((isinstance(_us, dict) and _us.get("totalPhysicalCapacityBytes"))
                       or disk.get("capacityInBytes") or disk.get("capacityBytes") or 0)
            # encryptionStatus is the ELF field name (string like "kEncrypted")
            enc_st  = disk.get("encryptionStatus") or ""
            enc     = bool(disk.get("encryptionEnabled") or disk.get("encrypted")
                           or (isinstance(enc_st, str) and "encrypt" in enc_st.lower()))
            tier    = disk.get("storageTier") or disk.get("diskTier") or ""
            dtype   = disk.get("diskType") or disk.get("type") or ""

            row = [cd["name"],
                   disk.get("_nodeId", ""), disk.get("_nodeIp", ""),
                   disk.get("diskId") or disk.get("id") or "",
                   dtype,
                   disk.get("diskModel") or disk.get("model") or "",
                   disk.get("diskSerial") or disk.get("serialNumber") or "",
                   status,
                   wear if wear is not None else "",
                   round(bytes_to_gb(cap_b), 2) if cap_b else "",
                   "Yes" if enc else "No",
                   tier]
            rn = ws.max_row + 1
            for c, val in enumerate(row, 1):
                ws.cell(row=rn, column=c, value=val).font = _font()

            # color Status cell (col 8)
            sc = ws.cell(row=rn, column=8)
            if status.lower() in ("kfailed", "failed", "kmissing", "missing"):
                sc.fill = _fill(RED);    sc.font = _font(bold=True, color=WHITE)
            elif status.lower() in ("kdegraded", "degraded", "kmarkedforremoval",
                                    "marked_for_removal"):
                sc.fill = _fill(ORANGE); sc.font = _font(bold=True)
            elif status.lower() in ("khealthy", "healthy", "kgood", "good"):
                sc.fill = _fill(LT_GREEN)

            # color SSD Wear cell (col 9)
            if wear is not None:
                wc = ws.cell(row=rn, column=9)
                if wear >= SSD_WEAR_CRIT_PCT:
                    wc.fill = _fill(RED);    wc.font = _font(bold=True, color=WHITE)
                elif wear >= SSD_WEAR_WARN_PCT:
                    wc.fill = _fill(ORANGE); wc.font = _font(bold=True)

    if not any_data:
        ws.cell(row=3, column=1,
                value="Disk detail not available via this API path (private endpoint).") \
           .font = _font(color="595959")

    auto_fit_columns(ws)


def _sheet_agent_health(wb, all_data):
    import datetime as _dt
    ws = wb.create_sheet("Agent Health")
    ws.freeze_panes = "A3"
    _title(ws, "Agent Health — Registered Host Agent Status & Versions", "I")

    cols = ["Cluster", "Host", "IP Address", "OS Type",
            "Agent Version", "Status", "Upgradable", "Cert Expiry", "Notes"]
    _hdr(ws, 2, cols)

    today    = _dt.date.today()
    any_data = False
    for cd in all_data:
        agents = cd.get("agents") or []
        if not agents:
            continue
        any_data = True
        for agent in agents:
            _hs2   = agent.get("healthStatus")
            status = ((isinstance(_hs2, dict) and _hs2.get("status"))
                      or (isinstance(_hs2, str) and _hs2)
                      or agent.get("agentStatus") or agent.get("status") or "")
            upg    = (agent.get("upgradability") or agent.get("upgradable") or "")

            # Certificate expiry (try usecs, then msecs)
            cert_exp = ""
            exp_raw  = (agent.get("certExpiryUsecs")
                        or agent.get("certExpiryTimeMsecs"))
            if exp_raw:
                if exp_raw > 1_000_000_000_000_000:
                    exp_dt = _dt.date.fromtimestamp(exp_raw / 1_000_000)
                elif exp_raw > 1_000_000_000_000:
                    exp_dt = _dt.date.fromtimestamp(exp_raw / 1_000)
                else:
                    exp_dt = _dt.date.fromtimestamp(exp_raw)
                days_left = (exp_dt - today).days
                cert_exp  = f"{exp_dt.isoformat()} ({days_left}d)"

            notes = ""
            err   = agent.get("lastUpgradeError")
            if err:
                notes = f"Upgrade error: {str(err)[:80]}"

            # Route hostname vs IP by content — the API sometimes puts FQDNs in hostIp
            _raw_host = agent.get("hostName") or agent.get("host") or ""
            _raw_ip   = agent.get("hostIp") or ""
            _host, _ip = _split_host_ip(_raw_host, _raw_ip)

            row = [cd["name"],
                   _host,
                   _ip,
                   agent.get("hostOsType") or agent.get("osType") or "",
                   agent.get("agentVersion") or "",
                   status, upg, cert_exp, notes]
            rn = ws.max_row + 1
            for c, val in enumerate(row, 1):
                ws.cell(row=rn, column=c, value=val).font = _font()

            sc = ws.cell(row=rn, column=6)
            if status.lower() in ("kunhealthy", "unhealthy", "unreachable",
                                  "kfailed", "failed"):
                sc.fill = _fill(RED);    sc.font = _font(bold=True, color=WHITE)
            elif status.lower() in ("kwarning", "warning", "kdegraded",
                                    "degraded"):
                sc.fill = _fill(ORANGE); sc.font = _font(bold=True)
            elif status.lower() in ("khealthy", "healthy", "krunning",
                                    "running"):
                sc.fill = _fill(LT_GREEN)

            # Upgradable flag
            uc = ws.cell(row=rn, column=7)
            if "upgradable" in upg.lower():
                uc.fill = _fill(YELLOW); uc.font = _font(bold=True)

    if not any_data:
        ws.cell(row=3, column=1,
                value="No agent data returned from /public/reports/agents.") \
           .font = _font(color="595959")

    auto_fit_columns(ws)
    ws.column_dimensions["B"].width = 35


def _sheet_source_coverage(wb, all_data):
    ws = wb.create_sheet("Source Coverage")
    ws.freeze_panes = "A3"
    _title(ws, "Source Coverage — Registered Sources & Protection Coverage", "H")

    cols = ["Cluster", "Source Name", "Environment", "Total Objects",
            "Protected", "Unprotected", "Coverage %", "Status"]
    _hdr(ws, 2, cols)

    for cd in all_data:
        sources = cd.get("sources") or []
        if not sources:
            continue
        for src in sources:
            if not isinstance(src, dict):
                continue
            _si  = src.get("sourceInfo")
            si   = _si if isinstance(_si, dict) else src
            name = (si.get("name") or si.get("sourceName")
                    or src.get("name") or "")
            env  = (si.get("environment") or si.get("sourceEnvironment")
                    or src.get("environment") or "")

            _ss  = src.get("stats");       _ss  = _ss  if isinstance(_ss,  dict) else {}
            _oc  = src.get("objectCount"); _oc  = _oc  if isinstance(_oc,  dict) else {}
            prot   = (src.get("protectedObjectsCount")
                      or src.get("numProtectedObjects")
                      or _ss.get("protectedObjectsCount")
                      or _oc.get("protectedCount") or 0)
            unprot = (src.get("unprotectedObjectsCount")
                      or src.get("numUnprotectedObjects")
                      or _ss.get("unprotectedObjectsCount")
                      or _oc.get("unprotectedCount") or 0)
            total  = prot + unprot
            pct    = round(prot / total * 100, 1) if total else ""

            if isinstance(pct, float):
                if pct >= 95:    flag = "Good"
                elif pct >= 75:  flag = f"Partial ({pct:.0f}%)"
                else:            flag = f"Low ({pct:.0f}%)"
            else:
                flag = "Unknown"

            row = [cd["name"], name, env,
                   total if total else "",
                   prot or "", unprot or "", pct, flag]
            rn = ws.max_row + 1
            for c, val in enumerate(row, 1):
                ws.cell(row=rn, column=c, value=val).font = _font()

            fc = ws.cell(row=rn, column=8)
            if isinstance(pct, float) and pct < 75:
                fc.fill = _fill(RED);    fc.font = _font(bold=True, color=WHITE)
            elif isinstance(pct, float) and pct < 95:
                fc.fill = _fill(ORANGE); fc.font = _font(bold=True)
            elif isinstance(pct, float):
                fc.fill = _fill(LT_GREEN)

    auto_fit_columns(ws)


def _sheet_user_security(wb, all_data):
    import datetime as _dt
    ws = wb.create_sheet("User Security")
    ws.freeze_panes = "A3"
    _title(ws, "User Security — Accounts, MFA Status & Access Control", "J")

    cols = ["Cluster", "Username", "Domain", "Type",
            "Roles", "Active", "Locked", "MFA Enabled",
            "Last Login", "Notes"]
    _hdr(ws, 2, cols)

    _ADMIN_ROLES = {"COHESITY_ADMIN", "Admin", "kAdmin", "kSuperAdmin",
                    "super_admin", "admin"}
    any_data = False
    for cd in all_data:
        users = cd.get("users") or []
        if not users:
            continue
        any_data = True
        for user in users:
            roles  = ", ".join(user.get("roles") or [])
            locked = bool(user.get("locked") or user.get("isLocked"))
            active = not bool(user.get("isDeleted") or user.get("isSuspended"))
            mfa    = bool(user.get("mfaEnabled") or user.get("isMfaEnabled"))

            last_ms = (user.get("lastLoginTimeMsecs")
                       or user.get("lastLoginTime") or 0)
            if last_ms:
                if last_ms > 1_000_000_000_000_000:  # usecs
                    last_ms = last_ms // 1000
                last_login = _dt.datetime.fromtimestamp(
                    last_ms / 1000, tz=_dt.timezone.utc).strftime("%Y-%m-%d")
            else:
                last_login = "Never"

            is_admin = bool(
                {r.lower() for r in (user.get("roles") or [])} &
                {r.lower() for r in _ADMIN_ROLES}
            )
            domain = user.get("domain") or ""
            utype  = (user.get("userType")
                      or ("Local" if domain.lower() in ("local", "")
                          else "AD/LDAP"))

            notes = []
            if locked:      notes.append("LOCKED")
            if not mfa:     notes.append("No MFA")

            row = [cd["name"], user.get("username", ""), domain, utype,
                   roles,
                   "Yes" if active else "No",
                   "Yes" if locked else "No",
                   "Yes" if mfa else "No",
                   last_login,
                   "; ".join(notes) if notes else "OK"]
            rn = ws.max_row + 1
            for c, val in enumerate(row, 1):
                ws.cell(row=rn, column=c, value=val).font = _font()

            if locked:
                ws.cell(row=rn, column=7).fill = _fill(RED)
                ws.cell(row=rn, column=7).font = _font(bold=True, color=WHITE)

            # MFA color: red for admins, yellow for non-admins without MFA
            if not mfa and is_admin:
                ws.cell(row=rn, column=8).fill = _fill(RED)
                ws.cell(row=rn, column=8).font = _font(bold=True, color=WHITE)
            elif not mfa:
                ws.cell(row=rn, column=8).fill = _fill(YELLOW)
                ws.cell(row=rn, column=8).font = _font(bold=True)

    if not any_data:
        ws.cell(row=3, column=1,
                value="No user data returned from /public/users.") \
           .font = _font(color="595959")

    auto_fit_columns(ws)
    ws.column_dimensions["E"].width = 30


# ─────────────────────────────────────────────────────────────────────────────
# TOPOLOGY DIAGRAM  (draw.io editable + matplotlib preview)
# ─────────────────────────────────────────────────────────────────────────────
#
# Cohesity brand palette used throughout:
_DG_TEAL    = "#96C73D"   # Cohesity green          — source clusters
_DG_TEAL_DK = "#6B8F2C"   # Darker green            — cluster stroke
_DG_NAVY    = "#1A3045"   # Dark navy               — header / background
_DG_REPL    = "#0062B1"   # Blue                   — replication targets
_DG_ARCH    = "#E07A00"   # Amber                  — archival vaults
_DG_FK      = "#1A5C3A"   # Dark green             — FortKnox vaults
_DG_FG      = "#FFFFFF"   # White text on dark boxes


def _topology_data(all_data):
    """Extract the topology graph from all_data.

    Returns a dict:
      clusters   — list of cluster dicts (id, name, groups, sources, workloads, cap_tb)
      repl_nodes — list of unique replication-target dicts (id, name)
      arch_nodes — list of unique archival-vault dicts (id, name)
      fk_nodes   — list of unique FortKnox-vault dicts (id, name)
      edges      — list of edge dicts (src, tgt, type="repl"|"arch"|"fk")
    """
    _fk_types = {"krpaas", "kfortknox", "kfort_knox", "krpaasarchival"}

    def _is_fk(tt, tname=""):
        tl = tt.lower(); nl = tname.lower()
        return any(x in tl for x in _fk_types) or "fortknox" in tl or "fortknox" in nl

    clusters = []
    all_repl  = {}   # name → node dict
    all_arch  = {}
    all_fk    = {}
    edges     = []

    for i, cd in enumerate(all_data):
        cid = f"c{i}"

        # Workload types (de-prefix the leading "k", title-case)
        workloads = sorted({
            g.get("environment", "").lstrip("k").title()
            for g in cd["groups"] if g.get("environment")
        })[:4]

        # Approximate protected capacity (sum logical bytes across groups)
        raw_cap = 0
        for g in cd["groups"]:
            st = g.get("latestSuccessfulRunStats") or g.get("lastSuccessfulRunStats") or {}
            raw_cap += st.get("logicalSizeBytes") or 0
        if raw_cap == 0:
            raw_cap = cd["info"].get("usedCapacityBytes") or 0
        cap_tb = raw_cap / (1024 ** 4) if raw_cap else 0

        for p in cd["policies"]:
            rtp = p.get("remoteTargetPolicy") or {}
            for t in (rtp.get("replicationTargets") or []):
                name = (t.get("targetName")
                        or (t.get("remoteTargetConfig") or {}).get("clusterName")
                        or (t.get("remoteTargetConfig") or {}).get("name") or "")
                if not name:
                    continue
                if name not in all_repl:
                    all_repl[name] = {"id": f"r{len(all_repl)}", "name": name}
                edges.append({"src": cid, "tgt": all_repl[name]["id"], "type": "repl"})

            for t in (rtp.get("archivalTargets") or []):
                cfg   = t.get("archivalTargetConfig") or {}
                tt    = cfg.get("targetType") or t.get("targetType") or ""
                tname = t.get("targetName") or cfg.get("name") or cfg.get("vaultName") or ""
                if not tname:
                    continue
                if _is_fk(tt, tname):
                    if tname not in all_fk:
                        all_fk[tname] = {"id": f"f{len(all_fk)}", "name": tname}
                    edges.append({"src": cid, "tgt": all_fk[tname]["id"], "type": "fk"})
                else:
                    if tname not in all_arch:
                        all_arch[tname] = {"id": f"a{len(all_arch)}", "name": tname}
                    edges.append({"src": cid, "tgt": all_arch[tname]["id"], "type": "arch"})

        # Pick up FK vaults that appear in the vaults list but not in policies
        for v in (cd.get("vaults") or []):
            vtype = v.get("vaultType") or ""
            vname = v.get("name") or ""
            if vname and _is_fk(vtype, vname) and vname not in all_fk:
                all_fk[vname] = {"id": f"f{len(all_fk)}", "name": vname}

        clusters.append({
            "id":       cid,
            "name":     cd["name"],
            "groups":   len(cd["groups"]),
            "sources":  len(cd.get("sources") or []),
            "workloads": workloads,
            "cap_tb":   cap_tb,
        })

    # Deduplicate edges (multiple policies can produce identical edges)
    seen = set()
    unique_edges = []
    for e in edges:
        k = (e["src"], e["tgt"], e["type"])
        if k not in seen:
            seen.add(k)
            unique_edges.append(e)

    return {
        "clusters":   clusters,
        "repl_nodes": list(all_repl.values()),
        "arch_nodes": list(all_arch.values()),
        "fk_nodes":   list(all_fk.values()),
        "edges":      unique_edges,
    }


# ── draw.io XML generator ─────────────────────────────────────────────────────

def _generate_topology_drawio(all_data, out_path):
    """Write an editable draw.io topology diagram to out_path + '.drawio'.

    Returns True on success, False on error.
    Node types and brand colors:
      Clusters       — Cohesity teal rounded rectangles (left column)
      Replication    — Blue rounded rectangles (2nd column)
      Archival Vault — Amber cylinders (3rd column)
      FortKnox/RPaaS — Dark-green rounded rectangles with lock symbol (4th column)
    """
    import html as _html

    def _he(s):
        return _html.escape(str(s), quote=False)

    def _av(html_label):
        """XML-escape an HTML label for use as a draw.io value attribute."""
        return (html_label
                .replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace('"', "&quot;"))

    topo       = _topology_data(all_data)
    clusters   = topo["clusters"]
    repl_nodes = topo["repl_nodes"]
    arch_nodes = topo["arch_nodes"]
    fk_nodes   = topo["fk_nodes"]
    edges      = topo["edges"]

    # ── Layout constants (draw.io units, 1 unit ≈ 1 px at 100% zoom) ────────
    MARGIN_X, MARGIN_Y = 40, 70
    C_W,  C_H  = 250, 110   # source cluster box
    T_W,  T_H  = 220, 80    # target node box
    V_GAP      = 22          # vertical gap between nodes in same column
    COL_GAP    = 80          # horizontal gap between columns

    col_c  = MARGIN_X
    col_r  = col_c + C_W + COL_GAP
    col_a  = col_r + T_W + COL_GAP
    col_f  = col_a + T_W + COL_GAP

    def _col_y(nodes, idx, node_h):
        return MARGIN_Y + 30 + idx * (node_h + V_GAP)

    cells  = []
    _ctr   = [2]

    def _nid():
        i = _ctr[0]; _ctr[0] += 1; return str(i)

    def _vertex(vid, av_label, x, y, w, h, style):
        cells.append(
            f'<mxCell id="{vid}" value="{av_label}" style="{style}" '
            f'vertex="1" parent="1">'
            f'<mxGeometry x="{x}" y="{y}" width="{w}" height="{h}" as="geometry"/>'
            f'</mxCell>'
        )

    def _edge(eid, src, tgt, style, label=""):
        cells.append(
            f'<mxCell id="{eid}" value="{label}" style="{style}" '
            f'edge="1" source="{src}" target="{tgt}" parent="1">'
            f'<mxGeometry relative="1" as="geometry"/>'
            f'</mxCell>'
        )

    # ── Styles ───────────────────────────────────────────────────────────────
    S_CLUSTER = (
        f"rounded=1;whiteSpace=wrap;html=1;"
        f"fillColor={_DG_TEAL};strokeColor={_DG_TEAL_DK};strokeWidth=2;"
        f"fontColor=#FFFFFF;fontSize=13;fontStyle=1;align=left;"
        f"spacingLeft=10;arcSize=6;"
    )
    S_REPL = (
        f"rounded=1;whiteSpace=wrap;html=1;"
        f"fillColor={_DG_REPL};strokeColor=#004A8A;strokeWidth=2;"
        f"fontColor=#FFFFFF;fontSize=12;align=center;arcSize=6;"
    )
    S_ARCH = (
        f"shape=cylinder3;whiteSpace=wrap;html=1;direction=south;"
        f"fillColor={_DG_ARCH};strokeColor=#B05A00;strokeWidth=2;"
        f"fontColor=#FFFFFF;fontSize=12;align=center;size=12;"
    )
    S_FK = (
        f"shape=cloud;whiteSpace=wrap;html=1;"
        f"fillColor={_DG_FK};strokeColor=#0E3B25;strokeWidth=2;"
        f"fontColor=#FFFFFF;fontSize=12;fontStyle=1;align=center;"
    )

    def _s_hdr(color):
        return (f"text;html=1;align=center;verticalAlign=middle;"
                f"fontStyle=1;fontSize=14;fontColor={color};")

    def _s_edge(color, dashed=False):
        d = "dashed=1;dashPattern=8 4;" if dashed else ""
        return (
            f"edgeStyle=elbowEdgeStyle;elbow=vertical;{d}"
            f"strokeColor={color};strokeWidth=2;fillColor={color};"
            f"exitX=1;exitY=0.5;exitDx=0;exitDy=0;"
            f"entryX=0;entryY=0.5;entryDx=0;entryDy=0;"
            f"endArrow=block;endFill=1;"
        )

    # ── Watermark / generation note ──────────────────────────────────────────
    nid = _nid()
    cells.append(
        f'<mxCell id="{nid}" value="Generated by Cohesity Health Check v{__version__}" '
        f'style="text;html=1;align=right;verticalAlign=top;fontSize=9;fontColor=#AAAAAA;" '
        f'vertex="1" parent="1">'
        f'<mxGeometry x="{col_f + T_W + 20}" y="5" width="270" height="20" as="geometry"/>'
        f'</mxCell>'
    )

    # ── Column headers ────────────────────────────────────────────────────────
    _vertex(_nid(), _av("Source Clusters"), col_c, MARGIN_Y - 28, C_W, 24, _s_hdr(_DG_TEAL))
    if repl_nodes:
        _vertex(_nid(), _av("Replication Targets"), col_r, MARGIN_Y - 28, T_W, 24, _s_hdr(_DG_REPL))
    if arch_nodes:
        _vertex(_nid(), _av("Archival Vaults"), col_a, MARGIN_Y - 28, T_W, 24, _s_hdr(_DG_ARCH))
    if fk_nodes:
        _vertex(_nid(), _av("FortKnox / RPaaS"), col_f, MARGIN_Y - 28, T_W, 24, _s_hdr(_DG_FK))

    # ── Source cluster nodes ──────────────────────────────────────────────────
    for idx, cl in enumerate(clusters):
        y   = _col_y(clusters, idx, C_H)
        wls = ", ".join(cl["workloads"]) if cl["workloads"] else ""
        det = f"{cl['groups']} groups"
        if cl["sources"]: det += f" &bull; {cl['sources']} sources"
        if cl["cap_tb"] > 0.01: det += f" &bull; {cl['cap_tb']:.1f} TB"
        lbl = (f"<b>{_he(cl['name'])}</b><br>"
               + (f"{_he(wls)}<br>" if wls else "")
               + det)
        _vertex(cl["id"], _av(lbl), col_c, y, C_W, C_H, S_CLUSTER)

    # ── Replication target nodes ──────────────────────────────────────────────
    for idx, n in enumerate(repl_nodes):
        y   = _col_y(repl_nodes, idx, T_H)
        lbl = f"<b>{_he(n['name'])}</b><br>Cluster"
        _vertex(n["id"], _av(lbl), col_r, y, T_W, T_H, S_REPL)

    # ── Archival vault nodes ──────────────────────────────────────────────────
    for idx, n in enumerate(arch_nodes):
        y   = _col_y(arch_nodes, idx, T_H)
        lbl = f"<b>{_he(n['name'])}</b><br>Archival Vault"
        _vertex(n["id"], _av(lbl), col_a, y, T_W, T_H, S_ARCH)

    # ── FortKnox nodes ────────────────────────────────────────────────────────
    for idx, n in enumerate(fk_nodes):
        y   = _col_y(fk_nodes, idx, T_H)
        lbl = f"&#x1F512; <b>{_he(n['name'])}</b><br>FortKnox / RPaaS"
        _vertex(n["id"], _av(lbl), col_f, y, T_W, T_H, S_FK)

    # ── Edges ─────────────────────────────────────────────────────────────────
    for e in edges:
        if e["type"] == "repl":
            style = _s_edge(_DG_REPL)
        elif e["type"] == "arch":
            style = _s_edge(_DG_ARCH)
        else:
            style = _s_edge(_DG_FK, dashed=True)
        _edge(_nid(), e["src"], e["tgt"], style)

    # ── Canvas dimensions ─────────────────────────────────────────────────────
    right_items = max(len(repl_nodes), len(arch_nodes), len(fk_nodes), 1)
    canvas_h = MARGIN_Y + 30 + max(len(clusters), right_items) * (C_H + V_GAP) + 60
    canvas_w = col_f + T_W + 60

    # ── Assemble XML ──────────────────────────────────────────────────────────
    cells_xml = "\n        ".join(cells)
    xml = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<mxfile host="app.diagrams.net" agent="Cohesity Health Check" '
        f'version="21.7.5" type="device">\n'
        '  <diagram id="cohesity-topology" name="Environment Topology">\n'
        f'    <mxGraphModel dx="1422" dy="762" grid="1" gridSize="10" '
        f'guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="0" '
        f'pageScale="1" pageWidth="{canvas_w}" pageHeight="{canvas_h}" '
        f'math="0" shadow="0">\n'
        '      <root>\n'
        '        <mxCell id="0" />\n'
        '        <mxCell id="1" parent="0" />\n'
        f'        {cells_xml}\n'
        '      </root>\n'
        '    </mxGraphModel>\n'
        '  </diagram>\n'
        '</mxfile>\n'
    )

    try:
        drawio_path = out_path + ".drawio"
        with open(drawio_path, "w", encoding="utf-8") as fh:
            fh.write(xml)
        return drawio_path
    except Exception as exc:
        print(f"  WARN: Could not write topology diagram: {exc}")
        return None


# ── Native Word shapes topology ──────────────────────────────────────────────

def _render_topology_word_shapes(doc, all_data):
    """Render environment topology using native Word shapes (DrawingML).

    Creates editable rounded rectangles (source clusters in Cohesity green),
    cloud shapes (FortKnox), cylinders (archive), and connector arrows.
    All shapes are anchored inside a single paragraph with wrapNone positioning.
    Returns True on success, False if python-docx OxmlElement API is unavailable.
    """
    try:
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn, nsmap as _nsmap
        from docx.shared import Emu, Pt
    except ImportError:
        return False

    # python-docx ≤1.2 does not register the wps namespace; patch it so
    # OxmlElement('wps:wsp') resolves correctly.  setdefault is idempotent.
    _nsmap.setdefault(
        'wps', 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape')

    try:
        return _render_topology_word_shapes_inner(doc, all_data,
                                                   OxmlElement, qn, Emu, Pt)
    except Exception as exc:
        print(f"  WARN: native Word shapes failed: {exc}")
        return False


def _render_topology_word_shapes_inner(doc, all_data, OxmlElement, qn, Emu, Pt):
    """Inner implementation — separated so the caller can catch all errors."""
    from docx.shared import RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    topo       = _topology_data(all_data)
    clusters   = topo["clusters"]
    repl_nodes = topo["repl_nodes"]
    arch_nodes = topo["arch_nodes"]
    fk_nodes   = topo["fk_nodes"]
    edges      = topo["edges"]

    if not clusters:
        return False

    # ── Constants (EMU: 914400 per inch) ─────────────────────────────────
    EMU = 914400

    CLR_CLUSTER = "96C73D";  CLR_CLUSTER_S = "6B8F2C"
    CLR_REPL    = "0062B1";  CLR_REPL_S    = "004A8A"
    CLR_ARCH    = "E07A00";  CLR_ARCH_S    = "B05A00"
    CLR_FK      = "1A5C3A";  CLR_FK_S      = "0E3B25"

    C_W  = int(2.05 * EMU);  C_H  = int(0.72 * EMU)
    T_W  = int(1.55 * EMU);  T_H  = int(0.58 * EMU)
    FK_W = int(1.70 * EMU);  FK_H = int(0.72 * EMU)
    V_GAP = int(0.20 * EMU)
    LABEL_H = int(0.28 * EMU)

    # ── Layout ───────────────────────────────────────────────────────────
    tcols = []  # (type, nodes, width, height, geom, fill, stroke, label)
    if repl_nodes:
        tcols.append(("repl", repl_nodes, T_W, T_H, "roundRect",
                       CLR_REPL, CLR_REPL_S, "Replication Targets"))
    if arch_nodes:
        tcols.append(("arch", arch_nodes, T_W, T_H, "can",
                       CLR_ARCH, CLR_ARCH_S, "Archival Vaults"))
    if fk_nodes:
        tcols.append(("fk", fk_nodes, FK_W, FK_H, "cloud",
                       CLR_FK, CLR_FK_S, "FortKnox"))

    # Column X positions (left edge)
    col_x = {"cluster": 0}
    x_cursor = C_W + int(0.55 * EMU)
    for ct, nodes, tw, *_ in tcols:
        col_x[ct] = x_cursor
        x_cursor += tw + int(0.30 * EMU)

    # Y positions
    SHAPES_TOP = LABEL_H + int(0.08 * EMU)

    def _ys(n, h):
        return [SHAPES_TOP + i * (h + V_GAP) for i in range(n)]

    cluster_ys = _ys(len(clusters), C_H)
    c_total_h  = len(clusters) * (C_H + V_GAP) - V_GAP if clusters else C_H

    tcol_ys = {}
    for ct, nodes, tw, th, *_ in tcols:
        ys = _ys(len(nodes), th)
        t_total = len(nodes) * (th + V_GAP) - V_GAP if nodes else th
        offset  = max((c_total_h - t_total) // 2, 0)
        tcol_ys[ct] = [y + offset for y in ys]

    max_right_h = max(
        (len(nodes) * (th + V_GAP) - V_GAP
         for _, nodes, _, th, *_ in tcols), default=0)
    total_h = SHAPES_TOP + max(c_total_h, max_right_h) + int(0.20 * EMU)

    # Position lookup: id → (left_x, top_y, width, height)
    pos = {}
    for i, cl in enumerate(clusters):
        pos[cl["id"]] = (col_x["cluster"], cluster_ys[i], C_W, C_H)
    for ct, nodes, tw, th, *_ in tcols:
        for i, n in enumerate(nodes):
            pos[n["id"]] = (col_x[ct], tcol_ys[ct][i], tw, th)

    # ── Shape ID counter ─────────────────────────────────────────────────
    _sid = [200]
    def _nid():
        _sid[0] += 1;  return _sid[0]

    # ── XML element helper ───────────────────────────────────────────────
    def _mk(tag, attrib=None, text=None):
        el = OxmlElement(tag)
        for k, v in (attrib or {}).items():
            el.set(k if ':' not in k else qn(k), str(v))
        if text is not None:
            el.text = str(text)
        return el

    def _anchor_frame(sid, x, y, w, h, z=0):
        """Build wp:anchor → a:graphic → a:graphicData; return (anchor, gd)."""
        anc = _mk('wp:anchor', {
            'distT': '0', 'distB': '0', 'distL': '0', 'distR': '0',
            'simplePos': '0',
            'relativeHeight': str(251660000 + z + sid),
            'behindDoc': '0', 'locked': '0',
            'layoutInCell': '1', 'allowOverlap': '1',
        })
        anc.append(_mk('wp:simplePos', {'x': '0', 'y': '0'}))
        ph = _mk('wp:positionH', {'relativeFrom': 'column'})
        ph.append(_mk('wp:posOffset', text=str(int(x))))
        anc.append(ph)
        pv = _mk('wp:positionV', {'relativeFrom': 'paragraph'})
        pv.append(_mk('wp:posOffset', text=str(int(y))))
        anc.append(pv)
        anc.append(_mk('wp:extent', {'cx': str(int(w)), 'cy': str(int(h))}))
        anc.append(_mk('wp:effectExtent',
                        {'l': '0', 't': '0', 'r': '19050', 'b': '19050'}))
        anc.append(_mk('wp:wrapNone'))
        anc.append(_mk('wp:docPr', {'id': str(sid), 'name': f'Shape{sid}'}))
        anc.append(_mk('wp:cNvGraphicFramePr'))
        gr  = _mk('a:graphic')
        gd  = _mk('a:graphicData', {
            'uri': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'})
        gr.append(gd)
        anc.append(gr)
        return anc, gd

    def _append_to_para(para, anchor):
        r   = _mk('w:r')
        dwg = _mk('w:drawing')
        dwg.append(anchor)
        r.append(dwg)
        para._p.append(r)

    # ── Add a filled shape with text ─────────────────────────────────────
    def _shape(para, x, y, w, h, geom, fill, stroke,
               text_lines, fsz=10, fclr="FFFFFF", z=100):
        sid = _nid()
        anc, gd = _anchor_frame(sid, x, y, w, h, z)

        wsp = _mk('wps:wsp')
        wsp.append(_mk('wps:cNvSpPr'))

        spPr = _mk('wps:spPr')
        xfrm = _mk('a:xfrm')
        xfrm.append(_mk('a:off', {'x': '0', 'y': '0'}))
        xfrm.append(_mk('a:ext', {'cx': str(int(w)), 'cy': str(int(h))}))
        spPr.append(xfrm)
        pg = _mk('a:prstGeom', {'prst': geom})
        pg.append(_mk('a:avLst'))
        spPr.append(pg)
        if fill == "NONE":
            spPr.append(_mk('a:noFill'))
        else:
            sf = _mk('a:solidFill');  sf.append(_mk('a:srgbClr', {'val': fill}))
            spPr.append(sf)
        if stroke == "NONE":
            ln = _mk('a:ln');  ln.append(_mk('a:noFill'));  spPr.append(ln)
        else:
            ln = _mk('a:ln', {'w': '12700'})
            lf = _mk('a:solidFill');  lf.append(_mk('a:srgbClr', {'val': stroke}))
            ln.append(lf);  spPr.append(ln)
        wsp.append(spPr)

        # Text body
        txbx = _mk('wps:txbx')
        tc   = _mk('w:txbxContent')
        for idx, line in enumerate(text_lines):
            wp  = _mk('w:p')
            pPr = _mk('w:pPr')
            pPr.append(_mk('w:jc', {'w:val': 'center'}))
            sp = _mk('w:spacing')
            sp.set(qn('w:after'), '0')
            sp.set(qn('w:line'), '216')
            sp.set(qn('w:lineRule'), 'auto')
            pPr.append(sp)
            wp.append(pPr)
            wr  = _mk('w:r')
            rPr = _mk('w:rPr')
            rPr.append(_mk('w:color', {'w:val': fclr}))
            fs = fsz if idx == 0 else max(fsz - 1, 7)
            rPr.append(_mk('w:sz',   {'w:val': str(fs * 2)}))
            rPr.append(_mk('w:szCs', {'w:val': str(fs * 2)}))
            if idx == 0:
                rPr.append(_mk('w:b'));  rPr.append(_mk('w:bCs'))
            wr.append(rPr)
            t = _mk('w:t')
            t.set(qn('xml:space'), 'preserve')
            t.text = line
            wr.append(t)
            wp.append(wr)
            tc.append(wp)
        txbx.append(tc)
        wsp.append(txbx)

        bPr = _mk('wps:bodyPr', {
            'anchor': 'ctr', 'anchorCtr': '0', 'wrap': 'square',
            'lIns': '91440', 'tIns': '36576',
            'rIns': '91440', 'bIns': '36576',
        })
        wsp.append(bPr)
        gd.append(wsp)
        _append_to_para(para, anc)

    # ── Add a connector arrow ────────────────────────────────────────────
    def _connector(para, x1, y1, x2, y2, color, z=50):
        sid  = _nid()
        bx   = min(x1, x2)
        by   = min(y1, y2)
        cx   = abs(x2 - x1) or 1
        cy   = abs(y2 - y1) or 1
        anc, gd = _anchor_frame(sid, bx, by, cx, cy, z)

        wsp = _mk('wps:wsp')
        wsp.append(_mk('wps:cNvCnPr'))
        spPr = _mk('wps:spPr')
        xfrm = _mk('a:xfrm')
        if x2 < x1: xfrm.set('flipH', '1')
        if y2 < y1: xfrm.set('flipV', '1')
        xfrm.append(_mk('a:off', {'x': '0', 'y': '0'}))
        xfrm.append(_mk('a:ext', {'cx': str(int(cx)), 'cy': str(int(cy))}))
        spPr.append(xfrm)
        pg = _mk('a:prstGeom', {'prst': 'straightConnector1'})
        pg.append(_mk('a:avLst'))
        spPr.append(pg)
        ln = _mk('a:ln', {'w': '19050'})
        lf = _mk('a:solidFill');  lf.append(_mk('a:srgbClr', {'val': color}))
        ln.append(lf)
        ln.append(_mk('a:tailEnd', {'type': 'triangle', 'w': 'med', 'len': 'med'}))
        spPr.append(ln)
        wsp.append(spPr)
        wsp.append(_mk('wps:bodyPr'))
        gd.append(wsp)
        _append_to_para(para, anc)

    # ── Column label (no-fill text shape) ────────────────────────────────
    def _label(para, x, w, text, color, z=200):
        _shape(para, x, 0, w, LABEL_H, "rect", "NONE", "NONE",
               [text], fsz=9, fclr=color, z=z)

    # ── Build the diagram ────────────────────────────────────────────────
    p_diag = doc.add_paragraph()
    p_diag.paragraph_format.space_after = Emu(total_h + int(0.15 * EMU))

    # Column labels
    _label(p_diag, col_x["cluster"], C_W, "Source Clusters", CLR_CLUSTER)
    for ct, _, tw, *rest in tcols:
        _label(p_diag, col_x[ct], tw, rest[-1], rest[2])  # rest[-1]=label, rest[2]=fill

    # Connectors (draw first → lower z-order)
    drawn = set()
    for e in edges:
        key = (e["src"], e["tgt"])
        if key in drawn:
            continue
        drawn.add(key)
        sp = pos.get(e["src"])
        tp = pos.get(e["tgt"])
        if not sp or not tp:
            continue
        sx = sp[0] + sp[2]            # right edge of source
        sy = sp[1] + sp[3] // 2       # vertical center
        tx = tp[0]                     # left edge of target
        ty = tp[1] + tp[3] // 2       # vertical center
        clr = {"repl": CLR_REPL, "arch": CLR_ARCH, "fk": CLR_FK}[e["type"]]
        _connector(p_diag, sx, sy, tx, ty, clr)

    # Source cluster shapes
    for i, cl in enumerate(clusters):
        wls = ", ".join(cl["workloads"]) if cl["workloads"] else ""
        det = f"{cl['groups']} groups"
        if cl["sources"]:
            det += f" \u00b7 {cl['sources']} sources"
        if cl["cap_tb"] > 0.01:
            det += f" \u00b7 {cl['cap_tb']:.1f} TB"
        lines = [cl["name"]]
        if wls:
            lines.append(wls)
        lines.append(det)
        _shape(p_diag, col_x["cluster"], cluster_ys[i], C_W, C_H,
               "roundRect", CLR_CLUSTER, CLR_CLUSTER_S, lines, fsz=10)

    # Target shapes
    for ct, nodes, tw, th, geom, fill, stroke, label in tcols:
        sub = {"repl": "Replication Target", "arch": "Archival Vault",
               "fk": "FortKnox (Indelible)"}[ct]
        for i, n in enumerate(nodes):
            _shape(p_diag, col_x[ct], tcol_ys[ct][i], tw, th,
                   geom, fill, stroke, [n["name"], sub], fsz=9)

    # Caption
    cap = doc.add_paragraph(
        "Figure: Environment Topology \u2014 source clusters (green), "
        "replication targets (blue), archival vaults (amber), "
        "FortKnox (dark-green cloud). All shapes are editable in Word. "
        "An editable draw.io diagram is also saved alongside this report."
    )
    cap.runs[0].font.size = Pt(8)
    cap.runs[0].font.color.rgb = RGBColor(0x80, 0x80, 0x80)
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return True


# ── matplotlib PNG preview (fallback) ────────────────────────────────────────

def _render_topology_png(all_data):
    """Render the environment topology as a PNG byte-string using matplotlib.

    Cohesity brand colors: teal clusters, blue replication, amber archival,
    dark-green FortKnox. Returns PNG bytes, or None if matplotlib is unavailable.
    """
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        from matplotlib.patches import FancyBboxPatch
        import io as _io
    except ImportError:
        return None

    topo       = _topology_data(all_data)
    clusters   = topo["clusters"]
    repl_nodes = topo["repl_nodes"]
    arch_nodes = topo["arch_nodes"]
    fk_nodes   = topo["fk_nodes"]
    edges      = topo["edges"]

    # ── Figure setup ─────────────────────────────────────────────────────────
    FIG_W, FIG_H = 16.0, 9.0
    HEADER_H = 0.0
    FOOTER_H = 0.60

    CT = FIG_H - 0.30               # content top (no header)
    CB = FOOTER_H + 0.15            # content bottom
    CH = CT - CB                    # content height

    C_W, C_H = 2.80, 1.05          # cluster box
    T_W, T_H = 2.30, 0.78          # target box

    fig, ax = plt.subplots(figsize=(FIG_W, FIG_H), dpi=130)
    fig.patch.set_facecolor("#F8F8F8")
    ax.set_facecolor("#F8F8F8")
    ax.set_xlim(0, FIG_W)
    ax.set_ylim(0, FIG_H)
    ax.axis("off")

    # ── Determine active right columns ────────────────────────────────────────
    right_cols = []
    if repl_nodes: right_cols.append(("repl", repl_nodes, _DG_REPL, "#004A8A", "Replication Targets"))
    if arch_nodes: right_cols.append(("arch", arch_nodes, _DG_ARCH, "#B05A00", "Archival Vaults"))
    if fk_nodes:   right_cols.append(("fk",   fk_nodes,  _DG_FK,   "#0E3B25", "FortKnox / RPaaS"))

    # Compute column X centers
    X_CLUSTER = 1.85
    RIGHT_START = X_CLUSTER + C_W / 2 + 1.4
    RIGHT_END   = FIG_W - 0.35
    n_right     = len(right_cols)
    if n_right > 0:
        step = (RIGHT_END - RIGHT_START - T_W) / max(n_right - 1, 1) if n_right > 1 else 0
        col_xs = [RIGHT_START + T_W / 2 + i * step for i in range(n_right)]
        # if only 1 column centre it in available space
        if n_right == 1:
            col_xs = [(RIGHT_START + RIGHT_END) / 2]
    else:
        col_xs = []

    # Build (type → x_center) lookup
    col_x_by_type = {}
    col_info_by_type = {}
    for idx, (ctype, nodes, fc, ec, header) in enumerate(right_cols):
        col_x_by_type[ctype] = col_xs[idx]
        col_info_by_type[ctype] = (nodes, fc, ec, header)

    def _ys(n, box_h):
        """Y-centers for n items, distributed evenly in the content area."""
        if n == 0: return []
        if n == 1: return [CB + CH / 2]
        step = (CH - box_h) / (n - 1)
        return [CT - box_h / 2 - i * step for i in range(n)]

    cluster_ys = _ys(len(clusters), C_H)
    type_ys    = {}
    for ctype, nodes, *_ in right_cols:
        type_ys[ctype] = _ys(len(nodes), T_H)

    # Build id → (cx, cy)
    pos = {}
    for i, cl in enumerate(clusters):
        pos[cl["id"]] = (X_CLUSTER, cluster_ys[i])
    for ctype, nodes, *_ in right_cols:
        cx = col_x_by_type[ctype]
        for i, n in enumerate(nodes):
            pos[n["id"]] = (cx, type_ys[ctype][i])

    # ── Title ─────────────────────────────────────────────────────────────────
    ax.text(FIG_W / 2, FIG_H - 0.15, "Environment Topology",
            fontsize=14, fontweight="bold", color="#333333",
            va="center", ha="center", zorder=6)

    # ── Footer / legend ───────────────────────────────────────────────────────
    ax.add_patch(plt.Rectangle((0, 0), FIG_W, FOOTER_H, color="#EDEDED", zorder=1))
    leg_items = [
        (_DG_TEAL, "Source Cluster"),
    ] + [(c[2], c[4]) for c in right_cols]
    lx = 0.5
    for col, lbl in leg_items:
        ax.add_patch(plt.Rectangle((lx, 0.18), 0.35, 0.22,
                                    color=col, zorder=2, linewidth=0))
        ax.text(lx + 0.45, 0.29, lbl, va="center", fontsize=8.0, color="#444444", zorder=2)
        lx += 2.8 if len(leg_items) <= 4 else 2.2

    # ── Column headers ────────────────────────────────────────────────────────
    def _col_hdr(x, label, color, bw):
        y = CT + 0.12
        ax.plot([x - bw / 2, x + bw / 2], [y - 0.05, y - 0.05],
                color=color, lw=1.0, alpha=0.35, zorder=1)
        ax.text(x, y + 0.02, label, ha="center", va="bottom",
                fontsize=8.5, fontweight="bold", color=color, zorder=2)

    _col_hdr(X_CLUSTER, "Source Clusters", _DG_TEAL, C_W)
    for ctype, nodes, fc, ec, header in right_cols:
        _col_hdr(col_x_by_type[ctype], header, fc, T_W)

    # ── Vertical separator line ───────────────────────────────────────────────
    sep_x = X_CLUSTER + C_W / 2 + 0.65
    ax.plot([sep_x, sep_x], [CB, CT], color="#CCCCCC", lw=1.0,
            linestyle="--", zorder=1)

    # ── Edges (drawn before boxes so they appear underneath) ─────────────────
    drawn_pairs = set()
    for e in edges:
        key = (e["src"], e["tgt"])
        if key in drawn_pairs: continue
        drawn_pairs.add(key)
        sp, tp = pos.get(e["src"]), pos.get(e["tgt"])
        if sp is None or tp is None: continue
        if   e["type"] == "repl": col = _DG_REPL; bw = T_W
        elif e["type"] == "arch": col = _DG_ARCH; bw = T_W
        else:                     col = _DG_FK;   bw = T_W;
        sx, sy = sp[0] + C_W / 2, sp[1]
        tx, ty = tp[0] - bw / 2,  tp[1]
        ax.annotate("", xy=(tx, ty), xytext=(sx, sy),
                    arrowprops=dict(arrowstyle="-|>", color=col, lw=1.5,
                                   mutation_scale=14,
                                   connectionstyle="arc3,rad=0.0"),
                    zorder=2)

    # ── Box drawing helpers ───────────────────────────────────────────────────
    def _box(cx, cy, bw, bh, fc, ec, lines, bold_first=True):
        x0, y0 = cx - bw / 2, cy - bh / 2
        # Drop shadow
        ax.add_patch(FancyBboxPatch((x0 + 0.035, y0 - 0.035), bw, bh,
                                    boxstyle="round,pad=0.06",
                                    facecolor="#00000018", edgecolor="none", zorder=2))
        # Main box
        ax.add_patch(FancyBboxPatch((x0, y0), bw, bh,
                                    boxstyle="round,pad=0.06",
                                    facecolor=fc, edgecolor=ec,
                                    linewidth=1.4, zorder=3))
        # Text lines (top → bottom)
        n = len(lines)
        line_h = bh / (n + 0.3)
        for j, line in enumerate(lines):
            ty = cy + (n - 1) / 2 * line_h - j * line_h
            fw = "bold" if (bold_first and j == 0) else "normal"
            fs = 10.5 if j == 0 else 9.0
            ax.text(cx, ty, line, ha="center", va="center",
                    fontsize=fs, fontweight=fw, color="white", zorder=4,
                    clip_on=True)

    # ── Source cluster boxes ──────────────────────────────────────────────────
    for i, cl in enumerate(clusters):
        cx, cy = pos[cl["id"]]
        wls = ", ".join(cl["workloads"]) if cl["workloads"] else ""
        det = f"{cl['groups']} groups"
        if cl["sources"]: det += f"  ·  {cl['sources']} sources"
        if cl["cap_tb"] > 0.01: det += f"  ·  {cl['cap_tb']:.1f} TB"
        lines = [cl["name"]] + ([wls] if wls else []) + [det]
        _box(cx, cy, C_W, C_H, _DG_TEAL, _DG_TEAL_DK, lines)

    # ── Replication target boxes ──────────────────────────────────────────────
    if "repl" in col_x_by_type:
        for i, n in enumerate(repl_nodes):
            cx, cy = pos[n["id"]]
            _box(cx, cy, T_W, T_H, _DG_REPL, "#004A8A",
                 [n["name"], "Cluster"])

    # ── Archival vault boxes ──────────────────────────────────────────────────
    if "arch" in col_x_by_type:
        for i, n in enumerate(arch_nodes):
            cx, cy = pos[n["id"]]
            _box(cx, cy, T_W, T_H, _DG_ARCH, "#B05A00",
                 [n["name"], "Archival Vault"])

    # ── FortKnox boxes ────────────────────────────────────────────────────────
    if "fk" in col_x_by_type:
        for i, n in enumerate(fk_nodes):
            cx, cy = pos[n["id"]]
            _box(cx, cy, T_W, T_H, _DG_FK, "#0E3B25",
                 [n["name"], "FortKnox / RPaaS (indelible)"])

    plt.subplots_adjust(left=0, right=1, top=1, bottom=0)
    buf = _io.BytesIO()
    plt.savefig(buf, format="png", dpi=130, bbox_inches="tight",
                facecolor=fig.get_facecolor(), pad_inches=0.0)
    plt.close(fig)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL ORCHESTRATOR
# ─────────────────────────────────────────────────────────────────────────────

def write_excel(all_data, args):
    if not EXCEL_OK:
        print("  WARN: openpyxl not installed — skipping Excel output.")
        print("        Install: pip install openpyxl")
        return None

    wb = Workbook()
    wb.remove(wb.active)

    print("  Writing Excel sheets...")
    _sheet_summary(wb, all_data)           # 1
    _sheet_infrastructure(wb, all_data)    # 2
    _sheet_hardware(wb, all_data)          # 3
    _sheet_disk_health(wb, all_data)       # 4
    _sheet_protection(wb, all_data)        # 5
    _sheet_storage(wb, all_data)           # 6
    _sheet_policies(wb, all_data)          # 7
    _sheet_policy_groups(wb, all_data)     # 8
    _sheet_alerts(wb, all_data)            # 9
    _sheet_security(wb, all_data)          # 10
    _sheet_agent_health(wb, all_data)      # 11
    _sheet_source_coverage(wb, all_data)   # 12
    _sheet_replication(wb, all_data)       # 13
    _sheet_fortknox_detail(wb, all_data)   # 14 NEW
    _sheet_views(wb, all_data)             # 15
    _sheet_coverage(wb, all_data)          # 16
    _sheet_user_security(wb, all_data)     # 17
    _sheet_trends(wb, all_data)            # 18
    _sheet_recommendations(wb, all_data)   # 19

    out = f"{args.output}.xlsx"
    wb.save(out)
    print(f"  Excel  → {out}")
    return out


# ─────────────────────────────────────────────────────────────────────────────
# WORD DOCUMENT
# ─────────────────────────────────────────────────────────────────────────────

def write_word(all_data, args):
    if not DOCX_OK:
        print("  WARN: python-docx not installed — skipping Word output.")
        print("        Install: pip install python-docx")
        return None

    doc      = Document()
    customer = args.customer or "Customer"
    today    = datetime.date.today().strftime("%B %d, %Y")

    # ── base font ──────────────────────────────────────────────────────────
    for style_name in ("Normal", "Heading 1", "Heading 2"):
        try:
            doc.styles[style_name].font.name = "Calibri"
        except Exception:
            pass

    def _h1(text):
        p = doc.add_heading(text, level=1)
        for run in p.runs:
            run.font.color.rgb = RGBColor(0x70, 0xAD, 0x47)
        return p

    def _h2(text):
        p = doc.add_heading(text, level=2)
        for run in p.runs:
            run.font.color.rgb = RGBColor(0x70, 0xAD, 0x47)
        return p

    def _cell_shade(cell, hex_color):
        """Apply background color to a docx table cell."""
        tc   = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd  = OxmlElement("w:shd")
        shd.set(qn("w:val"),   "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"),  hex_color)
        tcPr.append(shd)

    def _tbl_header(table, cols, bg=COH_GREEN):
        row = table.rows[0].cells
        for i, col in enumerate(cols):
            row[i].text = col
            _cell_shade(row[i], bg)
            for para in row[i].paragraphs:
                for run in para.runs:
                    run.font.bold  = True
                    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                    run.font.size  = Pt(9)

    # ── Cover page ─────────────────────────────────────────────────────────
    for _ in range(6):
        doc.add_paragraph("")

    cp = doc.add_paragraph()
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = cp.add_run("Cohesity Health Check Report")
    run.font.size  = Pt(28)
    run.font.bold  = True
    run.font.color.rgb = RGBColor(0x70, 0xAD, 0x47)

    doc.add_paragraph("")
    cp2 = doc.add_paragraph()
    cp2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = cp2.add_run(f"Prepared for: {customer}\n{today}")
    r2.font.size = Pt(16)

    doc.add_paragraph("")
    cp3 = doc.add_paragraph()
    cp3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r3 = cp3.add_run("CONFIDENTIAL — For authorized recipients only")
    r3.font.size   = Pt(10)
    r3.font.italic = True
    r3.font.color.rgb = RGBColor(0x59, 0x59, 0x59)

    doc.add_page_break()

    # ── 1. Executive Summary ───────────────────────────────────────────────
    _h1("1. Executive Summary")

    n_clusters = len(all_data)
    avg_score  = round(sum(cd["scores"]["overall"] for cd in all_data) / n_clusters, 1) \
                 if n_clusters else 0
    all_recs   = [(cd["name"], r) for cd in all_data for r in cd["recommendations"]]
    n_crit     = sum(1 for _, r in all_recs if r["priority"] == "CRITICAL")
    n_high     = sum(1 for _, r in all_recs if r["priority"] == "HIGH")

    summary = (
        f"This report covers a Cohesity environment comprising {n_clusters} "
        f"cluster(s) managed through Helios. The overall environment health score "
        f"is {avg_score}/100 ({grade(avg_score)}). "
    )
    if n_crit:
        summary += (f"There are {n_crit} critical finding(s) requiring immediate "
                    f"attention. ")
    if n_high:
        summary += f"{n_high} high-priority item(s) require action within 30 days. "
    if not n_crit and not n_high:
        summary += "No critical or high-priority gaps were identified. "
    summary += ("Detailed findings and prioritized recommendations are included "
                "in the sections below, with full data in the accompanying Excel workbook.")
    doc.add_paragraph(summary)

    t1 = doc.add_table(rows=1, cols=5)
    t1.style = "Table Grid"
    _tbl_header(t1, ["Cluster", "Health Score", "Grade", "Open Criticals", "Top Finding"])
    for cd in all_data:
        row = t1.add_row().cells
        crits = sum(1 for a in cd["alerts"] if a.get("severity") == "kCritical")
        top   = cd["recommendations"][0]["finding"] if cd["recommendations"] else "None"
        row[0].text = cd["name"]
        row[1].text = str(cd["scores"]["overall"])
        row[2].text = cd["scores"]["grade"]
        row[3].text = str(crits)
        row[4].text = top
        for para in row[1].paragraphs:
            for run in para.runs:
                run.font.bold = True
    doc.add_page_break()

    # ── 2. Environment Overview ────────────────────────────────────────────
    _h1("2. Environment Overview")
    doc.add_paragraph(
        f"The environment spans {n_clusters} cluster(s). "
        "The table below summarizes key infrastructure attributes for each cluster."
    )
    t2 = doc.add_table(rows=1, cols=7)
    t2.style = "Table Grid"
    _tbl_header(t2, ["Cluster", "Version", "Nodes", "Healthy",
                     "Cluster Enc", "SD Enc", "Status"])
    for cd in all_data:
        info  = cd["info"]
        nodes = cd["nodes"]
        healthy = sum(1 for n in nodes
                      if ((n.get("nodeStatus") or {}).get("overallStatus") == "kNormal"
                          or n.get("removalState") in (None, "kDontRemove")))
        cluster_enc_w, sd_enc_w = _sd_enc_status(cd)
        sd_enc_str_w = ("Yes" if sd_enc_w is True else
                        "No"  if sd_enc_w is False else "N/A")
        status = "OK" if healthy == len(nodes) else f"WARN: {len(nodes)-healthy} node(s)"
        row = t2.add_row().cells
        for i, val in enumerate([cd["name"], info.get("clusterSoftwareVersion", ""),
                                  len(nodes), healthy,
                                  "Yes" if cluster_enc_w else "No",
                                  sd_enc_str_w, status]):
            row[i].text = str(val)

    _h2("Software & Hardware Lifecycle")
    doc.add_paragraph(
        "The table below summarizes the software lifecycle status for each cluster. "
        "Cohesity software versions follow a feature-release (9-month support) and "
        "LTS (12+ month support) model. All current in-support versions (6.8.2 LTS, "
        "7.1.2_u2 LTS, and 7.2.x feature releases) reach end of support on "
        "June 8, 2026. Version 7.3 and later follow a new lifecycle policy."
    )
    # SW lifecycle table
    tbl = doc.add_table(rows=1, cols=4)
    tbl.style = "Table Grid"
    hdr = tbl.rows[0].cells
    for i, h in enumerate(["Cluster", "Version", "Lifecycle Status", "EOS Date"]):
        hdr[i].text = h
        for run in hdr[i].paragraphs[0].runs:
            run.font.bold = True
    for cd in all_data:
        sw_ver = (cd["info"].get("clusterSoftwareVersion") or "").strip()
        _, sw_eos_date, sw_label = _sw_eos(sw_ver)
        row = tbl.add_row().cells
        row[0].text = cd["name"]
        row[1].text = sw_ver
        row[2].text = sw_label
        row[3].text = sw_eos_date.isoformat() if sw_eos_date else "—"
    doc.add_paragraph("")

    doc.add_paragraph(
        "Hardware end-of-life (EOL) dates are provided by Cohesity's Platform EOL Terms. "
        "EOL hardware is no longer eligible for manufacturer support or standard replacement. "
        "Note: Cohesity C2000, C2500, and C3500 platform nodes reached EOL and are no longer "
        "eligible for support under the standard agreement. C4000/CN series hardware is in "
        "service until February 2027. See the Node Hardware sheet in the Excel workbook for "
        "per-node EOL status."
    )

    _h2("Environment Topology")
    doc.add_paragraph(
        "The table below shows each source cluster with its configured replication targets "
        "and archival/vault destinations, including FortKnox vaults."
    )
    t_topo = doc.add_table(rows=1, cols=4)
    t_topo.style = "Table Grid"
    _tbl_header(t_topo, ["Source Cluster", "Replication Targets",
                          "Archival Targets", "FortKnox Vaults"])
    for cd in all_data:
        policies_t = cd["policies"]
        vaults_t   = cd["vaults"]

        # Collect unique replication target names across all policies
        repl_names = sorted({
            (t.get("targetName")
             or (t.get("remoteTargetConfig") or {}).get("clusterName")
             or (t.get("remoteTargetConfig") or {}).get("name")
             or "Unknown")
            for p in policies_t
            for t in ((p.get("remoteTargetPolicy") or {}).get("replicationTargets") or [])
        })
        # Collect unique archival target names (non-FortKnox)
        arch_names = sorted({
            (t.get("targetName")
             or (t.get("archivalTargetConfig") or {}).get("name")
             or (t.get("archivalTargetConfig") or {}).get("vaultName")
             or "Unknown")
            for p in policies_t
            for t in ((p.get("remoteTargetPolicy") or {}).get("archivalTargets") or [])
            if (t.get("archivalTargetConfig") or {}).get("targetType", "") != "kFortKnox"
        })
        # FortKnox: use _fk_status() for active/idle/none, then collect
        # vault names using the same comprehensive detection logic.
        fk_st_t   = _fk_status(cd)
        _fkt_types = {"krpaas", "kfortknox", "kfort_knox", "krpaasarchival"}
        def _is_fk_t(tt, tname=""):
            tl = (tt or "").lower(); nl = (tname or "").lower()
            return any(x in tl for x in _fkt_types) or "fortknox" in tl or "fortknox" in nl
        fk_name_set = set()
        for v in vaults_t:
            vt = v.get("vaultType") or ""; vn = v.get("name") or ""
            if _is_fk_t(vt, vn):
                fk_name_set.add(vn or "FortKnox")
        for p in policies_t:
            for t in ((p.get("remoteTargetPolicy") or {}).get("archivalTargets") or []):
                cfg   = t.get("archivalTargetConfig") or {}
                tt    = cfg.get("targetType") or t.get("targetType") or ""
                tname = t.get("targetName") or cfg.get("name") or cfg.get("vaultName") or ""
                if tname and _is_fk_t(tt, tname):
                    fk_name_set.add(tname)
        if fk_st_t == "none":
            fk_cell_text = "No"
        elif fk_st_t == "active":
            fk_cell_text = "\n".join(sorted(fk_name_set)) if fk_name_set else "Active"
        else:  # idle — configured but no successful FK archival in lookback window
            fk_cell_text = (("\n".join(sorted(fk_name_set)) + "\n(Configured — not sending data)")
                            if fk_name_set else "Configured — not sending data")

        row_t = t_topo.add_row().cells
        row_t[0].text = cd["name"]
        row_t[1].text = "\n".join(repl_names) if repl_names else "None"
        row_t[2].text = "\n".join(arch_names) if arch_names else "None"
        row_t[3].text = fk_cell_text
        # Colour-code the FK cell: green=active, yellow=idle, no fill=none
        if fk_st_t == "active":
            _cell_shade(row_t[3], "E2EFDA")   # light green
        elif fk_st_t == "idle":
            _cell_shade(row_t[3], "FFF2CC")   # amber/yellow

    # ── Topology diagram ─────────────────────────────────────────────────────
    # Primary: native Word shapes (editable). Fallback: matplotlib PNG.
    if not _render_topology_word_shapes(doc, all_data):
        png_bytes = _render_topology_png(all_data)
        if png_bytes:
            from docx.shared import Inches
            import io as _io
            doc.add_paragraph("")
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_img = p_img.add_run()
            run_img.add_picture(_io.BytesIO(png_bytes), width=Inches(6.5))
            cap = doc.add_paragraph(
                "Figure: Environment Topology (PNG fallback)."
            )
            cap.runs[0].font.size = Pt(8)
            cap.runs[0].font.color.rgb = RGBColor(0x80, 0x80, 0x80)
            cap.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_page_break()

    # ── 3. Protection & Recovery ───────────────────────────────────────────
    _h1("3. Protection & Recovery Health")
    total_g   = sum(len(cd["groups"]) for cd in all_data)
    paused_g  = sum(sum(1 for g in cd["groups"] if g.get("isPaused"))
                    for cd in all_data)
    failed_g  = sum(sum(1 for g in cd["groups"]
                        if (g.get("lastRun") or {}).get("localBackupInfo", {})
                           .get("status", "") in ("kFailure", "Failed"))
                    for cd in all_data)
    doc.add_paragraph(
        f"The environment contains {total_g} active protection group(s) across all clusters. "
        f"{paused_g} group(s) are currently paused and {failed_g} have a failed last run. "
        "The table below summarizes protection health per cluster. "
        "Refer to the Protection Health and Coverage Gaps tabs in the Excel workbook "
        "for per-group detail."
    )
    t3 = doc.add_table(rows=1, cols=6)
    t3.style = "Table Grid"
    _tbl_header(t3, ["Cluster", "Groups", "Paused", "Failed Last Run",
                      "30d Success %", "SLA Score"])
    for cd in all_data:
        succ_pct, sla_pct = _success_stats(cd)
        paused = sum(1 for g in cd["groups"] if g.get("isPaused"))
        failed = sum(1 for g in cd["groups"]
                     if (g.get("lastRun") or {}).get("localBackupInfo", {})
                        .get("status", "") in ("kFailure", "Failed"))
        row = t3.add_row().cells
        for i, val in enumerate([cd["name"], len(cd["groups"]),
                                  paused, failed, f"{succ_pct}%",
                                  f"{cd['scores']['sla']}/100"]):
            row[i].text = str(val)
    doc.add_page_break()

    # ── 4. Storage & Capacity ──────────────────────────────────────────────
    _h1("4. Storage & Capacity")
    doc.add_paragraph(
        "Cohesity applies inline deduplication and compression to maximize storage "
        "efficiency. The data reduction ratio shows how much raw data has been saved "
        "compared to the logical data ingested."
    )
    t4 = doc.add_table(rows=1, cols=7)
    t4.style = "Table Grid"
    _tbl_header(t4, ["Cluster", "Usable (TB)", "Used (TB)", "Free (TB)",
                      "Used %", "Data Reduction", "Status"])
    for cd in all_data:
        info   = cd["info"]
        usable, used, cap_pct = _cap_stats(cd)
        free   = usable - used
        dr     = round((info.get("stats") or {}).get("dataReductionRatio") or 0, 2)
        pct_str = f"{cap_pct}%" if isinstance(cap_pct, float) else "N/A"
        status = ("N/A" if cap_pct == "N/A"
                  else "CRITICAL" if cap_pct >= CAPACITY_CRIT_PCT
                  else "WARNING" if cap_pct >= CAPACITY_WARN_PCT else "OK")
        row = t4.add_row().cells
        for i, val in enumerate([cd["name"],
                                  round(bytes_to_tb(usable), 2) or "N/A",
                                  round(bytes_to_tb(used), 2) or "N/A",
                                  round(bytes_to_tb(free), 2) or "N/A",
                                  pct_str, f"{dr}x", status]):
            row[i].text = str(val)
    doc.add_page_break()

    # ── 5. Security ────────────────────────────────────────────────────────
    _h1("5. Security Posture")
    doc.add_paragraph(
        "The following checklist evaluates each cluster against Cohesity security "
        "best practices. A fully hardened deployment enables cluster-wide and "
        "storage-domain encryption, at least one vault for archival, replication "
        "to a secondary cluster, and FortKnox or DataLock archival for indelible "
        "backups. Note: Cohesity's distributed filesystem is always immutable by "
        "design; DataLock (WORM) makes individual snapshots indelible — permanently "
        "retained even against administrator deletion."
    )
    t5 = doc.add_table(rows=1, cols=10)
    t5.style = "Table Grid"
    _tbl_header(t5, ["Cluster", "Cluster Enc", "SD Enc", "Vault",
                      "Indelible (FK)", "DataLock (WORM)",
                      "Replication", "Quorum", "Admins w/o MFA", "Score /100"])
    _admin_roles_w = {"COHESITY_ADMIN", "Admin", "kAdmin", "kSuperAdmin"}
    for cd in all_data:
        vaults   = cd["vaults"]
        policies = cd["policies"]
        cluster_enc_s, sd_enc_s = _sd_enc_status(cd)
        sd_enc_str_s = ("Yes" if sd_enc_s is True else
                        "No"  if sd_enc_s is False else "N/A")
        has_vault   = bool(vaults)
        fk_st_s     = _fk_status(cd)
        fk_label_s  = {"active": "Active", "idle": "Idle", "none": "No"}[fk_st_s]
        repl = any((p.get("remoteTargetPolicy") or {}).get("replicationTargets")
                   for p in policies)
        dl_status_s, dl_label_s, _, _ = _cluster_datalock(policies, cd["groups"])
        if cd.get("mode") == "direct":
            quorum_str_s = "N/A"
        else:
            quorum_str_s = "Yes" if cd.get("quorum_enabled") else "No"
        admins_no_mfa_w = sum(
            1 for u in (cd.get("users") or [])
            if isinstance(u, dict)
            and not (u.get("mfaEnabled") or u.get("isMfaEnabled"))
            and bool(set(u.get("roles") or []) & _admin_roles_w)
        )
        row = t5.add_row().cells
        yn  = lambda v: "Yes" if v else "No"
        for i, val in enumerate([cd["name"],
                                  yn(cluster_enc_s), sd_enc_str_s,
                                  yn(has_vault), fk_label_s, dl_label_s,
                                  yn(repl), quorum_str_s,
                                  admins_no_mfa_w,
                                  f"{cd['scores']['security']}/100"]):
            row[i].text = str(val)

        # ── Color-code every security column ──────────────────────────────
        def _sec_red(cell):
            _cell_shade(cell, "FF4C4C")
            for _r in cell.paragraphs[0].runs:
                _r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                _r.font.bold = True

        def _sec_green(cell):
            _cell_shade(cell, "E2EFDA")

        def _sec_amber(cell):
            _cell_shade(cell, "FFF2CC")
            for _r in cell.paragraphs[0].runs:
                _r.font.bold = True

        # Col 1 — Cluster Encryption
        (_sec_green if cluster_enc_s else _sec_red)(row[1])
        # Col 2 — SD Encryption
        if sd_enc_str_s == "Yes":   _sec_green(row[2])
        elif sd_enc_str_s == "No":  _sec_red(row[2])
        # Col 3 — Vault
        (_sec_green if has_vault else _sec_red)(row[3])
        # Col 4 — FortKnox (Indelible)
        if fk_st_s == "active":   _sec_green(row[4])
        elif fk_st_s == "idle":   _sec_amber(row[4])
        else:                     _sec_red(row[4])
        # Col 5 — DataLock (WORM)
        if dl_status_s in ("full_compliance", "full_any"):
            _sec_green(row[5])
        elif dl_status_s == "partial":
            _sec_amber(row[5])
        else:
            _sec_red(row[5])
        # Col 6 — Replication
        (_sec_green if repl else _sec_red)(row[6])
        # Col 7 — Quorum
        if quorum_str_s == "Yes":   _sec_green(row[7])
        elif quorum_str_s == "No":  _sec_red(row[7])
        # Col 8 — Admins w/o MFA
        if admins_no_mfa_w > 0:     _sec_red(row[8])
        else:                       _sec_green(row[8])
        # Col 9 — Score
        sec_val = cd["scores"]["security"]
        if sec_val >= 75:      _sec_green(row[9])
        elif sec_val >= 50:    _sec_amber(row[9])
        else:                  _sec_red(row[9])
    doc.add_page_break()

    # ── 6. Agent Health & Source Coverage ────────────────────────────────────
    _h1("6. Agent Health & Source Coverage")

    total_agents    = sum(len(cd.get("agents") or []) for cd in all_data)
    def _ag_st(a):
        _hs = a.get("healthStatus") if isinstance(a, dict) else None
        return ((isinstance(_hs, dict) and _hs.get("status"))
                or (isinstance(_hs, str) and _hs)
                or (isinstance(a, dict) and (a.get("agentStatus") or a.get("status")))
                or "")
    unhealthy_ag    = sum(
        sum(1 for a in (cd.get("agents") or [])
            if _ag_st(a).lower()
            in ("kunhealthy", "unhealthy", "unreachable", "kfailed", "failed"))
        for cd in all_data)
    upgradable_ag   = sum(
        sum(1 for a in (cd.get("agents") or [])
            if "upgradable" in (a.get("upgradability") or "").lower())
        for cd in all_data)

    doc.add_paragraph(
        f"The environment has {total_agents} registered host agent(s) across all clusters. "
        f"{unhealthy_ag} agent(s) are unhealthy or unreachable. "
        f"{upgradable_ag} agent(s) are eligible for upgrade. "
        "Unhealthy agents should be remediated promptly as they will prevent backups "
        "from running. Agents running older versions MAY cause issues with newer cluster "
        "features; Cohesity supports a limited range of back-versions — review the agent "
        "compatibility matrix at docs.cohesity.com before upgrading. "
        "Refer to the Agent Health sheet in the Excel workbook for per-host detail."
    )

    total_sources = sum(len(cd.get("sources") or []) for cd in all_data)
    low_cov_src   = sum(
        sum(1 for s in (cd.get("sources") or [])
            if (lambda p, u: p + u > 0 and (u / (p + u)) >= 0.25)(
                (s.get("protectedObjectsCount") or s.get("numProtectedObjects")
                 or (s.get("stats") or {}).get("protectedObjectsCount") or 0),
                (s.get("unprotectedObjectsCount") or s.get("numUnprotectedObjects")
                 or (s.get("stats") or {}).get("unprotectedObjectsCount") or 0)))
        for cd in all_data)
    doc.add_paragraph(
        f"There are {total_sources} registered source(s). "
        f"{low_cov_src} source(s) have ≥25% unprotected objects — "
        "these objects will be unrecoverable in the event of data loss. "
        "Review the Source Coverage sheet in the Excel workbook for per-source breakdown."
    )
    doc.add_page_break()

    # ── 7. User Security ───────────────────────────────────────────────────────
    _h1("7. User Security")

    total_users  = sum(len(cd.get("users") or []) for cd in all_data)
    _ADMIN_SET   = {"cohesity_admin", "admin", "kadmin", "ksuperadmin", "super_admin"}
    admins_no_mfa = sum(
        sum(1 for u in (cd.get("users") or [])
            if not (u.get("mfaEnabled") or u.get("isMfaEnabled"))
            and bool({r.lower() for r in (u.get("roles") or [])} & _ADMIN_SET))
        for cd in all_data)
    locked_users = sum(
        sum(1 for u in (cd.get("users") or [])
            if u.get("locked") or u.get("isLocked"))
        for cd in all_data)
    has_sso_any  = any(cd.get("idps") for cd in all_data)

    doc.add_paragraph(
        f"The environment has {total_users} user account(s) across all clusters. "
        f"{admins_no_mfa} admin account(s) do not have MFA enabled — these are at "
        "elevated risk from credential theft. "
        f"{locked_users} account(s) are currently locked. "
        f"SSO / IDP: {'Configured' if has_sso_any else 'Not configured — local credentials only'}. "
        "Refer to the User Security sheet in the Excel workbook for per-user detail."
    )

    t_usr = doc.add_table(rows=1, cols=5)
    t_usr.style = "Table Grid"
    _tbl_header(t_usr, ["Cluster", "Total Users", "Admins w/o MFA",
                         "Locked Accounts", "SSO Configured"])
    for cd in all_data:
        _A = {"cohesity_admin", "admin", "kadmin", "ksuperadmin", "super_admin"}
        ano_mfa = sum(1 for u in (cd.get("users") or [])
                      if not (u.get("mfaEnabled") or u.get("isMfaEnabled"))
                      and bool({r.lower() for r in (u.get("roles") or [])} & _A))
        locked  = sum(1 for u in (cd.get("users") or [])
                      if u.get("locked") or u.get("isLocked"))
        row = t_usr.add_row().cells
        for i, val in enumerate([cd["name"], len(cd.get("users") or []),
                                  ano_mfa, locked,
                                  "Yes" if cd.get("idps") else "No"]):
            row[i].text = str(val)
    doc.add_page_break()

    # ── 8. Recommendations ─────────────────────────────────────────────────
    _h1("8. Prioritized Recommendations")
    doc.add_paragraph(
        "The recommendations below are ranked by priority. Critical items should "
        "be addressed immediately; High-priority items within 30 days; Medium within "
        "90 days. Refer to the Recommendations tab in the Excel workbook for the "
        "full prioritized list."
    )
    sorted_recs = sorted(all_recs, key=lambda x: _PRIORITY_ORDER.get(x[1]["priority"], 99))
    if sorted_recs:
        t6 = doc.add_table(rows=1, cols=5)
        t6.style = "Table Grid"
        _tbl_header(t6, ["Priority", "Cluster", "Category", "Finding", "Action"])
        for cname, rec in sorted_recs:
            row = t6.add_row().cells
            for i, val in enumerate([rec["priority"], cname, rec["category"],
                                      rec["finding"], rec["action"]]):
                row[i].text = str(val)
                if i == 0:
                    for para in row[i].paragraphs:
                        for run in para.runs:
                            run.font.bold = True
    else:
        doc.add_paragraph("No recommendations generated — environment appears healthy.")

    doc.add_page_break()

    # ── Appendix ───────────────────────────────────────────────────────────
    _h1("Appendix — Methodology & Thresholds")
    _h2("Scoring Model")
    doc.add_paragraph(
        "Health scores (0-100) are computed from five weighted dimensions: "
        "Protection success rate (25%), SLA compliance (20%), "
        "Infrastructure health (15%), Storage capacity (15%), "
        "Security posture (15%), Alert health (10%). "
        "Within the Security posture dimension, the 100-point security score is "
        "calculated across six areas: cluster encryption (15 pts), storage-domain "
        "encryption (15 pts), vault configured (15 pts), replication (15 pts), "
        "FortKnox/indelible copy (20 pts), and DataLock/WORM (20 pts). "
        "A penalty of −10 pts is applied when any administrator account lacks MFA. "
        "DataLock coverage is measured at the protection-group level — only groups "
        "whose governing policy has DataLock (WORM) explicitly configured count; "
        "FortKnox archival (indelible by design) scores separately in the FK dimension."
    )
    _h2("Thresholds")
    notes = [
        f"Capacity: Warning ≥{CAPACITY_WARN_PCT}%, Critical ≥{CAPACITY_CRIT_PCT}%",
        f"Backup success rate: Warning <{SUCCESS_WARN_PCT}%, Critical <{SUCCESS_CRIT_PCT}% "
        f"— runs with status 'Warning' (kWarning) are counted as successful in the "
        f"Success % calculation. They complete with data but may have minor issues "
        f"(e.g. some objects skipped) and are shown separately in the Warning column "
        f"for visibility.",
        f"RPO: Warning >{RPO_WARN_HOURS}h, Critical >{RPO_CRIT_HOURS}h since last success",
        ("Security: cluster enc (15) + SD enc (15) + vault (15) + replication (15) "
         "+ FortKnox/indelible (20) + DataLock (20) = 100 pts. "
         "Partial DataLock coverage earns proportional credit."),
        f"Alert penalty: −10 pts per open critical, −2 pts per open warning",
        f"View quota: Warning ≥{VIEW_QUOTA_WARN_PCT}%, Critical ≥{VIEW_QUOTA_CRIT_PCT}%",
        f"Lookback period used for this report: {args.days} days",
        f"Report generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} (local time)",
    ]
    for note in notes:
        p = doc.add_paragraph(note, style="List Bullet")
        for run in p.runs:
            run.font.size = Pt(9)

    out = f"{args.output}.docx"
    doc.save(out)
    print(f"  Word   → {out}")
    return out


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description=f"Cohesity Health Check Report v{__version__}",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    # ── Helios API key auth ───────────────────────────────────────────────────
    parser.add_argument("--apikey",           help="Helios API key (stored in keychain if omitted)")
    # ── Direct cluster username/password auth (requires --cluster-host) ───────
    parser.add_argument("--username",         help="Cluster username (use with --cluster-host)")
    parser.add_argument("--password",         help="Password (prompted if omitted; saved to keychain)")
    parser.add_argument("--domain",           default="LOCAL",
                        help="Auth domain for cluster login: LOCAL or AD name (default: LOCAL)")
    parser.add_argument("--mfa-code",         dest="mfa_code",
                        help="TOTP/OTP code for MFA-enabled accounts")
    # ── Direct cluster mode ───────────────────────────────────────────────────
    parser.add_argument("--cluster-host",     dest="cluster_host",
                        help="Bypass Helios — connect directly to this cluster hostname/IP")
    # ── Credential management ─────────────────────────────────────────────────
    parser.add_argument("--clear-credentials", action="store_true",
                        help="Remove stored credentials from keychain and exit")
    # ── Helios cluster filter ─────────────────────────────────────────────────
    parser.add_argument("--cluster",          help="(Helios) Limit to one cluster (partial name match)")
    # ── Report options ────────────────────────────────────────────────────────
    parser.add_argument("--days",             type=int, default=30,
                        help="Lookback window in days (default 30)")
    parser.add_argument("--customer",         default="",
                        help="Customer name for report cover page")
    parser.add_argument("--output",           default="cohesity_health_check",
                        help="Output base path without extension "
                             "(default: cohesity_health_check)")
    parser.add_argument("--quick",            action="store_true",
                        help="Skip per-group run history; use last-run only (faster)")
    parser.add_argument("--excel-only",       action="store_true",
                        help="Generate Excel only, skip Word document")
    parser.add_argument("--word-only",        action="store_true",
                        help="Generate Word only, skip Excel")
    parser.add_argument("--debug",            action="store_true",
                        help="Print HTTP status for each API call")
    args = parser.parse_args()

    # ── Clear stored credentials and exit ─────────────────────────────────────
    if args.clear_credentials:
        clear_stored_credentials(
            cluster=args.cluster_host,
            username=args.username or "admin",
            domain=args.domain,
        )
        sys.exit(0)

    # ── Pre-requisite check ─────────────────────────────────────────────────
    # Build a full status table covering all packages (required + optional).
    # Always printed so the user can see at a glance what is and isn't available.
    def _pkg_ok(name):
        try:
            __import__(name)
            return True
        except ImportError:
            return False

    # requests is guaranteed present here (checked at module import level).
    # openpyxl / python-docx flags come from the module-level try/except.
    _PKG_TABLE = [
        # (name,          installed,    role,       description)
        ("requests",      True,         "required", "HTTP client for Cohesity REST API"),
        ("openpyxl",      EXCEL_OK,     "required", "Excel workbook output (skip with --word-only)"),
        ("python-docx",   DOCX_OK,      "required", "Word document generation (skip with --excel-only)"),
        ("matplotlib",    _pkg_ok("matplotlib"), "optional", "Topology PNG fallback for Word diagram"),
        ("keyring",       _pkg_ok("keyring"),    "optional", "OS keychain credential storage"),
    ]

    # Apply mode exemptions: some required packages are not needed in certain modes
    _exempt = set()
    if args.word_only:   _exempt.add("openpyxl")
    if args.excel_only:  _exempt.add("python-docx")

    _need_required = [(n, d) for n, ok, r, d in _PKG_TABLE
                      if r == "required" and not ok and n not in _exempt]
    _missing_opt   = [(n, d) for n, ok, r, d in _PKG_TABLE if r == "optional" and not ok]
    _any_issue     = bool(_need_required or _missing_opt)

    # Print the full status table whenever anything is missing or optional is absent
    if _any_issue:
        print()
        print("  Prerequisites")
        print("  " + "─" * 68)
        for name, ok, role, desc in _PKG_TABLE:
            if name in _exempt:
                mark   = "–"
                status = "skipped  "
                note   = f"  (not needed with {'--word-only' if name == 'openpyxl' else '--excel-only'})"
            else:
                mark   = "✓" if ok else "✗"
                status = "installed" if ok else "NOT FOUND"
                note   = ""
            role_tag = f"[{role}]"
            print(f"  {mark}  {name:16s}  {status}  {role_tag:11s}  {desc}{note}")
        print("  " + "─" * 68)
    else:
        print(f"[*] Prerequisites: all 5 packages OK")

    if _need_required:
        print()
        print("  ERROR: Missing required package(s). Install and re-run:")
        for n, d in _need_required:
            print(f"    pip install {n}")
        if len(_need_required) > 1:
            print(f"  All at once:  pip install {' '.join(n for n, d in _need_required)}")
        print()
        sys.exit(1)

    if _missing_opt:
        print()
        print("  Optional package(s) not installed — affected features will be skipped.")
        print(f"  To install:  pip install {' '.join(n for n, d in _missing_opt)}")
        print()

    # Build output filename: include customer name (if given) + local timestamp
    if args.output == "cohesity_health_check":
        ts    = datetime.datetime.now().strftime("%Y%m%d_%H%M")
        parts = ["cohesity_health_check", f"v{__version__}"]
        if args.customer:
            safe = "".join(c if c.isalnum() or c in "-_" else "_"
                           for c in args.customer).strip("_")
            parts.append(safe)
        parts.append(ts)
        args.output = "_".join(parts)

    # ── Authenticate and discover clusters ────────────────────────────────────
    session = requests.Session()
    api_key = None

    if args.cluster_host:
        # ── Direct cluster mode ───────────────────────────────────────────────
        if not args.username:
            print("ERROR: --username is required with --cluster-host.")
            sys.exit(1)
        pwd   = get_cluster_password(args.cluster_host, args.username, args.domain,
                                     cli_password=args.password)
        token = get_auth_token(args.cluster_host, args.username, pwd, args.domain,
                               mfa_code=args.mfa_code)
        session.base_url = f"https://{args.cluster_host}"
        # Fetch cluster name from the cluster itself
        try:
            import urllib3 as _u3; _u3.disable_warnings()
            _r = session.get(
                f"https://{args.cluster_host}/irisservices/api/v1/public/cluster",
                headers=make_headers(token), verify=False, timeout=30)
            _r.raise_for_status()
            cname = _r.json().get("name", args.cluster_host)
        except Exception:
            cname = args.cluster_host
        clusters = [{"name": cname, "clusterId": None,
                     "mode": "direct", "host": args.cluster_host, "token": token}]
        auth_desc = f"Direct ({args.cluster_host})"

    elif args.username and not args.cluster_host:
        # --username without --cluster-host makes no sense: Helios does not
        # expose a scriptable username/password login endpoint (it uses
        # OIDC/OAuth2 web-based auth). Direct the user to create an API key.
        print(
            "ERROR: --username is only supported with --cluster-host.\n"
            "\n"
            "  Helios does not provide a scriptable username/password login\n"
            "  endpoint. Its web login uses OIDC/OAuth2 and cannot be driven\n"
            "  from a script.\n"
            "\n"
            "  To use Helios, generate an API key and pass it with --apikey:\n"
            "    helios.cohesity.com → Settings → Access Management → API Keys\n"
            "\n"
            "  To run against a single cluster directly (no Helios required):\n"
            "    --cluster-host <ip/hostname> --username <user>"
        )
        sys.exit(1)

    else:
        # ── Helios API key (default) ──────────────────────────────────────────
        api_key = get_api_key(args.apikey)
        session.base_url = f"https://{HELIOS_HOST}"
        clusters = get_helios_clusters(api_key)
        if not clusters:
            print("ERROR: No clusters returned from Helios.")
            sys.exit(1)
        auth_desc = "Helios API key"
        # Quorum is a Helios control-plane feature configured in Security Center →
        # Security Posture → Quorum. When any enabled quorum group exists, the
        # quorum requirement applies to all clusters managed by that Helios tenant.
        # We therefore set quorum_enabled=True for every cluster when at least one
        # enabled group is returned by the API (rather than matching per clusterId,
        # which is unreliable when clusterIdentifiers is unpopulated).
        _quorum_enabled = False
        try:
            for _qg in _helios_quorum_groups(session, api_key, args.debug):
                if _qg.get("isEnabled"):
                    _quorum_enabled = True
                    break
        except Exception:
            pass
        for _c in clusters:
            _c["quorum_enabled"] = _quorum_enabled

    # Apply --cluster name filter (Helios modes only)
    if args.cluster and not args.cluster_host:
        clusters = [c for c in clusters
                    if args.cluster.lower() in c.get("name", "").lower()]
        if not clusters:
            print(f"ERROR: No cluster matching '{args.cluster}'")
            sys.exit(1)

    print(f"\nCohesity Health Check Report  v{__version__}")
    print(f"  Auth     : {auth_desc}")
    print(f"  Customer : {args.customer or '(not specified)'}")
    print(f"  Lookback : {args.days} days")
    print(f"  Mode     : {'Quick (last-run only)' if args.quick else 'Full (run history)'}")
    print(f"  Clusters : {', '.join(c.get('name', '?') for c in clusters)}")
    print()

    all_data = []
    for cluster in clusters:
        try:
            cd = collect_cluster(session, api_key, cluster, args)
            s  = cd["scores"]
            print(f"  [{cd['name']}] score={s['overall']}/100 ({s['grade']}) "
                  f"| {len(cd['recommendations'])} recommendation(s)")
            all_data.append(cd)
        except Exception as exc:
            import traceback
            print(f"  ERROR collecting [{cluster.get('name','?')}]: {exc}")
            if args.debug:
                traceback.print_exc()
            continue

    if not all_data:
        print("ERROR: No data collected from any cluster.")
        sys.exit(1)

    print()
    if not args.word_only:
        write_excel(all_data, args)
    if not args.excel_only:
        write_word(all_data, args)

    # Generate editable topology diagram (draw.io)
    drawio_path = _generate_topology_drawio(all_data, args.output)
    if drawio_path:
        print(f"\nComplete.  Output: {args.output}.xlsx / {args.output}.docx / {drawio_path}")
    else:
        print(f"\nComplete.  Output: {args.output}.xlsx / {args.output}.docx")


if __name__ == "__main__":
    main()
