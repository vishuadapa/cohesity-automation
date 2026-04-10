#!/usr/bin/env python3
# =============================================================================
# DISCLAIMER: This code is provided on a best-effort basis and is not in any
# way officially supported or sanctioned by Cohesity. The code is intentionally
# kept simple to retain value as example code. The code in this repository is
# provided as-is and the author accepts no liability for damages resulting
# from its use.
# =============================================================================
"""
cohesity_auth.py
----------------
Unified authentication helper for Cohesity scripts.

Supports two modes:
  1. Helios (SaaS)   — API key against helios.cohesity.com
                       Helios does not expose a scriptable username/password
                       endpoint; its web login uses an OIDC/OAuth2 flow.
                       Generate an API key at:
                         helios.cohesity.com → Settings → Access Management → API Keys
  2. Direct cluster  — username/password (+ optional MFA code) → Bearer token
                       via POST /v2/users/sessions

Credential storage:
  - Helios API key stored in OS keychain under service "cohesity_helios" / user "apikey"
  - Cluster passwords stored under "cohesity_cluster" / user "<ip>:<domain>:<user>"
  - Both fall back to interactive prompt if keyring is unavailable or nothing stored
  - clear_stored_credentials() removes stored creds

Typical import pattern:
  from utils.cohesity_auth import (
      get_api_key, get_cluster_password, get_auth_token,
      make_headers, make_helios_headers,
      get_helios_clusters, clear_stored_credentials, HELIOS_HOST,
  )

Version history:
  1.5 (2026-04-10) — fix: get_auth_token() now tries v1 endpoint first
                     (POST /irisservices/api/v1/public/accessTokens) then
                     falls back to v2 (POST /v2/users/sessions). The v1
                     endpoint has the widest compatibility across Cohesity
                     versions and account types. v2 /users/sessions returns
                     400 "KValidationError Access denied" on some clusters.
  1.4 (2026-04-09) — fix: removed helios_login() / get_helios_password() —
                     Helios login endpoints require OIDC/OAuth2 (not plain
                     username/password). The /irisservices/api/v1/mcm/login
                     endpoint returned 401 "No open ID token found in header".
                     Helios access requires an API key; username/password auth
                     is only supported for direct cluster (--cluster-host).
                     Removed helios_user param from clear_stored_credentials().
  1.3 (2026-04-09) — fix: helios_login() /mcm/login returns 404; added
                     fallback to /irisservices/api/v1/mcm/login (removed in
                     v1.4 as the endpoint requires OIDC, not credentials).
  1.2 (2026-04-09) — feat: Helios username/password + MFA login attempt;
                     mfa_code param for get_auth_token(); cli_password param
                     for get_cluster_password().
  1.1 (2026-04-07) — fix: parse response JSON once in get_auth_token;
                     validate access token is non-empty before returning.
  1.0 (2026-04-06) — feat: initial module. Extracted shared auth patterns from
                     fortknox_vault_report.py and protection_group_report.py
                     into a reusable module so individual scripts stay lean.
"""

__version__ = "1.5"

import getpass
import sys
import urllib3

import requests

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

HELIOS_HOST     = "helios.cohesity.com"
_KR_SVC_HELIOS  = "cohesity_helios"
_KR_USER_HELIOS = "apikey"
_KR_SVC_CLUSTER = "cohesity_cluster"


# ---------------------------------------------------------------------------
# Keyring helpers
# ---------------------------------------------------------------------------

def _keyring_available() -> bool:
    try:
        import keyring          # noqa: F401
        return True
    except ImportError:
        return False


# ---------------------------------------------------------------------------
# Helios API key
# ---------------------------------------------------------------------------

def get_api_key(cli_key: str = None) -> str:
    """
    Return the Helios API key. Priority:
      1. --apikey CLI argument (one-time use, not stored)
      2. Key stored in the OS keychain
      3. Interactive prompt — key is saved to keychain for future runs

    The key is never written to disk in plaintext and never appears
    in the process list.
    """
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
        print("          Install with:  pip install keyring")

    key = getpass.getpass("    Enter Helios API key (will be saved to keychain): ").strip()
    if not key:
        print("ERROR: API key cannot be empty.")
        sys.exit(1)

    if _keyring_available():
        import keyring
        keyring.set_password(_KR_SVC_HELIOS, _KR_USER_HELIOS, key)
        print("[*] Helios API key saved to system keychain.")

    return key


# ---------------------------------------------------------------------------
# Direct cluster password
# ---------------------------------------------------------------------------

def get_cluster_password(cluster: str, username: str, domain: str,
                          cli_password: str = None) -> str:
    """
    Return the password for direct cluster auth.
    cli_password is used as-is (not stored). Otherwise checks keychain,
    then prompts interactively and saves the result.
    """
    if cli_password:
        return cli_password
    kr_user = f"{cluster}:{domain}:{username}"
    if _keyring_available():
        import keyring
        stored = keyring.get_password(_KR_SVC_CLUSTER, kr_user)
        if stored:
            print(f"[*] Using stored password for {domain}\\{username}@{cluster}")
            return stored

    pwd = getpass.getpass(f"Password for {domain}\\{username}@{cluster}: ").strip()
    if not pwd:
        print("ERROR: Password cannot be empty.")
        sys.exit(1)

    if _keyring_available():
        import keyring
        keyring.set_password(_KR_SVC_CLUSTER, kr_user, pwd)
        print(f"[*] Password saved to keychain for {domain}\\{username}@{cluster}")

    return pwd


# ---------------------------------------------------------------------------
# Clear stored credentials
# ---------------------------------------------------------------------------

def clear_stored_credentials(cluster: str = None, username: str = "admin",
                              domain: str = "LOCAL"):
    """
    Remove stored credentials from the OS keychain.

    Without --cluster-host: clears the Helios API key.
    With --cluster-host:    clears the stored direct-cluster password.
    """
    if not _keyring_available():
        print("ERROR: 'keyring' package not installed. Nothing to clear.")
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
            print("[*] Stored Helios API key removed from system keychain.")
        else:
            print("[*] No stored Helios API key found — nothing to clear.")


# ---------------------------------------------------------------------------
# Direct cluster auth
# ---------------------------------------------------------------------------

def get_auth_token(cluster: str, username: str, password: str, domain: str,
                   mfa_code: str = None) -> str:
    """
    Authenticate to a Cohesity cluster and return a Bearer token.

    Tries two endpoints in order:
      1. POST /irisservices/api/v1/public/accessTokens  (classic, widest compat)
      2. POST /v2/users/sessions                         (v2 API, newer clusters)

    domain: 'LOCAL' for local accounts, or AD domain name (e.g. 'CORP').
    mfa_code: optional TOTP code for MFA-enabled accounts.
    """
    hdrs = {"Content-Type": "application/json", "Accept": "application/json"}

    # ── Attempt 1 — v1 /public/accessTokens (works on all versions) ──────
    url_v1     = f"https://{cluster}/irisservices/api/v1/public/accessTokens"
    payload_v1 = {"username": username, "password": password, "domain": domain}
    try:
        r = requests.post(url_v1, json=payload_v1, headers=hdrs,
                          verify=False, timeout=30)
        r.raise_for_status()
        data         = r.json()
        access_token = data.get("accessToken", "")
        if access_token:
            token = f"{data.get('tokenType', 'Bearer')} {access_token}"
            print(f"[+] Authenticated to {cluster} as {domain}\\{username} (v1)")
            return token
    except requests.exceptions.ConnectionError:
        print(f"ERROR: Cannot connect to '{cluster}'. Check hostname/IP and network.")
        sys.exit(1)
    except requests.exceptions.HTTPError:
        pass  # fall through to v2

    # ── Attempt 2 — v2 /users/sessions ───────────────────────────────────
    url_v2     = f"https://{cluster}/v2/users/sessions"
    payload_v2 = {"username": username, "password": password, "domain": domain}
    if mfa_code:
        payload_v2["otpCode"] = mfa_code
        payload_v2["otpType"] = "Totp"
    try:
        r = requests.post(url_v2, json=payload_v2, headers=hdrs,
                          verify=False, timeout=30)
        r.raise_for_status()
    except requests.exceptions.ConnectionError:
        print(f"ERROR: Cannot connect to '{cluster}'. Check hostname/IP and network.")
        sys.exit(1)
    except requests.exceptions.HTTPError as e:
        body = r.text[:400]
        if r.status_code == 401 and "otp" in body.lower():
            print("ERROR: Cluster requires an MFA code. Re-run with --mfa-code <code>.")
            sys.exit(1)
        print(f"ERROR: Authentication failed on both v1 and v2 endpoints.\n"
              f"       v2 response: {e}\n       Body: {body}\n"
              f"       Check username/password/domain and that the account is active.")
        sys.exit(1)
    data         = r.json()
    access_token = data.get("accessToken", "")
    if not access_token:
        print(f"ERROR: No access token in response: {r.text[:200]}")
        sys.exit(1)
    token = f"{data.get('tokenType', 'Bearer')} {access_token}"
    print(f"[+] Authenticated to {cluster} as {domain}\\{username} (v2)")
    return token


def make_headers(token: str) -> dict:
    """Build request headers for direct cluster auth."""
    return {
        "Authorization": token,
        "Accept":        "application/json",
        "Content-Type":  "application/json",
    }


# ---------------------------------------------------------------------------
# Helios auth
# ---------------------------------------------------------------------------

def make_helios_headers(api_key: str, cluster_id: int = None) -> dict:
    """
    Build Helios request headers.
    'accessClusterId' routes the call to a specific on-prem cluster.
    (Confirmed from cohesity-api.ps1 — NOT 'clusterIdentifier'.)
    """
    h = {
        "apiKey":        api_key,
        "Accept":        "application/json",
        "Content-Type":  "application/json",
    }
    if cluster_id is not None:
        h["accessClusterId"] = str(cluster_id)
    return h


def get_helios_clusters(api_key: str) -> list:
    """
    Return [{name, clusterId}, ...] for all Helios-connected clusters.
    Endpoint: GET /mcm/clusters/connectionStatus
    """
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
