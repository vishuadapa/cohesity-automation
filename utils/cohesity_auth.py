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
  1.8 (2026-04-25) — feat: Helios bearer token auth via username/password.
                     Added get_helios_bearer_token() and make_helios_bearer_headers()
                     using POST /irisservices/api/v1/public/mcm/createAccessToken.
                     clear_stored_credentials() now accepts helios_user/helios_domain
                     to wipe stored Helios passwords from the OS keychain.
  1.7 (2026-04-17) — fix(security): raw API response bodies removed from
                     authentication error messages to prevent token/credential
                     fragments leaking into terminal history and CI logs.
                     Error output now shows only HTTP status codes.
  1.6 (2026-04-10) — fix: get_auth_token() v1 failure is now printed as
                     "[!] v1 endpoint failed (HTTP NNN) — trying v2 ..."
                     instead of being silently swallowed. Final error message
                     includes the v1 response and actionable troubleshooting
                     hints (--clear-credentials, --domain, account locked).
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

__version__ = "1.8"

import getpass
import sys

import requests

HELIOS_HOST       = "helios.cohesity.com"
_KR_SVC_HELIOS    = "cohesity_helios"
_KR_USER_HELIOS   = "apikey"
_KR_SVC_CLUSTER   = "cohesity_cluster"
_KR_SVC_HELIOS_PW = "cohesity_helios_password"


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
                              domain: str = "LOCAL",
                              helios_user: str = None,
                              helios_domain: str = "cohesity.com"):
    """
    Remove stored credentials from the OS keychain.

    helios_user set:    clears the stored Helios bearer-token password.
    cluster set:        clears the stored direct-cluster password.
    Neither set:        clears the Helios API key.
    """
    if not _keyring_available():
        print("ERROR: 'keyring' package not installed. Nothing to clear.")
        sys.exit(1)
    import keyring
    if helios_user:
        kr_user = f"{helios_domain}:{helios_user}"
        if keyring.get_password(_KR_SVC_HELIOS_PW, kr_user):
            keyring.delete_password(_KR_SVC_HELIOS_PW, kr_user)
            print(f"[*] Stored Helios password for {helios_user}@{helios_domain} removed.")
        else:
            print(f"[*] No stored Helios password found for {helios_user}@{helios_domain}.")
    elif cluster:
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
                   mfa_code: str = None, verify=True) -> str:
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
    _v1_note   = ""
    try:
        r = requests.post(url_v1, json=payload_v1, headers=hdrs,
                          verify=verify, timeout=30)
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
        _v1_note = f"HTTP {r.status_code}"
        print(f"  [!] v1 endpoint failed ({r.status_code}) — trying v2 ...")

    # ── Attempt 2 — v2 /users/sessions ───────────────────────────────────
    url_v2     = f"https://{cluster}/v2/users/sessions"
    payload_v2 = {"username": username, "password": password, "domain": domain}
    if mfa_code:
        payload_v2["otpCode"] = mfa_code
        payload_v2["otpType"] = "Totp"
    try:
        r = requests.post(url_v2, json=payload_v2, headers=hdrs,
                          verify=verify, timeout=30)
        r.raise_for_status()
    except requests.exceptions.ConnectionError:
        print(f"ERROR: Cannot connect to '{cluster}'. Check hostname/IP and network.")
        sys.exit(1)
    except requests.exceptions.HTTPError as e:
        _body = r.text[:400]
        if r.status_code == 401 and "otp" in _body.lower():
            print("ERROR: Cluster requires an MFA code. Re-run with --mfa-code <code>.")
            sys.exit(1)
        print(f"ERROR: Authentication failed on both endpoints.")
        if _v1_note:
            print(f"  v1: {_v1_note}")
        print(f"  v2: HTTP {r.status_code}")
        print()
        print("  Troubleshooting:")
        print("    • Wrong password?   Run --clear-credentials to wipe the stored")
        print("      password, then retry (you will be prompted for the correct one).")
        print(f"    • Wrong domain?    Use --domain <AD_DOMAIN> for AD accounts")
        print(f"      (current: {domain}). Use LOCAL for local cluster accounts.")
        print("    • Account locked?  Check Access Management in the Cohesity UI.")
        sys.exit(1)
    data         = r.json()
    access_token = data.get("accessToken", "")
    if not access_token:
        print("ERROR: No access token in response.")
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


def get_helios_bearer_token(username: str, domain: str = "cohesity.com",
                            cli_password: str = None, verify=True) -> str:
    """
    Authenticate to Helios with username/password and return a Bearer token string.
    Endpoint: POST https://helios.cohesity.com/irisservices/api/v1/public/mcm/createAccessToken
    Credentials are stored in the OS keychain under service 'cohesity_helios_password'.
    """
    kr_user  = f"{domain}:{username}"
    password = cli_password

    if not password and _keyring_available():
        import keyring
        password = keyring.get_password(_KR_SVC_HELIOS_PW, kr_user)
        if password:
            print(f"[*] Using stored Helios password for {username}@{domain}")

    if not password:
        password = getpass.getpass(f"    Helios password for {username}@{domain}: ").strip()
        if not password:
            print("ERROR: Password cannot be empty.")
            sys.exit(1)
        if _keyring_available():
            import keyring
            keyring.set_password(_KR_SVC_HELIOS_PW, kr_user, password)
            print(f"[*] Helios password saved to keychain for {username}@{domain}")

    url     = f"https://{HELIOS_HOST}/irisservices/api/v1/public/mcm/createAccessToken"
    payload = {"username": username, "password": password, "domain": domain}
    hdrs    = {"Content-Type": "application/json", "Accept": "application/json"}
    try:
        r = requests.post(url, json=payload, headers=hdrs, verify=verify, timeout=30)
        r.raise_for_status()
    except requests.exceptions.ConnectionError:
        print("ERROR: Cannot connect to Helios. Check your network/VPN.")
        sys.exit(1)
    except requests.exceptions.HTTPError:
        print(f"ERROR: Helios bearer token auth failed (HTTP {r.status_code}). Check credentials.")
        sys.exit(1)

    data         = r.json()
    access_token = data.get("accessToken", "")
    if not access_token:
        print("ERROR: No access token returned from Helios.")
        sys.exit(1)

    print(f"[+] Helios bearer token obtained for {username}@{domain}")
    return f"{data.get('tokenType', 'Bearer')} {access_token}"


def make_helios_bearer_headers(token: str, cluster_id: int = None) -> dict:
    """Build Helios request headers using a Bearer token (instead of apiKey)."""
    h = {
        "Authorization": token,
        "Accept":        "application/json",
        "Content-Type":  "application/json",
    }
    if cluster_id is not None:
        h["accessClusterId"] = str(cluster_id)
    return h


def get_helios_clusters(api_key: str, verify=True) -> list:
    """
    Return [{name, clusterId}, ...] for all Helios-connected clusters.
    Endpoint: GET /mcm/clusters/connectionStatus
    """
    url = f"https://{HELIOS_HOST}/mcm/clusters/connectionStatus"
    try:
        r = requests.get(url, headers=make_helios_headers(api_key),
                         verify=verify, timeout=30)
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
