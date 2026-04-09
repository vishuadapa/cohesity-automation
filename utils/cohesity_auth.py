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

Supports three modes:
  1. Direct cluster  — username/password (+ optional MFA code) → Bearer token
                       via POST /v2/users/sessions
  2. Helios API key  — API key against helios.cohesity.com
  3. Helios user     — username/password (+ optional MFA/TOTP code)
                       via POST /mcm/login → token used as apiKey header

Credential storage:
  - Helios API key stored in OS keychain under service "cohesity_helios" / user "apikey"
  - Helios user passwords stored under "cohesity_helios_user" / user "<username>"
  - Cluster passwords stored under "cohesity_cluster" / user "<ip>:<domain>:<user>"
  - All fall back to interactive prompt if keyring is unavailable or nothing stored
  - clear_stored_credentials() removes stored creds

Typical import pattern:
  from utils.cohesity_auth import (
      get_api_key, get_helios_password, helios_login,
      get_cluster_password, get_auth_token,
      make_headers, make_helios_headers,
      get_helios_clusters, clear_stored_credentials, HELIOS_HOST,
  )

Version history:
  1.3 (2026-04-09) — fix: helios_login() tried /mcm/login which returns 404.
                     Rewritten to try /irisservices/api/v1/mcm/login first
                     (emailId payload field), then /mcm/login (username field)
                     as fallback. Also extracts token from session cookie for
                     Helios versions that return cookie-based sessions.
  1.2 (2026-04-09) — feat: Helios username/password + MFA login via
                     helios_login() / get_helios_password(). Added mfa_code
                     parameter to get_auth_token() for direct-cluster MFA.
                     Added cli_password parameter to get_cluster_password().
                     Updated clear_stored_credentials() to accept helios_user
                     for clearing Helios user passwords from keychain.
  1.1 (2026-04-07) — fix: parse response JSON once in get_auth_token;
                     validate access token is non-empty before returning.
  1.0 (2026-04-06) — feat: initial module. Extracted shared auth patterns from
                     fortknox_vault_report.py and protection_group_report.py
                     into a reusable module so individual scripts stay lean.
"""

__version__ = "1.3"

import getpass
import sys
import urllib3

import requests

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

HELIOS_HOST       = "helios.cohesity.com"
_KR_SVC_HELIOS    = "cohesity_helios"
_KR_USER_HELIOS   = "apikey"
_KR_SVC_HELIOS_US = "cohesity_helios_user"   # Helios username/password
_KR_SVC_CLUSTER   = "cohesity_cluster"


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
                              domain: str = "LOCAL", helios_user: str = None):
    """
    Remove stored credentials from the OS keychain.

    Behavior (evaluated in order, first match wins):
      helios_user set  → clear Helios username/password for that user
      cluster set      → clear direct-cluster password for cluster:domain:username
      neither set      → clear the Helios API key
    """
    if not _keyring_available():
        print("ERROR: 'keyring' package not installed. Nothing to clear.")
        sys.exit(1)
    import keyring
    if helios_user:
        kr_user = helios_user.lower()
        if keyring.get_password(_KR_SVC_HELIOS_US, kr_user):
            keyring.delete_password(_KR_SVC_HELIOS_US, kr_user)
            print(f"[*] Stored Helios password for '{helios_user}' removed from keychain.")
        else:
            print(f"[*] No stored Helios password found for '{helios_user}'.")
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
# Helios username/password auth
# ---------------------------------------------------------------------------

def get_helios_password(username: str, cli_password: str = None) -> str:
    """
    Return the Helios password for username/password login.
    cli_password is used as-is (not stored). Otherwise checks keychain under
    service "cohesity_helios_user", then prompts and saves.
    """
    if cli_password:
        return cli_password
    kr_user = username.lower()
    if _keyring_available():
        import keyring
        stored = keyring.get_password(_KR_SVC_HELIOS_US, kr_user)
        if stored:
            print(f"[*] Using stored Helios password for '{username}'.")
            return stored
    pwd = getpass.getpass(f"    Helios password for '{username}': ").strip()
    if not pwd:
        print("ERROR: Password cannot be empty.")
        sys.exit(1)
    if _keyring_available():
        import keyring
        keyring.set_password(_KR_SVC_HELIOS_US, kr_user, pwd)
        print(f"[*] Helios password saved to keychain for '{username}'.")
    return pwd


def helios_login(username: str, password: str, mfa_code: str = None) -> str:
    """
    Login to Helios with username/password and optional MFA/TOTP code.
    Returns a token that is used in the apiKey request header (same as
    a regular Helios API key).

    Tries these endpoints in order until one succeeds (not 404):
      1. POST /irisservices/api/v1/mcm/login  — primary path (emailId field)
      2. POST /mcm/login                       — alternate path (username field)

    MFA: otpCode + otpType="Totp" added to payload when mfa_code is given.
    Token extracted from response body (token / accessToken / sessionToken)
    or from the session cookie if none of those fields are present.
    """
    hdrs = {"Content-Type": "application/json", "Accept": "application/json"}

    # Endpoint candidates: (path, username_field_name)
    endpoints = [
        (f"https://{HELIOS_HOST}/irisservices/api/v1/mcm/login", "emailId"),
        (f"https://{HELIOS_HOST}/mcm/login",                     "username"),
    ]

    last_status = None
    last_body   = ""
    for url, user_field in endpoints:
        payload = {user_field: username, "password": password}
        if mfa_code:
            payload["otpCode"] = mfa_code
            payload["otpType"] = "Totp"
        try:
            r = requests.post(url, json=payload, headers=hdrs,
                              verify=False, timeout=30)
        except requests.exceptions.ConnectionError:
            print("ERROR: Cannot connect to Helios. Check your network/VPN.")
            sys.exit(1)

        if r.status_code == 404:
            continue           # try next endpoint

        if r.status_code in (401, 403):
            body = r.text[:400]
            if "otp" in body.lower():
                print("ERROR: Helios requires an MFA code. Re-run with --mfa-code <code>.")
                sys.exit(1)
            print(f"ERROR: Helios login failed ({r.status_code}). Check credentials.\n  {body}")
            sys.exit(1)

        if not r.ok:
            last_status = r.status_code
            last_body   = r.text[:400]
            continue           # try next endpoint

        # ── Successful response — extract token ───────────────────────────
        try:
            data = r.json()
        except ValueError:
            data = {}

        token = (data.get("token") or data.get("accessToken")
                 or data.get("sessionToken") or "")

        # Some Helios versions return the token as a session cookie
        if not token and r.cookies:
            for name in ("cohesity-session", "sessionToken", "token"):
                if name in r.cookies:
                    token = r.cookies[name]
                    break

        if not token:
            print(f"ERROR: No token in Helios login response: {r.text[:200]}")
            sys.exit(1)

        print(f"[+] Logged in to Helios as '{username}'")
        return token

    # All endpoints failed
    print(f"ERROR: Helios login failed — no valid endpoint found "
          f"(last status {last_status}).\n  {last_body}")
    sys.exit(1)


# ---------------------------------------------------------------------------
# Direct cluster auth
# ---------------------------------------------------------------------------

def get_auth_token(cluster: str, username: str, password: str, domain: str,
                   mfa_code: str = None) -> str:
    """
    Authenticate to a Cohesity cluster and return a Bearer token.
    Endpoint: POST /v2/users/sessions
    domain: 'LOCAL' for local accounts, or AD domain name (e.g. 'CORP').
    mfa_code: optional TOTP code for MFA-enabled accounts.
    """
    url     = f"https://{cluster}/v2/users/sessions"
    payload = {"username": username, "password": password, "domain": domain}
    if mfa_code:
        payload["otpCode"] = mfa_code
        payload["otpType"] = "Totp"
    hdrs    = {"Content-Type": "application/json", "Accept": "application/json"}
    try:
        r = requests.post(url, json=payload, headers=hdrs, verify=False, timeout=30)
        r.raise_for_status()
    except requests.exceptions.ConnectionError:
        print(f"ERROR: Cannot connect to '{cluster}'. Check hostname/IP and network.")
        sys.exit(1)
    except requests.exceptions.HTTPError as e:
        body = r.text[:400]
        if r.status_code == 401 and "otp" in body.lower():
            print("ERROR: Cluster requires an MFA code. Re-run with --mfa-code <code>.")
            sys.exit(1)
        print(f"ERROR: Authentication failed: {e}\n       Response: {body}")
        sys.exit(1)
    data         = r.json()
    access_token = data.get("accessToken", "")
    if not access_token:
        print(f"ERROR: No access token in response: {r.text[:200]}")
        sys.exit(1)
    token = f"{data.get('tokenType', 'Bearer')} {access_token}"
    print(f"[+] Authenticated to {cluster} as {domain}\\{username}")
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
