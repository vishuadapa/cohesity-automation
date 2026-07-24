#!/usr/bin/env python3
"""
Cohesity MCP v3 — One-time credential setup
Stores credentials securely in the OS keyring:
  Windows  → Credential Manager
  macOS    → Keychain
  Linux    → SecretService / KWallet

Run this once per machine per user before starting cohesity_mcp_v3.py.
Nothing sensitive is written to any file.

Requires: pip install httpx keyring pyotp

MFA support:
  TOTP (Google/Microsoft Authenticator): store the secret key once → codes
       generated automatically at runtime, no manual input needed.
  Email OTP: prompted once per setup run; the resulting session token is
       cached for 23 h. Re-run this script when the token expires.
"""

import getpass
import json
import os
import sys
from pathlib import Path

try:
    import httpx
    import keyring
except ImportError:
    sys.exit(
        "Missing dependencies.\n"
        "Run:  pip install httpx keyring pyotp\n"
        "Then re-run this script."
    )

try:
    import pyotp
    _PYOTP = True
except ImportError:
    _PYOTP = False

KEYRING_SVC = "cohesity-mcp"

_CONFIG_SEARCH = [
    Path(os.environ.get("COHESITY_CONFIG", "__none__")),
    Path.home() / ".cohesity-mcp" / "config.json",
    Path(__file__).parent / "cohesity_clusters.json",
]

_STARTER_CONFIG = {
    "default": "cluster1",
    "clusters": [
        {
            "alias":      "cluster1",
            "type":       "direct",
            "host":       "your-cluster.corp.local",
            "domain":     "CORP",
            "username":   "your-ad-username",
            "verify_ssl": False,
        },
        {
            "alias":      "helios",
            "type":       "helios",
            "host":       "helios.cohesity.com",
            "verify_ssl": True,
        },
    ],
}


# ── Config helpers ──────────────────────────────────────────────────────────

def _find_config() -> tuple[Path | None, dict]:
    for p in _CONFIG_SEARCH:
        if p.exists():
            return p, json.loads(p.read_text())
    return None, {}


def _create_starter_config() -> dict:
    cfg_path = Path.home() / ".cohesity-mcp" / "config.json"
    cfg_path.parent.mkdir(parents=True, exist_ok=True)
    cfg_path.write_text(json.dumps(_STARTER_CONFIG, indent=2))
    print(f"  Created starter config: {cfg_path}")
    print()
    print("  Edit that file to add your cluster addresses, domains, and usernames,")
    print("  then re-run this script to store credentials.")
    print()
    print("  See also: cohesity_clusters.example.json for a fuller example.")
    sys.exit(0)


# ── Keyring helpers ─────────────────────────────────────────────────────────

def _keyring_key_direct(username: str, host: str) -> str:
    return f"{username}@{host}"


def _keyring_key_helios(host: str) -> str:
    return f"helios-apikey@{host}"


def _has_cred(key: str) -> bool:
    return bool(keyring.get_password(KEYRING_SVC, key))


def _store(key: str, value: str) -> None:
    keyring.set_password(KEYRING_SVC, key, value)


def _delete_cred(key: str) -> None:
    try:
        keyring.delete_password(KEYRING_SVC, key)
    except keyring.errors.PasswordDeleteError:
        pass


# ── Per-cluster setup ───────────────────────────────────────────────────────

def _auth_attempt(host: str, username: str, password: str,
                  domain: str, verify: bool,
                  otp_code: str = "", otp_type: str = "") -> dict:
    """POST to accessTokens; returns the decoded JSON on success, raises on error."""
    body: dict = {"username": username, "password": password, "domain": domain}
    if otp_code:
        body["otpCode"]  = otp_code
        body["otpType"]  = otp_type   # "Totp" | "Email"
    resp = httpx.post(
        f"https://{host}/irisservices/api/v1/public/accessTokens",
        json=body, verify=verify, timeout=30,
    )
    resp.raise_for_status()
    return resp.json()


def _is_mfa_error(exc: httpx.HTTPStatusError) -> bool:
    """Return True if the cluster is asking for an MFA OTP code.

    Cohesity signals MFA in several ways depending on version:
      - errorCode KMFA / KOTP / KMULTIFACTOR
      - errorCode KValidationError + "mandatory parameters"
        (otpCode is treated as a mandatory field when MFA is enabled)
    """
    try:
        body = exc.response.json()
        code = body.get("errorCode", "").upper()
        msg  = body.get("message", "").lower()
        return (
            "MFA" in code
            or "OTP" in code
            or "MULTIFACTOR" in code
            or "two-factor" in msg
            or "otp" in msg
            or "one-time" in msg
            # Cohesity 7.x returns KValidationError when otpCode is absent
            or (code == "KVALIDATIONERROR" and "mandatory" in msg)
        )
    except Exception:
        return False


def _prompt_otp() -> tuple[str, str]:
    """Ask user for OTP type and code; return (otp_type, otp_code)."""
    print()
    print("  MFA required. Choose type:")
    print("    1  TOTP  (Google / Microsoft Authenticator)")
    print("    2  Email OTP")
    choice = input("  Type [1]: ").strip() or "1"
    otp_type = "Totp" if choice != "2" else "Email"
    otp_code = input(f"  Enter {'authenticator' if otp_type == 'Totp' else 'email'} OTP code: ").strip()
    return otp_type, otp_code



def _setup_direct(cluster: dict) -> None:
    alias    = cluster["alias"]
    host     = cluster["host"]
    domain   = cluster.get("domain", "LOCAL")
    username = cluster.get("username", "").strip()
    verify   = cluster.get("verify_ssl", False)

    # Derive a NetBIOS-style short domain to try as fallback
    # e.g. "talabs.local" → "TALABS"
    short_domain = domain.split(".")[0].upper() if "." in domain else domain

    print(f"\n── {alias}  ({host},  domain={domain}) ──")
    if short_domain != domain and short_domain != domain.upper():
        print(f"  Note: if auth fails try changing domain to '{short_domain}' in config.json")

    if not username:
        username = input("  AD username: ").strip()
        if not username:
            print("  Skipped (no username).")
            return

    key = _keyring_key_direct(username, host)

    if _has_cred(key):
        ans = input(f"  Credential already stored for {username}@{host}. Update? [y/N]: ").strip().lower()
        if ans != "y":
            print("  Skipped.")
            return

    password = getpass.getpass(f"  Password for {username} (hidden): ")
    if not password:
        print("  Skipped (empty password).")
        return

    # ── First auth attempt (no OTP) ─────────────────────────────────────
    print("  Verifying...", end=" ", flush=True)
    tok      = None
    otp_type = ""
    used_otp = False
    try:
        tok = _auth_attempt(host, username, password, domain, verify)
        print(f"OK  (token type: {tok.get('tokenType', 'unknown')})")

    except httpx.HTTPStatusError as exc:
        try:
            body = exc.response.json()
            err_msg = body.get("message", "")
        except Exception:
            body    = {}
            err_msg = exc.response.text

        # "domain does not exist" is a hard config error — don't treat as MFA
        domain_error = "does not exist" in err_msg.lower() and "domain" in err_msg.lower()

        if exc.response.status_code in (400, 401) and not domain_error:
            # Any 400/401 that isn't a domain problem = treat as MFA challenge.
            # Cohesity 7.x returns KValidationError+"mandatory parameters" when
            # otpCode is missing; older builds return KMFA. Catch all variants.
            print("MFA required.")
            otp_type, otp_code = _prompt_otp()
            print("  Verifying with OTP...", end=" ", flush=True)
            try:
                tok      = _auth_attempt(host, username, password, domain,
                                         verify, otp_code, otp_type)
                used_otp = True
                print(f"OK  (token type: {tok.get('tokenType', 'unknown')})")
            except httpx.HTTPStatusError as exc2:
                print(f"FAILED  (HTTP {exc2.response.status_code})")
                print(f"  Response: {exc2.response.text[:300]}")
                ans = input("  Store credential anyway? [y/N]: ").strip().lower()
                if ans != "y":
                    print("  Credential not stored.")
                    return
            except Exception as exc2:
                print(f"FAILED  ({exc2})")
                ans = input("  Store credential anyway? [y/N]: ").strip().lower()
                if ans != "y":
                    print("  Credential not stored.")
                    return
        else:
            print(f"FAILED  (HTTP {exc.response.status_code})")
            print(f"  {err_msg}")
            if domain_error:
                print(f"  Check the 'domain' field in your config.json.")
                print(f"  Tried: '{domain}'  —  ask your Cohesity admin for the registered AD domain name.")
            ans = input("  Store credential anyway? [y/N]: ").strip().lower()
            if ans != "y":
                print("  Credential not stored.")
                return

    except Exception as exc:
        print(f"FAILED  ({exc})")
        ans = input("  Store credential anyway? [y/N]: ").strip().lower()
        if ans != "y":
            print("  Credential not stored.")
            return

    _store(key, password)
    print(f"  ✓ Password stored:  service={KEYRING_SVC!r}  key={key!r}")

    # ── Cache the session token so the MCP server can use it immediately ─
    # The MCP server reads this token directly — no re-auth (and no MFA
    # challenge) until it expires after 23 hours.
    if tok:
        import time as _time
        token_str = f"{tok['tokenType']} {tok['accessToken']}"
        expiry    = int(_time.time() + 23 * 3600)
        _store(f"session-token@{host}", f"{token_str}|{expiry}")
        print(f"  ✓ Session token cached (valid ~23 h) — MCP server ready to use immediately")



def _setup_helios(cluster: dict) -> None:
    alias = cluster["alias"]
    host  = cluster.get("host", "helios.cohesity.com")
    key   = _keyring_key_helios(host)

    print(f"\n── {alias}  ({host}) ──")
    print("  Get your API key from: https://helios.cohesity.com → Account → API Keys")

    if _has_cred(key):
        ans = input("  Helios API key already stored. Update? [y/N]: ").strip().lower()
        if ans != "y":
            print("  Skipped.")
            return

    api_key = getpass.getpass("  Helios API key (hidden): ")
    if not api_key:
        print("  Skipped (empty key).")
        return

    print("  Verifying...", end=" ", flush=True)
    try:
        resp = httpx.get(
            f"https://{host}/irisservices/api/v1/mcm/clusters/info",
            headers={"apiKey": api_key, "Accept": "application/json"},
            verify=True,
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        n = len(data) if isinstance(data, list) else len(data.get("clusterList", []))
        print(f"OK  ({n} clusters connected to Helios)")
    except httpx.HTTPStatusError as exc:
        print(f"FAILED  (HTTP {exc.response.status_code})")
        ans = input("  Store credential anyway? [y/N]: ").strip().lower()
        if ans != "y":
            print("  Credential not stored.")
            return
    except Exception as exc:
        print(f"FAILED  ({exc})")
        ans = input("  Store credential anyway? [y/N]: ").strip().lower()
        if ans != "y":
            print("  Credential not stored.")
            return

    _store(key, api_key)
    print(f"  ✓ Stored:  service={KEYRING_SVC!r}  key={key!r}")


# ── Status display ──────────────────────────────────────────────────────────

def _status(clusters: list[dict]) -> None:
    print()
    print("Keyring status:")
    for c in clusters:
        alias = c["alias"]
        ctype = c.get("type", "direct")
        if ctype == "helios":
            host = c.get("host", "helios.cohesity.com")
            key  = _keyring_key_helios(host)
            mark = "✓" if _has_cred(key) else "✗"
            print(f"  {mark}  {alias:<18}  {host}  [Helios API key]")
        else:
            host = c["host"]
            user = c.get("username", "<no username>")
            key  = _keyring_key_direct(user, host)
            mark = "✓" if _has_cred(key) else "✗"
            print(f"  {mark}  {alias:<18}  {user}@{host}")


# ── Delete flow ─────────────────────────────────────────────────────────────

def _delete_flow(clusters: list[dict]) -> None:
    print()
    for i, c in enumerate(clusters, 1):
        alias = c["alias"]
        ctype = c.get("type", "direct")
        if ctype == "helios":
            host = c.get("host", "helios.cohesity.com")
            key  = _keyring_key_helios(host)
        else:
            host = c["host"]
            user = c.get("username", "")
            key  = _keyring_key_direct(user, host)
        mark = "stored" if _has_cred(key) else "not stored"
        print(f"  {i}. {alias}  ({mark})")

    nums = input("\nDelete credentials for which clusters? (comma-separated numbers, or Enter to cancel): ").strip()
    if not nums:
        return
    for part in nums.split(","):
        try:
            idx = int(part.strip()) - 1
            c   = clusters[idx]
        except (ValueError, IndexError):
            print(f"  Invalid selection: {part.strip()!r}")
            continue
        alias = c["alias"]
        ctype = c.get("type", "direct")
        if ctype == "helios":
            host = c.get("host", "helios.cohesity.com")
            key  = _keyring_key_helios(host)
        else:
            host = c["host"]
            user = c.get("username", "")
            key  = _keyring_key_direct(user, host)
        _delete_cred(key)
        print(f"  Deleted: {alias}  ({key})")


# ── Main ────────────────────────────────────────────────────────────────────

def main() -> None:
    print("═" * 60)
    print("  Cohesity MCP v3 — Credential Setup")
    print("  Passwords/keys stored in OS keyring only — never in files.")
    print("═" * 60)

    cfg_path, cfg = _find_config()
    if cfg_path:
        print(f"\nConfig: {cfg_path}")
    else:
        print("\nNo config file found.")
        ans = input("Create a starter config at ~/.cohesity-mcp/config.json? [Y/n]: ").strip().lower()
        if ans not in ("", "y"):
            print("Aborted.")
            sys.exit(0)
        _create_starter_config()

    clusters: list[dict] = cfg.get("clusters", [])
    if not clusters:
        print("No clusters defined in config. Edit the config file and re-run.")
        sys.exit(1)

    _status(clusters)

    print()
    print("Options:")
    print("  1. Configure / update credentials")
    print("  2. Delete stored credentials")
    print("  3. Show status and exit")
    choice = input("\nChoice [1]: ").strip() or "1"

    if choice == "1":
        for cluster in clusters:
            ctype = cluster.get("type", "direct")
            if ctype == "helios":
                _setup_helios(cluster)
            else:
                _setup_direct(cluster)
        _status(clusters)
        print()
        print("Done. Start the MCP server with:  python cohesity_mcp_v3.py")

    elif choice == "2":
        _delete_flow(clusters)
        _status(clusters)

    elif choice == "3":
        _status(clusters)

    else:
        print("Unknown option.")


if __name__ == "__main__":
    main()
