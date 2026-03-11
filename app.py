"""
EXO Tools – Reverse Mailbox Permission Search
Flask web application entry point.
"""
import json
import logging
import os
import re
import subprocess
import uuid
from functools import wraps
from urllib.parse import urlparse

import msal
import requests
from cachelib import FileSystemCache
from flask import (
    Flask,
    jsonify,
    redirect,
    render_template,
    request,
    session,
    url_for,
)
from flask_session import Session

from config import Config

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(name)s %(levelname)s %(message)s",
)
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# App setup
# ---------------------------------------------------------------------------
app = Flask(__name__)
app.config.from_object(Config)

# Ensure the session directory exists and is writable before Flask-Session
# initialises.  /home is a persistent Azure Files mount on App Service but
# some tiers or Zip-deployed apps make it read-only at startup; /tmp is
# always writable and is the safe fallback (non-persistent, but sessions
# are short-lived anyway).
def _find_session_dir(preferred: str) -> str:
    """Return the first writable candidate directory, creating it if needed."""
    for candidate in [preferred, "/tmp/flask_session"]:
        try:
            os.makedirs(candidate, exist_ok=True)
            probe = os.path.join(candidate, ".write_probe")
            with open(probe, "w") as _f:
                pass
            os.remove(probe)
            return candidate
        except OSError:
            logger.warning("Session directory not writable: %s – trying next", candidate)
    raise OSError(
        "No writable session directory found. "
        "Tried: %s and /tmp/flask_session. "
        "Set SESSION_FILE_DIR to a writable path." % preferred
    )


_session_dir = _find_session_dir(
    app.config.get("SESSION_FILE_DIR", "/home/flask_session")
)
logger.info("Using session directory: %s", _session_dir)

# Wire up the cachelib FileSystemCache directly so Flask-Session uses the
# modern, non-deprecated SESSION_TYPE="cachelib" path.
app.config["SESSION_CACHELIB"] = FileSystemCache(
    cache_dir=_session_dir,
    threshold=500,
    mode=0o600,
)
Session(app)

# Separate filesystem store for MSAL token caches.  Keeping tokens out of
# the Flask session avoids the 4 KB browser cookie size limit entirely – the
# session cookie carries only a tiny session-ID / user-identity dict while
# the bulky MSAL blobs (refresh tokens, account metadata, etc.) live on disk.
_msal_cache_dir = os.path.join(_session_dir, "msal_tokens")
os.makedirs(_msal_cache_dir, exist_ok=True)
_msal_store = FileSystemCache(
    cache_dir=_msal_cache_dir,
    threshold=500,
    mode=0o600,
)

# Warn loudly at startup if REDIRECT_URI is misconfigured (a common mistake
# that causes a login loop: Azure AD redirects to the wrong URL and the
# auth code is never handled).
_redirect_uri = app.config.get("REDIRECT_URI", "")
if not urlparse(_redirect_uri).path.rstrip("/").endswith("/callback"):
    logger.warning(
        "REDIRECT_URI (%s) does not end with /callback. "
        "Authentication will fail. Set REDIRECT_URI to "
        "https://<your-app>.azurewebsites.net/callback",
        _redirect_uri,
    )


# ---------------------------------------------------------------------------
# MSAL helpers
# ---------------------------------------------------------------------------

def _build_msal_app(cache: msal.SerializableTokenCache | None = None):
    return msal.ConfidentialClientApplication(
        app.config["CLIENT_ID"],
        authority=app.config["AUTHORITY"],
        client_credential=app.config["CLIENT_SECRET"],
        token_cache=cache,
        # Skip authority validation network call at construction time.
        # The authority URL format is pre-validated via TENANT_ID configuration.
        validate_authority=False,
    )


def _load_cache(oid: str | None = None) -> msal.SerializableTokenCache:
    """Deserialise the MSAL token cache from the filesystem store."""
    cache = msal.SerializableTokenCache()
    oid = oid or session.get("user", {}).get("oid")
    if oid:
        raw = _msal_store.get(f"msal:{oid}")
        if raw:
            cache.deserialize(raw)
    return cache


def _save_cache(cache: msal.SerializableTokenCache, oid: str | None = None) -> None:
    """Persist the MSAL token cache to the filesystem store (never the session)."""
    if not cache.has_state_changed:
        return
    oid = oid or session.get("user", {}).get("oid")
    if not oid:
        logger.warning("Cannot persist MSAL token cache: no user OID available")
        return
    # Strip large short-lived token blobs.  MSAL can silently re-acquire
    # access tokens using the refresh token, and the ID token is already
    # decoded into session["user"].
    cache_data = json.loads(cache.serialize())
    cache_data.pop("AccessToken", None)
    cache_data.pop("IdToken", None)
    _msal_store.set(f"msal:{oid}", json.dumps(cache_data), timeout=86400)


def _get_token_from_cache(scopes: list[str] | None = None):
    """Return a valid access token from the MSAL cache (refresh if needed)."""
    cache = _load_cache()
    cca = _build_msal_app(cache=cache)
    accounts = cca.get_accounts()
    if accounts:
        result = cca.acquire_token_silent(
            scopes or app.config["SCOPES"], account=accounts[0]
        )
        _save_cache(cache)
        return result
    return None


# ---------------------------------------------------------------------------
# Access-control helpers
# ---------------------------------------------------------------------------

def _check_group_membership(access_token: str) -> bool:
    """Return True when the signed-in user belongs to the required group."""
    group_id = app.config["ACCESS_GROUP_ID"]
    if not group_id:
        return True  # no restriction configured

    headers = {"Authorization": f"Bearer {access_token}"}

    # Fast path: check the `groups` claim already in the session's ID token
    groups_in_token = session.get("user", {}).get("groups", [])
    if groups_in_token and group_id in groups_in_token:
        return True

    # Graph API fallback (handles groups-overage and missing claim)
    try:
        resp = requests.post(
            f"{app.config['GRAPH_API_BASE']}/me/checkMemberObjects",
            headers=headers,
            json={"ids": [group_id]},
            timeout=10,
        )
        if resp.status_code == 200:
            return group_id in resp.json().get("value", [])
    except Exception as exc:
        logger.warning("Group membership check failed: %s", exc)

    return False


# ---------------------------------------------------------------------------
# Auth decorator
# ---------------------------------------------------------------------------

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("user"):
            session["next"] = request.url
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated


# ---------------------------------------------------------------------------
# Routes – authentication
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    if not session.get("user"):
        return redirect(url_for("login"))
    return render_template("dashboard.html", user=session["user"])


@app.route("/login")
def login():
    if session.get("user"):
        return redirect(url_for("index"))

    state = str(uuid.uuid4())
    session["state"] = state

    auth_url = None
    error = None
    try:
        cca = _build_msal_app()
        auth_url = cca.get_authorization_request_url(
            app.config["SCOPES"],
            state=state,
            redirect_uri=app.config["REDIRECT_URI"],
        )
    except Exception as exc:
        logger.error("Failed to build auth URL: %s", exc)
        error = "Could not reach the Microsoft login service. Please check the application configuration."

    return render_template("login.html", auth_url=auth_url, error=error)


@app.route("/callback")
def callback():
    # CSRF guard
    if request.args.get("state") != session.get("state"):
        return render_template("login.html", auth_url=None, error="State mismatch. Please try again."), 400

    if request.args.get("error"):
        desc = request.args.get("error_description", "Unknown error")
        return render_template("login.html", auth_url=None, error=f"Authentication failed: {desc}"), 400

    cache = _load_cache()
    cca = _build_msal_app(cache=cache)
    result = cca.acquire_token_by_authorization_code(
        request.args["code"],
        scopes=app.config["SCOPES"],
        redirect_uri=app.config["REDIRECT_URI"],
    )

    if "error" in result:
        return render_template(
            "login.html",
            auth_url=None,
            error=f"Could not acquire token: {result.get('error_description', result.get('error'))}",
        ), 400

    # Group membership check
    if app.config["ACCESS_GROUP_ID"]:
        if not _check_group_membership(result["access_token"]):
            return render_template(
                "login.html",
                auth_url=None,
                error="Access denied. Your account is not in the required group.",
            ), 403

    # Store only the essential identity claims.  The raw id_token_claims
    # dict can contain dozens of Azure AD-specific entries (group GUIDs,
    # xms_* extension claims, nonce, etc.).  We persist only what the
    # application actually reads – this keeps the session cookie tiny.
    claims = result.get("id_token_claims", {})
    session["user"] = {
        "name": claims.get("name", ""),
        "preferred_username": claims.get("preferred_username", ""),
        "oid": claims.get("oid", ""),
        "tid": claims.get("tid", ""),
        # Keep the groups claim for the fast-path group-membership check.
        # Azure AD only includes groups here when the count is small (<= 6 by
        # default); if the user is in many groups the claim is absent and we
        # fall through to the Graph API check automatically.
        "groups": claims.get("groups", []),
    }

    # Persist the MSAL token cache to the *filesystem* (not the session).
    # This is the key change that prevents the Set-Cookie header from
    # exceeding the 4 KB browser limit.
    _save_cache(cache, oid=claims.get("oid", ""))

    next_url = session.pop("next", url_for("index"))
    return redirect(next_url)


@app.route("/logout")
def logout():
    # Remove the user's MSAL token cache from the filesystem store.
    oid = session.get("user", {}).get("oid")
    if oid:
        _msal_store.delete(f"msal:{oid}")
    session.clear()
    post_logout = url_for("index", _external=True)
    return redirect(
        f"{app.config['AUTHORITY']}/oauth2/v2.0/logout"
        f"?post_logout_redirect_uri={post_logout}"
    )


# ---------------------------------------------------------------------------
# Routes – API
# ---------------------------------------------------------------------------

@app.route("/api/search-users")
@login_required
def search_users():
    """Search M365 users via the Microsoft Graph API (People Picker)."""
    query = request.args.get("q", "").strip()
    if len(query) < 2:
        return jsonify({"users": []})

    token_result = _get_token_from_cache()
    if not token_result:
        return jsonify({"error": "Authentication session expired. Please refresh the page."}), 401

    headers = {
        "Authorization": f"Bearer {token_result['access_token']}",
        "ConsistencyLevel": "eventual",
    }
    params = {
        "$search": f'"displayName:{query}" OR "userPrincipalName:{query}"',
        "$select": "id,displayName,userPrincipalName,mail,jobTitle,department",
        "$top": "15",
        "$count": "true",
    }

    try:
        resp = requests.get(
            f"{app.config['GRAPH_API_BASE']}/users",
            headers=headers,
            params=params,
            timeout=10,
        )
        if resp.status_code == 200:
            users = [
                {
                    "id": u.get("id"),
                    "displayName": u.get("displayName", ""),
                    "userPrincipalName": u.get("userPrincipalName", ""),
                    "mail": u.get("mail") or u.get("userPrincipalName", ""),
                    "jobTitle": u.get("jobTitle") or "",
                    "department": u.get("department") or "",
                }
                for u in resp.json().get("value", [])
                if u.get("userPrincipalName")
            ]
            return jsonify({"users": users})

        logger.error("Graph API %s: %s", resp.status_code, resp.text[:200])
        return jsonify({"error": "Failed to search users", "users": []}), resp.status_code

    except requests.exceptions.Timeout:
        return jsonify({"error": "Search timed out", "users": []}), 504
    except Exception as exc:
        logger.exception("User search error: %s", exc)
        return jsonify({"error": "Internal server error", "users": []}), 500


@app.route("/api/get-permissions", methods=["POST"])
@login_required
def get_permissions():
    """Run the PowerShell EXO permission-check script for the given UPN."""
    data = request.get_json(silent=True) or {}
    upn = (data.get("userPrincipalName") or "").strip()

    # Input validation – strict UPN format to prevent injection
    if not upn or not re.fullmatch(
        r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", upn
    ):
        return jsonify({"error": "A valid userPrincipalName is required."}), 400

    # Check EXO configuration
    cfg_errors = []
    if not app.config["EXO_APP_ID"]:
        cfg_errors.append("EXO_APP_ID is not set.")
    if not app.config["EXO_ORGANIZATION"]:
        cfg_errors.append("EXO_ORGANIZATION is not set.")
    if not os.path.isfile(app.config["EXO_CERT_PATH"]):
        cfg_errors.append(
            f"EXO certificate not found at: {app.config['EXO_CERT_PATH']}"
        )
    if cfg_errors:
        return jsonify({"error": "Exchange Online is not configured: " + " ".join(cfg_errors)}), 503

    script_path = os.path.join(
        os.path.dirname(__file__), "scripts", "Get-MailboxPermissions.ps1"
    )
    if not os.path.isfile(script_path):
        return jsonify({"error": "PowerShell script not found on server."}), 500

    cmd = [
        app.config["PWSH_PATH"],
        "-NonInteractive",
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", script_path,
        "-UserPrincipalName", upn,
        "-AppId", app.config["EXO_APP_ID"],
        "-CertificatePath", app.config["EXO_CERT_PATH"],
        "-Organization", app.config["EXO_ORGANIZATION"],
    ]
    if app.config["EXO_CERT_PASSWORD"]:
        cmd += ["-CertificatePassword", app.config["EXO_CERT_PASSWORD"]]

    logger.info("Starting permission check for: %s", upn)
    try:
        proc = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=app.config["PS_TIMEOUT"],
        )

        if proc.returncode != 0:
            err = proc.stderr.strip() or "PowerShell exited with a non-zero status."
            logger.error("PowerShell stderr: %s", err[:500])
            return jsonify({"error": f"Permission check failed: {err[:300]}"}), 500

        stdout = proc.stdout.strip()
        permissions = _parse_ps_json(stdout)
        logger.info("Permission check complete. Found %d permissions.", len(permissions))
        return jsonify({"permissions": permissions, "user": upn, "count": len(permissions)})

    except subprocess.TimeoutExpired:
        return jsonify({
            "error": (
                "The permission check timed out. "
                "This usually happens on very large tenants. "
                "Consider increasing PS_TIMEOUT."
            )
        }), 504
    except FileNotFoundError:
        return jsonify({
            "error": (
                f"PowerShell not found at '{app.config['PWSH_PATH']}'. "
                "Please install PowerShell Core and set the PWSH_PATH environment variable."
            )
        }), 500
    except Exception as exc:
        logger.exception("Unexpected error during permission check: %s", exc)
        return jsonify({"error": f"Internal server error: {exc}"}), 500


def _parse_ps_json(output: str) -> list:
    """Extract and parse the JSON array/object from PowerShell stdout."""
    if not output:
        return []

    # PowerShell may emit log lines before the JSON – find the last JSON block
    for start_char, end_char in (("[", "]"), ("{", "}")):
        start = output.rfind(start_char)
        end = output.rfind(end_char)
        if start != -1 and end > start:
            try:
                parsed = json.loads(output[start : end + 1])
                return parsed if isinstance(parsed, list) else [parsed]
            except json.JSONDecodeError:
                pass

    logger.warning("Could not parse PowerShell output as JSON: %s", output[:300])
    return []


# ---------------------------------------------------------------------------
# Health check
# ---------------------------------------------------------------------------

@app.route("/health")
def health():
    return jsonify({"status": "ok", "version": "1.0.0"})


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_ENV", "production") == "development"
    app.run(host="0.0.0.0", port=port, debug=debug)
