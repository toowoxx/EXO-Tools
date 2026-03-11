"""
Configuration for EXO Tools web application.
All settings are loaded from environment variables.
"""
import os


class Config:
    # Flask
    SECRET_KEY = os.environ.get("SECRET_KEY", os.urandom(24).hex())
    SESSION_COOKIE_HTTPONLY = True
    SESSION_COOKIE_SAMESITE = "Lax"
    # Use secure cookies in production (HTTPS)
    SESSION_COOKIE_SECURE = os.environ.get("FLASK_ENV", "production") == "production"

    # Server-side session storage (Flask-Session).
    # Keeps the session cookie small by storing data on the server filesystem
    # rather than encoding it into the cookie itself.
    # Set SESSION_FILE_DIR to a persistent, non-world-readable directory in
    # production – e.g. /home/flask_session on Azure App Service (the /home
    # mount persists across restarts and is not accessible to other apps).
    # Flask-Session creates each session file with mode 0600 (owner-only).
    SESSION_TYPE = "filesystem"
    SESSION_FILE_DIR = os.environ.get("SESSION_FILE_DIR", "/tmp/flask_session")
    SESSION_PERMANENT = False

    # ----------------------------------------------------------------
    # Azure AD App Registration (user authentication + Graph API)
    # ----------------------------------------------------------------
    CLIENT_ID = os.environ.get("CLIENT_ID", "")
    CLIENT_SECRET = os.environ.get("CLIENT_SECRET", "")
    TENANT_ID = os.environ.get("TENANT_ID", "")
    REDIRECT_URI = os.environ.get("REDIRECT_URI", "http://localhost:5000/callback")
    AUTHORITY = f"https://login.microsoftonline.com/{os.environ.get('TENANT_ID', '')}"

    # Delegated scopes for Graph API (admin consent required for the tenant)
    SCOPES = ["User.Read", "User.ReadBasic.All", "GroupMember.Read.All"]

    # ----------------------------------------------------------------
    # Access Control
    # ----------------------------------------------------------------
    # Object ID of the Entra ID security group allowed to use this app.
    # Leave empty to allow all authenticated users in the tenant.
    ACCESS_GROUP_ID = os.environ.get("ACCESS_GROUP_ID", "")

    # ----------------------------------------------------------------
    # Exchange Online App Registration (PowerShell app-only auth)
    # ----------------------------------------------------------------
    EXO_APP_ID = os.environ.get("EXO_APP_ID", "")
    # Path to the PFX certificate file used for EXO app-only authentication
    EXO_CERT_PATH = os.environ.get("EXO_CERT_PATH", "/app/certs/exo.pfx")
    # Password for the PFX certificate (leave empty if the cert has no password)
    EXO_CERT_PASSWORD = os.environ.get("EXO_CERT_PASSWORD", "")
    # Your M365 tenant's onmicrosoft.com domain, e.g. contoso.onmicrosoft.com
    EXO_ORGANIZATION = os.environ.get("EXO_ORGANIZATION", "")

    # ----------------------------------------------------------------
    # PowerShell
    # ----------------------------------------------------------------
    # Path to the PowerShell Core executable
    PWSH_PATH = os.environ.get("PWSH_PATH", "pwsh")
    # Maximum seconds to wait for the PowerShell permission-check script
    PS_TIMEOUT = int(os.environ.get("PS_TIMEOUT", "600"))

    # ----------------------------------------------------------------
    # Microsoft Graph
    # ----------------------------------------------------------------
    GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"
