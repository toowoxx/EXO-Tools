/**
 * Configuration for EXO Tools web application.
 * All settings are loaded from environment variables.
 */

require("dotenv").config();

const config = {
  // Express
  SECRET_KEY: process.env.SECRET_KEY || require("crypto").randomBytes(24).toString("hex"),
  SESSION_COOKIE_HTTPONLY: true,
  SESSION_COOKIE_SAMESITE: "lax",
  // Use secure cookies in production (HTTPS)
  SESSION_COOKIE_SECURE: (process.env.NODE_ENV || "production") === "production",

  // Server-side session storage (session-file-store).
  // Keeps the session cookie small by storing data on the server filesystem
  // rather than encoding it into the cookie itself.
  //
  // On Azure App Service the /home mount is persistent across restarts and
  // shared across all instances, so it is used as the default.  For Docker
  // or local dev you can override via the SESSION_FILE_DIR env var.
  SESSION_FILE_DIR: process.env.SESSION_FILE_DIR || "/home/flask_session",
  SESSION_PERMANENT: false,

  // ----------------------------------------------------------------
  // Azure AD App Registration (user authentication + Graph API)
  // ----------------------------------------------------------------
  CLIENT_ID: process.env.CLIENT_ID || "",
  CLIENT_SECRET: process.env.CLIENT_SECRET || "",
  TENANT_ID: process.env.TENANT_ID || "",
  REDIRECT_URI: process.env.REDIRECT_URI || "http://localhost:5000/callback",
  AUTHORITY: `https://login.microsoftonline.com/${process.env.TENANT_ID || ""}`,

  // Delegated scopes for Graph API (admin consent required for the tenant)
  SCOPES: ["User.Read", "User.ReadBasic.All", "GroupMember.Read.All"],

  // ----------------------------------------------------------------
  // Access Control
  // ----------------------------------------------------------------
  // Object ID of the Entra ID security group allowed to use this app.
  // Leave empty to allow all authenticated users in the tenant.
  ACCESS_GROUP_ID: process.env.ACCESS_GROUP_ID || "",

  // ----------------------------------------------------------------
  // Exchange Online App Registration (PowerShell app-only auth)
  // ----------------------------------------------------------------
  EXO_APP_ID: process.env.EXO_APP_ID || "",
  // Path to the PFX certificate file used for EXO app-only authentication
  EXO_CERT_PATH: process.env.EXO_CERT_PATH || "/app/certs/exo.pfx",
  // Password for the PFX certificate (leave empty if the cert has no password)
  EXO_CERT_PASSWORD: process.env.EXO_CERT_PASSWORD || "",
  // Your M365 tenant's onmicrosoft.com domain, e.g. contoso.onmicrosoft.com
  EXO_ORGANIZATION: process.env.EXO_ORGANIZATION || "",

  // ----------------------------------------------------------------
  // PowerShell
  // ----------------------------------------------------------------
  // Path to the PowerShell Core executable
  PWSH_PATH: process.env.PWSH_PATH || "pwsh",
  // Maximum seconds to wait for the PowerShell permission-check script
  PS_TIMEOUT: parseInt(process.env.PS_TIMEOUT || "600", 10),

  // ----------------------------------------------------------------
  // Microsoft Graph
  // ----------------------------------------------------------------
  GRAPH_API_BASE: "https://graph.microsoft.com/v1.0",
};

module.exports = config;
