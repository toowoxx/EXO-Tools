# EXO Tools – Reverse Mailbox Permission Search

A lightweight web application for Microsoft 365 administrators to quickly see **which mailboxes a specific user has delegated access to** in Exchange Online — the reverse of what the EAC shows ou[...]  

> [!WARNING]
> ## ⚠️ AI-Generated Code — Security Notice
>
> This project was **primarily created using AI-assisted development (Vibe Coding)** and has been processed through multiple AI agents and automated stages to review for security vulnerabilities.
>
> **We do not provide any guarantee that this code is free of security vulnerabilities.** Use it at your own risk, and ensure you perform your own security review before deploying it in a production or sensitive environment.
>
> **Responsible / confidential reporting of security vulnerabilities is welcome.**
> Please report any discovered vulnerabilities privately — do **not** open a public issue.
> You can use [GitHub's private vulnerability reporting](https://github.com/toowoxx/EXO-Tools/security/advisories/new) to submit a report confidentially.

---

## Features

| Feature | Details |
|---|---|
| **People Picker** | Autocomplete user search via Microsoft Graph API |
| **Full Access** | Finds every mailbox the user can open and manage |
| **Send As** | Finds every mailbox the user can send email as |
| **Send on Behalf** | Finds every mailbox the user can send on behalf of |
| **M365 Login** | Azure AD authentication; access restricted to a configurable Entra ID group |
| **Filter & Search** | Filter results by permission type or mailbox name |
| **Export CSV** | One-click export of all results |
| **Azure-ready** | Runs on Azure App Service (free tier) or in Docker |

---

## Architecture

```
Browser
  │
  ├─► Express (Node.js) ──► Microsoft Graph API   (user search / autocomplete)
  │         │
  │         └──► PowerShell (ExchangeOnlineManagement) ──► Exchange Online
  │
  └─ Azure AD (MSAL) ──► Entra ID group check ──► allow / deny
```

Two Azure AD app registrations are required:

| Registration | Purpose | Auth type |
|---|---|---|
| **Web App** | User login + Graph API (People Picker) | Delegated (user signs in) |
| **EXO App** | PowerShell app-only access to Exchange Online | Application (certificate) |

---

## Prerequisites

- Node.js 18+ (for local dev or Azure App Service)  
  *or* Docker (for containerised deployment)
- PowerShell Core 7+ (`pwsh`) with the `ExchangeOnlineManagement` module ≥ 3.0.0
- Two Azure AD app registrations (see below)
- An EXO certificate (PFX) for the Exchange Online app registration

---

## Setup Guide

### Step 1 – Web App Registration (user authentication)

1. **Azure Portal → Azure Active Directory → App registrations → New registration**
2. Name it, e.g. `EXO Tools`.
3. Supported account type: **Accounts in this organizational directory only (Single tenant)**.
4. Redirect URI → **Web** → `https://<your-app>.azurewebsites.net/callback`  
   (add `http://localhost:5000/callback` for local development as well)
5. Note the **Application (client) ID** and **Directory (tenant) ID**.
6. **Certificates & secrets → New client secret** → copy the value immediately.
7. **API permissions → Add a permission → Microsoft Graph → Delegated:**
   - `User.Read`
   - `User.ReadBasic.All`
   - `GroupMember.Read.All`  *(only needed if you configure `ACCESS_GROUP_ID`)*
8. **Grant admin consent** for the tenant.
9. *(Optional)* **Token configuration → Add groups claim → Security groups** – adds the `groups` claim to the ID token for a faster group membership check.

### Step 2 – EXO App Registration (PowerShell app-only)

1. **New registration** → name it, e.g. `EXO Tools – EXO Access`.
2. Note the **Application (client) ID**.
3. **API permissions → Add a permission → APIs my organization uses →**  
   Search `Office 365 Exchange Online` → **Application permissions** →  
   `Exchange.ManageAsApp` → Add.
4. **Grant admin consent**.
5. Assign the **Exchange Administrator** role to the app in Exchange Online:

   ```powershell
   # Connect as a Global / Exchange Administrator first
   Connect-ExchangeOnline -UserPrincipalName admin@yourtenant.onmicrosoft.com

   New-ManagementRoleAssignment `
       -App "<EXO App client ID>" `
       -Role "View-Only Organization Management"
   # For full read access you may also add:
   # -Role "Mailbox Search"
   ```

   > **Minimum roles needed:** `View-Only Organization Management` (for `Get-Mailbox`, `Get-User`) and `View-Only Recipients` (for `Get-RecipientPermission`).

### Step 3 – Generate the EXO Certificate

```powershell
# Run in PowerShell on your local machine

# Generate a 2-year self-signed certificate
$cert = New-SelfSignedCertificate `
    -Subject "CN=EXOTools" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -KeyAlgorithm RSA `
    -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddYears(2)

# Export the PUBLIC key (.cer) – upload this to the Azure AD app registration
Export-Certificate -Cert $cert -FilePath "exo-cert.cer"

# Export the PRIVATE key (.pfx) – place this on the web server, never commit it
$pwd = ConvertTo-SecureString -String "YourCertPassword" -Force -AsPlainText
Export-PfxCertificate -Cert $cert -FilePath "exo-cert.pfx" -Password $pwd
```

Upload `exo-cert.cer` to the EXO app registration under **Certificates & secrets → Certificates**.  
Keep `exo-cert.pfx` secure – it goes on the server only.

### Step 4 – Configure Environment Variables

```bash
cp .env.example .env
# Edit .env and fill in all values
```

| Variable | Description |
|---|---|
| `SECRET_KEY` | Random string for session signing |
| `CLIENT_ID` | Web App registration client ID |
| `CLIENT_SECRET` | Web App client secret |
| `TENANT_ID` | Azure AD tenant ID |
| `REDIRECT_URI` | OAuth redirect URI (must match app registration) |
| `ACCESS_GROUP_ID` | Entra ID group Object ID (leave empty for all users) |
| `EXO_APP_ID` | EXO app registration client ID |
| `EXO_CERT_PATH` | Absolute path to the PFX file |
| `EXO_CERT_PASSWORD` | PFX password (empty if none) |
| `EXO_ORGANIZATION` | `yourtenant.onmicrosoft.com` |
| `PWSH_PATH` | Path to `pwsh` executable (default: `pwsh`) |
| `PS_TIMEOUT` | Seconds before the PS script times out (default: `600`) |

---

## Deployment

### Option A – Docker (recommended for Azure Container Apps / self-hosted)

```bash
# 1. Build and start
cp .env.example .env        # fill in your values
mkdir -p certs
cp /path/to/exo-cert.pfx certs/exo.pfx

docker compose up -d

# 2. Open http://localhost:5000
```

### Option B – Azure App Service (Free / Basic tier, Linux)

```bash
# Install Azure CLI if needed: https://aka.ms/install-az-cli

az login

# Create resource group and App Service plan (free tier)
az group create --name rg-exotools --location westeurope
az appservice plan create \
    --name asp-exotools \
    --resource-group rg-exotools \
    --sku F1 \
    --is-linux

# Create the web app (Node.js 22)
az webapp create \
    --name exo-tools \
    --resource-group rg-exotools \
    --plan asp-exotools \
    --runtime "NODE:22-lts"

# Set startup command
az webapp config set \
    --name exo-tools \
    --resource-group rg-exotools \
    --startup-file "bash startup.sh"

# Upload environment variables (repeat for each variable)
az webapp config appsettings set \
    --name exo-tools \
    --resource-group rg-exotools \
    --settings \
        SECRET_KEY="..." \
        CLIENT_ID="..." \
        CLIENT_SECRET="..." \
        TENANT_ID="..." \
        REDIRECT_URI="https://exo-tools.azurewebsites.net/callback" \
        EXO_APP_ID="..." \
        EXO_ORGANIZATION="yourtenant.onmicrosoft.com" \
        EXO_CERT_PATH="/home/site/wwwroot/certs/exo.pfx" \
        EXO_CERT_PASSWORD="..."

# Deploy code
az webapp deploy \
    --name exo-tools \
    --resource-group rg-exotools \
    --src-path . \
    --type zip

# Upload the certificate via Kudu (replace values)
curl -X PUT \
    "https://exo-tools.scm.azurewebsites.net/api/vfs/site/wwwroot/certs/exo.pfx" \
    -u "<deployment-user>:<password>" \
    --upload-file /path/to/exo-cert.pfx
```

> **Tip:** For production, store the certificate in **Azure Key Vault** and retrieve it at startup, instead of uploading the PFX directly.

### Option C – Local Development

```bash
# Requires Node.js 18+ and PowerShell Core (pwsh) on PATH

npm install

cp .env.example .env             # fill in values
mkdir -p certs
cp /path/to/exo-cert.pfx certs/exo.pfx

# Install the EXO PowerShell module (once)
pwsh -Command "Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force"

NODE_ENV=development node app.js
# Open http://localhost:5000
```

---

## Usage

1. Navigate to the web app URL.
2. Click **Sign in with Microsoft 365** and authenticate with your admin account.
3. In the **User** field, start typing a name or email address.
4. Select the user from the autocomplete suggestions.
5. Click **Check Permissions**.
6. The app queries Exchange Online in the background (this may take several minutes on large tenants).
7. Results appear in a table grouped by permission type.
8. Use the **type filter buttons** or the **search box** to narrow results.
9. Click **Export CSV** to download the full results.

---

## Performance Notes

| Operation | Complexity | Notes |
|---|---|---|
| **Send As** | O(1) – single EXO query | Fast; uses `Get-RecipientPermission -Trustee` |
| **Full Access** | O(n mailboxes) | Iterates every mailbox; slow on large tenants |
| **Send on Behalf** | O(n mailboxes) | Iterates every mailbox to check `GrantSendOnBehalfTo` |

For tenants with **> 500 mailboxes**, the check typically takes **5–15 minutes**.  
Increase `PS_TIMEOUT` (default 600 s) if you see timeouts.

---

## Security

- All traffic should be served over HTTPS (Azure App Service enforces this).
- The EXO certificate private key (PFX) never leaves the server.
- The `userPrincipalName` input is validated against a strict regex before being passed to PowerShell.
- Sessions are signed (HMAC) and marked `HttpOnly` + `SameSite=Lax`.
- Set `SESSION_COOKIE_SECURE=True` (the default when `NODE_ENV=production`) to enforce HTTPS-only cookies.
- Restrict app access to authorised administrators via `ACCESS_GROUP_ID`.

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| "PowerShell not found" | `pwsh` not on PATH | Install PowerShell Core or set `PWSH_PATH` |
| "Failed to connect to Exchange Online" | Wrong cert / app ID | Check `EXO_APP_ID`, `EXO_CERT_PATH`, `EXO_ORGANIZATION` |
| "Access denied. Your account is not in the required group." | User not in `ACCESS_GROUP_ID` group | Add user to the group or clear `ACCESS_GROUP_ID` |
| "Could not acquire token" | Wrong `CLIENT_ID` / `CLIENT_SECRET` | Verify the app registration values |
| Timeout after 10 minutes | Very large tenant | Increase `PS_TIMEOUT`; consider off-hours scheduling |
| No users found in People Picker | Missing admin consent | Grant admin consent for `User.ReadBasic.All` |