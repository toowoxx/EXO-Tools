<#
.SYNOPSIS
    Retrieves all mailbox permissions granted to a specific user in Exchange Online.

.DESCRIPTION
    This script connects to Exchange Online using app-only authentication (certificate)
    and finds all mailboxes where the specified user has been delegated access via:
      - Full Access (open / manage mailbox)
      - Send As (send email as the mailbox owner)
      - Send on Behalf (send email on behalf of the mailbox owner)

    Results are written to stdout as a JSON array for consumption by the Flask web app.

.PARAMETER UserPrincipalName
    The UPN of the user whose mailbox permissions should be enumerated.

.PARAMETER AppId
    The Application (client) ID of the Azure AD app registration used for EXO app-only auth.

.PARAMETER CertificatePath
    Path to the PFX certificate file for app-only authentication.

.PARAMETER CertificatePassword
    Password for the PFX certificate (optional).

.PARAMETER Organization
    The primary domain of the M365 tenant, e.g. contoso.onmicrosoft.com

.NOTES
    Requires ExchangeOnlineManagement module v3.0.0 or later.
    The app registration must have the Exchange.ManageAsApp permission granted
    and the app must be assigned the "Exchange Administrator" (or equivalent) role.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $true)]
    [string]$AppId,

    [Parameter(Mandatory = $true)]
    [string]$CertificatePath,

    [Parameter(Mandatory = $false)]
    [string]$CertificatePassword = "",

    [Parameter(Mandatory = $true)]
    [string]$Organization
)

$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# Ensure the ExchangeOnlineManagement module is available
# ---------------------------------------------------------------------------
$minVersion = [Version]"3.0.0"
$module = Get-Module -ListAvailable -Name ExchangeOnlineManagement |
    Sort-Object Version -Descending |
    Select-Object -First 1

if (-not $module -or $module.Version -lt $minVersion) {
    Write-Host "Installing ExchangeOnlineManagement module (minimum v3.0.0)..." -ForegroundColor Yellow
    try {
        Install-Module -Name ExchangeOnlineManagement `
            -MinimumVersion $minVersion.ToString() `
            -Force -AllowClobber -Scope CurrentUser -Repository PSGallery
    } catch {
        Write-Error "Failed to install ExchangeOnlineManagement: $_"
        exit 1
    }
}

Import-Module ExchangeOnlineManagement -MinimumVersion $minVersion.ToString() -ErrorAction Stop

# ---------------------------------------------------------------------------
# Connect to Exchange Online (app-only)
# ---------------------------------------------------------------------------
try {
    if ($CertificatePassword) {
        $secPwd = ConvertTo-SecureString -String $CertificatePassword -AsPlainText -Force
        Connect-ExchangeOnline `
            -AppId $AppId `
            -CertificateFilePath $CertificatePath `
            -CertificatePassword $secPwd `
            -Organization $Organization `
            -ShowBanner:$false `
            -ErrorAction Stop
    } else {
        Connect-ExchangeOnline `
            -AppId $AppId `
            -CertificateFilePath $CertificatePath `
            -Organization $Organization `
            -ShowBanner:$false `
            -ErrorAction Stop
    }
} catch {
    Write-Error "Failed to connect to Exchange Online: $_"
    exit 1
}

# ---------------------------------------------------------------------------
# Helper – resolve the target user's identity properties
# ---------------------------------------------------------------------------
try {
    $targetUser = Get-User -Identity $UserPrincipalName -ErrorAction Stop
} catch {
    Write-Error "User '$UserPrincipalName' not found in Exchange Online: $_"
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    exit 1
}

$targetName        = $targetUser.Name
$targetDisplayName = $targetUser.DisplayName
$targetDN          = $targetUser.DistinguishedName

$results = [System.Collections.Generic.List[hashtable]]::new()

# ---------------------------------------------------------------------------
# Retrieve all mailboxes once (reused for all three permission types)
# ---------------------------------------------------------------------------
Write-Host "Retrieving all mailboxes..." -ForegroundColor Cyan
try {
    $allMailboxes = Get-Mailbox -ResultSize Unlimited -ErrorAction Stop
} catch {
    Write-Error "Failed to retrieve mailboxes: $_"
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    exit 1
}
Write-Host "Checking permissions across $($allMailboxes.Count) mailbox(es)..." -ForegroundColor Cyan

# ---------------------------------------------------------------------------
# 1. Full Access
# ---------------------------------------------------------------------------
Write-Host "Checking Full Access permissions..." -ForegroundColor Yellow
$fullAccessCount = 0

foreach ($mbx in $allMailboxes) {
    try {
        $perms = Get-MailboxPermission -Identity $mbx.Identity -User $UserPrincipalName -ErrorAction SilentlyContinue
        if ($perms) {
            foreach ($p in $perms) {
                if (($p.AccessRights -contains "FullAccess") -and ($p.Deny -eq $false) -and ($p.IsInherited -eq $false)) {
                    $results.Add(@{
                        MailboxDisplayName = $mbx.DisplayName
                        MailboxUPN         = $mbx.UserPrincipalName
                        MailboxType        = $mbx.RecipientTypeDetails.ToString()
                        PermissionType     = "FullAccess"
                        AccessRights       = ($p.AccessRights -join ", ")
                        IsInherited        = $false
                    })
                    $fullAccessCount++
                }
            }
        }
    } catch {
        # Skip individual mailbox errors silently
    }
}
Write-Host "  Full Access: $fullAccessCount permission(s) found." -ForegroundColor Green

# ---------------------------------------------------------------------------
# 2. Send As  (Get-RecipientPermission is efficient – single query by trustee)
# ---------------------------------------------------------------------------
Write-Host "Checking Send As permissions..." -ForegroundColor Yellow
$sendAsCount = 0

try {
    $sendAsPerms = Get-RecipientPermission -Trustee $UserPrincipalName -ResultSize Unlimited -ErrorAction Stop
    foreach ($p in $sendAsPerms) {
        if ($p.AccessControlType -eq "Allow") {
            # Look up the matching mailbox for display name / type
            $mbxInfo = $allMailboxes | Where-Object { $_.Identity -eq $p.Identity } | Select-Object -First 1
            $results.Add(@{
                MailboxDisplayName = if ($mbxInfo) { $mbxInfo.DisplayName } else { $p.Identity }
                MailboxUPN         = if ($mbxInfo) { $mbxInfo.UserPrincipalName } else { $p.Identity }
                MailboxType        = if ($mbxInfo) { $mbxInfo.RecipientTypeDetails.ToString() } else { "Unknown" }
                PermissionType     = "SendAs"
                AccessRights       = "SendAs"
                IsInherited        = $false
            })
            $sendAsCount++
        }
    }
} catch {
    Write-Warning "Could not retrieve Send As permissions: $_"
}
Write-Host "  Send As: $sendAsCount permission(s) found." -ForegroundColor Green

# ---------------------------------------------------------------------------
# 3. Send on Behalf  (stored in GrantSendOnBehalfTo on each mailbox)
# ---------------------------------------------------------------------------
Write-Host "Checking Send on Behalf permissions..." -ForegroundColor Yellow
$sobCount = 0

# Build a set of possible identifiers for the target user
$identifiers = @(
    $UserPrincipalName,
    $targetName,
    $targetDisplayName,
    $targetDN
) | Where-Object { $_ } | ForEach-Object { $_.ToLower() }

foreach ($mbx in $allMailboxes) {
    if (-not $mbx.GrantSendOnBehalfTo -or $mbx.GrantSendOnBehalfTo.Count -eq 0) { continue }

    $isDelegate = $false
    foreach ($entry in $mbx.GrantSendOnBehalfTo) {
        $entryStr = $entry.ToString().ToLower()
        if ($identifiers | Where-Object { $entryStr -eq $_ -or $entryStr -like "*$_*" }) {
            $isDelegate = $true
            break
        }
    }

    if ($isDelegate) {
        $results.Add(@{
            MailboxDisplayName = $mbx.DisplayName
            MailboxUPN         = $mbx.UserPrincipalName
            MailboxType        = $mbx.RecipientTypeDetails.ToString()
            PermissionType     = "SendOnBehalf"
            AccessRights       = "SendOnBehalf"
            IsInherited        = $false
        })
        $sobCount++
    }
}
Write-Host "  Send on Behalf: $sobCount permission(s) found." -ForegroundColor Green

Write-Host "Total permissions found: $($results.Count)" -ForegroundColor Cyan

# ---------------------------------------------------------------------------
# Disconnect and output JSON
# ---------------------------------------------------------------------------
try {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
} catch { }

# Emit results as JSON (the Flask app parses the last JSON block in stdout)
$results | ConvertTo-Json -Depth 3 -Compress
