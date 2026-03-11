#!/usr/bin/env bash
# ===========================================================================
# startup.sh – Azure App Service (Linux) startup script
#
# Installs PowerShell Core and the ExchangeOnlineManagement module on first
# run. Both are cached in /home to survive application restarts.
# Azure App Service free/shared tiers persist /home across restarts.
#
# Set this as the startup command in App Service:
#   bash /home/site/wwwroot/startup.sh
# ===========================================================================

set -euo pipefail

SITE_ROOT="/home/site/wwwroot"
PWSH_DIR="${SITE_ROOT}/.pwsh"
PS_MODULE_DIR="${SITE_ROOT}/.ps-modules"
PWSH_BIN="${PWSH_DIR}/pwsh"
# Allow the PS version to be overridden via an environment variable
PWSH_VERSION="${PWSH_VERSION:-7.4.6}"

echo "=== EXO Tools startup ==="

# ---------------------------------------------------------------------------
# 1. Install PowerShell Core if not already present
# ---------------------------------------------------------------------------
if [ ! -x "${PWSH_BIN}" ]; then
    echo "[startup] Installing PowerShell Core ${PWSH_VERSION}..."
    mkdir -p "${PWSH_DIR}"
    wget -q \
        "https://github.com/PowerShell/PowerShell/releases/download/v${PWSH_VERSION}/powershell-${PWSH_VERSION}-linux-x64.tar.gz" \
        -O /tmp/pwsh.tar.gz
    tar -xzf /tmp/pwsh.tar.gz -C "${PWSH_DIR}"
    chmod +x "${PWSH_BIN}"
    rm -f /tmp/pwsh.tar.gz
    echo "[startup] PowerShell Core installed at ${PWSH_BIN}"
fi

# ---------------------------------------------------------------------------
# 2. Install ExchangeOnlineManagement module if not present
# ---------------------------------------------------------------------------
if [ ! -d "${PS_MODULE_DIR}/ExchangeOnlineManagement" ]; then
    echo "[startup] Installing ExchangeOnlineManagement PowerShell module..."
    mkdir -p "${PS_MODULE_DIR}"
    "${PWSH_BIN}" -NonInteractive -NoProfile -Command "
        \$env:PSModulePath = '${PS_MODULE_DIR}:' + \$env:PSModulePath
        Install-Module -Name ExchangeOnlineManagement \
            -MinimumVersion '3.0.0' \
            -Force -AllowClobber \
            -Scope CurrentUser \
            -Repository PSGallery \
            -ErrorAction Stop
        Write-Host 'ExchangeOnlineManagement installed.'
    "
    echo "[startup] ExchangeOnlineManagement module installed."
fi

# ---------------------------------------------------------------------------
# 3. Ensure the Flask server-side session directory exists
# ---------------------------------------------------------------------------
SESSION_DIR="${SESSION_FILE_DIR:-/home/flask_session}"
mkdir -p "${SESSION_DIR}"
chmod 700 "${SESSION_DIR}"
echo "[startup] SESSION_FILE_DIR=${SESSION_DIR}"

# ---------------------------------------------------------------------------
# 4. Export PWSH_PATH so the Flask app can find it
# ---------------------------------------------------------------------------
export PWSH_PATH="${PWSH_BIN}"
echo "[startup] PWSH_PATH=${PWSH_PATH}"

# ---------------------------------------------------------------------------
# 5. Start the Flask application with gunicorn
# ---------------------------------------------------------------------------
cd "${SITE_ROOT}"
echo "[startup] Starting gunicorn on port ${PORT:-8000}..."
exec gunicorn \
    --bind "0.0.0.0:${PORT:-8000}" \
    --workers 2 \
    --timeout 660 \
    --access-logfile - \
    --error-logfile - \
    app:app
