# ============================================================
# EXO Tools – Dockerfile
# Builds a container with Node.js 22, PowerShell Core, and
# the ExchangeOnlineManagement module pre-installed.
# ============================================================

FROM node:22-slim

# ---------------------------------------------------------------------------
# System dependencies + PowerShell Core
# ---------------------------------------------------------------------------
RUN apt-get update && apt-get install -y --no-install-recommends \
        wget \
        curl \
        apt-transport-https \
        software-properties-common \
        ca-certificates \
    && wget -q https://packages.microsoft.com/config/debian/12/packages-microsoft-prod.deb \
         -O /tmp/packages-microsoft-prod.deb \
    && dpkg -i /tmp/packages-microsoft-prod.deb \
    && rm /tmp/packages-microsoft-prod.deb \
    && apt-get update \
    && apt-get install -y --no-install-recommends powershell \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# ---------------------------------------------------------------------------
# Pre-install ExchangeOnlineManagement PowerShell module
# ---------------------------------------------------------------------------
RUN pwsh -NonInteractive -NoProfile -Command \
    "Install-Module -Name ExchangeOnlineManagement -MinimumVersion '3.0.0' -Force -AllowClobber -Scope AllUsers -Repository PSGallery"

# ---------------------------------------------------------------------------
# Node.js application
# ---------------------------------------------------------------------------
WORKDIR /app

COPY package.json package-lock.json ./
RUN npm ci --omit=dev

COPY . .

# Directory for the EXO certificate (mount at runtime)
RUN mkdir -p /app/certs

# Non-root user for security
RUN useradd -m -r appuser && chown -R appuser /app
USER appuser

EXPOSE 5000

CMD ["node", "app.js"]
