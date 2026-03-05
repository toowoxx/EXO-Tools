# ============================================================
# EXO Tools – Dockerfile
# Builds a container with Python 3.12, PowerShell Core, and
# the ExchangeOnlineManagement module pre-installed.
# ============================================================

FROM python:3.12-slim

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
# Python application
# ---------------------------------------------------------------------------
WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Directory for the EXO certificate (mount at runtime)
RUN mkdir -p /app/certs

# Non-root user for security
RUN useradd -m -r appuser && chown -R appuser /app
USER appuser

EXPOSE 5000

# gunicorn: 2 workers, 10-minute timeout for long-running PS scripts
CMD ["gunicorn", \
     "--bind", "0.0.0.0:5000", \
     "--workers", "2", \
     "--timeout", "660", \
     "--access-logfile", "-", \
     "--error-logfile", "-", \
     "app:app"]
