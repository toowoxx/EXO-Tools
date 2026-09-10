# ============================================================
# EXO Tools – Dockerfile
# Builds a container with Node.js 22, PowerShell Core, and
# the ExchangeOnlineManagement module pre-installed.
# ============================================================

FROM node:22-slim

ARG GIT_COMMIT=unknown
ARG BUILD_TIME=unknown
ARG POWERSHELL_DEB_SHA256=79642721f0bc9baf07dafaab68ece1cbd822f86722492acf9b4031d41029a735
ARG EXO_MODULE_SHA256=40d5d1d2c926c7a1318e9070492c9c462a134d7465cc8330756e23ac72a0ea6b

ENV GIT_COMMIT=${GIT_COMMIT} \
    BUILD_TIME=${BUILD_TIME}

# ---------------------------------------------------------------------------
# System dependencies + PowerShell Core
# ---------------------------------------------------------------------------
RUN apt-get update && apt-get install -y --no-install-recommends \
        curl \
        ca-certificates \
        unzip \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

ADD https://github.com/PowerShell/PowerShell/releases/download/v7.4.6/powershell_7.4.6-1.deb_amd64.deb /tmp/powershell.deb

# ---------------------------------------------------------------------------
# Pre-install ExchangeOnlineManagement PowerShell module
# ---------------------------------------------------------------------------
ADD https://www.powershellgallery.com/api/v2/package/ExchangeOnlineManagement/3.0.0 /tmp/ExchangeOnlineManagement.3.0.0.nupkg

RUN echo "${POWERSHELL_DEB_SHA256}  /tmp/powershell.deb" | sha256sum -c - \
    && echo "${EXO_MODULE_SHA256}  /tmp/ExchangeOnlineManagement.3.0.0.nupkg" | sha256sum -c - \
    && dpkg-deb --info /tmp/powershell.deb >/dev/null \
    && unzip -tq /tmp/ExchangeOnlineManagement.3.0.0.nupkg >/dev/null \
    && apt-get update \
    && apt-get install -y --no-install-recommends /tmp/powershell.deb \
    && mkdir -p /tmp/exo-module \
    && mkdir -p /usr/local/share/powershell/Modules/ExchangeOnlineManagement/3.0.0 \
    && unzip -q /tmp/ExchangeOnlineManagement.3.0.0.nupkg -d /tmp/exo-module \
    && cp -R /tmp/exo-module/. /usr/local/share/powershell/Modules/ExchangeOnlineManagement/3.0.0/ \
    && pwsh -NonInteractive -NoProfile -Command "Import-Module ExchangeOnlineManagement; exit 0" \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/* /tmp/powershell.deb /tmp/ExchangeOnlineManagement.3.0.0.nupkg /tmp/exo-module

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
