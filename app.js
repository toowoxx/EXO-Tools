/**
 * EXO Tools – Reverse Mailbox Permission Search
 * Express web application entry point.
 */

"use strict";

const path = require("path");
const fs = require("fs");
const { URL } = require("url");
const { execFile } = require("child_process");
const crypto = require("crypto");

const express = require("express");
const session = require("express-session");
const FileStore = require("session-file-store")(session);
const msal = require("@azure/msal-node");
const axios = require("axios");

const config = require("./config");

// ---------------------------------------------------------------------------
// Logging
// ---------------------------------------------------------------------------
const logger = {
  info: (...args) => console.log(new Date().toISOString(), "INFO", ...args),
  warn: (...args) => console.warn(new Date().toISOString(), "WARN", ...args),
  error: (...args) => console.error(new Date().toISOString(), "ERROR", ...args),
};

// ---------------------------------------------------------------------------
// App setup
// ---------------------------------------------------------------------------
const app = express();
app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));

// Trust the first proxy (for Azure App Service / reverse proxy setups)
app.set("trust proxy", 1);

// Parse JSON request bodies
app.use(express.json());

// Serve static files
app.use("/static", express.static(path.join(__dirname, "static")));

// ---------------------------------------------------------------------------
// Session directory setup
// ---------------------------------------------------------------------------

/**
 * Return the first writable candidate directory, creating it if needed.
 */
function findSessionDir(preferred) {
  for (const candidate of [preferred, "/tmp/flask_session"]) {
    try {
      fs.mkdirSync(candidate, { recursive: true });
      const probe = path.join(candidate, ".write_probe");
      fs.writeFileSync(probe, "");
      fs.unlinkSync(probe);
      return candidate;
    } catch {
      logger.warn(`Session directory not writable: ${candidate} – trying next`);
    }
  }
  throw new Error(
    `No writable session directory found. Tried: ${preferred} and /tmp/flask_session. ` +
      "Set SESSION_FILE_DIR to a writable path."
  );
}

const sessionDir = findSessionDir(config.SESSION_FILE_DIR);
logger.info(`Using session directory: ${sessionDir}`);

// ---------------------------------------------------------------------------
// Session middleware
// ---------------------------------------------------------------------------
app.use(
  session({
    store: new FileStore({
      path: sessionDir,
      ttl: 86400,
      retries: 0,
      fileExtension: ".json",
    }),
    secret: config.SECRET_KEY,
    resave: false,
    saveUninitialized: false,
    cookie: {
      httpOnly: config.SESSION_COOKIE_HTTPONLY,
      sameSite: config.SESSION_COOKIE_SAMESITE,
      secure: config.SESSION_COOKIE_SECURE,
    },
  })
);

// ---------------------------------------------------------------------------
// MSAL token cache (filesystem-backed)
// ---------------------------------------------------------------------------

// Separate directory for MSAL token caches to keep session cookies small.
const msalCacheDir = path.join(sessionDir, "msal_tokens");
fs.mkdirSync(msalCacheDir, { recursive: true });

/** How long MSAL token caches persist on disk (ms). */
const MSAL_CACHE_TTL = 86400 * 1000; // 24 hours

function getMsalCachePath(oid) {
  // Sanitise OID to prevent path traversal
  const safeOid = oid.replace(/[^a-zA-Z0-9-]/g, "");
  return path.join(msalCacheDir, `msal_${safeOid}.json`);
}

function loadMsalCache(oid) {
  if (!oid) return null;
  const filePath = getMsalCachePath(oid);
  try {
    if (fs.existsSync(filePath)) {
      const stat = fs.statSync(filePath);
      // Check TTL
      if (Date.now() - stat.mtimeMs > MSAL_CACHE_TTL) {
        fs.unlinkSync(filePath);
        return null;
      }
      return fs.readFileSync(filePath, "utf8");
    }
  } catch (err) {
    logger.warn(`Failed to load MSAL cache for ${oid}: ${err.message}`);
  }
  return null;
}

function saveMsalCache(oid, cacheData) {
  if (!oid || !cacheData) return;
  try {
    // Strip large short-lived token blobs. MSAL can silently re-acquire
    // access tokens using the refresh token, and the ID token is already
    // decoded into session user.
    const parsed = JSON.parse(cacheData);
    delete parsed.AccessToken;
    delete parsed.IdToken;
    fs.writeFileSync(getMsalCachePath(oid), JSON.stringify(parsed), {
      mode: 0o600,
    });
  } catch (err) {
    logger.warn(`Failed to save MSAL cache for ${oid}: ${err.message}`);
  }
}

function deleteMsalCache(oid) {
  if (!oid) return;
  try {
    const filePath = getMsalCachePath(oid);
    if (fs.existsSync(filePath)) {
      fs.unlinkSync(filePath);
    }
  } catch (err) {
    logger.warn(`Failed to delete MSAL cache for ${oid}: ${err.message}`);
  }
}

// ---------------------------------------------------------------------------
// MSAL helpers
// ---------------------------------------------------------------------------

function buildMsalApp(cachePlugin) {
  const msalConfig = {
    auth: {
      clientId: config.CLIENT_ID,
      authority: config.AUTHORITY,
      clientSecret: config.CLIENT_SECRET,
    },
  };
  if (cachePlugin) {
    msalConfig.cache = { cachePlugin };
  }
  return new msal.ConfidentialClientApplication(msalConfig);
}

/**
 * Create a cache plugin for MSAL that reads/writes to the filesystem.
 */
function createCachePlugin(oid) {
  if (!oid) return null;

  const beforeCacheAccess = async (cacheContext) => {
    const data = loadMsalCache(oid);
    if (data) {
      cacheContext.tokenCache.deserialize(data);
    }
  };

  const afterCacheAccess = async (cacheContext) => {
    if (cacheContext.cacheHasChanged) {
      saveMsalCache(oid, cacheContext.tokenCache.serialize());
    }
  };

  return { beforeCacheAccess, afterCacheAccess };
}

async function getTokenFromCache(req) {
  const oid = req.session.user && req.session.user.oid;
  if (!oid) return null;

  const cachePlugin = createCachePlugin(oid);
  const cca = buildMsalApp(cachePlugin);

  const accounts = await cca.getTokenCache().getAllAccounts();
  if (accounts.length > 0) {
    try {
      const result = await cca.acquireTokenSilent({
        scopes: config.SCOPES,
        account: accounts[0],
      });
      return result;
    } catch (err) {
      logger.warn(`Token refresh failed: ${err.message}`);
      return null;
    }
  }
  return null;
}

// ---------------------------------------------------------------------------
// Access-control helpers
// ---------------------------------------------------------------------------

async function checkGroupMembership(accessToken) {
  const groupId = config.ACCESS_GROUP_ID;
  if (!groupId) return true; // no restriction configured

  // Graph API check
  try {
    const resp = await axios.post(
      `${config.GRAPH_API_BASE}/me/checkMemberObjects`,
      { ids: [groupId] },
      {
        headers: { Authorization: `Bearer ${accessToken}` },
        timeout: 10000,
      }
    );
    if (resp.status === 200) {
      return (resp.data.value || []).includes(groupId);
    }
  } catch (err) {
    logger.warn(`Group membership check failed: ${err.message}`);
  }
  return false;
}

// ---------------------------------------------------------------------------
// Auth middleware
// ---------------------------------------------------------------------------

function loginRequired(req, res, next) {
  if (!req.session.user) {
    req.session.next = req.originalUrl;
    return res.redirect("/login");
  }
  next();
}

// ---------------------------------------------------------------------------
// Warn at startup if REDIRECT_URI is misconfigured
// ---------------------------------------------------------------------------
try {
  const redirectPath = new URL(config.REDIRECT_URI).pathname.replace(/\/+$/, "");
  if (!redirectPath.endsWith("/callback")) {
    logger.warn(
      `REDIRECT_URI (${config.REDIRECT_URI}) does not end with /callback. ` +
        "Authentication will fail. Set REDIRECT_URI to " +
        "https://<your-app>.azurewebsites.net/callback"
    );
  }
} catch {
  logger.warn(`REDIRECT_URI is not a valid URL: ${config.REDIRECT_URI}`);
}

// ---------------------------------------------------------------------------
// Routes – authentication
// ---------------------------------------------------------------------------

app.get("/", (req, res) => {
  if (!req.session.user) {
    return res.redirect("/login");
  }
  res.render("dashboard", { title: "EXO Tools – Mailbox Permission Search", user: req.session.user });
});

app.get("/login", async (req, res) => {
  if (req.session.user) {
    return res.redirect("/");
  }

  const state = crypto.randomUUID();
  req.session.state = state;

  let authUrl = null;
  let error = null;
  try {
    const cca = buildMsalApp();
    const authCodeUrlParams = {
      scopes: config.SCOPES,
      redirectUri: config.REDIRECT_URI,
      state,
    };
    authUrl = await cca.getAuthCodeUrl(authCodeUrlParams);
  } catch (err) {
    logger.error(`Failed to build auth URL: ${err.message}`);
    error =
      "Could not reach the Microsoft login service. Please check the application configuration.";
  }

  res.render("login", { title: "EXO Tools – Sign In", authUrl, error });
});

app.get("/callback", async (req, res) => {
  // CSRF guard
  if (req.query.state !== req.session.state) {
    return res
      .status(400)
      .render("login", { title: "EXO Tools – Sign In",
        authUrl: null,
        error: "State mismatch. Please try again.",
      });
  }

  if (req.query.error) {
    const desc = req.query.error_description || "Unknown error";
    return res
      .status(400)
      .render("login", { title: "EXO Tools – Sign In",
        authUrl: null,
        error: `Authentication failed: ${desc}`,
      });
  }

  try {
    const cca = buildMsalApp();
    const tokenRequest = {
      code: req.query.code,
      scopes: config.SCOPES,
      redirectUri: config.REDIRECT_URI,
    };
    const result = await cca.acquireTokenByCode(tokenRequest);

    // Group membership check
    if (config.ACCESS_GROUP_ID) {
      const isMember = await checkGroupMembership(result.accessToken);
      if (!isMember) {
        return res.status(403).render("login", { title: "EXO Tools – Sign In",
          authUrl: null,
          error: "Access denied. Your account is not in the required group.",
        });
      }
    }

    // Extract identity claims from the account object
    const account = result.account;
    if (!account || !account.localAccountId) {
      logger.error(
        "Azure AD token is missing the account identifier – cannot establish session"
      );
      return res.status(400).render("login", { title: "EXO Tools – Sign In",
        authUrl: null,
        error:
          "Authentication failed: the identity token is missing required claims.",
      });
    }

    const userOid = account.localAccountId;

    req.session.user = {
      name: account.name || "",
      preferred_username: account.username || "",
      oid: userOid,
      tid: account.tenantId || "",
      groups: [],
    };

    // Persist the MSAL token cache to the filesystem (not the session).
    const cacheData = cca.getTokenCache().serialize();
    saveMsalCache(userOid, cacheData);

    const nextUrl = req.session.next || "/";
    delete req.session.next;
    res.redirect(nextUrl);
  } catch (err) {
    logger.error(`Token acquisition failed: ${err.message}`);
    return res.status(400).render("login", { title: "EXO Tools – Sign In",
      authUrl: null,
      error: `Could not acquire token: ${err.message}`,
    });
  }
});

app.get("/logout", (req, res) => {
  // Remove the user's MSAL token cache from the filesystem store
  const oid = req.session.user && req.session.user.oid;
  if (oid) {
    deleteMsalCache(oid);
  }

  req.session.destroy(() => {
    const postLogout = `${req.protocol}://${req.get("host")}/`;
    res.redirect(
      `${config.AUTHORITY}/oauth2/v2.0/logout?post_logout_redirect_uri=${encodeURIComponent(postLogout)}`
    );
  });
});

// ---------------------------------------------------------------------------
// Routes – API
// ---------------------------------------------------------------------------

app.get("/api/search-users", loginRequired, async (req, res) => {
  const query = (req.query.q || "").trim();
  if (query.length < 2) {
    return res.json({ users: [] });
  }

  const tokenResult = await getTokenFromCache(req);
  if (!tokenResult) {
    return res
      .status(401)
      .json({ error: "Authentication session expired. Please refresh the page." });
  }

  try {
    const resp = await axios.get(`${config.GRAPH_API_BASE}/users`, {
      headers: {
        Authorization: `Bearer ${tokenResult.accessToken}`,
        ConsistencyLevel: "eventual",
      },
      params: {
        $search: `"displayName:${query}" OR "userPrincipalName:${query}"`,
        $select:
          "id,displayName,userPrincipalName,mail,jobTitle,department",
        $top: "15",
        $count: "true",
      },
      timeout: 10000,
    });

    const users = (resp.data.value || [])
      .filter((u) => u.userPrincipalName)
      .map((u) => ({
        id: u.id,
        displayName: u.displayName || "",
        userPrincipalName: u.userPrincipalName || "",
        mail: u.mail || u.userPrincipalName || "",
        jobTitle: u.jobTitle || "",
        department: u.department || "",
      }));

    return res.json({ users });
  } catch (err) {
    if (err.code === "ECONNABORTED") {
      return res.status(504).json({ error: "Search timed out", users: [] });
    }
    if (err.response) {
      logger.error(
        `Graph API ${err.response.status}: ${JSON.stringify(err.response.data).substring(0, 200)}`
      );
      return res
        .status(err.response.status)
        .json({ error: "Failed to search users", users: [] });
    }
    logger.error(`User search error: ${err.message}`);
    return res.status(500).json({ error: "Internal server error", users: [] });
  }
});

app.post("/api/get-permissions", loginRequired, (req, res) => {
  const upn = ((req.body && req.body.userPrincipalName) || "").trim();

  // Input validation – strict UPN format to prevent injection
  const upnRegex = /^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$/;
  if (!upn || !upnRegex.test(upn)) {
    return res
      .status(400)
      .json({ error: "A valid userPrincipalName is required." });
  }

  // Check EXO configuration
  const cfgErrors = [];
  if (!config.EXO_APP_ID) cfgErrors.push("EXO_APP_ID is not set.");
  if (!config.EXO_ORGANIZATION) cfgErrors.push("EXO_ORGANIZATION is not set.");
  if (!fs.existsSync(config.EXO_CERT_PATH)) {
    cfgErrors.push(`EXO certificate not found at: ${config.EXO_CERT_PATH}`);
  }
  if (cfgErrors.length) {
    return res.status(503).json({
      error: "Exchange Online is not configured: " + cfgErrors.join(" "),
    });
  }

  const scriptPath = path.join(__dirname, "scripts", "Get-MailboxPermissions.ps1");
  if (!fs.existsSync(scriptPath)) {
    return res
      .status(500)
      .json({ error: "PowerShell script not found on server." });
  }

  const args = [
    "-NonInteractive",
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    "-File",
    scriptPath,
    "-UserPrincipalName",
    upn,
    "-AppId",
    config.EXO_APP_ID,
    "-CertificatePath",
    config.EXO_CERT_PATH,
    "-Organization",
    config.EXO_ORGANIZATION,
  ];
  if (config.EXO_CERT_PASSWORD) {
    args.push("-CertificatePassword", config.EXO_CERT_PASSWORD);
  }

  logger.info(`Starting permission check for: ${upn}`);

  const child = execFile(
    config.PWSH_PATH,
    args,
    { timeout: config.PS_TIMEOUT * 1000, maxBuffer: 50 * 1024 * 1024 },
    (err, stdout, stderr) => {
      if (err) {
        if (err.killed || err.signal === "SIGTERM") {
          return res.status(504).json({
            error:
              "The permission check timed out. " +
              "This usually happens on very large tenants. " +
              "Consider increasing PS_TIMEOUT.",
          });
        }
        if (err.code === "ENOENT") {
          return res.status(500).json({
            error: `PowerShell not found at '${config.PWSH_PATH}'. ` +
              "Please install PowerShell Core and set the PWSH_PATH environment variable.",
          });
        }

        const errMsg = (stderr || "").trim() || "PowerShell exited with a non-zero status.";
        logger.error(`PowerShell stderr: ${errMsg.substring(0, 500)}`);
        return res.status(500).json({
          error: `Permission check failed: ${errMsg.substring(0, 300)}`,
        });
      }

      const output = (stdout || "").trim();
      const permissions = parsePsJson(output);
      logger.info(
        `Permission check complete. Found ${permissions.length} permissions.`
      );
      return res.json({
        permissions,
        user: upn,
        count: permissions.length,
      });
    }
  );

  // Handle spawn errors (e.g. ENOENT if pwsh not found)
  child.on("error", (err) => {
    if (err.code === "ENOENT") {
      return res.status(500).json({
        error: `PowerShell not found at '${config.PWSH_PATH}'. ` +
          "Please install PowerShell Core and set the PWSH_PATH environment variable.",
      });
    }
    logger.error(`Unexpected error during permission check: ${err.message}`);
    return res.status(500).json({ error: `Internal server error: ${err.message}` });
  });
});

/**
 * Extract and parse the JSON array/object from PowerShell stdout.
 */
function parsePsJson(output) {
  if (!output) return [];

  // PowerShell may emit log lines before the JSON – find the last JSON block
  for (const [startChar, endChar] of [
    ["[", "]"],
    ["{", "}"],
  ]) {
    const start = output.lastIndexOf(startChar);
    const end = output.lastIndexOf(endChar);
    if (start !== -1 && end > start) {
      try {
        const parsed = JSON.parse(output.substring(start, end + 1));
        return Array.isArray(parsed) ? parsed : [parsed];
      } catch {
        // Continue to next pattern
      }
    }
  }

  logger.warn(
    `Could not parse PowerShell output as JSON: ${output.substring(0, 300)}`
  );
  return [];
}

// ---------------------------------------------------------------------------
// Health check
// ---------------------------------------------------------------------------

app.get("/health", (_req, res) => {
  res.json({ status: "ok", version: "1.0.0" });
});

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

const port = parseInt(process.env.PORT || "5000", 10);

app.listen(port, "0.0.0.0", () => {
  logger.info(`EXO Tools listening on http://0.0.0.0:${port}`);
});

module.exports = app;
