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
const rateLimit = require("express-rate-limit");
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

// CSRF protection for API mutation endpoints.
// Verify that POST/PUT/PATCH/DELETE requests to /api/* include a JSON
// Content-Type header.  Browsers cannot send cross-origin requests with
// this content type from plain HTML forms, so this check effectively
// prevents CSRF attacks on the API endpoints.
app.use("/api", (req, res, next) => {
  if (["POST", "PUT", "PATCH", "DELETE"].includes(req.method)) {
    const ct = req.headers["content-type"] || "";
    if (!ct.includes("application/json")) {
      return res.status(415).json({ error: "Content-Type must be application/json" });
    }
  }
  next();
});

// ---------------------------------------------------------------------------
// Session directory setup
// ---------------------------------------------------------------------------

/**
 * Return the first writable candidate directory, creating it if needed.
 */
function findSessionDir(preferred) {
  for (const candidate of [preferred, "/tmp/exo_session"]) {
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
    `No writable session directory found. Tried: ${preferred} and /tmp/exo_session. ` +
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
// Rate limiting
// ---------------------------------------------------------------------------

// Azure App Service (and some reverse proxies) include the client port in the
// X-Forwarded-For header (e.g. "91.37.41.151:59928").  With trust proxy: 1,
// Express propagates that raw string as req.ip, which express-rate-limit
// rejects as an invalid IP address.  This helper strips the port so every
// rate-limiter receives a plain IP address as its key.
function ipKey(req) {
  const ip = req.ip || req.socket.remoteAddress || "";
  // Bracket-notation IPv6 with port: "[2001:db8::1]:12345" → "[2001:db8::1]"
  if (ip.startsWith("[")) {
    return ip.replace(/\]:.*$/, "]");
  }
  // IPv4 with port: "1.2.3.4:12345" → "1.2.3.4" (exactly one colon)
  const colonCount = (ip.match(/:/g) || []).length;
  if (colonCount === 1) {
    return ip.substring(0, ip.lastIndexOf(":"));
  }
  // Pure IPv4 or bare IPv6 – return as-is.
  // An empty string here (no remote address at all) is an extreme edge case;
  // Express always populates req.socket.remoteAddress for live TCP connections.
  return ip || "unknown";
}

// General API rate limiter: 60 requests per minute per IP
const apiLimiter = rateLimit({
  windowMs: 60 * 1000,
  max: 60,
  standardHeaders: true,
  legacyHeaders: false,
  keyGenerator: ipKey,
  message: { error: "Too many requests, please try again later." },
});
app.use("/api/", apiLimiter);

// Stricter limiter for the permission-check endpoint (expensive PowerShell call).
// Only applied to the POST (job-start) route, not the GET (poll) route.
const permissionsLimiter = rateLimit({
  windowMs: 60 * 1000,
  max: 5,
  standardHeaders: true,
  legacyHeaders: false,
  keyGenerator: ipKey,
  message: { error: "Too many permission check requests, please try again later." },
});

// ---------------------------------------------------------------------------
// Async permission-check job store
// ---------------------------------------------------------------------------

/**
 * In-memory map of jobId → job object.
 * Each job: { status, upn, oid, startedAt, result, error }
 * status: "pending" | "done" | "error"
 */
const jobs = new Map();

// Purge completed/failed jobs older than 1 hour to prevent unbounded growth.
// Only removes non-pending jobs so active long-running checks are never evicted
// mid-flight (PS_TIMEOUT caps execution at ≤ 600 s, so in practice no pending
// job survives 1 hour, but we guard defensively).
setInterval(() => {
  const cutoff = Date.now() - 60 * 60 * 1000;
  for (const [id, job] of jobs) {
    if (job.status !== "pending" && job.startedAt < cutoff) jobs.delete(id);
  }
}, 10 * 60 * 1000).unref();

// Auth route rate limiter: prevent brute-force / enumeration on login and
// callback endpoints (30 attempts per 15 minutes per IP).
const authLimiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 30,
  standardHeaders: true,
  legacyHeaders: false,
  keyGenerator: ipKey,
  message: "Too many authentication attempts. Please try again later.",
});

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

  logger.info(`Checking group membership against ACCESS_GROUP_ID: ${groupId}`);

  // Graph API check – throws on network/API error so callers can surface a
  // meaningful message instead of silently treating failures as "not a member".
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
      const isMember = (resp.data.value || []).includes(groupId);
      logger.info(`Group membership result: ${isMember ? "member – access granted" : "not a member – access denied"}`);
      return isMember;
    }
    throw new Error(`Unexpected HTTP status ${resp.status} from checkMemberObjects`);
  } catch (err) {
    logger.warn(`Group membership check failed: ${err.message}`);
    throw err;
  }
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
// Auth URL helper
// ---------------------------------------------------------------------------

/**
 * Generate a fresh Microsoft login URL and store the CSRF state token in the
 * session.  Returns null when MSAL cannot be reached (misconfiguration).
 */
async function buildAuthUrl(req) {
  try {
    const state = crypto.randomUUID();
    req.session.state = state;
    const cca = buildMsalApp();
    return await cca.getAuthCodeUrl({
      scopes: config.SCOPES,
      redirectUri: config.REDIRECT_URI,
      state,
    });
  } catch (err) {
    logger.error(`Failed to build auth URL: ${err.message}`);
    return null;
  }
}

// ---------------------------------------------------------------------------
// HTTP request logging
// ---------------------------------------------------------------------------

app.use((req, res, next) => {
  const start = Date.now();
  res.on("finish", () => {
    const ms = Date.now() - start;
    const user =
      req.session && req.session.user
        ? req.session.user.preferred_username || req.session.user.oid
        : "anonymous";
    logger.info(`${req.method} ${req.path} ${res.statusCode} ${ms}ms user=${user}`);
  });
  next();
});

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

app.get("/login", authLimiter, async (req, res) => {
  if (req.session.user) {
    return res.redirect("/");
  }

  logger.info("Login page requested – generating auth URL");

  let error = null;
  const authUrl = await buildAuthUrl(req);
  if (!authUrl) {
    error =
      "Could not reach the Microsoft login service. Please check the application configuration.";
  }

  res.render("login", { title: "EXO Tools – Sign In", authUrl, error });
});

app.get("/callback", authLimiter, async (req, res) => {
  logger.info("Auth callback received");

  // CSRF guard
  if (req.query.state !== req.session.state) {
    logger.warn("State mismatch in auth callback – possible CSRF or expired session");
    const authUrl = await buildAuthUrl(req);
    return res
      .status(400)
      .render("login", {
        title: "EXO Tools – Sign In",
        authUrl,
        error: "Session expired or state mismatch. Please sign in again.",
      });
  }

  if (req.query.error) {
    const desc = req.query.error_description || "Unknown error";
    logger.warn(`Azure AD returned an error: ${req.query.error} – ${desc}`);
    const authUrl = await buildAuthUrl(req);
    return res
      .status(400)
      .render("login", {
        title: "EXO Tools – Sign In",
        authUrl,
        error: `Authentication failed: ${desc}`,
      });
  }

  // Exchange the auth code for tokens
  let result;
  const cca = buildMsalApp();
  try {
    result = await cca.acquireTokenByCode({
      code: req.query.code,
      scopes: config.SCOPES,
      redirectUri: config.REDIRECT_URI,
    });
    logger.info("Token acquired successfully");
  } catch (err) {
    logger.error(`Token acquisition failed: ${err.message}`);
    const authUrl = await buildAuthUrl(req);
    return res
      .status(400)
      .render("login", {
        title: "EXO Tools – Sign In",
        authUrl,
        error: `Could not acquire token: ${err.message}`,
      });
  }

  // Group membership check (separate try/catch to distinguish API errors from
  // a definitive "not a member" result)
  if (config.ACCESS_GROUP_ID) {
    try {
      const isMember = await checkGroupMembership(result.accessToken);
      if (!isMember) {
        const authUrl = await buildAuthUrl(req);
        return res
          .status(403)
          .render("login", {
            title: "EXO Tools – Sign In",
            authUrl,
            error: "Access denied. Your account is not in the required group.",
          });
      }
    } catch (err) {
      logger.error(`Group membership check error: ${err.message}`);
      const authUrl = await buildAuthUrl(req);
      return res
        .status(503)
        .render("login", {
          title: "EXO Tools – Sign In",
          authUrl,
          error:
            "Could not verify group membership. Please try again later.",
        });
    }
  }

  // Extract identity claims from the account object
  const account = result.account;
  if (!account || !account.localAccountId) {
    logger.error(
      "Azure AD token is missing the account identifier – cannot establish session"
    );
    const authUrl = await buildAuthUrl(req);
    return res.status(400).render("login", {
      title: "EXO Tools – Sign In",
      authUrl,
      error:
        "Authentication failed: the identity token is missing required claims.",
    });
  }

  const userOid = account.localAccountId;
  logger.info(`User authenticated: ${account.username || userOid}`);

  req.session.user = {
    name: account.name || "",
    preferred_username: account.username || "",
    oid: userOid,
    tid: account.tenantId || "",
    groups: [],
  };

  // Persist the MSAL token cache to the filesystem (not the session).
  saveMsalCache(userOid, cca.getTokenCache().serialize());

  const nextUrl = req.session.next || "/";
  delete req.session.next;
  logger.info(`Redirecting authenticated user to: ${nextUrl}`);
  res.redirect(nextUrl);
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

// POST – start a permission-check job (returns immediately with a jobId).
app.post("/api/get-permissions", permissionsLimiter, loginRequired, (req, res) => {
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

  // Create the job record and return its ID immediately so the HTTP connection
  // is not held open while PowerShell runs (avoids 504 Gateway Timeout).
  const jobId = crypto.randomUUID();
  const oid = req.session.user.oid;
  jobs.set(jobId, { status: "pending", upn, oid, startedAt: Date.now(), result: null, error: null });

  logger.info(`Starting permission check for: ${upn} (jobId=${jobId})`);

  const child = execFile(
    config.PWSH_PATH,
    args,
    { timeout: config.PS_TIMEOUT * 1000, maxBuffer: 50 * 1024 * 1024 },
    (err, stdout, stderr) => {
      const job = jobs.get(jobId);
      if (!job) return; // job was already cleaned up

      if (err) {
        if (err.killed || err.signal === "SIGTERM") {
          job.status = "error";
          job.error =
            "The permission check timed out. " +
            "This usually happens on very large tenants. " +
            "Consider increasing PS_TIMEOUT.";
          return;
        }
        if (err.code === "ENOENT") {
          job.status = "error";
          job.error =
            `PowerShell not found at '${config.PWSH_PATH}'. ` +
            "Please install PowerShell Core and set the PWSH_PATH environment variable.";
          return;
        }

        const errMsg = (stderr || "").trim() || "PowerShell exited with a non-zero status.";
        logger.error(`PowerShell stderr: ${errMsg.substring(0, 500)}`);
        job.status = "error";
        job.error = `Permission check failed: ${errMsg.substring(0, 300)}`;
        return;
      }

      const output = (stdout || "").trim();
      const permissions = parsePsJson(output);
      logger.info(
        `Permission check complete. Found ${permissions.length} permissions. (jobId=${jobId})`
      );
      job.status = "done";
      job.result = { permissions, user: upn, count: permissions.length };
    }
  );

  // Handle spawn errors (e.g. ENOENT if pwsh not found before the process starts)
  child.on("error", (err) => {
    const job = jobs.get(jobId);
    if (!job) return;
    if (err.code === "ENOENT") {
      job.status = "error";
      job.error =
        `PowerShell not found at '${config.PWSH_PATH}'. ` +
        "Please install PowerShell Core and set the PWSH_PATH environment variable.";
      return;
    }
    logger.error(`Unexpected error during permission check: ${err.message}`);
    job.status = "error";
    job.error = `Internal server error: ${err.message}`;
  });

  // Return 202 Accepted with the job ID – the client will poll for results.
  return res.status(202).json({ jobId });
});

// GET – poll the status of an existing permission-check job.
app.get("/api/get-permissions/:jobId", loginRequired, (req, res) => {
  const { jobId } = req.params;

  // Validate jobId is a well-formed UUID to guard against path traversal etc.
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(jobId)) {
    return res.status(400).json({ error: "Invalid job ID." });
  }

  const job = jobs.get(jobId);
  if (!job) {
    return res.status(404).json({ error: "Job not found or expired." });
  }

  // Prevent one user from polling another user's job.
  if (job.oid !== req.session.user.oid) {
    return res.status(403).json({ error: "Access denied." });
  }

  if (job.status === "pending") {
    return res.json({ status: "pending" });
  }
  if (job.status === "error") {
    return res.status(500).json({ status: "error", error: job.error });
  }
  // done
  return res.json({ status: "done", ...job.result });
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
  // Startup configuration summary – visible in Azure App Service Log Stream
  // Sensitive values are masked to show only the first 8 characters.
  const mask = (v) => (v && v.length > 0 ? `${v.substring(0, 8)}…` : "NOT SET");
  logger.info("=".repeat(60));
  logger.info(`EXO Tools listening on http://0.0.0.0:${port}`);
  logger.info(`NODE_ENV          : ${process.env.NODE_ENV || "not set (defaults to production)"}`);
  logger.info(`Session directory : ${sessionDir}`);
  logger.info("-".repeat(60));
  logger.info("Azure AD / Auth");
  logger.info(`  CLIENT_ID       : ${mask(config.CLIENT_ID)}`);
  logger.info(`  TENANT_ID       : ${mask(config.TENANT_ID)}`);
  logger.info(`  CLIENT_SECRET   : ${config.CLIENT_SECRET ? "set" : "NOT SET"}`);
  logger.info(`  REDIRECT_URI    : ${config.REDIRECT_URI}`);
  logger.info(`  ACCESS_GROUP_ID : ${config.ACCESS_GROUP_ID || "not set (all authenticated users allowed)"}`);
  logger.info("-".repeat(60));
  logger.info("Exchange Online");
  logger.info(`  EXO_APP_ID      : ${mask(config.EXO_APP_ID)}`);
  logger.info(`  EXO_ORGANIZATION: ${config.EXO_ORGANIZATION || "NOT SET"}`);
  logger.info(`  EXO_CERT_PATH   : ${config.EXO_CERT_PATH || "NOT SET"} (exists: ${config.EXO_CERT_PATH ? fs.existsSync(config.EXO_CERT_PATH) : false})`);
  logger.info(`  EXO_CERT_PW     : ${config.EXO_CERT_PASSWORD ? "set" : "not set"}`);
  logger.info(`  PWSH_PATH       : ${config.PWSH_PATH}`);
  logger.info("=".repeat(60));
});

module.exports = app;
