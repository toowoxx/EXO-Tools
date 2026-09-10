"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const request = require("supertest");

const appModule = require("../app");
const app = appModule;

const BUILD_ENV_KEYS = ["GIT_COMMIT", "BUILD_TIME"];

/**
 * Runs `fn` with the given build-metadata env vars applied, then restores the
 * previous values. A value of `undefined` means "unset for the duration".
 *
 * app.js reads process.env per request, so stubbing here is sufficient — the
 * app does not need to be re-required.
 */
async function withBuildEnv(overrides, fn) {
  const saved = Object.fromEntries(
    BUILD_ENV_KEYS.map((key) => [key, process.env[key]]),
  );

  const apply = (values) => {
    for (const key of BUILD_ENV_KEYS) {
      if (values[key] === undefined) {
        delete process.env[key];
      } else {
        process.env[key] = values[key];
      }
    }
  };

  apply(overrides);
  try {
    return await fn();
  } finally {
    apply(saved);
  }
}

test("GET /health falls back to 'unknown' build metadata when unset", async () => {
  await withBuildEnv({ GIT_COMMIT: undefined, BUILD_TIME: undefined }, async () => {
    const response = await request(app)
      .get("/health")
      .expect(200)
      .expect("Content-Type", /json/);

    assert.deepEqual(response.body, {
      status: "ok",
      version: "1.0.0",
      commit: "unknown",
      buildTime: "unknown",
    });
  });
});

test("GET /health returns configured build metadata", async () => {
  await withBuildEnv(
    { GIT_COMMIT: "deadbeef", BUILD_TIME: "2026-09-10T12:00:00Z" },
    async () => {
      const response = await request(app)
        .get("/health")
        .expect(200)
        .expect("Content-Type", /json/);

      assert.deepEqual(response.body, {
        status: "ok",
        version: "1.0.0",
        commit: "deadbeef",
        buildTime: "2026-09-10T12:00:00Z",
      });
    },
  );
});

test("unauthenticated GET / redirects to /login", async () => {
  const response = await request(app).get("/").expect(302);

  assert.equal(response.headers.location, "/login");
});

test("POST /api/get-permissions rejects non-JSON requests before auth", async () => {
  const response = await request(app)
    .post("/api/get-permissions")
    .type("form")
    .send({ userPrincipalName: "user@example.com" })
    .expect(415)
    .expect("Content-Type", /json/);

  assert.equal(response.body.error, "Content-Type must be application/json");
});

test("GET /api/get-permissions/:jobId rejects invalid job IDs", async () => {
  const response = await request(app)
    .get("/api/get-permissions/not-a-uuid")
    .expect(400)
    .expect("Content-Type", /json/);

  assert.equal(response.body.error, "Invalid job ID.");
});

test("safeRedirectUrl keeps same-origin relative paths", () => {
  assert.equal(appModule.safeRedirectUrl("/dashboard"), "/dashboard");
});

test("safeRedirectUrl rejects absolute and protocol-relative URLs", () => {
  assert.equal(appModule.safeRedirectUrl("https://evil.example"), "/");
  assert.equal(appModule.safeRedirectUrl("//evil.example"), "/");
});

test("parsePsJson parses a trailing JSON array from mixed stdout", () => {
  const parsed = appModule.parsePsJson("Started\nMore logs\n[{\"MailboxUPN\":\"user@example.com\"}]");

  assert.deepEqual(parsed, [{ MailboxUPN: "user@example.com" }]);
});

test("parsePsJson wraps a trailing JSON object in an array", () => {
  const parsed = appModule.parsePsJson("noise\n{\"MailboxUPN\":\"user@example.com\"}");

  assert.deepEqual(parsed, [{ MailboxUPN: "user@example.com" }]);
});

test("parsePsJson returns an empty array for invalid output", () => {
  assert.deepEqual(appModule.parsePsJson("noise\n{not-json}"), []);
  assert.deepEqual(appModule.parsePsJson(""), []);
});
