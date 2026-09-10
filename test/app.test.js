"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const request = require("supertest");

const originalGitCommit = process.env.GIT_COMMIT;
const originalBuildTime = process.env.BUILD_TIME;
delete process.env.GIT_COMMIT;
delete process.env.BUILD_TIME;

const appModule = require("../app");
const app = appModule;

if (originalGitCommit === undefined) {
  delete process.env.GIT_COMMIT;
} else {
  process.env.GIT_COMMIT = originalGitCommit;
}

if (originalBuildTime === undefined) {
  delete process.env.BUILD_TIME;
} else {
  process.env.BUILD_TIME = originalBuildTime;
}

test("GET /health returns the expected health payload", async () => {
  const response = await request(app)
    .get("/health")
    .expect(200)
    .expect("Content-Type", /json/);

  assert.equal(response.body.status, "ok");
  assert.equal(response.body.version, "1.0.0");
  assert.equal(response.body.commit, "unknown");
  assert.equal(response.body.buildTime, "unknown");
});

test("GET /health returns configured build metadata", { concurrency: false }, async () => {
  const originalGitCommit = process.env.GIT_COMMIT;
  const originalBuildTime = process.env.BUILD_TIME;
  process.env.GIT_COMMIT = "deadbeef";
  process.env.BUILD_TIME = "2026-09-10T12:00:00Z";

  try {
    const response = await request(app)
      .get("/health")
      .expect(200)
      .expect("Content-Type", /json/);

    assert.equal(response.body.commit, "deadbeef");
    assert.equal(response.body.buildTime, "2026-09-10T12:00:00Z");
  } finally {
    if (originalGitCommit === undefined) {
      delete process.env.GIT_COMMIT;
    } else {
      process.env.GIT_COMMIT = originalGitCommit;
    }

    if (originalBuildTime === undefined) {
      delete process.env.BUILD_TIME;
    } else {
      process.env.BUILD_TIME = originalBuildTime;
    }
  }
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
