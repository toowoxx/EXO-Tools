"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const request = require("supertest");

const appModule = require("../app");
const app = appModule;

test("GET /health returns the expected health payload", async () => {
  const response = await request(app)
    .get("/health")
    .expect(200)
    .expect("Content-Type", /json/);

  assert.equal(response.body.status, "ok");
  assert.equal(response.body.version, "1.0.0");
  assert.equal(typeof response.body.commit, "string");
  assert.equal(typeof response.body.buildTime, "string");
  assert.ok(response.body.commit.length > 0);
  assert.ok(response.body.buildTime.length > 0);
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
