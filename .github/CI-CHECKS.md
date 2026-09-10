# CI & checks

## What runs

- **CI (`.github/workflows/ci.yml`)** runs on push to `main` and on pull requests.
- The test job runs on **Node.js 18, 20, and 22**, then executes the app import smoke check and `npm test`.
- The dependency audit job runs `npm ci` and `npm audit --audit-level=high`; high-severity vulnerabilities fail the workflow.
- A **gitleaks** job scans pushes to `main` and pull requests for leaked secrets.
- **CodeQL (`.github/workflows/codeql.yml`)** runs on push to `main`, pull requests, and a weekly schedule.

## Local baseline

```bash
npm ci && npm test
```
