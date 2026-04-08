---
name: dev-test-next
description: "Local development, testing, and verification workflow for core-next and cli-next packages. Use when: building core-next or cli-next, running unit/integration/e2e tests, verifying a feature end-to-end, running lifecycle tests locally, setting up developer credentials, debugging test failures, checking test results."
argument-hint: "Describe what you want to build, test, or verify"
---

# Dev Test Next — Local Development & Verification

## When to Use

- Building or testing `packages/core-next` or `packages/cli-next`
- Running unit, integration, or E2E tests locally
- Verifying a feature works end-to-end before pushing
- Debugging test failures from CI

## Package Overview

| Package | Path | Role |
|---------|------|------|
| **core-next** | `packages/core-next` | Engine: templates, drivers, lifecycle operations, project scaffolding |
| **cli-next** | `packages/cli-next` | CLI: commands, auth, output formatting, E2E tests |

Dependency: `core-next` → `cli-next` (cli-next depends on core-next)

**TypeScript:** Both packages use TypeScript 6.0 (`~6.0.0`) with `moduleResolution: "bundler"` and `module: "commonjs"`. The deprecated `node10` resolution is no longer used.

## Quick Reference

```bash
# Build both packages (always build core-next first)
cd <repo-root>
npm run setup:next                            # install + build both

# Or build individually — postbuild auto-runs eslint --fix + prettier
pnpm --filter ./packages/core-next build
pnpm --filter ./packages/cli-next build

# Link CLI globally for local testing
cd packages/cli-next && pnpm link --global    # makes `atk` command available
```

> **Note:** `build` includes a `postbuild` hook that automatically runs `eslint --fix` and
> `prettier --write` on `src/` and `tests/`. No separate format step needed after building.

## Test Pyramid

### 1. Unit Tests (no credentials needed)

```bash
# core-next — ~500 unit tests
cd packages/core-next && npm run test:unit

# cli-next — ~80 unit + ~55 integration tests
cd packages/cli-next && npm run test:unit
```

- Framework: Mocha + Chai + Sinon, ts-node/register via `.mocharc.js`
- Coverage: NYC (reports inline after test run)
- Pattern: `tests/unit/**/*.tests.ts`

### 2. Integration Tests (no credentials needed)

```bash
cd packages/cli-next && npm run test:integration
```

- Pattern: `tests/integration/**/*.tests.ts`
- Tests cross-module wiring without Azure or M365 services

### 3. E2E Tests

Two test suites with different credential requirements:

| Suite | Command | Credentials | What it tests |
|-------|---------|-------------|---------------|
| **CLI syntax** | `npm run test:e2e:cli` | None (just needs `atk` on PATH) | Binary arg parsing, `--version`, `--help`, scaffold |
| **Lifecycle** | `npm run test:e2e:lifecycle` | Azure + M365 login | Full scaffold → provision → deploy → validate per template |
| **All E2E** | `npm run test:e2e` | Azure + M365 login | Both suites |
| **Cleanup** | `npm run test:e2e:clean` | Azure login | Deletes stale Azure resource groups |

Test files: `packages/cli-next/tests/e2e/`

## Local E2E Setup

### CLI Syntax Tests (no credentials)

```bash
cd packages/cli-next
pnpm link --global          # register `atk` command
npm run test:e2e:cli        # runs cli-syntax.tests.ts
```

### Lifecycle Tests (credentials required)

**Step 1 — Login once using the CLI:**

```bash
atk auth login azure        # opens browser, caches token to ~/.fx/account/
atk auth login m365          # opens browser, caches M365 token
```

Tokens persist in `~/.fx/account/` (AES-encrypted MSAL cache). No env vars needed locally.

**Step 2 — Run:**

```bash
npm run test:e2e:lifecycle
```

**How auth works locally:**
- Lifecycle operations (provision, deploy) use `createTokenProvider()` which picks
  up cached tokens from `atk auth login` via `AzureAccountManager` (MSAL silent flow).
- Resource group management uses `DefaultAzureCredential` which discovers the same
  cached credentials automatically.
- Subscription defaults to `af46c703-f714-4f4c-af42-835a673c2b13`
  (Teams Cloud – E2E Testing, TTL = 1 day). Override with `AZURE_SUBSCRIPTION_ID` env var.

### Run a Specific Template

Use Mocha `--grep` to filter by template ID (tests are named `E2E lifecycle: <id> [<lang>]`):

```bash
# Single template
npx mocha --require tests/e2e/setup.ts --timeout 1200000 \
  --grep "bot/echo.*TypeScript" "tests/e2e/lifecycle.tests.ts"

# All bot templates
npx mocha --require tests/e2e/setup.ts --timeout 1200000 \
  --grep "bot/" "tests/e2e/lifecycle.tests.ts"

# All TypeScript templates
npx mocha --require tests/e2e/setup.ts --timeout 1200000 \
  --grep "TypeScript" "tests/e2e/lifecycle.tests.ts"
```

### Cleanup Stale Resources

```bash
npm run test:e2e:clean                    # sweep resources older than 2h
npx ts-node tests/e2e/clean.ts --run-id <id>  # cleanup a specific run
```

Resources are tagged with `atk-test=true` and `created-at` timestamp for reliable tag-based cleanup.

## CI Pipelines

| Workflow | File | Trigger | What it runs |
|----------|------|---------|--------------|
| **CI Next** | `.github/workflows/ci-next.yml` | PR/push to `dev`/`release/**` | Build → Lint → Format → Unit tests → Integration tests |
| **E2E Test Next** | `.github/workflows/e2e-test-next.yml` | PR/push, nightly, manual | Matrix of E2E test files with Azure creds |

### CI Next Verification

CI Next runs automatically on every PR touching `packages/core-next/**` or `packages/cli-next/**`.
To replicate locally:

```bash
# Full CI check
pnpm --filter ./packages/core-next build
pnpm --filter ./packages/cli-next build
pnpm --filter ./packages/core-next lint
pnpm --filter ./packages/cli-next lint
pnpm --filter ./packages/core-next format:check
pnpm --filter ./packages/cli-next format:check
pnpm --filter ./packages/core-next test:unit
pnpm --filter ./packages/cli-next test:unit
pnpm --filter ./packages/cli-next test:integration
```

### E2E in CI vs Local

| Aspect | CI (`CI_ENABLED=true`) | Local (default) |
|--------|------------------------|-----------------|
| Auth | Env vars (service principal / username+password) | `atk auth login` cached tokens |
| Subscription | `AZURE_SUBSCRIPTION_ID` env var | `af46c703-...` (E2E Testing, TTL 1d) |
| Credential for RG mgmt | `UsernamePasswordCredential` / `ClientSecretCredential` | `DefaultAzureCredential` |
| Env vars | All required, throws on missing | All optional |
| Retry | Mocha `.retries(1)` + workflow rerun up to 5× | Mocha `.retries(1)` + checkpoint resume |
| Cleanup | Per-test + tear-down job + stale sweep | Per-test + manual `npm run test:e2e:clean` |

## Verification Checklist

Before pushing a change to core-next or cli-next, verify:

1. **Build** — `pnpm --filter ./packages/core-next build && pnpm --filter ./packages/cli-next build`
   (auto-runs lint fix + format via `postbuild` hook)
2. **Lint check** — `pnpm --filter ./packages/core-next lint && pnpm --filter ./packages/cli-next lint`
   (should show 0 errors; warnings are acceptable)
3. **Format check** — `pnpm --filter ./packages/core-next format:check && pnpm --filter ./packages/cli-next format:check`
4. **Unit tests** — `pnpm --filter ./packages/core-next test:unit && pnpm --filter ./packages/cli-next test:unit`
5. **Integration tests** — `pnpm --filter ./packages/cli-next test:integration`
6. **(Optional) E2E CLI** — `cd packages/cli-next && npm run test:e2e:cli`
7. **(Optional) E2E Lifecycle** — `cd packages/cli-next && npm run test:e2e:lifecycle` (after `atk auth login`)

Steps 1–5 match what CI Next runs. Steps 6–7 match what E2E Test Next runs.

> **Shortcut:** Since `build` now auto-formats, steps 1–3 can be collapsed into just
> `build` + `lint`. Only run `format:check` separately if you skipped the build.

## E2E Architecture

```
tests/e2e/
├── infra/
│   ├── config.ts        — Auth config (CI vs local mode)
│   ├── tracer.ts        — TestTracer (telemetry capture) + TestProgress + verifyTelemetry()
│   ├── checkpoint.ts    — Resume-from-failure retry (skip completed phases)
│   ├── azure.ts         — Tagged resource group create/delete/sweep
│   ├── validators.ts    — Tag-driven validators (teamsApp, bot, tab, aad, etc.)
│   └── testContext.ts   — Creates AtkContext with real auth + test instrumentation
├── setup.ts             — Mocha root hook (keytar mock, log dir, driver registration)
├── clean.ts             — Standalone cleanup script
├── lifecycle.tests.ts   — Data-driven: templateRegistry.list() × languages
└── cli-syntax.tests.ts  — Subprocess CLI verification
```

**Key design decisions:**
- **Programmatic API primary** — tests call `runOperation(provisionOp, ctx, input)` directly, not subprocesses
- **Data-driven** — `templateRegistry.list()` generates tests; new templates get E2E coverage automatically
- **Tag-based cleanup** — Azure RGs tagged with `atk-test`, `test-run-id`, `created-at` for reliable sweep
- **Checkpoint retry** — on Mocha retry, completed phases (scaffold, create-rg) are skipped
- **Telemetry verification** — `verifyTelemetry()` runs 6 contract rules after each test

## Troubleshooting

| Symptom | Fix |
|---------|-----|
| `atk: command not found` | `cd packages/cli-next && pnpm link --global` |
| `Cannot find module '@microsoft/teamsfx-core'` | Build core-next first: `pnpm --filter ./packages/core-next build` |
| `E2E config: missing required environment variable` | You're in CI mode. Unset `CI_ENABLED` or set all required vars |
| `DefaultAzureCredential: no credential` | Run `atk auth login azure` first |
| Tests hang at provision | Check `atk auth login azure` is still valid; tokens expire after ~1h |
| `SyntaxError: Unexpected token ':'` | Missing ts-node/register — `.mocharc.js` should load it automatically |
| Resource group not cleaned up | Run `npm run test:e2e:clean` to sweep stale resources |
| Coverage threshold failure (exit code 1) | Unit tests passed but NYC coverage < 50% — this is pre-existing, not blocking |
