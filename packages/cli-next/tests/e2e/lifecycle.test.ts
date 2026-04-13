// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Data-driven E2E lifecycle tests.
 *
 * Iterates over templateRegistry.list(), generating a Mocha describe/it block
 * for each testable template × language combination.
 *
 * Each test:
 *   1. scaffold via createProjectOp (programmatic)
 *   2. create tagged resource group
 *   3. provision via provisionOp (programmatic, with skipConsent)
 *   4. deploy via deployOp (programmatic)
 *   5. validate via tag-driven validators
 *   6. verify telemetry (contract tests)
 *   7. cleanup via Promise.allSettled (tagged RG deletion)
 *
 * Checkpoint-based retry: on Mocha retry, skip already‐completed phases.
 */

import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import { expect } from "chai";
import { describe, it, afterEach } from "mocha";
import { v4 as uuidv4 } from "uuid";
import * as yaml from "js-yaml";
import {
  templateRegistry,
  registerBuiltinTemplates,
  runOperation,
  project,
  provisionOp,
  deployOp,
} from "@microsoft/teamsfx-core-next";
import type { TemplateDescriptor } from "@microsoft/teamsfx-core-next";

// Populate the registry — templates don't self-register at import time
registerBuiltinTemplates();
import { readEnvFile } from "@microsoft/teamsfx-core-next/build/environment/envManager";
import { createTestContext } from "./infra/testContext";
import { TestCheckpoint } from "./infra/checkpoint";
import { StepLogger, verifyTelemetry } from "./infra/tracer";
import { createResourceGroup, deleteResourceGroup } from "./infra/azure";
import { runValidators, allPassed } from "./infra/validators";
import { getConfig } from "./infra/config";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Set an env var in a dotenv-style string.
 * If the key already exists (even with an empty value), replace the line.
 * Otherwise append a new line.
 */
function upsertEnvVar(content: string, key: string, value: string): string {
  const re = new RegExp(`^${key}=.*$`, "m");
  if (re.test(content)) {
    return content.replace(re, `${key}=${value}`);
  }
  return content.trimEnd() + `\n${key}=${value}\n`;
}

function getTestFolder(): string {
  const folder = path.resolve(os.homedir(), "atk-e2e-tests");
  if (!fs.existsSync(folder)) {
    fs.mkdirSync(folder, { recursive: true });
  }
  return folder;
}

function getUniqueAppName(): string {
  return "atkE2E" + Date.now().toString() + uuidv4().slice(0, 2);
}

async function loadEnvMap(projectPath: string, envName: string): Promise<Map<string, string>> {
  const result = await readEnvFile(projectPath, envName);
  if (result.isOk()) {
    return new Map(Object.entries(result.value));
  }
  return new Map();
}

/**
 * Determine which validators to run based on template tags
 * and which lifecycle phases to execute based on template YAML.
 */
function getValidationTags(template: TemplateDescriptor): string[] {
  return template.tags ?? [];
}

/**
 * Driver ID prefixes that require an Azure resource group.
 * Templates whose provision/deploy steps only use teamsApp/* or cli/* drivers
 * don't need Azure infra and can skip RG creation entirely.
 */
const AZURE_DRIVER_PREFIXES = ["arm/", "azureFunctions/", "azureAppService/", "azureStorage/"];

/**
 * Parse m365agents.yml and check whether any provision/deploy step uses
 * a driver that requires Azure infrastructure (e.g. arm/deploy).
 */
function yamlNeedsAzure(yamlContent: string): boolean {
  try {
    const doc = yaml.load(yamlContent) as Record<string, unknown> | undefined;
    if (!doc) return false;
    const steps = [
      ...((doc.provision as Array<{ uses?: string }>) ?? []),
      ...((doc.deploy as Array<{ uses?: string }>) ?? []),
    ];
    return steps.some(
      (s) => s.uses && AZURE_DRIVER_PREFIXES.some((prefix) => s.uses!.startsWith(prefix))
    );
  } catch {
    return false;
  }
}

// ---------------------------------------------------------------------------
// Data-driven test generation
// ---------------------------------------------------------------------------

const templates = templateRegistry.list().filter((t) => t.testable !== false);

if (templates.length === 0) {
  throw new Error(
    "templateRegistry is empty — registerBuiltinTemplates() may have failed. " +
      "Cannot generate lifecycle tests with 0 templates."
  );
}

for (const template of templates) {
  for (const lang of template.languages) {
    describe(`E2E lifecycle: ${template.id} [${lang}]`, function () {
      this.timeout(20 * 60 * 1000); // 20 minutes
      this.retries(1); // 1 retry with checkpoint resume

      const appName = getUniqueAppName();
      const projectPath = path.join(getTestFolder(), appName);
      const rgName = `${appName}-rg`;
      const envName = "dev";
      const testId = `${template.id}/${lang}`;
      const cfg = getConfig();

      const checkpoint = new TestCheckpoint(testId);
      const logger = new StepLogger(testId);

      afterEach(async function () {
        // GUARANTEED cleanup — runs even on test failure
        // Uses Promise.allSettled so one failure doesn't block others
        await Promise.allSettled([
          deleteResourceGroup(rgName),
          fs.promises.rm(projectPath, { recursive: true, force: true }).catch(() => {}),
        ]);
        checkpoint.reset();
        await logger.flush();
      });

      it(`scaffold → provision → deploy → validate → telemetry check`, async function () {
        const { ctx, tracer, progress: _progress } = createTestContext(projectPath);

        // --- Phase 1: Scaffold ---
        await checkpoint.runPhase("scaffold", async () => {
          await logger.wrapStep("scaffold", async () => {
            const result = await runOperation(project.createProjectOp, ctx, {
              templateId: template.id,
              projectName: appName,
              language: lang,
              destinationPath: getTestFolder(),
            });
            expect(result.isOk(), `scaffold failed: ${result.isErr() ? result.error.message : ""}`)
              .to.be.true;
            expect(fs.existsSync(projectPath), "project folder must exist").to.be.true;
          });
        });

        // --- Lifecycle detection ---
        // Some templates (e.g. da/typespec, da/mcp-local) don't produce
        // m365agents.yml and therefore have no provision/deploy lifecycle.
        // Skip those phases gracefully instead of failing.
        const yamlPath = path.join(projectPath, "m365agents.yml");
        const yamlExists = fs.existsSync(yamlPath);
        const yamlContent = yamlExists ? fs.readFileSync(yamlPath, "utf-8") : "";
        const hasProvisionLifecycle = yamlExists && yamlContent.includes("provision:");
        const hasDeployLifecycle = yamlExists && yamlContent.includes("deploy:");
        const needsAzure = yamlExists && yamlNeedsAzure(yamlContent);

        // --- Phase 2: Create resource group (only when needed) ---
        if (needsAzure) {
          await checkpoint.runPhase("create-rg", async () => {
            await logger.wrapStep("create-rg", async () => {
              const ok = await createResourceGroup({
                name: rgName,
                location: "westus",
                templateId: template.id,
                runId: cfg.githubRunId,
              });
              expect(ok, "resource group creation must succeed").to.be.true;
            });
          });
        }

        // --- Phase 3: Provision ---
        if (hasProvisionLifecycle) {
          await checkpoint.runPhase("provision", async () => {
            await logger.wrapStep("provision", async () => {
              // Inject resource group name and subscription ID into env file.
              // Templates scaffold .env.dev with empty placeholders like
              // AZURE_RESOURCE_GROUP_NAME= so we must replace the value,
              // not just check for key presence.
              // Some templates may not scaffold .env.dev at all, so create it
              // when missing to ensure env vars are always available.
              const envDir = path.join(projectPath, "env");
              const envFilePath = path.join(envDir, `.env.${envName}`);
              if (!fs.existsSync(envDir)) {
                fs.mkdirSync(envDir, { recursive: true });
              }
              let content = fs.existsSync(envFilePath) ? fs.readFileSync(envFilePath, "utf-8") : "";
              content = upsertEnvVar(content, "AZURE_RESOURCE_GROUP_NAME", rgName);
              content = upsertEnvVar(content, "AZURE_SUBSCRIPTION_ID", cfg.azureSubscriptionId);
              fs.writeFileSync(envFilePath, content);

              const result = await runOperation(provisionOp, ctx, {
                projectPath,
                envName,
                skipConsent: true,
              });
              expect(
                result.isOk(),
                `provision failed: ${result.isErr() ? result.error.message : ""}`
              ).to.be.true;

              // Validate provision
              const envMap = await loadEnvMap(projectPath, envName);
              const assertions = await runValidators(
                getValidationTags(template),
                envMap,
                projectPath
              );
              return { assertions, env: Object.fromEntries(envMap) };
            });
          });
        } else {
          await logger.wrapStep("provision", async () => {
            // No provision lifecycle — skip
            return undefined;
          });
        }

        // --- Phase 4: Deploy (skip if template has no deploy lifecycle) ---
        if (hasDeployLifecycle) {
          await checkpoint.runPhase("deploy", async () => {
            await logger.wrapStep("deploy", async () => {
              const result = await runOperation(deployOp, ctx, {
                projectPath,
                envName,
                skipConsent: true,
              });
              expect(result.isOk(), `deploy failed: ${result.isErr() ? result.error.message : ""}`)
                .to.be.true;
            });
          });
        } else {
          await logger.wrapStep("deploy", async () => {
            // No deploy lifecycle — skip
            return undefined;
          });
        }

        // --- Phase 5: Final validation ---
        await logger.wrapStep("validate", async () => {
          const envMap = await loadEnvMap(projectPath, envName);
          const validationTags = getValidationTags(template);
          // Only check teamsApp when provision lifecycle was executed
          if (hasProvisionLifecycle) {
            validationTags.push("teamsApp");
          }
          const assertions = await runValidators(validationTags, envMap, projectPath);
          const failed = assertions.filter((a) => !a.passed);
          expect(
            allPassed(assertions),
            `Validation failed:\n${failed.map((a) => `  ✗ ${a.name}: expected=${a.expected}, actual=${a.actual}`).join("\n")}`
          ).to.be.true;
          return { assertions, env: Object.fromEntries(envMap) };
        });

        // --- Phase 6: Telemetry verification ---
        const issues = verifyTelemetry(tracer.spans);
        logger.logTelemetryCheck(issues, tracer.spans.length);
        expect(
          issues,
          `Telemetry issues:\n${issues.map((i) => `  ✗ [${i.rule}] ${i.message}`).join("\n")}`
        ).to.have.length(0);
      });
    });
  }
}
