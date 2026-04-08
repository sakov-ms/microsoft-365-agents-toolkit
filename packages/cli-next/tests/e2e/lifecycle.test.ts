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
import { v4 as uuidv4 } from "uuid";
import {
  templateRegistry,
  runOperation,
  project,
  provisionOp,
  deployOp,
} from "@microsoft/teamsfx-core-next";
import type { TemplateDescriptor } from "@microsoft/teamsfx-core-next";
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

// ---------------------------------------------------------------------------
// Data-driven test generation
// ---------------------------------------------------------------------------

const templates = templateRegistry.list().filter((t) => t.testable !== false);

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

        // --- Phase 2: Create resource group ---
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

        // --- Phase 3: Provision ---
        await checkpoint.runPhase("provision", async () => {
          await logger.wrapStep("provision", async () => {
            // Inject resource group name into env file
            const envDir = path.join(projectPath, "env");
            const envFilePath = path.join(envDir, `.env.${envName}`);
            if (fs.existsSync(envFilePath)) {
              let content = fs.readFileSync(envFilePath, "utf-8");
              if (!content.includes("AZURE_RESOURCE_GROUP_NAME")) {
                content += `\nAZURE_RESOURCE_GROUP_NAME=${rgName}\n`;
                fs.writeFileSync(envFilePath, content);
              }
            }

            const result = await runOperation(provisionOp, ctx, {
              projectPath,
              envName,
              skipConsent: true,
            });
            expect(result.isOk(), `provision failed: ${result.isErr() ? result.error.message : ""}`)
              .to.be.true;

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

        // --- Phase 4: Deploy (skip if template has no deploy lifecycle) ---
        const yamlPath = path.join(projectPath, "m365agents.yml");
        const hasDeployLifecycle =
          fs.existsSync(yamlPath) && fs.readFileSync(yamlPath, "utf-8").includes("deploy:");

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
          const assertions = await runValidators(getValidationTags(template), envMap, projectPath);
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
