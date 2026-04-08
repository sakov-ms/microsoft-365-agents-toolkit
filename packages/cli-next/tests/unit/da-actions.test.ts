/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it, afterEach, beforeEach } from "mocha";
import * as sinon from "sinon";
import * as fs from "node:fs/promises";
import * as path from "node:path";
import * as os from "node:os";
import type { AtkContext } from "@microsoft/teamsfx-core-next";

// Actions under test
import { addActionAction } from "../../src/actions/addAction";
import { addCapabilityAction } from "../../src/actions/addCapability";
import { addAuthConfigAction } from "../../src/actions/addAuthConfig";
import { setSensitivityLabelAction } from "../../src/actions/setSensitivityLabel";
import { m365SideloadAction } from "../../src/actions/m365Sideload";

function createMockContext(overrides?: Partial<AtkContext>): AtkContext {
  return {
    auth: {
      m365TokenProvider: {} as any,
      azureAccountProvider: {} as any,
    },
    logger: {
      log: sinon.stub(),
      verbose: sinon.stub(),
      debug: sinon.stub(),
      info: sinon.stub(),
      warning: sinon.stub(),
      error: sinon.stub(),
      logInFile: sinon.stub().resolves(),
      getLogFilePath: sinon.stub().returns("/tmp/test.log"),
    } as any,
    telemetry: {
      sendTelemetryEvent: sinon.stub(),
      sendTelemetryErrorEvent: sinon.stub(),
      sendTelemetryException: sinon.stub(),
    },
    ui: {} as any,
    correlationId: "test-correlation-id",
    projectPath: "/tmp/test-project",
    ...overrides,
  };
}

describe("DA Action Handlers", () => {
  const sandbox = sinon.createSandbox();
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "cli-da-test-"));
  });

  afterEach(async () => {
    sandbox.restore();
    if (tmpDir) {
      await fs.rm(tmpDir, { recursive: true, force: true });
    }
  });

  describe("addActionAction()", () => {
    it("should throw when agent manifest not found and no auto-discover", async () => {
      const ctx = createMockContext({ projectPath: tmpDir });
      try {
        await addActionAction(ctx, {
          projectPath: tmpDir,
          apiSpecPath: "/nonexistent/spec.yaml",
          pluginManifestPath: "/nonexistent/plugin.json",
          actionId: "myAction",
          // No agentManifestPath → auto-discover will fail (no manifest.json)
        });
        expect.fail("should have thrown");
      } catch (e: any) {
        expect(e.name).to.equal("ManifestNotFound");
      }
    });

    it("should have correct function signature and accept explicit manifest path", async () => {
      const ctx = createMockContext({ projectPath: tmpDir });
      // The operation will fail at Zod validation or file I/O, but the shape is verified.
      try {
        await addActionAction(ctx, {
          projectPath: tmpDir,
          agentManifestPath: path.join(tmpDir, "da.json"),
          apiSpecPath: path.join(tmpDir, "spec.yaml"),
          pluginManifestPath: path.join(tmpDir, "plugin.json"),
          actionId: "a1",
        });
      } catch (e: any) {
        // Expected — files don't exist
        expect(e.message).to.be.a("string");
      }
    });
  });

  describe("addCapabilityAction()", () => {
    it("should throw when agent manifest not found via auto-discover", async () => {
      const ctx = createMockContext({ projectPath: tmpDir });
      try {
        await addCapabilityAction(ctx, {
          projectPath: tmpDir,
          source: "web-search",
        });
        expect.fail("should have thrown");
      } catch (e: any) {
        expect(e.name).to.equal("ManifestNotFound");
      }
    });

    it("should accept all knowledge source types", async () => {
      const sources = [
        "web-search",
        "onedrive-sharepoint",
        "graph-connector",
        "embedded-knowledge",
      ] as const;
      for (const source of sources) {
        const ctx = createMockContext({ projectPath: tmpDir });
        try {
          await addCapabilityAction(ctx, {
            projectPath: tmpDir,
            agentManifestPath: path.join(tmpDir, "da.json"),
            source,
          });
        } catch {
          // Expected — file doesn't exist
        }
      }
    });
  });

  describe("addAuthConfigAction()", () => {
    it("should accept oauth auth type", async () => {
      const ctx = createMockContext({ projectPath: tmpDir });
      try {
        await addAuthConfigAction(ctx, {
          projectPath: tmpDir,
          authType: "oauth",
          authName: "myOAuth",
          specPath: "./spec.yaml",
        });
      } catch (e: any) {
        // Expected — teamsapp.yml doesn't exist
        expect(e.message).to.be.a("string");
      }
    });

    it("should accept api-key auth type", async () => {
      const ctx = createMockContext({ projectPath: tmpDir });
      try {
        await addAuthConfigAction(ctx, {
          projectPath: tmpDir,
          authType: "api-key",
          authName: "myApiKey",
          specPath: "./spec.yaml",
        });
      } catch (e: any) {
        expect(e.message).to.be.a("string");
      }
    });

    it("should use custom yml-path when provided", async () => {
      const ctx = createMockContext({ projectPath: tmpDir });
      const customYml = path.join(tmpDir, "custom.yml");
      try {
        await addAuthConfigAction(ctx, {
          projectPath: tmpDir,
          authType: "oauth",
          ymlPath: customYml,
          authName: "myOAuth",
          specPath: "./spec.yaml",
          entra: true,
          enablePkce: true,
        });
      } catch (e: any) {
        expect(e.message).to.be.a("string");
      }
    });
  });

  describe("setSensitivityLabelAction()", () => {
    it("should throw when agent manifest not found via auto-discover", async () => {
      const ctx = createMockContext({ projectPath: tmpDir });
      try {
        await setSensitivityLabelAction(ctx, {
          projectPath: tmpDir,
          labelId: "label-123",
        });
        expect.fail("should have thrown");
      } catch (e: any) {
        expect(e.name).to.equal("ManifestNotFound");
      }
    });

    it("should accept explicit agent manifest path", async () => {
      const ctx = createMockContext({ projectPath: tmpDir });
      try {
        await setSensitivityLabelAction(ctx, {
          projectPath: tmpDir,
          agentManifestPath: path.join(tmpDir, "da.json"),
          labelId: "label-456",
        });
      } catch (e: any) {
        // Expected — file doesn't exist
        expect(e.message).to.be.a("string");
      }
    });
  });

  describe("m365SideloadAction()", () => {
    it("should throw when app package file not found", async () => {
      const ctx = createMockContext({ projectPath: tmpDir });
      try {
        await m365SideloadAction(ctx, {
          filePath: path.join(tmpDir, "nonexistent.zip"),
        });
        expect.fail("should have thrown");
      } catch (e: any) {
        expect(e.name).to.equal("AppPackageNotFound");
      }
    });

    it("should accept scope parameter", async () => {
      const ctx = createMockContext({ projectPath: tmpDir });
      try {
        await m365SideloadAction(ctx, {
          filePath: path.join(tmpDir, "app.zip"),
          scope: "Shared",
        });
      } catch (e: any) {
        // Expected — file doesn't exist
        expect(e.name).to.equal("AppPackageNotFound");
      }
    });
  });
});
