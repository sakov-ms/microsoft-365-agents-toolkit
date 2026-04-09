// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * MCP scaffold integration tests — subprocess layer.
 *
 * Verifies that `atk new da mcp-remote` and `atk add action` with MCP flags
 * produce the correct project structure (ai-plugin.json, declarativeAgent.json,
 * .vscode/mcp.json). Does NOT create real Azure resources.
 *
 * Maps to ADO test plan suite "MCP - Declarative Agent with Remote MCP Server"
 * (37357417).
 */

import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import { expect } from "chai";
import { describe, it, beforeEach, afterEach } from "mocha";
import { promisify } from "util";
import { exec } from "child_process";

const execAsync = promisify(exec);

const ATK_BIN = process.env.ATK_BIN || "atk";
const TIMEOUT = 120_000; // 2 min per command
const MCP_SERVER_URL = "https://learn.microsoft.com/api/mcp";

function tmpDir(): string {
  const dir = path.join(os.tmpdir(), `atk-mcp-test-${Date.now()}`);
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

async function run(
  command: string,
  cwd?: string
): Promise<{ stdout: string; stderr: string; exitCode: number }> {
  try {
    const { stdout, stderr } = await execAsync(command, {
      cwd,
      timeout: TIMEOUT,
      env: { ...process.env, CI_ENABLED: "true" },
    });
    return { stdout, stderr, exitCode: 0 };
  } catch (e: any) {
    return {
      stdout: e.stdout ?? "",
      stderr: e.stderr ?? "",
      exitCode: e.code ?? 1,
    };
  }
}

function readJson(filePath: string): any {
  return JSON.parse(fs.readFileSync(filePath, "utf8"));
}

describe("MCP scaffold integration", function () {
  this.timeout(5 * 60 * 1000); // 5 min total

  let dir: string;

  beforeEach(function () {
    dir = tmpDir();
  });

  afterEach(async function () {
    await fs.promises.rm(dir, { recursive: true, force: true }).catch(() => {});
  });

  describe("atk new da mcp-remote (no auth)", function () {
    it("scaffolds project with MCP server URL", async function () {
      const result = await run(
        `${ATK_BIN} new da mcp-remote --name McpApp --mcpServerUrl ${MCP_SERVER_URL} --folder ${dir} --non-interactive`,
        dir
      );
      expect(result.exitCode, `stderr: ${result.stderr}`).to.equal(0);

      const projectPath = path.join(dir, "McpApp");
      expect(fs.existsSync(projectPath), "project dir should exist").to.be.true;

      // Verify core project files
      expect(fs.existsSync(path.join(projectPath, "m365agents.yml")), "m365agents.yml").to.be.true;
      expect(fs.existsSync(path.join(projectPath, "env", ".env.dev")), ".env.dev").to.be.true;

      // Verify appPackage structure
      const appPackage = path.join(projectPath, "appPackage");
      expect(fs.existsSync(path.join(appPackage, "ai-plugin.json")), "ai-plugin.json").to.be.true;
      expect(fs.existsSync(path.join(appPackage, "declarativeAgent.json")), "declarativeAgent.json")
        .to.be.true;
      expect(fs.existsSync(path.join(appPackage, "manifest.json")), "manifest.json").to.be.true;

      // Verify declarativeAgent.json references ai-plugin as action
      const daManifest = readJson(path.join(appPackage, "declarativeAgent.json"));
      expect(daManifest.actions).to.be.an("array").that.is.not.empty;
      expect(daManifest.actions[0].file).to.equal("ai-plugin.json");

      // Verify .vscode/mcp.json has the remote URL
      const mcpConfig = readJson(path.join(projectPath, ".vscode", "mcp.json"));
      expect(mcpConfig.servers).to.be.an("object");
      const serverEntry = Object.values(mcpConfig.servers)[0] as any;
      expect(serverEntry.url).to.equal(MCP_SERVER_URL);
      expect(serverEntry.type).to.equal("http");
    });

    it("scaffolds project without URL — mcp.json has empty/no URL", async function () {
      const result = await run(
        `${ATK_BIN} new da mcp-remote --name McpNoUrl --folder ${dir} --non-interactive`,
        dir
      );
      // May exit 0 (scaffold with placeholder) or non-zero (validation failure)
      const projectPath = path.join(dir, "McpNoUrl");
      if (result.exitCode === 0 && fs.existsSync(projectPath)) {
        // If scaffold succeeds without URL, verify no URL appears in mcp.json
        const mcpConfigPath = path.join(projectPath, ".vscode", "mcp.json");
        if (fs.existsSync(mcpConfigPath)) {
          const mcpConfig = readJson(mcpConfigPath);
          const serverEntry = Object.values(mcpConfig.servers)[0] as any;
          expect(serverEntry.url).to.satisfy(
            (url: string) => !url || url === "" || url.includes("{{"),
            "URL should be empty or a placeholder"
          );
        }
      } else {
        // Scaffold failed without URL — acceptable behavior
        expect(result.exitCode).to.not.equal(0);
      }
    });
  });

  describe("atk new da mcp-remote --help", function () {
    it("shows --mcpServerUrl option", async function () {
      const result = await run(`${ATK_BIN} new da mcp-remote --help`);
      expect(result.exitCode).to.equal(0);
      expect(result.stdout.toLowerCase()).to.include("mcpserverurl");
    });
  });

  describe("atk add action on MCP project", function () {
    it("add action --help exits 0 and shows MCP-relevant flags", async function () {
      const result = await run(`${ATK_BIN} add action --help`);
      expect(result.exitCode).to.equal(0);
      expect(result.stdout).to.include("--api-spec-path");
      expect(result.stdout).to.include("--plugin-manifest-path");
      expect(result.stdout).to.include("--action-id");
    });

    it("add action fails gracefully without required flags", async function () {
      // scaffold a DA first
      const scaffoldResult = await run(
        `${ATK_BIN} new da basic --name McpBase --folder ${dir} --non-interactive`,
        dir
      );
      expect(scaffoldResult.exitCode, `scaffold stderr: ${scaffoldResult.stderr}`).to.equal(0);

      const projectPath = path.join(dir, "McpBase");
      // add action without required flags → should fail with helpful error
      const result = await run(`${ATK_BIN} add action`, projectPath);
      expect(result.exitCode).to.not.equal(0);
      // Should mention missing required option
      expect(result.stderr).to.match(/required|missing|api-spec-path/i);
    });
  });
});
