// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Tag-driven validators for E2E tests.
 *
 * Each validator is a pure function: (envMap, projectPath) → AssertionResult[].
 * Template tags (e.g. ["bot", "tab", "aad"]) select which validators to run.
 * Adding a new validator = add a function + register it in VALIDATOR_MAP.
 */

import * as fs from "fs";
import * as path from "path";
import type { AssertionResult } from "./tracer";

type EnvMap = Map<string, string>;

/** A validator function returns assertion results without throwing. */
type ValidatorFn = (envMap: EnvMap, projectPath: string) => Promise<AssertionResult[]>;

// ---------------------------------------------------------------------------
// Individual validators
// ---------------------------------------------------------------------------

async function validateTeamsApp(envMap: EnvMap): Promise<AssertionResult[]> {
  const appId = envMap.get("TEAMS_APP_ID");
  return [
    {
      name: "TEAMS_APP_ID is defined",
      passed: !!appId,
      expected: "non-empty string",
      actual: appId ?? "undefined",
    },
  ];
}

async function validateBot(envMap: EnvMap): Promise<AssertionResult[]> {
  const botId = envMap.get("BOT_ID");
  const botEndpoint = envMap.get("BOT_ENDPOINT") ?? envMap.get("BOT_DOMAIN");
  return [
    {
      name: "BOT_ID is defined",
      passed: !!botId,
      expected: "non-empty string",
      actual: botId ?? "undefined",
    },
    {
      name: "BOT_ENDPOINT or BOT_DOMAIN is defined",
      passed: !!botEndpoint,
      expected: "non-empty string",
      actual: botEndpoint ?? "undefined",
    },
  ];
}

async function validateTab(envMap: EnvMap): Promise<AssertionResult[]> {
  const endpoint = envMap.get("TAB_ENDPOINT") ?? envMap.get("FRONTEND_ENDPOINT");
  return [
    {
      name: "TAB_ENDPOINT or FRONTEND_ENDPOINT is defined",
      passed: !!endpoint,
      expected: "non-empty string",
      actual: endpoint ?? "undefined",
    },
  ];
}

async function validateAad(envMap: EnvMap): Promise<AssertionResult[]> {
  const clientId = envMap.get("AAD_APP_CLIENT_ID");
  return [
    {
      name: "AAD_APP_CLIENT_ID is defined",
      passed: !!clientId,
      expected: "non-empty string",
      actual: clientId ?? "undefined",
    },
  ];
}

async function validateFunction(envMap: EnvMap): Promise<AssertionResult[]> {
  const endpoint = envMap.get("API_FUNCTION_ENDPOINT");
  return [
    {
      name: "API_FUNCTION_ENDPOINT is defined",
      passed: !!endpoint,
      expected: "non-empty string",
      actual: endpoint ?? "undefined",
    },
  ];
}

async function validateProjectStructure(
  _envMap: EnvMap,
  projectPath: string
): Promise<AssertionResult[]> {
  const results: AssertionResult[] = [];

  const envDir = path.join(projectPath, "env");
  results.push({
    name: "env/ directory exists",
    passed: fs.existsSync(envDir),
    expected: "directory exists",
    actual: fs.existsSync(envDir) ? "exists" : "missing",
  });

  return results;
}

// ---------------------------------------------------------------------------
// Validator registry — maps tags to validator functions
// ---------------------------------------------------------------------------

const VALIDATOR_MAP: Record<string, ValidatorFn> = {
  teamsApp: validateTeamsApp,
  bot: validateBot,
  tab: validateTab,
  aad: validateAad,
  function: validateFunction,
  project: validateProjectStructure,
};

/**
 * Run validators selected by tags.
 * Always includes "project" validator. Additional validators (teamsApp, bot,
 * tab, etc.) are selected by the caller based on template lifecycle.
 */
export async function runValidators(
  tags: string[],
  envMap: EnvMap,
  projectPath: string
): Promise<AssertionResult[]> {
  const allTags = new Set(["project", ...tags]);
  const results: AssertionResult[] = [];
  for (const tag of allTags) {
    const fn = VALIDATOR_MAP[tag];
    if (fn) {
      results.push(...(await fn(envMap, projectPath)));
    }
  }
  return results;
}

/**
 * Check if all assertions passed.
 */
export function allPassed(results: AssertionResult[]): boolean {
  return results.every((r) => r.passed);
}
