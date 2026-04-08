// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * E2E test root hook — runs before all tests.
 * Mock keytar to prevent native module issues on Linux CI.
 * Initialize log directory.
 */

import * as fs from "fs";
import * as path from "path";

// Mock keytar before any other import can trigger it
const Module = require("module");
const originalRequire = Module.prototype.require;
Module.prototype.require = function (id: string, ...args: any[]) {
  if (id === "keytar") {
    return {
      getPassword: async () => null,
      setPassword: async () => {},
      deletePassword: async () => true,
      findPassword: async () => null,
      findCredentials: async () => [],
    };
  }
  return originalRequire.apply(this, [id, ...args]);
};

// Ensure log directory exists
const logDir = path.join(__dirname, "logs");
if (!fs.existsSync(logDir)) {
  fs.mkdirSync(logDir, { recursive: true });
}

// Register builtin drivers for programmatic tests
import { registerBuiltinDrivers } from "@microsoft/teamsfx-core-next";
registerBuiltinDrivers();
