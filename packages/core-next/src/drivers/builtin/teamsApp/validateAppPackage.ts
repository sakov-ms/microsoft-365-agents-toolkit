// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as fs from "fs";
import AdmZip from "adm-zip";
import { z } from "zod";
import { ok, err } from "neverthrow";
import { TeamsManifestWrapper } from "@microsoft/app-manifest";
import { createDriver } from "../../createDriver";
import { userError, systemError } from "../../../core/error";

const inputSchema = z.object({
  /** Path to the .zip app package to validate */
  packagePath: z.string().min(1),
});

/**
 * Driver: teamsApp/validateAppPackage
 *
 * Opens a Teams app package (.zip), extracts manifest.json,
 * validates it, and checks that referenced icons exist in the archive.
 *
 * Outputs:
 * - TEAMS_APP_PACKAGE_VALID: "true" or "false"
 */
export const validateAppPackageDriver = createDriver({
  id: "teamsApp/validateAppPackage",
  name: "Validate App Package",
  inputSchema,
  execute: async (_ctx, config) => {
    try {
      await fs.promises.access(config.packagePath);
    } catch {
      return err(
        userError("PackageNotFound", `App package not found: ${config.packagePath}`, {
          source: "teamsApp/validateAppPackage",
        })
      );
    }

    try {
      const zip = new AdmZip(config.packagePath);
      const manifestEntry = zip.getEntry("manifest.json");
      if (!manifestEntry) {
        return ok({
          outputs: { TEAMS_APP_PACKAGE_VALID: "false" },
        });
      }

      const manifestContent = manifestEntry.getData().toString("utf8");
      const wrapper = TeamsManifestWrapper.fromJSON(manifestContent);
      const errors = await wrapper.validate();

      // Check icons exist in the ZIP
      for (const iconName of [wrapper.icons.color, wrapper.icons.outline]) {
        if (!zip.getEntry(iconName)) {
          errors.push(`Icon file "${iconName}" not found in app package`);
        }
      }

      const valid = errors.length === 0;
      return ok({
        outputs: { TEAMS_APP_PACKAGE_VALID: String(valid) },
      });
    } catch (e) {
      return err(
        systemError("ValidateAppPackageFailed", `Failed to validate app package: ${e}`, {
          source: "teamsApp/validateAppPackage",
          inner: e instanceof Error ? e : new Error(String(e)),
        })
      );
    }
  },
});
