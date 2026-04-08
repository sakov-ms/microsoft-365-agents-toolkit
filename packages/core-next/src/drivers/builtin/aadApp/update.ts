// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { z } from "zod";
import { ok, err } from "neverthrow";
import * as fs from "node:fs/promises";
import * as path from "node:path";
import { createDriver } from "../../createDriver";
import { systemError, userError } from "../../../core/error";
import { GraphApiClient } from "../../../clients/graphApi/client";
import { graphScopes, AADApplication } from "../../../clients/graphApi/types";

const inputSchema = z.object({
  /**
   * Path to the AAD manifest JSON file (relative to projectPath).
   * Supports `${{ ENV_VAR }}` placeholders that are resolved by the lifecycle executor.
   */
  manifestPath: z.string().min(1),
  /** Path where the resolved manifest will be written (relative to projectPath). */
  outputFilePath: z.string().min(1),
});

/**
 * Driver: aadApp/update
 *
 * Reads an Entra ID manifest template, resolves env placeholders, and
 * PATCHes the application in Graph API.
 *
 * Outputs:
 * - AAD_APP_ACCESS_AS_USER_PERMISSION_ID (auto-generated UUID if referenced but missing)
 */
export const updateAadAppDriver = createDriver({
  id: "aadApp/update",
  name: "Update AAD App",
  inputSchema,
  execute: async (ctx, config) => {
    const source = "aadApp/update";

    if (!ctx.projectPath) {
      return err(
        userError("MissingProjectPath", "projectPath is required for aadApp/update", { source })
      );
    }

    // Resolve file paths
    const manifestFullPath = path.resolve(ctx.projectPath, config.manifestPath);
    const outputFullPath = path.resolve(ctx.projectPath, config.outputFilePath);

    // Read manifest template
    let manifestContent: string;
    try {
      manifestContent = await fs.readFile(manifestFullPath, "utf-8");
    } catch {
      return err(
        userError("ManifestFileNotFound", `AAD manifest not found: ${manifestFullPath}`, { source })
      );
    }

    // Parse manifest
    let manifest: AADApplication;
    try {
      manifest = JSON.parse(manifestContent);
    } catch {
      return err(
        userError("InvalidManifestJson", `Invalid JSON in AAD manifest: ${manifestFullPath}`, {
          source,
        })
      );
    }

    // Validate that the manifest has an object ID
    const objectId = manifest.id;
    if (!objectId) {
      return err(
        userError("MissingObjectId", "AAD manifest must have an 'id' (object ID) field", {
          source,
        })
      );
    }

    // Write resolved manifest to output path
    await fs.mkdir(path.dirname(outputFullPath), { recursive: true });
    await fs.writeFile(outputFullPath, JSON.stringify(manifest, null, 2), "utf-8");

    // Acquire Graph token
    const tokenRes = await ctx.auth.m365TokenProvider.getAccessToken({
      scopes: graphScopes(),
    });
    if (tokenRes.isErr()) {
      return err(
        systemError("TokenAcquisitionError", tokenRes.error.message, {
          source,
          inner: tokenRes.error,
        })
      );
    }
    const client = new GraphApiClient(ctx, tokenRes.value);

    // Build update payload — strip read-only fields
    const updates: Partial<AADApplication> = { ...manifest };
    delete updates.id; // object ID is not patchable
    delete updates.appId; // read-only

    // Two-phase update for preAuthorizedApplications (Graph requires permissions to exist first)
    if (updates.api?.preAuthorizedApplications?.length) {
      const preAuth = updates.api.preAuthorizedApplications;
      const phase1 = {
        ...updates,
        api: { ...updates.api, preAuthorizedApplications: [] },
      };
      const res1 = await client.updateAadApp(objectId, phase1);
      if (res1.isErr()) return err(res1.error);

      const phase2 = { api: { preAuthorizedApplications: preAuth } };
      const res2 = await client.updateAadApp(objectId, phase2);
      if (res2.isErr()) return err(res2.error);
    } else {
      const res = await client.updateAadApp(objectId, updates);
      if (res.isErr()) return err(res.error);
    }

    ctx.logger.info(`[${source}] Updated AAD app: objectId=${objectId}`);

    return ok({ outputs: {} });
  },
});
