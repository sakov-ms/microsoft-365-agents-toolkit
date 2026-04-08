// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { z } from "zod";
import { ok, err } from "neverthrow";
import * as fs from "node:fs/promises";
import AdmZip from "adm-zip";
import { createDriver } from "../../createDriver";
import { userError, systemError } from "../../../core/error";
import { TeamsDevPortalClient } from "../../../clients/teamsDevPortal/client";
import { appStudioScopes } from "../../../clients/teamsDevPortal/types";

const inputSchema = z.object({
  /** Path to the zipped app package (.zip) */
  appPackagePath: z.string().min(1),
});

/**
 * Driver: teamsApp/publishAppPackage
 *
 * Publishes a Teams app to the organization catalog.
 * Handles first-publish and update scenarios automatically:
 * - If the app is not yet published, calls the initial publish endpoint.
 * - If the app is already published, updates the published version.
 *
 * Outputs:
 * - TEAMS_APP_PUBLISHED_APP_ID: the published app ID from the catalog
 */
export const publishAppPackageDriver = createDriver({
  id: "teamsApp/publishAppPackage",
  name: "Publish App Package",
  inputSchema,
  execute: async (ctx, config) => {
    // Read the app package directly (no separate existence check to avoid TOCTOU race)
    let archivedFile: Buffer;
    try {
      archivedFile = Buffer.from(await fs.readFile(config.appPackagePath));
    } catch (e: unknown) {
      if ((e as NodeJS.ErrnoException).code === "ENOENT") {
        return err(
          userError("AppPackageNotFound", `App package not found: ${config.appPackagePath}`, {
            source: "teamsApp/publishAppPackage",
          })
        );
      }
      throw e;
    }

    // Extract app ID from manifest
    const zip = new AdmZip(archivedFile);
    const manifestEntry = zip.getEntry("manifest.json");
    if (!manifestEntry) {
      return err(
        userError("ManifestNotFound", "manifest.json not found inside app package", {
          source: "teamsApp/publishAppPackage",
        })
      );
    }
    const manifest = JSON.parse(manifestEntry.getData().toString("utf8")) as { id?: string };
    const teamsAppId = manifest.id;
    if (!teamsAppId) {
      return err(
        userError("InvalidTeamsAppId", "manifest.json missing 'id' field", {
          source: "teamsApp/publishAppPackage",
        })
      );
    }

    // Acquire M365 token
    const tokenRes = await ctx.auth.m365TokenProvider.getAccessToken({
      scopes: appStudioScopes(),
    });
    if (tokenRes.isErr()) {
      return err(
        systemError("TokenAcquisitionError", tokenRes.error.message, {
          source: "teamsApp/publishAppPackage",
          inner: tokenRes.error,
        })
      );
    }
    const client = new TeamsDevPortalClient(ctx, tokenRes.value);

    // Check if already published
    const stagedRes = await client.getStagedApp(teamsAppId);
    if (stagedRes.isErr()) return err(stagedRes.error);
    const staged = stagedRes.value;

    let publishedAppId: string;

    if (staged) {
      // Already published — update
      ctx.logger.info(
        `[teamsApp/publishAppPackage] App ${teamsAppId} already published (state: ${staged.publishingState}). Updating…`
      );
      const updateRes = await client.publishTeamsAppUpdate(teamsAppId, archivedFile);
      if (updateRes.isErr()) return err(updateRes.error);
      publishedAppId = updateRes.value;
    } else {
      // First publish
      ctx.logger.info(
        `[teamsApp/publishAppPackage] Publishing app ${teamsAppId} for the first time`
      );
      const publishRes = await client.publishTeamsApp(teamsAppId, archivedFile);
      if (publishRes.isErr()) return err(publishRes.error);
      publishedAppId = publishRes.value;
    }

    ctx.logger.info(`[teamsApp/publishAppPackage] Published app ID: ${publishedAppId}`);

    return ok({
      outputs: {
        TEAMS_APP_PUBLISHED_APP_ID: publishedAppId,
      },
    });
  },
});
