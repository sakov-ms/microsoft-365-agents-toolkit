// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { z } from "zod";
import { ok, err } from "neverthrow";
import { createDriver } from "../../createDriver";
import { systemError, userError } from "../../../core/error";
import { TeamsDevPortalClient } from "../../../clients/teamsDevPortal/client";
import {
  appStudioScopes,
  ApiSecretRegistration,
  ApiSecretRegistrationAppType,
  ApiSecretRegistrationTargetAudience,
  ApiSecretRegistrationClientSecret,
} from "../../../clients/teamsDevPortal/types";
import { graphScopes } from "../../../clients/graphApi/types";

const MAX_NAME_LENGTH = 128;
const MIN_SECRET_LENGTH = 10;
const MAX_SECRET_LENGTH = 512;

const secretSchema = z.string().min(MIN_SECRET_LENGTH).max(MAX_SECRET_LENGTH);

const httpsUrl = z
  .string()
  .url()
  .refine((u) => u.startsWith("https://"), {
    message: "URL must use HTTPS",
  });

const inputSchema = z.object({
  /** Display name for the API key registration (max 128 chars) */
  name: z.string().min(1).max(MAX_NAME_LENGTH),
  /** Teams app ID */
  appId: z.string().min(1),
  /** Primary client secret (10-512 chars) */
  primaryClientSecret: secretSchema.optional(),
  /** Secondary client secret (10-512 chars) */
  secondaryClientSecret: secretSchema.optional(),
  /** Base URL of the API (HTTPS required) */
  baseUrl: httpsUrl,
  /** Who can use this registration */
  applicableToApps: z.nativeEnum(ApiSecretRegistrationAppType).optional(),
  /** Tenant audience */
  targetAudience: z.nativeEnum(ApiSecretRegistrationTargetAudience).optional(),
  /** Existing registration ID for idempotency */
  existingRegistrationId: z.string().optional(),
});

/**
 * Driver: apiKey/register
 *
 * Creates an API key (secret) registration in the Teams Developer Portal.
 *
 * Outputs:
 * - API_KEY_REGISTRATION_ID
 */
export const apiKeyRegisterDriver = createDriver({
  id: "apiKey/register",
  name: "Register API Key",
  inputSchema,
  execute: async (ctx, config) => {
    const source = "apiKey/register";

    // Acquire M365 token for TDP
    const tokenRes = await ctx.auth.m365TokenProvider.getAccessToken({
      scopes: appStudioScopes(),
    });
    if (tokenRes.isErr()) {
      return err(
        systemError("TokenAcquisitionError", tokenRes.error.message, {
          source,
          inner: tokenRes.error,
        })
      );
    }
    const client = new TeamsDevPortalClient(ctx, tokenRes.value);

    // Idempotency: skip creation if existing registration ID is provided and found
    if (config.existingRegistrationId) {
      const existingRes = await client.getApiKeyRegistration(config.existingRegistrationId);
      if (existingRes.isOk() && existingRes.value) {
        ctx.logger.info(
          `[${source}] API key registration ${config.existingRegistrationId} already exists — skipping creation`
        );
        return ok({
          outputs: {
            API_KEY_REGISTRATION_ID: config.existingRegistrationId,
          },
        });
      }
    }

    // Get current user's OID for manageableByUsers
    const graphTokenRes = await ctx.auth.m365TokenProvider.getJsonObject({
      scopes: graphScopes(),
    });
    let userId: string | undefined;
    if (graphTokenRes.isOk() && graphTokenRes.value) {
      userId = graphTokenRes.value["oid"] as string | undefined;
    }

    // Build client secrets array
    const clientSecrets: ApiSecretRegistrationClientSecret[] = [];
    if (config.primaryClientSecret) {
      clientSecrets.push({
        value: config.primaryClientSecret,
        description: config.name,
        priority: 0,
        isValueRedacted: false,
      });
    }
    if (config.secondaryClientSecret) {
      clientSecrets.push({
        value: config.secondaryClientSecret,
        description: config.name,
        priority: 1,
        isValueRedacted: false,
      });
    }

    if (clientSecrets.length === 0) {
      return err(
        userError(
          "NoClientSecretProvided",
          "At least one client secret (primaryClientSecret or secondaryClientSecret) is required",
          { source }
        )
      );
    }

    // Build the registration payload
    const applicableToApps = config.applicableToApps ?? ApiSecretRegistrationAppType.AnyApp;
    const targetAudience = config.targetAudience ?? ApiSecretRegistrationTargetAudience.AnyTenant;

    const registration: ApiSecretRegistration = {
      description: config.name,
      targetUrlsShouldStartWith: [config.baseUrl],
      applicableToApps,
      specificAppId:
        applicableToApps === ApiSecretRegistrationAppType.SpecificApp ? config.appId : "",
      targetAudience,
      clientSecrets,
      manageableByUsers: userId ? [{ userId, accessType: "ReadWrite" }] : undefined,
    };

    const createRes = await client.createApiKeyRegistration(registration);
    if (createRes.isErr()) return err(createRes.error);

    const registrationId = createRes.value.id;
    if (!registrationId) {
      return err(
        systemError("MissingRegistrationId", "API key registration created but no ID returned", {
          source,
        })
      );
    }

    ctx.logger.info(`[${source}] Created API key registration: ${registrationId}`);

    return ok({
      outputs: {
        API_KEY_REGISTRATION_ID: registrationId,
      },
    });
  },
});
