// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { z } from "zod";
import { ok, err } from "neverthrow";
import { createDriver } from "../../createDriver";
import { systemError, userError } from "../../../core/error";
import { TeamsDevPortalClient } from "../../../clients/teamsDevPortal/client";
import {
  appStudioScopes,
  OauthRegistration,
  OauthRegistrationAppType,
  OauthRegistrationTargetAudience,
  TokenExchangeMethodType,
} from "../../../clients/teamsDevPortal/types";

const MAX_NAME_LENGTH = 128;

const httpsUrl = z
  .string()
  .url()
  .refine((u) => u.startsWith("https://"), {
    message: "URL must use HTTPS",
  });

const inputSchema = z.object({
  /** Display name for the OAuth configuration (max 128 chars) */
  name: z.string().min(1).max(MAX_NAME_LENGTH),
  /** Teams app ID */
  appId: z.string().min(1),
  /** OAuth flow — only "authorizationCode" is supported */
  flow: z.literal("authorizationCode"),
  /** OAuth client ID */
  clientId: z.string().min(1),
  /** OAuth client secret (required for Custom provider unless PKCE is enabled) */
  clientSecret: z.string().optional(),
  /** Whether PKCE is enabled */
  isPKCEEnabled: z.boolean().optional(),
  /** Identity provider: "Custom" or "MicrosoftEntra" */
  identityProvider: z.enum(["Custom", "MicrosoftEntra"]).optional(),
  /** Token exchange method */
  tokenExchangeMethodType: z.nativeEnum(TokenExchangeMethodType).optional(),
  /** Who can use this registration */
  applicableToApps: z.nativeEnum(OauthRegistrationAppType).optional(),
  /** Tenant audience */
  targetAudience: z.nativeEnum(OauthRegistrationTargetAudience).optional(),
  /** Base URL of the API — required */
  baseUrl: httpsUrl,
  /** Authorization URL — required for Custom provider */
  authorizationUrl: httpsUrl.optional(),
  /** Token URL — required for Custom provider */
  tokenUrl: httpsUrl.optional(),
  /** Refresh URL */
  refreshUrl: httpsUrl.optional(),
  /** OAuth scopes, comma-separated */
  scope: z.string().optional(),
  /** Existing configuration ID for idempotency */
  existingConfigurationId: z.string().optional(),
});

/**
 * Driver: oauth/register
 *
 * Creates an OAuth configuration in the Teams Developer Portal.
 *
 * Outputs:
 * - OAUTH2_CONFIGURATION_ID
 */
export const oauthRegisterDriver = createDriver({
  id: "oauth/register",
  name: "Register OAuth Configuration",
  inputSchema,
  execute: async (ctx, config) => {
    const source = "oauth/register";

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

    // Idempotency: skip creation if existing config ID is provided and found
    if (config.existingConfigurationId) {
      const existingRes = await client.getOauthRegistration(config.existingConfigurationId);
      if (existingRes.isOk() && existingRes.value) {
        ctx.logger.info(
          `[${source}] OAuth config ${config.existingConfigurationId} already exists — skipping creation`
        );
        return ok({
          outputs: {
            OAUTH2_CONFIGURATION_ID: config.existingConfigurationId,
          },
        });
      }
    }

    // Validate that Custom provider has authorizationUrl and tokenUrl
    const isCustom = !config.identityProvider || config.identityProvider === "Custom";
    if (isCustom) {
      if (!config.authorizationUrl) {
        return err(
          userError(
            "MissingAuthorizationUrl",
            "authorizationUrl is required for Custom identity provider",
            { source }
          )
        );
      }
      if (!config.tokenUrl) {
        return err(
          userError("MissingTokenUrl", "tokenUrl is required for Custom identity provider", {
            source,
          })
        );
      }
    }

    // Build the registration payload
    const applicableToApps = config.applicableToApps ?? OauthRegistrationAppType.AnyApp;
    const targetAudience = config.targetAudience ?? OauthRegistrationTargetAudience.AnyTenant;
    const tokenExchangeMethodType =
      config.tokenExchangeMethodType ?? TokenExchangeMethodType.BasicAuthorizationHeader;

    const registration: OauthRegistration =
      config.identityProvider === "MicrosoftEntra"
        ? {
            description: config.name,
            targetUrlsShouldStartWith: [config.baseUrl],
            applicableToApps,
            m365AppId:
              applicableToApps === OauthRegistrationAppType.SpecificApp ? config.appId : "",
            targetAudience,
            clientId: config.clientId,
            clientSecret: "",
            identityProvider: "MicrosoftEntra",
            tokenExchangeMethodType,
          }
        : {
            description: config.name,
            targetUrlsShouldStartWith: [config.baseUrl],
            applicableToApps,
            m365AppId:
              applicableToApps === OauthRegistrationAppType.SpecificApp ? config.appId : "",
            targetAudience,
            clientId: config.clientId,
            clientSecret: config.clientSecret ?? "",
            isPKCEEnabled: !!config.isPKCEEnabled,
            authorizationEndpoint: config.authorizationUrl,
            tokenExchangeEndpoint: config.tokenUrl,
            tokenRefreshEndpoint: config.refreshUrl,
            scopes: config.scope ? config.scope.split(",").map((s) => s.trim()) : [],
            identityProvider: "Custom",
            tokenExchangeMethodType,
          };

    const createRes = await client.createOauthRegistration(registration);
    if (createRes.isErr()) return err(createRes.error);

    const configId = createRes.value.configurationRegistrationId.oAuthConfigId;
    ctx.logger.info(`[${source}] Created OAuth config: ${configId}`);

    return ok({
      outputs: {
        OAUTH2_CONFIGURATION_ID: configId,
      },
    });
  },
});
