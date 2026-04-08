// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AxiosInstance } from "axios";
import { Result, ok, err } from "neverthrow";
import { AtkContext } from "../../core/context";
import { AtkError, userError, systemError } from "../../core/error";
import { createHttpClient } from "../../http/httpClient";
import { sendWithRetry } from "../../http/retry";
import { AADApplication, AadOwner, PasswordCredential, GRAPH_BASE_URL } from "./types";

/**
 * Client for Microsoft Graph API — Entra ID (Azure AD) application operations.
 *
 * Instantiated per-operation with an already-acquired Graph bearer token.
 */
export class GraphApiClient {
  private readonly axios: AxiosInstance;

  constructor(ctx: AtkContext, token: string) {
    this.axios = createHttpClient(ctx, { baseURL: GRAPH_BASE_URL });
    this.axios.defaults.headers.common["Authorization"] = `Bearer ${token}`;
    this.axios.defaults.headers.common["Content-Type"] = "application/json";
  }

  /**
   * Create a new Entra ID application.
   */
  async createAadApp(
    displayName: string,
    signInAudience?: string,
    serviceManagementReference?: string
  ): Promise<Result<AADApplication, AtkError>> {
    const body: Record<string, unknown> = { displayName };
    if (signInAudience) body.signInAudience = signInAudience;
    if (serviceManagementReference) body.serviceManagementReference = serviceManagementReference;

    try {
      const response = await sendWithRetry(() => this.axios.post("/applications", body));
      return ok(response.data as AADApplication);
    } catch (e: unknown) {
      return err(this.wrapError("createAadApp", e));
    }
  }

  /**
   * Generate a client secret (password credential) for an application.
   *
   * Uses extra retries for 404 to handle Entra replication delay.
   */
  async generateClientSecret(
    objectId: string,
    expiresInDays = 180,
    description = "default"
  ): Promise<Result<string, AtkError>> {
    const now = new Date();
    const end = new Date(now);
    end.setDate(end.getDate() + expiresInDays);

    const body = {
      passwordCredential: {
        displayName: description,
        startDateTime: now.toISOString(),
        endDateTime: end.toISOString(),
      },
    };

    try {
      const response = await sendWithRetry(
        () => this.axios.post(`/applications/${objectId}/addPassword`, body),
        5 // extra retries for Entra sync delay
      );
      const cred = response.data as PasswordCredential;
      if (!cred.secretText) {
        return err(
          systemError("EmptyClientSecret", "Graph API returned an empty secretText", {
            source: "GraphApiClient",
          })
        );
      }
      return ok(cred.secretText);
    } catch (e: unknown) {
      return err(this.wrapError("generateClientSecret", e));
    }
  }

  /**
   * Update an Entra ID application (PATCH semantics).
   */
  async updateAadApp(
    objectId: string,
    updates: Partial<AADApplication>
  ): Promise<Result<void, AtkError>> {
    try {
      await sendWithRetry(
        // lgtm[js/file-access-to-http] updates is a programmatic object, not raw file data
        () => this.axios.patch(`/applications/${objectId}`, updates),
        5 // extra retries for 404/400 race conditions during permission sync
      );
      return ok(undefined);
    } catch (e: unknown) {
      return err(this.wrapError("updateAadApp", e));
    }
  }

  /**
   * Get the owners of an application.
   */
  async getOwners(objectId: string): Promise<Result<AadOwner[], AtkError>> {
    try {
      const response = await sendWithRetry(() =>
        this.axios.get(`/applications/${objectId}/owners`)
      );
      return ok((response.data?.value ?? []) as AadOwner[]);
    } catch (e: unknown) {
      return err(this.wrapError("getOwners", e));
    }
  }

  /**
   * Add an owner to an application.
   */
  async addOwner(objectId: string, userObjectId: string): Promise<Result<void, AtkError>> {
    try {
      await sendWithRetry(() =>
        this.axios.post(`/applications/${objectId}/owners/$ref`, {
          "@odata.id": `${GRAPH_BASE_URL}/directoryObjects/${userObjectId}`,
        })
      );
      return ok(undefined);
    } catch (e: unknown) {
      return err(this.wrapError("addOwner", e));
    }
  }

  /* ------------------------------------------------------------------ */
  /*  Error mapping                                                      */
  /* ------------------------------------------------------------------ */

  private wrapError(apiName: string, e: unknown): AtkError {
    if (e && typeof e === "object" && "response" in e) {
      const resp = (e as any).response;
      const status: number = resp?.status ?? 0;
      const graphCode: string = resp?.data?.error?.code ?? "";
      const graphMsg: string = resp?.data?.error?.message ?? String(e);

      if (status >= 400 && status < 500) {
        return userError(`GraphApiError_${graphCode || status}`, `[${apiName}] ${graphMsg}`, {
          source: "GraphApiClient",
        });
      }
      return systemError(`GraphApiError_${status}`, `[${apiName}] ${graphMsg}`, {
        source: "GraphApiClient",
      });
    }
    return systemError(
      "GraphApiUnexpectedError",
      `[${apiName}] ${e instanceof Error ? e.message : String(e)}`,
      { source: "GraphApiClient", inner: e instanceof Error ? e : undefined }
    );
  }
}
