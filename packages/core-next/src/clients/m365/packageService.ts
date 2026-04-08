// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import axios, { AxiosInstance } from "axios";
import * as fs from "node:fs/promises";
import AdmZip from "adm-zip";
import FormData from "form-data";

/** MOS3 endpoint for public cloud. */
const MOS_ENDPOINT = "https://titles.prod.mos.microsoft.com";

/** M365 sideloading token scope. */
export const mosServiceScopes = (): string[] => [`${MOS_ENDPOINT}/.default`];

/** App scope for sideloading. */
export enum AppScope {
  Personal = "Personal",
  Shared = "Shared",
  Tenant = "Tenant",
}

/**
 * Check if a manifest describes a Declarative Agent app.
 */
function isDeclarativeAgentManifest(manifest: Record<string, unknown>): boolean {
  const agents = (manifest as { copilotAgents?: { declarativeAgents?: unknown[] } }).copilotAgents
    ?.declarativeAgents;
  return Array.isArray(agents) && agents.length > 0;
}

/**
 * Minimal M365 PackageService client for sideloading apps.
 * Implements the V1 and V2 sideloading flows used by the
 * teamsApp/extendToM365 driver.
 */
export class M365PackageService {
  private readonly http: AxiosInstance;
  private readonly initEndpoint: string;

  constructor(endpoint = MOS_ENDPOINT) {
    this.initEndpoint = endpoint;
    this.http = axios.create({ timeout: 60_000 });
  }

  /**
   * Resolve the actual titles service URL from the MOS config endpoint.
   */
  private async getTitleServiceUrl(token: string): Promise<string> {
    const res = await this.http.get("/config/v1/environment", {
      baseURL: this.initEndpoint,
      headers: { Authorization: `Bearer ${token}` },
    });
    return res.data.titlesServiceUrl as string;
  }

  /**
   * Sideload an app package. Returns [titleId, appId, shareLink].
   */
  async sideLoad(
    token: string,
    packagePath: string,
    appScope = AppScope.Personal
  ): Promise<[string, string, string]> {
    const zip = new AdmZip(packagePath);
    const manifestEntry = zip.getEntry("manifest.json");
    if (!manifestEntry) {
      throw new Error("Invalid app package zip — manifest.json is missing");
    }
    const manifest = JSON.parse(manifestEntry.getData().toString("utf8")) as Record<
      string,
      unknown
    >;

    if (isDeclarativeAgentManifest(manifest)) {
      const [titleId, appId] = await this.sideLoadV2(token, packagePath, appScope);
      let shareLink = "";
      if (appScope === AppScope.Shared) {
        shareLink = await this.getShareLink(token, titleId);
      }
      return [titleId, appId, shareLink];
    } else {
      const [titleId, appId] = await this.sideLoadV1(token, packagePath);
      return [titleId, appId, ""];
    }
  }

  /**
   * V2 sideloading (Builder API) — used for declarative agents.
   */
  private async sideLoadV2(
    token: string,
    packagePath: string,
    appScope: AppScope
  ): Promise<[string, string]> {
    const data = await fs.readFile(packagePath);
    const form = new FormData();
    form.append("package", data);
    form.append("info", JSON.stringify({ builderName: "TeamsToolKit" }));
    const serviceUrl = await this.getTitleServiceUrl(token);

    const uploadRes = await this.http.post("/builder/v1/users/packages", form, {
      baseURL: serviceUrl,
      headers: {
        ...form.getHeaders(),
        Authorization: `Bearer ${token}`,
      },
      params: { scope: appScope },
    });

    const statusId: string = uploadRes.data.statusId;

    // Poll for completion
    const maxAttempts = 60;
    for (let i = 0; i < maxAttempts; i++) {
      const statusRes = await this.http.get(`/builder/v1/users/packages/status/${statusId}`, {
        baseURL: serviceUrl,
        headers: { Authorization: `Bearer ${token}` },
      });
      if (statusRes.status === 200 && statusRes.data.titleId) {
        return [statusRes.data.titleId as string, statusRes.data.appId as string];
      }
      await delay(5_000);
    }
    throw new Error("Sideloading timed out waiting for package status.");
  }

  /**
   * V1 sideloading — used for classic Teams apps.
   */
  private async sideLoadV1(token: string, packagePath: string): Promise<[string, string]> {
    const data = await fs.readFile(packagePath);
    const form = new FormData();
    form.append("package", data);
    const serviceUrl = await this.getTitleServiceUrl(token);

    const uploadRes = await this.http.post("/dev/v1/users/packages", form.getBuffer(), {
      baseURL: serviceUrl,
      headers: {
        ...form.getHeaders(),
        Authorization: `Bearer ${token}`,
      },
    });

    const operationId: string = uploadRes.data.operationId;

    // Acquire
    const acquireRes = await this.http.post(
      "/dev/v1/users/packages/acquisitions",
      { operationId },
      {
        baseURL: serviceUrl,
        headers: { Authorization: `Bearer ${token}` },
      }
    );
    const statusId: string = acquireRes.data.statusId;

    // Poll for completion
    const maxAttempts = 60;
    for (let i = 0; i < maxAttempts; i++) {
      const statusRes = await this.http.get(`/dev/v1/users/packages/status/${statusId}`, {
        baseURL: serviceUrl,
        headers: { Authorization: `Bearer ${token}` },
      });
      if (statusRes.status === 200 && statusRes.data.titleId) {
        return [statusRes.data.titleId as string, statusRes.data.appId as string];
      }
      await delay(5_000);
    }
    throw new Error("Sideloading timed out waiting for package status.");
  }

  /**
   * Get a share link for a sideloaded title.
   */
  private async getShareLink(token: string, titleId: string): Promise<string> {
    try {
      const serviceUrl = await this.getTitleServiceUrl(token);
      const res = await this.http.get(`/marketplace/v1/users/titles/${titleId}/sharingInfo`, {
        baseURL: serviceUrl,
        headers: { Authorization: `Bearer ${token}` },
      });
      return (res.data?.shareLink as string) ?? "";
    } catch {
      return "";
    }
  }
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
