// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as appInsights from "applicationinsights";

/**
 * Thin wrapper around Application Insights TelemetryClient.
 * All methods are no-ops until `init()` is called with a valid key.
 */
export class AppInsightsTransport {
  private client: appInsights.TelemetryClient | undefined;

  /**
   * Initialise the App Insights client.  Must be called once before any
   * tracking.  If called with an empty key the transport stays inert.
   */
  init(instrumentationKey: string, commonProperties: Record<string, string>): void {
    if (!instrumentationKey) return;

    if (appInsights.defaultClient) {
      this.client = new appInsights.TelemetryClient(instrumentationKey);
      this.client.channel.setUseDiskRetryCaching(true);
    } else {
      appInsights
        .setup(instrumentationKey)
        .setAutoCollectRequests(false)
        .setAutoCollectPerformance(false)
        .setAutoCollectExceptions(false)
        .setAutoCollectDependencies(false)
        .setAutoDependencyCorrelation(false)
        .setAutoCollectConsole(false)
        .setUseDiskRetryCaching(true)
        .start();
      this.client = appInsights.defaultClient;
    }

    this.client.commonProperties = commonProperties;
  }

  trackEvent(
    name: string,
    properties?: Record<string, string>,
    measurements?: Record<string, number>
  ): void {
    this.client?.trackEvent({ name, properties, measurements });
  }

  trackException(
    error: Error,
    properties?: Record<string, string>,
    measurements?: Record<string, number>
  ): void {
    this.client?.trackException({ exception: error, properties, measurements });
  }

  flush(): Promise<void> {
    if (!this.client) return Promise.resolve();
    return new Promise<void>((resolve) => {
      this.client!.flush({ callback: () => resolve() });
    });
  }
}
