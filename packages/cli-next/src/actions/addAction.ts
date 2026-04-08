// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { AtkContext, AtkError } from "@microsoft/teamsfx-core";
import { runOperation, declarativeAgent } from "@microsoft/teamsfx-core";

export interface AddActionInput {
  projectPath: string;
  agentManifestPath?: string;
  apiSpecPath: string;
  pluginManifestPath: string;
  actionId: string;
}

export async function addActionAction(ctx: AtkContext, input: AddActionInput): Promise<void> {
  const agentManifestPath = await resolveAgentManifestPath(
    input.agentManifestPath,
    input.projectPath
  );

  const result = await runOperation(declarativeAgent.addExistingPluginOp, ctx, {
    agentManifestPath,
    pluginManifestPath: input.pluginManifestPath,
    apiSpecPath: input.apiSpecPath,
    actionId: input.actionId,
  });
  if (result.isErr()) throw toError(result.error);
}

async function resolveAgentManifestPath(
  explicit: string | undefined,
  projectPath: string
): Promise<string> {
  if (explicit) return explicit;
  const result = await declarativeAgent.getAgentManifestPath(projectPath);
  if (result.isErr()) throw toError(result.error);
  return result.value;
}

function toError(atkError: AtkError): Error {
  const err = new Error(atkError.message);
  err.name = atkError.code;
  if (atkError.inner) {
    err.cause = atkError.inner;
  }
  return err;
}
