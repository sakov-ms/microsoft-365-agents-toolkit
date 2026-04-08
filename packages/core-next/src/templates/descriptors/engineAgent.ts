// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { TemplateDescriptor, TemplateActionOptions } from "../types";
import type { AtkContext } from "../../core/context";
import { scaffoldTemplates } from "../scaffold/scaffolder";
import { getTemplateReplaceMap } from "../scaffold/replaceMap";
import { type TemplateInfo, convertToLangKey } from "../scaffold/types";

/**
 * Custom Engine Agent template artifact names matching the template repository folder names.
 */
export const EngineAgentTemplateNames = {
  Basic: "basic-custom-engine-agent",
  FunctionCalling: "weather-agent",
  MultiAgent: "travel-agent",
  TeamsCollaborator: "teams-collaborator-agent",
} as const;

/**
 * Create a standard scaffold function for Custom Engine Agent templates.
 */
function makeEngineAgentScaffoldFn(templateName: string) {
  return async (ctx: AtkContext, opts: TemplateActionOptions) => {
    const replaceMap = getTemplateReplaceMap({
      appName: opts.projectName,
      ...opts,
    });

    const tplInfo: TemplateInfo = {
      templateName,
      language: convertToLangKey(opts.language),
      replaceMap,
    };

    const result = await scaffoldTemplates(ctx, [tplInfo], opts.destinationPath);
    return result.map((files) => ({
      projectPath: opts.destinationPath,
      warnings: files.length === 0 ? ["No files were scaffolded"] : undefined,
    }));
  };
}

/**
 * All Custom Engine Agent template descriptors.
 */
export const engineAgentTemplateDescriptors: TemplateDescriptor[] = [
  {
    id: "engine-agent/basic",
    name: "Basic Agent",
    description: "A basic custom engine agent",
    category: "custom-engine-agent",
    languages: ["typescript", "javascript", "python"],
    scaffoldFn: makeEngineAgentScaffoldFn(EngineAgentTemplateNames.Basic),
    displayOrder: 1,
  },
  {
    id: "engine-agent/function-calling",
    name: "Agent with Function Calling",
    description: "A custom engine agent with function calling capabilities",
    category: "custom-engine-agent",
    languages: ["typescript", "javascript", "python", "csharp"],
    scaffoldFn: makeEngineAgentScaffoldFn(EngineAgentTemplateNames.FunctionCalling),
    displayOrder: 2,
  },
  {
    id: "engine-agent/multi-agent",
    name: "Multi-Agent Orchestration",
    description: "Multiple custom engine agents working together",
    category: "custom-engine-agent",
    languages: ["csharp"],
    scaffoldFn: makeEngineAgentScaffoldFn(EngineAgentTemplateNames.MultiAgent),
    displayOrder: 3,
  },
  {
    id: "engine-agent/teams-collaborator",
    name: "Teams Collaborator Agent",
    description: "A custom engine agent for Teams collaboration",
    category: "custom-engine-agent",
    languages: ["typescript", "csharp"],
    scaffoldFn: makeEngineAgentScaffoldFn(EngineAgentTemplateNames.TeamsCollaborator),
    displayOrder: 4,
  },
];
