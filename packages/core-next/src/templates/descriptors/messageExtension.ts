// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { TemplateDescriptor, TemplateActionOptions } from "../types";
import type { AtkContext } from "../../core/context";
import { scaffoldTemplates } from "../scaffold/scaffolder";
import { getTemplateReplaceMap } from "../scaffold/replaceMap";
import { type TemplateInfo, convertToLangKey } from "../scaffold/types";

/**
 * Message Extension template artifact names matching the template repository folder names.
 * These must match the directory names inside template zip archives.
 */
export const MessageExtensionTemplateNames = {
  SearchBased: "message-extension-v2",
  /** API key variant — no dedicated local folder; remote only */
  SearchApiKey: "copilot-plugin-from-scratch-api-key",
  /** API SSO variant — no dedicated local folder; remote only */
  ApiSSO: "api-message-extension-sso",
  /** M365 variant — uses same message-extension-v2 folder */
  M365: "message-extension-v2",
  /** Action variant — no dedicated local folder; remote only */
  Action: "message-extension-action",
  /** Link unfurling — no dedicated local folder; remote only */
  LinkUnfurling: "link-unfurling",
} as const;

/**
 * Create a standard scaffold function for message extension templates.
 */
function makeMEScaffoldFn(templateName: string) {
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
 * All Message Extension template descriptors.
 */
export const messageExtensionTemplateDescriptors: TemplateDescriptor[] = [
  {
    id: "me/search-based",
    name: "Search-based Message Extension",
    description: "A search-based message extension for Copilot",
    category: "message-extension",
    languages: ["typescript", "javascript", "csharp"],
    scaffoldFn: makeMEScaffoldFn(MessageExtensionTemplateNames.SearchBased),
    displayOrder: 1,
  },
  {
    id: "me/search-api-key",
    name: "Search Message Extension (API Key)",
    description: "A search-based message extension with API key authentication",
    category: "message-extension",
    languages: ["typescript", "javascript"],
    scaffoldFn: makeMEScaffoldFn(MessageExtensionTemplateNames.SearchApiKey),
    displayOrder: 2,
  },
  {
    id: "me/api-sso",
    name: "API Message Extension (SSO)",
    description: "An API-based message extension with Single Sign-On",
    category: "message-extension",
    languages: ["typescript"],
    scaffoldFn: makeMEScaffoldFn(MessageExtensionTemplateNames.ApiSSO),
    displayOrder: 3,
  },
  {
    id: "me/m365",
    name: "M365 Message Extension",
    description: "A message extension for Microsoft 365",
    category: "message-extension",
    languages: ["typescript", "python"],
    scaffoldFn: makeMEScaffoldFn(MessageExtensionTemplateNames.M365),
    displayOrder: 4,
  },
  {
    id: "me/action",
    name: "Action Message Extension",
    description: "A message extension that performs actions",
    category: "message-extension",
    languages: ["typescript"],
    scaffoldFn: makeMEScaffoldFn(MessageExtensionTemplateNames.Action),
    displayOrder: 5,
  },
  {
    id: "me/link-unfurling",
    name: "Link Unfurling",
    description: "A message extension that unfurls links into rich previews",
    category: "message-extension",
    languages: ["typescript"],
    scaffoldFn: makeMEScaffoldFn(MessageExtensionTemplateNames.LinkUnfurling),
    displayOrder: 6,
  },
];
