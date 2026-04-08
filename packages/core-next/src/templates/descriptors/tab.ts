// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { TemplateDescriptor, TemplateActionOptions } from "../types";
import type { AtkContext } from "../../core/context";
import { scaffoldTemplates } from "../scaffold/scaffolder";
import { getTemplateReplaceMap } from "../scaffold/replaceMap";
import { type TemplateInfo, convertToLangKey } from "../scaffold/types";

/**
 * Tab template artifact names matching the template repository folder names.
 * These must match the directory names inside template zip archives.
 */
export const TabTemplateNames = {
  Basic: "basic-tab",
  /** SSO NAA tab — no local folder yet; uses remote download */
  SSO: "sso-tab-naa",
  /** Dashboard tab — no local folder yet; uses remote download */
  Dashboard: "dashboard-tab",
  SSRBasic: "non-sso-tab-ssr",
  SSRSSO: "sso-tab-ssr",
} as const;

/**
 * Create a standard scaffold function for tab templates.
 */
function makeTabScaffoldFn(templateName: string) {
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
 * All Tab template descriptors.
 */
export const tabTemplateDescriptors: TemplateDescriptor[] = [
  {
    id: "tab/basic",
    name: "Basic Tab",
    description: "A simple tab app without SSO",
    category: "tab",
    languages: ["typescript", "csharp"],
    scaffoldFn: makeTabScaffoldFn(TabTemplateNames.Basic),
    displayOrder: 1,
  },
  {
    id: "tab/sso",
    name: "SSO Tab",
    description: "A tab app with Single Sign-On using NAA",
    category: "tab",
    languages: ["typescript"],
    scaffoldFn: makeTabScaffoldFn(TabTemplateNames.SSO),
    displayOrder: 2,
  },
  {
    id: "tab/dashboard",
    name: "Dashboard Tab",
    description: "A tab app with a dashboard layout",
    category: "tab",
    languages: ["typescript", "javascript"],
    scaffoldFn: makeTabScaffoldFn(TabTemplateNames.Dashboard),
    displayOrder: 3,
  },
  {
    id: "tab/ssr-basic",
    name: "SSR Tab (Basic)",
    description: "A server-side rendered tab without SSO",
    category: "tab",
    languages: ["csharp"],
    scaffoldFn: makeTabScaffoldFn(TabTemplateNames.SSRBasic),
    displayOrder: 4,
  },
  {
    id: "tab/ssr-sso",
    name: "SSR Tab (SSO)",
    description: "A server-side rendered tab with Single Sign-On",
    category: "tab",
    languages: ["csharp"],
    scaffoldFn: makeTabScaffoldFn(TabTemplateNames.SSRSSO),
    displayOrder: 5,
  },
];
