// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { TemplateDescriptor, TemplateActionOptions } from "../types";
import type { AtkContext } from "../../core/context";
import { scaffoldTemplates } from "../scaffold/scaffolder";
import { getTemplateReplaceMap } from "../scaffold/replaceMap";
import { type TemplateInfo, convertToLangKey } from "../scaffold/types";

/**
 * Bot template artifact names matching the template repository folder names.
 */
export const BotTemplateNames = {
  Echo: "default-bot",
  NotificationExpress: "notification-express",
  NotificationWebApi: "notification-webapi",
  NotificationHttpTrigger: "notification-http-trigger",
  NotificationTimerTrigger: "notification-timer-trigger",
  NotificationHttpAndTimer: "notification-http-timer-trigger",
  CommandResponse: "command-and-response",
  Workflow: "workflow",
} as const;

/**
 * Create a standard scaffold function for bot templates.
 * All bot templates share the same scaffold pipeline with no extra variables.
 */
function makeBotScaffoldFn(templateName: string) {
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
 * All Bot template descriptors.
 */
export const botTemplateDescriptors: TemplateDescriptor[] = [
  {
    id: "bot/echo",
    name: "Echo Bot",
    description: "A simple bot that echoes user messages",
    category: "bot",
    languages: ["typescript", "javascript", "python", "csharp"],
    scaffoldFn: makeBotScaffoldFn(BotTemplateNames.Echo),
    displayOrder: 1,
  },
  {
    id: "bot/notification-express",
    name: "Notification Bot (Express)",
    description: "Send proactive notifications using Express.js",
    category: "bot",
    languages: ["typescript", "javascript"],
    scaffoldFn: makeBotScaffoldFn(BotTemplateNames.NotificationExpress),
    displayOrder: 2,
  },
  {
    id: "bot/notification-webapi",
    name: "Notification Bot (Web API)",
    description: "Send proactive notifications using ASP.NET Web API",
    category: "bot",
    languages: ["csharp"],
    scaffoldFn: makeBotScaffoldFn(BotTemplateNames.NotificationWebApi),
    displayOrder: 3,
  },
  {
    id: "bot/notification-http-trigger",
    name: "Notification Bot (HTTP Trigger)",
    description: "Send proactive notifications using Azure Functions HTTP trigger",
    category: "bot",
    languages: ["typescript", "javascript"],
    scaffoldFn: makeBotScaffoldFn(BotTemplateNames.NotificationHttpTrigger),
    displayOrder: 4,
  },
  {
    id: "bot/notification-timer-trigger",
    name: "Notification Bot (Timer Trigger)",
    description: "Send proactive notifications using Azure Functions timer trigger",
    category: "bot",
    languages: ["typescript", "javascript"],
    scaffoldFn: makeBotScaffoldFn(BotTemplateNames.NotificationTimerTrigger),
    displayOrder: 5,
  },
  {
    id: "bot/notification-http-and-timer",
    name: "Notification Bot (HTTP + Timer Trigger)",
    description: "Send proactive notifications using both HTTP and timer triggers",
    category: "bot",
    languages: ["typescript", "javascript"],
    scaffoldFn: makeBotScaffoldFn(BotTemplateNames.NotificationHttpAndTimer),
    displayOrder: 6,
  },
  {
    id: "bot/command-response",
    name: "Command and Response Bot",
    description: "A bot that responds to specific commands",
    category: "bot",
    languages: ["typescript", "javascript"],
    scaffoldFn: makeBotScaffoldFn(BotTemplateNames.CommandResponse),
    displayOrder: 7,
  },
  {
    id: "bot/workflow",
    name: "Workflow Bot",
    description: "A bot with multi-step Adaptive Card workflows",
    category: "bot",
    languages: ["typescript", "javascript"],
    scaffoldFn: makeBotScaffoldFn(BotTemplateNames.Workflow),
    displayOrder: 8,
  },
];
