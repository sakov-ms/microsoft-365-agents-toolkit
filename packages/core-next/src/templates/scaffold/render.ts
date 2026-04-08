// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import Mustache from "mustache";

/** Extension for template files that get Mustache processing. */
const TEMPLATE_FILE_EXT = ".tpl";

/** Standard Mustache delimiters. */
const DELIMITERS: [string, string] = ["{{", "}}"];

/** Legacy delimiters (pre-v3). */
const LEGACY_DELIMITERS: [string, string] = ["{%", "%}"];

/**
 * Render Mustache variables in a file name and strip the .tpl extension.
 */
export function renderTemplateFileName(
  fileName: string,
  _data: Buffer,
  variables: Record<string, string>
): string {
  // Disable HTML escaping for file names
  const savedEscape = Mustache.escape;
  Mustache.escape = (value: unknown) => String(value);
  try {
    const rendered = Mustache.render(fileName, variables, undefined, DELIMITERS);
    return rendered.endsWith(TEMPLATE_FILE_EXT)
      ? rendered.slice(0, -TEMPLATE_FILE_EXT.length)
      : rendered;
  } finally {
    Mustache.escape = savedEscape;
  }
}

/**
 * Render Mustache variables in file content.
 *
 * Only processes files ending in `.tpl`. Non-tpl files are returned as raw Buffer.
 * Undefined variables are preserved as literal `{{var}}` text (not blanked out).
 */
export function renderTemplateFileData(
  fileName: string,
  data: Buffer,
  variables: Record<string, string>
): string | Buffer {
  if (!fileName.endsWith(TEMPLATE_FILE_EXT)) {
    return data; // binary/non-template file — pass through
  }

  const template = data.toString("utf-8");

  // Disable HTML escaping for template content
  const savedEscape = Mustache.escape;
  Mustache.escape = (value: unknown) => String(value);
  try {
    // First pass: standard delimiters
    const pass1 = renderPreservingUndefined(template, variables, DELIMITERS);
    // Second pass: legacy delimiters
    return renderPreservingUndefined(pass1, variables, LEGACY_DELIMITERS);
  } finally {
    Mustache.escape = savedEscape;
  }
}

/**
 * Render Mustache template while preserving undefined/null variables as literal text.
 *
 * Parses the template, converts tokens for missing variables from "name" to "text"
 * so they render as the raw placeholder (e.g. `{{appName}}` stays if appName is undefined).
 */
function renderPreservingUndefined(
  template: string,
  variables: Record<string, string>,
  delimiters: [string, string]
): string {
  const tokens = Mustache.parse(template, delimiters);
  escapeUndefinedTokens(tokens, variables, delimiters);
  // Reconstruct template from modified tokens and render
  const escaped = rebuildTemplate(tokens, delimiters);
  return Mustache.render(escaped, variables, undefined, delimiters);
}

/**
 * Walk the token tree and convert "name" tokens to "text" tokens
 * if the variable is not defined in the view.
 */
function escapeUndefinedTokens(
  tokens: Array<[string, string, number, number, ...unknown[]]>,
  variables: Record<string, string>,
  delimiters: [string, string]
): void {
  for (const token of tokens) {
    const [type, name] = token;
    if (type === "name") {
      if (!(name in variables) || variables[name] === undefined || variables[name] === null) {
        // Convert to text token — will emit the literal placeholder
        token[0] = "text";
        token[1] = `${delimiters[0]}${name}${delimiters[1]}`;
      }
    } else if (type === "#" || type === "^") {
      // Section/inverted section — recurse into children
      const children = token[4] as
        | Array<[string, string, number, number, ...unknown[]]>
        | undefined;
      if (children) {
        escapeUndefinedTokens(children, variables, delimiters);
      }
    }
  }
}

/**
 * Rebuild a template string from parsed tokens.
 */
function rebuildTemplate(
  tokens: Array<[string, string, number, number, ...unknown[]]>,
  delimiters: [string, string]
): string {
  let result = "";
  for (const token of tokens) {
    const [type, value] = token;
    if (type === "text") {
      result += value;
    } else if (type === "name") {
      result += `${delimiters[0]}${value}${delimiters[1]}`;
    } else if (type === "#") {
      const children = token[4] as Array<[string, string, number, number, ...unknown[]]>;
      result += `${delimiters[0]}#${value}${delimiters[1]}`;
      result += rebuildTemplate(children, delimiters);
      result += `${delimiters[0]}/${value}${delimiters[1]}`;
    } else if (type === "^") {
      const children = token[4] as Array<[string, string, number, number, ...unknown[]]>;
      result += `${delimiters[0]}^${value}${delimiters[1]}`;
      result += rebuildTemplate(children, delimiters);
      result += `${delimiters[0]}/${value}${delimiters[1]}`;
    }
  }
  return result;
}
