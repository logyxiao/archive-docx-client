import type { ProcessTemplate } from "./types";

export function withPreferredUserTemplates(userTemplates: ProcessTemplate[], baseTemplates: ProcessTemplate[]): ProcessTemplate[] {
  if (userTemplates.length === 0) {
    return baseTemplates;
  }

  const result = [...userTemplates];
  const seen = new Set(result.map(templateIdentity));
  for (const template of baseTemplates) {
    const identity = templateIdentity(template);
    if (seen.has(identity)) {
      continue;
    }
    seen.add(identity);
    result.push(template);
  }
  return result;
}

export function matchingUserTemplates(title: string, templates: ProcessTemplate[]): ProcessTemplate[] {
  const normalizedTitle = normalizeTemplateMatchText(title);
  if (!normalizedTitle) {
    return [];
  }

  return templates.filter((template) => {
    if (!template.userTemplatePath || template.enabled === false) {
      return false;
    }

    const keywords = normalizeTemplateKeywords(template.matchKeywords);
    if (keywords.length > 0) {
      return template.matchMode === "all"
        ? keywords.every((keyword) => normalizedTitle.includes(keyword))
        : keywords.some((keyword) => normalizedTitle.includes(keyword));
    }

    return [template.templateFile, template.originalName].some((name) => {
      const normalizedTemplateName = normalizeTemplateMatchText(fileNameWithoutExtension(name));
      return normalizedTemplateName.length >= 2
        && (normalizedTitle.includes(normalizedTemplateName) || normalizedTemplateName.includes(normalizedTitle));
    });
  });
}

function normalizeTemplateKeywords(keywords: string[] | undefined): string[] {
  if (!Array.isArray(keywords)) {
    return [];
  }

  const result: string[] = [];
  for (const keyword of keywords) {
    const normalized = normalizeTemplateMatchText(keyword);
    if (!normalized || result.includes(normalized)) {
      continue;
    }
    result.push(normalized);
  }
  return result;
}

function fileNameWithoutExtension(value: string): string {
  return value.replace(/\.[^.]+$/, "");
}

function normalizeTemplateMatchText(value: string): string {
  return value
    .replace(/\{\{[^}]+}}/g, "")
    .replace(/^[0-9]+[-_、.\s]*/, "")
    .replace(/（.*?）|\(.*?\)/g, "")
    .replace(/[^\p{Script=Han}A-Za-z0-9]/gu, "")
    .toLowerCase();
}

function templateIdentity(template: ProcessTemplate): string {
  return template.userTemplatePath ?? `${template.templateFile}\u0000${template.originalName}\u0000${template.outputFileCodeOverride ?? ""}`;
}
