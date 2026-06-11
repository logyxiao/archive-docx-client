import { loadUserProcessTemplates, readBinaryFile } from "../tauriFiles";
import { PROCESS_TEMPLATE_ROOT } from "./constants";
import type { ProcessTemplate, ProcessTemplateManifest } from "./types";

export async function loadProcessManifest(): Promise<ProcessTemplateManifest> {
  const response = await fetch(`${PROCESS_TEMPLATE_ROOT}/manifest.json`);
  if (!response.ok) {
    throw new Error("无法加载过程资料模板清单");
  }

  const manifest = await response.json() as ProcessTemplateManifest;
  let userTemplates: ProcessTemplate[] = [];
  try {
    userTemplates = await loadUserProcessTemplates();
  } catch {
    userTemplates = [];
  }

  return {
    templates: mergeProcessTemplates(manifest.templates, userTemplates),
  };
}

export async function loadProcessTemplate(template: ProcessTemplate | string): Promise<ArrayBuffer> {
  if (typeof template !== "string" && template.userTemplatePath) {
    const bytes = await readBinaryFile(template.userTemplatePath);
    return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  }

  const templateFile = typeof template === "string" ? template : template.templateFile;
  const response = await fetch(`${PROCESS_TEMPLATE_ROOT}/${encodeURIComponent(templateFile)}`);
  if (!response.ok) {
    throw new Error(`无法加载过程资料模板：${templateFile}`);
  }

  return response.arrayBuffer();
}

function mergeProcessTemplates(builtInTemplates: ProcessTemplate[], userTemplates: ProcessTemplate[]): ProcessTemplate[] {
  const result = [...builtInTemplates];
  const seenUserPaths = new Set<string>();
  for (const template of userTemplates) {
    if (!template.userTemplatePath || seenUserPaths.has(template.userTemplatePath)) {
      continue;
    }
    seenUserPaths.add(template.userTemplatePath);
    result.push(template);
  }
  return result;
}
