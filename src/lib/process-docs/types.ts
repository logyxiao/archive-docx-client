import type { GeneratedFile } from "../types";
import type { ProcessTemplateCategoryId } from "./categories";

export type ProcessTemplateModule = "process" | "switch-station" | "collector-line";

export interface ProcessTemplateManifest {
  templates: ProcessTemplate[];
}

export interface ProcessTemplate {
  sequence: number;
  kind: "docx" | "xlsx";
  originalName: string;
  templateFile: string;
  outputExtension: ".docx" | ".xlsx";
  outputFileCodeOverride?: string;
}

export interface GenerateProcessOptions {
  selectedCodes: string[];
  outputDir: string;
  userFields?: ProcessUserFields;
  selectedTemplateCategories?: ProcessTemplateCategoryId[];
  templateModule?: ProcessTemplateModule;
}

export interface ProcessTemplateMatch {
  templateModule: ProcessTemplateModule;
  template: ProcessTemplate;
}

export interface ProcessTemplateSelection extends ProcessTemplateMatch {
  archiveCode: string;
  fileCode: string;
  sequence: string;
}

export interface GenerateSelectedProcessOptions {
  outputDir: string;
  userFields?: ProcessUserFields;
  selections: ProcessTemplateSelection[];
}

export interface ProcessGenerationResult {
  files: GeneratedFile[];
  skipped: string[];
  errors: string[];
}

export interface ProcessUserFields {
  projectName?: string;
  projectManager?: string;
  projectTechnicalLeader?: string;
  generalContractorUnit?: string;
  generalContractorProjectManager?: string;
  generalContractorTechnicalLeader?: string;
  constructionUnit?: string;
  constructionProjectManager?: string;
  constructionTechnicalLeader?: string;
  subcontractorUnit?: string;
  subcontractorProjectManager?: string;
  subcontractorContent?: string;
  supervisionDepartment?: string;
}

export interface ResolvedProcessFields {
  generalContractorUnit: string;
  generalContractorProjectManager: string;
  generalContractorTechnicalLeader: string;
  constructionUnit: string;
  constructionProjectManager: string;
  constructionTechnicalLeader: string;
  subcontractorUnit: string;
  subcontractorProjectManager: string;
  subcontractorContent: string;
  supervisionDepartment: string;
}
