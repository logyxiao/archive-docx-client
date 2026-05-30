import type { GeneratedFile } from "../types";
import type { ProcessTemplateCategoryId } from "./categories";

export interface ProcessTemplateManifest {
  templates: ProcessTemplate[];
}

export interface ProcessTemplate {
  sequence: number;
  kind: "docx" | "xlsx";
  originalName: string;
  templateFile: string;
  outputExtension: ".docx" | ".xlsx";
}

export interface GenerateProcessOptions {
  selectedCodes: string[];
  outputDir: string;
  userFields?: ProcessUserFields;
  selectedTemplateCategories?: ProcessTemplateCategoryId[];
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
