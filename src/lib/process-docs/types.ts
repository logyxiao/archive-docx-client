import type { GeneratedFile } from "../types";

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
}

export interface ProcessGenerationResult {
  files: GeneratedFile[];
  skipped: string[];
  errors: string[];
}

export interface ProcessUserFields {
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
  subcontractorTechnicalLeader?: string;
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
  subcontractorTechnicalLeader: string;
  supervisionDepartment: string;
}
