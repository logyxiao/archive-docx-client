import type { ProcessUserFields, ResolvedProcessFields } from "./types";

export function resolveProcessFields(userFields: ProcessUserFields, defaultUnit: string): ResolvedProcessFields {
  return {
    generalContractorUnit: firstNonBlank(userFields.generalContractorUnit, defaultUnit),
    generalContractorProjectManager: firstNonBlank(userFields.generalContractorProjectManager, userFields.projectManager),
    generalContractorTechnicalLeader: firstNonBlank(userFields.generalContractorTechnicalLeader, userFields.projectTechnicalLeader),
    constructionUnit: firstNonBlank(userFields.constructionUnit, defaultUnit),
    constructionProjectManager: firstNonBlank(userFields.constructionProjectManager, userFields.projectManager),
    constructionTechnicalLeader: firstNonBlank(userFields.constructionTechnicalLeader, userFields.projectTechnicalLeader),
    subcontractorUnit: firstNonBlank(userFields.subcontractorUnit),
    subcontractorProjectManager: firstNonBlank(userFields.subcontractorProjectManager),
    subcontractorContent: firstNonBlank(userFields.subcontractorContent),
    supervisionDepartment: firstNonBlank(userFields.supervisionDepartment),
  };
}

function firstNonBlank(...values: Array<string | undefined>): string {
  return values.find((value) => value?.trim())?.trim() ?? "";
}
