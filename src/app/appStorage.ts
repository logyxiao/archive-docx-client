import {
  normalizeProcessTemplateCategories,
  PROCESS_TEMPLATE_CATEGORY_IDS,
  type ProcessTemplateCategoryId,
} from "../lib/processDocs";
import { PROCESS_FIELDS_KEY } from "./appConstants";

export interface SavedProcessFields {
  projectName: string;
  generalContractorUnit: string;
  generalContractorManager: string;
  generalContractorTechnicalLeader: string;
  constructionUnit: string;
  constructionManager: string;
  constructionTechnicalLeader: string;
  subcontractorUnit: string;
  subcontractorManager: string;
  subcontractorContent: string;
  supervisionDepartment: string;
}

export const EMPTY_PROCESS_FIELDS: SavedProcessFields = {
  projectName: "",
  generalContractorUnit: "",
  generalContractorManager: "",
  generalContractorTechnicalLeader: "",
  constructionUnit: "",
  constructionManager: "",
  constructionTechnicalLeader: "",
  subcontractorUnit: "",
  subcontractorManager: "",
  subcontractorContent: "",
  supervisionDepartment: "",
};

export function loadSavedProcessFields(): SavedProcessFields {
  try {
    const rawValue = localStorage.getItem(PROCESS_FIELDS_KEY);
    if (!rawValue) {
      return EMPTY_PROCESS_FIELDS;
    }

    const parsed = JSON.parse(rawValue) as Partial<SavedProcessFields> & { subcontractorTechnicalLeader?: string };
    return {
      ...EMPTY_PROCESS_FIELDS,
      ...Object.fromEntries(
        Object.entries(parsed).map(([key, value]) => [key, typeof value === "string" ? value : ""]),
      ),
      subcontractorContent:
        typeof parsed.subcontractorContent === "string"
          ? parsed.subcontractorContent
          : typeof parsed.subcontractorTechnicalLeader === "string"
            ? parsed.subcontractorTechnicalLeader
            : "",
    };
  } catch {
    return EMPTY_PROCESS_FIELDS;
  }
}

export function loadSavedProcessTemplateCategories(storageKey: string): ProcessTemplateCategoryId[] {
  try {
    const rawValue = localStorage.getItem(storageKey);
    if (!rawValue) {
      return [...PROCESS_TEMPLATE_CATEGORY_IDS];
    }

    const parsed = JSON.parse(rawValue);
    if (!Array.isArray(parsed)) {
      return [...PROCESS_TEMPLATE_CATEGORY_IDS];
    }

    return [...normalizeProcessTemplateCategories(parsed)];
  } catch {
    return [...PROCESS_TEMPLATE_CATEGORY_IDS];
  }
}
