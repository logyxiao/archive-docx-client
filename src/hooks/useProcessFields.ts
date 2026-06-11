import { useEffect, useMemo, useState } from "react";
import { PROCESS_FIELDS_KEY } from "../app/appConstants";
import { loadSavedProcessFields, type SavedProcessFields } from "../app/appStorage";

export function useProcessFields() {
  const savedProcessFields = useMemo(loadSavedProcessFields, []);
  const [projectName, setProjectName] = useState(savedProcessFields.projectName);
  const [generalContractorUnit, setGeneralContractorUnit] = useState(savedProcessFields.generalContractorUnit);
  const [generalContractorManager, setGeneralContractorManager] = useState(savedProcessFields.generalContractorManager);
  const [generalContractorTechnicalLeader, setGeneralContractorTechnicalLeader] = useState(
    savedProcessFields.generalContractorTechnicalLeader,
  );
  const [constructionUnit, setConstructionUnit] = useState(savedProcessFields.constructionUnit);
  const [constructionManager, setConstructionManager] = useState(savedProcessFields.constructionManager);
  const [constructionTechnicalLeader, setConstructionTechnicalLeader] = useState(
    savedProcessFields.constructionTechnicalLeader,
  );
  const [subcontractorUnit, setSubcontractorUnit] = useState(savedProcessFields.subcontractorUnit);
  const [subcontractorManager, setSubcontractorManager] = useState(savedProcessFields.subcontractorManager);
  const [subcontractorContent, setSubcontractorContent] = useState(savedProcessFields.subcontractorContent);
  const [supervisionDepartment, setSupervisionDepartment] = useState(savedProcessFields.supervisionDepartment);

  const fields: SavedProcessFields = {
    projectName,
    generalContractorUnit,
    generalContractorManager,
    generalContractorTechnicalLeader,
    constructionUnit,
    constructionManager,
    constructionTechnicalLeader,
    subcontractorUnit,
    subcontractorManager,
    subcontractorContent,
    supervisionDepartment,
  };

  useEffect(() => {
    localStorage.setItem(PROCESS_FIELDS_KEY, JSON.stringify(fields));
  }, [
    projectName,
    generalContractorUnit,
    generalContractorManager,
    generalContractorTechnicalLeader,
    constructionUnit,
    constructionManager,
    constructionTechnicalLeader,
    subcontractorUnit,
    subcontractorManager,
    subcontractorContent,
    supervisionDepartment,
  ]);

  function processUserFields() {
    return {
      projectName: projectName.trim(),
      generalContractorUnit: generalContractorUnit.trim(),
      generalContractorProjectManager: generalContractorManager.trim(),
      generalContractorTechnicalLeader: generalContractorTechnicalLeader.trim(),
      constructionUnit: constructionUnit.trim(),
      constructionProjectManager: constructionManager.trim(),
      constructionTechnicalLeader: constructionTechnicalLeader.trim(),
      subcontractorUnit: subcontractorUnit.trim(),
      subcontractorProjectManager: subcontractorManager.trim(),
      subcontractorContent: subcontractorContent.trim(),
      supervisionDepartment: supervisionDepartment.trim(),
    };
  }

  return {
    fields,
    setters: {
      setProjectName,
      setGeneralContractorUnit,
      setGeneralContractorManager,
      setGeneralContractorTechnicalLeader,
      setConstructionUnit,
      setConstructionManager,
      setConstructionTechnicalLeader,
      setSubcontractorUnit,
      setSubcontractorManager,
      setSubcontractorContent,
      setSupervisionDepartment,
    },
    processUserFields,
  };
}
