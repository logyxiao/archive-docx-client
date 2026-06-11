import { SlidersHorizontal, X } from "lucide-react";

interface ProcessInfoModalProps {
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
  onProjectNameChange: (value: string) => void;
  onGeneralContractorUnitChange: (value: string) => void;
  onGeneralContractorManagerChange: (value: string) => void;
  onGeneralContractorTechnicalLeaderChange: (value: string) => void;
  onConstructionUnitChange: (value: string) => void;
  onConstructionManagerChange: (value: string) => void;
  onConstructionTechnicalLeaderChange: (value: string) => void;
  onSubcontractorUnitChange: (value: string) => void;
  onSubcontractorManagerChange: (value: string) => void;
  onSubcontractorContentChange: (value: string) => void;
  onSupervisionDepartmentChange: (value: string) => void;
  onClose: () => void;
}

export function ProcessInfoModal({
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
  onProjectNameChange,
  onGeneralContractorUnitChange,
  onGeneralContractorManagerChange,
  onGeneralContractorTechnicalLeaderChange,
  onConstructionUnitChange,
  onConstructionManagerChange,
  onConstructionTechnicalLeaderChange,
  onSubcontractorUnitChange,
  onSubcontractorManagerChange,
  onSubcontractorContentChange,
  onSupervisionDepartmentChange,
  onClose,
}: ProcessInfoModalProps) {
  return (
    <div className="modal-backdrop" onClick={onClose}>
      <section
        className="process-info-modal"
        role="dialog"
        aria-modal="true"
        aria-labelledby="process-info-title"
        onClick={(event) => event.stopPropagation()}
      >
        <div className="modal-heading">
          <div className="section-heading">
            <SlidersHorizontal size={19} />
            <h2 id="process-info-title">过程资料生成信息</h2>
          </div>
          <button className="icon-button" onClick={onClose} title="关闭" aria-label="关闭生成信息">
            <X size={18} />
          </button>
        </div>
        <div className="process-field-grid modal-process-field-grid">
          <label className="text-field">
            <span>工程名称</span>
            <input value={projectName} onChange={(event) => onProjectNameChange(event.currentTarget.value)} />
          </label>
          <label className="text-field">
            <span>总承包单位</span>
            <input
              value={generalContractorUnit}
              onChange={(event) => onGeneralContractorUnitChange(event.currentTarget.value)}
            />
          </label>
          <label className="text-field">
            <span>总承包单位项目负责人</span>
            <input
              value={generalContractorManager}
              onChange={(event) => onGeneralContractorManagerChange(event.currentTarget.value)}
            />
          </label>
          <label className="text-field">
            <span>总承包单位项目技术负责人</span>
            <input
              value={generalContractorTechnicalLeader}
              onChange={(event) => onGeneralContractorTechnicalLeaderChange(event.currentTarget.value)}
            />
          </label>
          <label className="text-field">
            <span>施工单位</span>
            <input value={constructionUnit} onChange={(event) => onConstructionUnitChange(event.currentTarget.value)} />
          </label>
          <label className="text-field">
            <span>施工单位项目负责人</span>
            <input value={constructionManager} onChange={(event) => onConstructionManagerChange(event.currentTarget.value)} />
          </label>
          <label className="text-field">
            <span>施工单位项目技术负责人</span>
            <input
              value={constructionTechnicalLeader}
              onChange={(event) => onConstructionTechnicalLeaderChange(event.currentTarget.value)}
            />
          </label>
          <label className="text-field">
            <span>分包单位</span>
            <input value={subcontractorUnit} onChange={(event) => onSubcontractorUnitChange(event.currentTarget.value)} />
          </label>
          <label className="text-field">
            <span>分包单位项目负责人</span>
            <input
              value={subcontractorManager}
              onChange={(event) => onSubcontractorManagerChange(event.currentTarget.value)}
            />
          </label>
          <label className="text-field">
            <span>分包内容</span>
            <input
              value={subcontractorContent}
              onChange={(event) => onSubcontractorContentChange(event.currentTarget.value)}
            />
          </label>
          <label className="text-field">
            <span>监理项目部</span>
            <input
              value={supervisionDepartment}
              onChange={(event) => onSupervisionDepartmentChange(event.currentTarget.value)}
            />
          </label>
        </div>
        <div className="modal-footer">
          <button className="primary-button" type="button" onClick={onClose}>
            完成
          </button>
        </div>
      </section>
    </div>
  );
}
