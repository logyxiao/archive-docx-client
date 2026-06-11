import { Info, X } from "lucide-react";
import {
  PROCESS_TEMPLATE_PLACEHOLDERS,
  TEMPLATE_CONFIG_FIELDS,
  TEMPLATE_IMPORT_STEPS,
} from "../app/appConstants";

interface TemplateGuideModalProps {
  copiedPlaceholder: string;
  onCopyPlaceholder: (value: string) => void;
  onClose: () => void;
}

export function TemplateGuideModal({ copiedPlaceholder, onCopyPlaceholder, onClose }: TemplateGuideModalProps) {
  return (
    <div className="modal-backdrop" onClick={onClose}>
      <section
        className="template-guide-modal"
        role="dialog"
        aria-modal="true"
        aria-labelledby="template-guide-title"
        onClick={(event) => event.stopPropagation()}
      >
        <div className="modal-heading">
          <div className="section-heading">
            <Info size={19} />
            <h2 id="template-guide-title">导入教程</h2>
          </div>
          <button className="icon-button" onClick={onClose} title="关闭" aria-label="关闭导入教程">
            <X size={18} />
          </button>
        </div>
        <div className="template-guide-section">
          <div className="template-guide-card">
            <div className="template-section-heading">
              <span>导入步骤</span>
            </div>
            <ol className="template-guide-list">
              {TEMPLATE_IMPORT_STEPS.map((step) => (
                <li key={step}>{step}</li>
              ))}
            </ol>
          </div>
          <div className="template-guide-card">
            <div className="template-section-heading">
              <span>可填写配置</span>
            </div>
            <ul className="template-guide-list">
              {TEMPLATE_CONFIG_FIELDS.map((field) => (
                <li key={field}>{field}</li>
              ))}
            </ul>
          </div>
          <div className="template-guide-card">
            <div className="template-section-heading">
              <span>常用占位符</span>
            </div>
            <div className="placeholder-grid">
              {PROCESS_TEMPLATE_PLACEHOLDERS.map((placeholder) => {
                const placeholderText = `{{${placeholder}}}`;
                const isCopied = copiedPlaceholder === placeholderText;
                return (
                  <button
                    key={placeholder}
                    type="button"
                    className={`placeholder-copy ${isCopied ? "copied" : ""}`}
                    onClick={() => onCopyPlaceholder(placeholderText)}
                    title={`复制 ${placeholderText}`}
                  >
                    <code>{placeholderText}</code>
                    <span>{isCopied ? "已复制" : "复制"}</span>
                  </button>
                );
              })}
            </div>
          </div>
        </div>
      </section>
    </div>
  );
}
