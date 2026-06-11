import { FolderOpen, Info, Loader2, Save, Search, Settings2, Trash2, Upload, X } from "lucide-react";
import {
  draftFromUserTemplate,
  userTemplateKey,
  type TemplateDirectoryEntry,
  type UserTemplateDraft,
} from "../app/templateOptions";
import {
  PROCESS_TEMPLATE_CATEGORIES,
  type ProcessTemplate,
  type ProcessTemplateCategoryId,
  type ProcessTemplateMatchMode,
  type ProcessTemplateModule,
} from "../lib/processDocs";

interface TemplateManagerModalProps {
  isImporting: boolean;
  templateDirectoryEntries: TemplateDirectoryEntry[];
  filteredTemplateDirectoryEntries: TemplateDirectoryEntry[];
  templateDirectoryError: string;
  templateDirectoryQuery: string;
  userProcessTemplates: ProcessTemplate[];
  templateDrafts: Record<string, UserTemplateDraft>;
  savingTemplateKeys: Record<string, boolean>;
  deletingTemplateKeys: Record<string, boolean>;
  onClose: () => void;
  onOpenGuide: () => void;
  onImportTemplates: () => void;
  onTemplateDirectoryQueryChange: (value: string) => void;
  onOpenTemplateDirectory: (entry: TemplateDirectoryEntry) => void;
  onUpdateTemplateDraft: (template: ProcessTemplate, patch: Partial<UserTemplateDraft>) => void;
  onDeleteTemplate: (template: ProcessTemplate) => void;
  onSaveTemplate: (template: ProcessTemplate) => void;
}

export function TemplateManagerModal({
  isImporting,
  templateDirectoryEntries,
  filteredTemplateDirectoryEntries,
  templateDirectoryError,
  templateDirectoryQuery,
  userProcessTemplates,
  templateDrafts,
  savingTemplateKeys,
  deletingTemplateKeys,
  onClose,
  onOpenGuide,
  onImportTemplates,
  onTemplateDirectoryQueryChange,
  onOpenTemplateDirectory,
  onUpdateTemplateDraft,
  onDeleteTemplate,
  onSaveTemplate,
}: TemplateManagerModalProps) {
  return (
    <div className="modal-backdrop" onClick={onClose}>
      <section
        className="template-manager-modal"
        role="dialog"
        aria-modal="true"
        aria-labelledby="template-manager-title"
        onClick={(event) => event.stopPropagation()}
      >
        <div className="modal-heading">
          <div className="section-heading">
            <Settings2 size={19} />
            <h2 id="template-manager-title">模板管理</h2>
          </div>
          <div className="modal-actions">
            <button type="button" className="secondary-button modal-nav-button" onClick={onOpenGuide}>
              <Info size={15} />
              导入教程
            </button>
            <button
              type="button"
              className="secondary-button modal-nav-button"
              onClick={onImportTemplates}
              disabled={isImporting}
            >
              {isImporting ? <Loader2 className="spin" size={15} /> : <Upload size={15} />}
              导入模板
            </button>
            <button className="icon-button" onClick={onClose} title="关闭" aria-label="关闭模板管理">
              <X size={18} />
            </button>
          </div>
        </div>

        <section className="template-directory-section" aria-label="当前模板目录">
          <div className="template-section-heading">
            <span>当前模板目录</span>
            <span>
              {filteredTemplateDirectoryEntries.length} / {templateDirectoryEntries.length} 个模板
            </span>
          </div>
          <label className="template-directory-search">
            <Search size={15} />
            <input
              value={templateDirectoryQuery}
              onChange={(event) => onTemplateDirectoryQueryChange(event.currentTarget.value)}
              placeholder="搜索模板名称、来源、类型或路径"
            />
          </label>
          {templateDirectoryError ? <div className="template-directory-error">{templateDirectoryError}</div> : null}
          <div className="template-directory-list">
            {templateDirectoryEntries.length === 0 ? (
              <div className="empty-state detail-empty">暂无模板</div>
            ) : filteredTemplateDirectoryEntries.length === 0 ? (
              <div className="empty-state detail-empty">没有匹配的模板</div>
            ) : (
              filteredTemplateDirectoryEntries.map((entry) => (
                <div key={entry.key} className="template-directory-row">
                  <div className="template-directory-name">
                    <strong>{entry.name}</strong>
                    <span>
                      {entry.source} · {entry.kind}
                    </span>
                  </div>
                  <code>{entry.path || "内置模板路径加载中"}</code>
                  <button
                    type="button"
                    className="template-dir-button"
                    onClick={() => onOpenTemplateDirectory(entry)}
                    disabled={!entry.path}
                  >
                    <FolderOpen size={14} />
                    打开目录
                  </button>
                </div>
              ))
            )}
          </div>
        </section>

        <section className="template-user-config-section" aria-label="导入模板配置">
          <div className="template-section-heading">
            <span>导入模板配置</span>
            <span>{userProcessTemplates.length} 个导入模板</span>
          </div>
          {userProcessTemplates.length === 0 ? (
            <div className="empty-state detail-empty">导入模板后可在这里填写匹配配置</div>
          ) : (
            <div className="template-manager-list">
              {userProcessTemplates.map((template) => {
                const key = userTemplateKey(template);
                const draft = templateDrafts[key] ?? draftFromUserTemplate(template);
                const isSaving = Boolean(savingTemplateKeys[key]);
                const isDeleting = Boolean(deletingTemplateKeys[key]);
                return (
                  <article key={key} className={`template-manager-item ${draft.enabled ? "" : "disabled"}`}>
                    <div className="template-manager-item-heading">
                      <div>
                        <strong>{template.templateFile}</strong>
                        <span>{template.userTemplatePath}</span>
                      </div>
                      <label className="template-enabled-toggle">
                        <input
                          type="checkbox"
                          checked={draft.enabled}
                          onChange={(event) => onUpdateTemplateDraft(template, { enabled: event.currentTarget.checked })}
                        />
                        启用
                      </label>
                    </div>
                    <div className="template-manager-grid">
                      <label className="text-field">
                        <span>显示名称</span>
                        <input
                          value={draft.displayName}
                          onChange={(event) => onUpdateTemplateDraft(template, { displayName: event.currentTarget.value })}
                        />
                      </label>
                      <label className="text-field">
                        <span>所属模块</span>
                        <select
                          value={draft.templateModule}
                          onChange={(event) =>
                            onUpdateTemplateDraft(template, { templateModule: event.currentTarget.value as ProcessTemplateModule })
                          }
                        >
                          <option value="process">过程资料</option>
                          <option value="switch-station">开关站电气设备安装</option>
                          <option value="collector-line">集电线路安装工程</option>
                        </select>
                      </label>
                      <label className="text-field">
                        <span>模板分类</span>
                        <select
                          value={draft.category}
                          onChange={(event) =>
                            onUpdateTemplateDraft(template, { category: event.currentTarget.value as ProcessTemplateCategoryId | "" })
                          }
                        >
                          <option value="">自动识别</option>
                          {PROCESS_TEMPLATE_CATEGORIES.map((category) => (
                            <option key={category.id} value={category.id}>
                              {category.label}
                            </option>
                          ))}
                        </select>
                      </label>
                      <label className="text-field">
                        <span>关键词模式</span>
                        <select
                          value={draft.matchMode}
                          onChange={(event) =>
                            onUpdateTemplateDraft(template, { matchMode: event.currentTarget.value as ProcessTemplateMatchMode })
                          }
                        >
                          <option value="any">任一命中</option>
                          <option value="all">全部命中</option>
                        </select>
                      </label>
                    </div>
                    <label className="text-field template-keyword-field">
                      <span>匹配关键词</span>
                      <textarea
                        value={draft.matchKeywordsText}
                        onChange={(event) => onUpdateTemplateDraft(template, { matchKeywordsText: event.currentTarget.value })}
                        placeholder="每行一个关键词，也可以用逗号分隔；留空时按模板文件名自动匹配"
                      />
                    </label>
                    <div className="template-manager-actions">
                      <button
                        type="button"
                        className="secondary-button modal-nav-button"
                        onClick={() => onDeleteTemplate(template)}
                        disabled={isDeleting || isSaving}
                      >
                        {isDeleting ? <Loader2 className="spin" size={15} /> : <Trash2 size={15} />}
                        删除
                      </button>
                      <button
                        type="button"
                        className="primary-button modal-nav-button"
                        onClick={() => onSaveTemplate(template)}
                        disabled={isSaving || isDeleting}
                      >
                        {isSaving ? <Loader2 className="spin" size={15} /> : <Save size={15} />}
                        保存
                      </button>
                    </div>
                  </article>
                );
              })}
            </div>
          )}
        </section>
      </section>
    </div>
  );
}
