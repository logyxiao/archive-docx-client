export const DEFAULT_BACKUP_NOTE = "";
export const LAST_OUTPUT_DIR_KEY = "archive-docx-client:last-output-dir";
export const PROCESS_FIELDS_KEY = "archive-docx-client:process-fields";
export const PROCESS_TEMPLATE_CATEGORIES_KEY = "archive-docx-client:process-template-categories";
export const SWITCH_STATION_TEMPLATE_CATEGORIES_KEY = "archive-docx-client:switch-station-template-categories";
export const COLLECTOR_LINE_TEMPLATE_CATEGORIES_KEY = "archive-docx-client:collector-line-template-categories";
export const ARCHIVE_DOCX_TAB = "archive-docx";
export const ALL_PROCESS_DOCS_TAB = "all-process-docs";
export const PROCESS_DOCS_TAB = "process-docs";
export const SWITCH_STATION_TAB = "switch-station-process-docs";
export const COLLECTOR_LINE_TAB = "collector-line-process-docs";
export const NO_PROCESS_TEMPLATE_KEY = "__none__";

export const TEMPLATE_IMPORT_STEPS = [
  "在 Word 或 Excel 模板中写入占位符，例如 {{项目名}}、{{文件题名}}、{{文件日期中文}}。",
  "Excel 质量结果可写 {{质量验收结果:<2}} 或 {{质量验收结果:≥6}}；下限型标准默认填 6~9，明确范围可写 {{质量验收结果:6~10}}。",
  "点击“导入模板”，选择 .docx 或 .xlsx 文件，软件会复制到导入模板目录。",
  "在导入模板配置里填写匹配关键词；生成时文件题名命中关键词就会默认选中该模板。",
];

export const TEMPLATE_CONFIG_FIELDS = [
  "显示名称：模板管理页和下拉框里看到的名称。",
  "所属模块：过程资料、开关站电气设备安装、集电线路安装工程。",
  "模板分类：用于普通过程资料页的分类筛选；不填则自动识别。",
  "匹配关键词：每行一个，也可用逗号分隔。",
  "关键词模式：任一命中或全部命中。",
  "启用状态：关闭后该模板不参与默认匹配和下拉选择。",
  "质量验收结果：支持 ±2、<2、≤2、0~2、≥6 这类标准，生成时自动随机填充。",
];

export const PROCESS_TEMPLATE_PLACEHOLDERS = [
  "项目名",
  "工程名称",
  "工程编号",
  "档号",
  "案卷题名",
  "文件编号",
  "文件题名",
  "文件题名去项目",
  "文件题名主题",
  "验收项目名称",
  "文件日期",
  "文件日期中文",
  "责任者",
  "编制单位",
  "立卷单位",
  "总承包单位",
  "施工单位",
  "监理单位",
  "监理项目部",
  "分包单位",
  "分包内容",
  "质量验收结果:<2",
  "质量验收结果:≥6",
  "质量验收结果:6~10",
];
