import { describe, expect, it } from "vitest";
import {
  generateProcessDocs,
  generateSelectedProcessDocs,
  loadProcessManifest,
  matchingAllProcessTemplates,
  type ProcessTemplate,
} from "../lib/processDocs";
import { createSubunitSummaryFixtureItems, createSuixiGridConnectedElectricalRecord, fileNames, processRecord, stubProcessFetch } from "./processDocsTestUtils";

describe("template matching", () => {
  it("matches process templates only by explicit title rules", async () => {
    stubProcessFetch();
    const paths: string[] = [];

    const result = await generateProcessDocs(
      [processRecord],
      {
        selectedCodes: [processRecord.archiveCode],
        outputDir: "/tmp/archive-output",
        userFields: {
          generalContractorUnit: "测试总承包单位",
          generalContractorProjectManager: "总包负责人",
          generalContractorTechnicalLeader: "总包技术负责人",
          constructionUnit: "测试施工单位",
          constructionProjectManager: "施工负责人",
          constructionTechnicalLeader: "施工技术负责人",
          subcontractorUnit: "测试分包单位",
          subcontractorProjectManager: "分包负责人",
          subcontractorContent: "支架安装",
          supervisionDepartment: "测试监理项目部",
        },
      },
      async (path) => {
        paths.push(path);
      },
    );

    expect(result.files).toHaveLength(6);
    expect(paths).toHaveLength(6);
    expect(fileNames(paths).filter((name) => name.includes("开工报审"))).toHaveLength(4);
    expect(fileNames(paths).filter((name) => name.includes("子单位工程质量验收记录"))).toHaveLength(1);
    expect(fileNames(paths).filter((name) => name.includes("子单位工程报验申请单"))).toHaveLength(1);
    expect(result.skipped).toEqual([]);
    expect(result.errors).toEqual([]);
    expect(paths[0]).toContain("/过程资料/5028G01-0011-8312-001");
  });

  it("resolves all-module matches for preview and generates selected manual templates", async () => {
    stubProcessFetch();
    const manifest = await loadProcessManifest();
    const startReportMatches = matchingAllProcessTemplates(processRecord, processRecord.items[0], manifest.templates);
    const unmatchedItem = {
      ...processRecord.items[4],
      sequence: "99",
      fileCode: "5028G01-SG-ZHHC-TEST-099",
      title: "高明分布式项目 需要人工选择模板的文件",
    };
    const record = {
      ...processRecord,
      items: [unmatchedItem],
    };
    const manualTemplate = manifest.templates.find((template) => template.templateFile === "开工报审.docx")!;
    const paths: string[] = [];

    expect(startReportMatches.map((match) => match.template.templateFile)).toEqual(["开工报审.docx"]);
    expect(matchingAllProcessTemplates(record, unmatchedItem, manifest.templates)).toEqual([]);

    const result = await generateSelectedProcessDocs(
      [record],
      {
        outputDir: "/tmp/archive-output",
        selections: [
          {
            archiveCode: record.archiveCode,
            fileCode: unmatchedItem.fileCode,
            sequence: unmatchedItem.sequence,
            templateModule: "process",
            template: manualTemplate,
          },
        ],
      },
      async (path) => {
        paths.push(path);
      },
    );

    expect(result.files).toHaveLength(1);
    expect(result.skipped).toEqual([]);
    expect(result.errors).toEqual([]);
    expect(fileNames(paths)).toEqual(["99、5028G01-SG-ZHHC-TEST-099高明分布式项目 需要人工选择模板的文件.docx"]);
  });

  it("keeps all default matches for quality application and acceptance record titles", async () => {
    stubProcessFetch();
    const manifest = await loadProcessManifest();
    const record = {
      ...processRecord,
      items: createSubunitSummaryFixtureItems(),
    };
    const item = record.items.find((entry) => entry.title.includes("1#厂房支架安装"))!;

    const matches = matchingAllProcessTemplates(record, item, manifest.templates).map((match) => match.template.templateFile);

    expect(matches).toEqual([
      "厂房支架安装分项工程质量验收表.xlsx",
      "分项工程报验申请单.docx",
    ]);
  });

  it("prioritizes enabled user templates by configured keywords", () => {
    const record = createSuixiGridConnectedElectricalRecord();
    const item = record.items[0];
    const userTemplate: ProcessTemplate = {
      sequence: 0,
      kind: "docx",
      originalName: "自定义电缆终端制作模板.docx",
      templateFile: "自定义电缆终端制作模板.docx",
      outputExtension: ".docx",
      userTemplatePath: "/tmp/process-templates/自定义电缆终端制作模板.docx",
      displayName: "自定义电缆终端制作模板",
      matchKeywords: ["电缆终端制作"],
      matchMode: "any",
      templateModule: "process",
      enabled: true,
    };
    const builtInTemplate: ProcessTemplate = {
      sequence: 0,
      kind: "xlsx",
      originalName: "电力电缆终端制作安装分项工程质量验收表.xlsx",
      templateFile: "电力电缆终端制作安装分项工程质量验收表.xlsx",
      outputExtension: ".xlsx",
    };

    const matches = matchingAllProcessTemplates(record, item, [builtInTemplate, userTemplate]);

    expect(matches.map((match) => match.template.templateFile)).toEqual([
      "自定义电缆终端制作模板.docx",
      "电力电缆终端制作安装分项工程质量验收表.xlsx",
    ]);
  });

  it("ignores disabled user templates and supports legacy filename matching", () => {
    const item = {
      ...processRecord.items[4],
      title: "高明分布式项目 需要人工选择模板的文件",
    };
    const record = { ...processRecord, items: [item] };
    const disabledTemplate: ProcessTemplate = {
      sequence: 0,
      kind: "docx",
      originalName: "需要人工选择模板.docx",
      templateFile: "需要人工选择模板.docx",
      outputExtension: ".docx",
      userTemplatePath: "/tmp/process-templates/需要人工选择模板.docx",
      matchKeywords: ["需要人工选择模板"],
      templateModule: "process",
      enabled: false,
    };
    const legacyTemplate: ProcessTemplate = {
      sequence: 0,
      kind: "docx",
      originalName: "需要人工选择模板.docx",
      templateFile: "需要人工选择模板.docx",
      outputExtension: ".docx",
      userTemplatePath: "/tmp/process-templates/legacy/需要人工选择模板.docx",
    };

    const matches = matchingAllProcessTemplates(record, item, [disabledTemplate, legacyTemplate]);

    expect(matches.map((match) => match.template.userTemplatePath)).toEqual([
      "/tmp/process-templates/legacy/需要人工选择模板.docx",
    ]);
  });

  it("requires every keyword when user templates use all-keyword matching", () => {
    const record = createSuixiGridConnectedElectricalRecord();
    const item = record.items[0];
    const template: ProcessTemplate = {
      sequence: 0,
      kind: "docx",
      originalName: "严格匹配模板.docx",
      templateFile: "严格匹配模板.docx",
      outputExtension: ".docx",
      userTemplatePath: "/tmp/process-templates/严格匹配模板.docx",
      matchKeywords: ["电缆终端制作", "不存在关键词"],
      matchMode: "all",
      templateModule: "process",
    };

    expect(matchingAllProcessTemplates(record, item, [template])).toEqual([]);
  });
});
