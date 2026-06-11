import PizZip from "pizzip";
import type { ProcessTemplateModule } from "./types";

export function preserveProcessWorkbookPrintLayout(
  bytes: Uint8Array,
  template: ArrayBuffer | Uint8Array,
  templateModule: ProcessTemplateModule,
): Uint8Array {
  const zip = new PizZip(bytes);
  const templateZip = new PizZip(template);
  const worksheetPaths = Object.keys(zip.files).filter((name) => /^xl\/worksheets\/sheet\d+\.xml$/.test(name));

  for (const path of worksheetPaths) {
    const sheet = zip.file(path);
    if (!sheet) {
      continue;
    }
    const templateSheet = templateZip.file(path);
    let sheetXml = templateModule === "collector-line" ? sheet.asText() : forceOnePageWorksheetLayout(sheet.asText());
    if (templateSheet) {
      sheetXml = restoreTemplateSheetFormatting(sheetXml, templateSheet.asText());
    }
    zip.file(path, sheetXml);
  }

  const templateStylesFile = templateZip.file("xl/styles.xml");
  const generatedStylesFile = zip.file("xl/styles.xml");
  if (templateStylesFile && generatedStylesFile) {
    zip.file("xl/styles.xml", restoreTemplateDefaultFont(generatedStylesFile.asText(), templateStylesFile.asText()));
  }

  preserveProcessWorkbookPrintArea(zip, templateZip);

  return zip.generate({ type: "uint8array", compression: "DEFLATE" });
}

function forceOnePageWorksheetLayout(xml: string): string {
  let nextXml = upsertFitToPage(xml);
  nextXml = upsertPageSetup(nextXml);
  nextXml = normalizeHorizontalPageMargins(nextXml);
  return nextXml;
}

/**
 * Copy <sheetFormatPr> and <cols> from template to generated sheet XML
 * so ExcelJS-injected defaultColWidth and altered column widths are reverted.
 */
function restoreTemplateSheetFormatting(generatedXml: string, templateXml: string): string {
  let result = generatedXml;

  result = restoreTemplateBlockElement(result, templateXml, "sheetViews", ["<sheetFormatPr", "<cols", "<sheetData"]);

  // Restore <sheetFormatPr .../> — remove ExcelJS's version, insert template's (or nothing)
  const templateSheetFormatPr = templateXml.match(/<sheetFormatPr[^>]*\/>/) ?.[0];
  if (templateSheetFormatPr) {
    // Replace generated sheetFormatPr with template's
    if (/<sheetFormatPr[^>]*\/>/.test(result)) {
      result = result.replace(/<sheetFormatPr[^>]*\/>/, templateSheetFormatPr);
    } else {
      result = result.replace(/(<sheetData\b)/, `${templateSheetFormatPr}$1`);
    }
  } else {
    // Template has no sheetFormatPr — strip ExcelJS's injected one
    result = result.replace(/<sheetFormatPr[^>]*\/>/, '');
  }

  // Restore <cols>...</cols> — replace generated with template's
  const templateCols = templateXml.match(/<cols>[\s\S]*?<\/cols>/) ?.[0];
  const generatedHasCols = /<cols>[\s\S]*?<\/cols>/.test(result);
  if (templateCols) {
    if (generatedHasCols) {
      result = result.replace(/<cols>[\s\S]*?<\/cols>/, templateCols);
    } else {
      result = result.replace(/(<sheetData\b)/, `${templateCols}$1`);
    }
  } else if (generatedHasCols) {
    result = result.replace(/<cols>[\s\S]*?<\/cols>/, '');
  }

  result = restoreTemplateBlockElement(result, templateXml, "mergeCells", [
    "<phoneticPr",
    "<conditionalFormatting",
    "<dataValidations",
    "<printOptions",
    "<pageMargins",
    "<pageSetup",
    "<headerFooter",
    "</worksheet>",
  ]);
  result = restoreTemplateSelfClosingElement(result, templateXml, "printOptions", ["<pageMargins", "<pageSetup", "</worksheet>"]);
  result = restoreTemplateSelfClosingElement(result, templateXml, "pageMargins", ["<pageSetup", "</worksheet>"]);
  result = restoreTemplateSelfClosingElement(result, templateXml, "pageSetup", ["<headerFooter", "</worksheet>"]);

  return result;
}

function restoreTemplateDefaultFont(generatedStylesXml: string, templateStylesXml: string): string {
  const templateDefaultFont = firstFontXml(templateStylesXml);
  if (!templateDefaultFont) {
    return generatedStylesXml;
  }

  return generatedStylesXml.replace(/(<fonts\b[^>]*>)([\s\S]*?<\/font>)/, `$1${templateDefaultFont}`);
}

function firstFontXml(stylesXml: string): string | undefined {
  return stylesXml.match(/<fonts\b[^>]*>\s*(<font>[\s\S]*?<\/font>)/)?.[1];
}

function upsertFitToPage(xml: string): string {
  if (/<sheetPr\b[^>]*>[\s\S]*?<\/sheetPr>/.test(xml)) {
    return xml.replace(/<sheetPr\b([^>]*)>([\s\S]*?)<\/sheetPr>/, (_match, attrs: string, body: string) => {
      const nextBody = /<pageSetUpPr\b[^>]*\/>/.test(body)
        ? body.replace(/<pageSetUpPr\b([^>]*)\/>/, (_pageSetupMatch: string, pageSetupAttrs: string) => {
            return `<pageSetUpPr${upsertXmlAttribute(pageSetupAttrs, "fitToPage", "1")}/>`;
          })
        : `<pageSetUpPr fitToPage="1"/>${body}`;
      return `<sheetPr${attrs}>${nextBody}</sheetPr>`;
    });
  }

  if (/<sheetPr\b[^>]*\/>/.test(xml)) {
    return xml.replace(/<sheetPr\b([^>]*)\/>/, (_match, attrs: string) => `<sheetPr${attrs}><pageSetUpPr fitToPage="1"/></sheetPr>`);
  }

  return xml.replace(/(<worksheet\b[^>]*>)/, '$1<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>');
}

function upsertPageSetup(xml: string): string {
  const existing = /<pageSetup\b([^>]*)\/>/;
  if (existing.test(xml)) {
    return xml.replace(existing, (_match, attrs: string) => {
      const cleanAttrs = attrs.replace(/\s+scale="[^"]*"/g, "");
      const nextAttrs = upsertXmlAttribute(
        upsertXmlAttribute(
          upsertXmlAttribute(
            upsertXmlAttribute(cleanAttrs, "paperSize", "9"),
            "fitToWidth",
            "1",
          ),
          "fitToHeight",
          "1",
        ),
        "usePrinterDefaults",
        "0",
      );
      return `<pageSetup${nextAttrs}/>`;
    });
  }

  return upsertSelfClosingElement(
    xml,
    "pageSetup",
    '<pageSetup paperSize="9" fitToWidth="1" fitToHeight="1" usePrinterDefaults="0"/>',
    "</worksheet>",
  );
}

function restoreTemplateBlockElement(
  generatedXml: string,
  templateXml: string,
  tagName: string,
  insertBeforeCandidates: string[],
): string {
  const elementPattern = new RegExp(`<${tagName}\\b[\\s\\S]*?</${tagName}>|<${tagName}\\b[^>]*/>`);
  const templateElement = templateXml.match(elementPattern)?.[0];

  if (!templateElement) {
    return generatedXml.replace(elementPattern, "");
  }

  if (elementPattern.test(generatedXml)) {
    return generatedXml.replace(elementPattern, templateElement);
  }

  const insertIndex = firstExistingIndex(generatedXml, insertBeforeCandidates);
  if (insertIndex === -1) {
    return generatedXml;
  }
  return `${generatedXml.slice(0, insertIndex)}${templateElement}${generatedXml.slice(insertIndex)}`;
}

function restoreTemplateSelfClosingElement(
  generatedXml: string,
  templateXml: string,
  tagName: string,
  insertBeforeCandidates: string[],
): string {
  const templateElement = templateXml.match(new RegExp(`<${tagName}\\b[^>]*/>`))?.[0];
  const generatedElement = new RegExp(`<${tagName}\\b[^>]*/>`);

  if (!templateElement) {
    return generatedXml.replace(generatedElement, "");
  }

  if (generatedElement.test(generatedXml)) {
    return generatedXml.replace(generatedElement, templateElement);
  }

  const insertIndex = firstExistingIndex(generatedXml, insertBeforeCandidates);
  if (insertIndex === -1) {
    return generatedXml;
  }
  return `${generatedXml.slice(0, insertIndex)}${templateElement}${generatedXml.slice(insertIndex)}`;
}

function firstExistingIndex(xml: string, needles: string[]): number {
  const indexes = needles.map((needle) => xml.indexOf(needle)).filter((index) => index >= 0);
  return indexes.length > 0 ? Math.min(...indexes) : -1;
}

function normalizeHorizontalPageMargins(xml: string): string {
  return xml.replace(/<pageMargins\b([^>]*)\/>/, (match: string, attrs: string) => {
    const left = xmlAttribute(attrs, "left");
    const right = xmlAttribute(attrs, "right");
    const leftNumber = left === undefined ? undefined : Number(left);
    const rightNumber = right === undefined ? undefined : Number(right);

    if (left && right !== undefined && rightNumber === 0 && leftNumber !== undefined && leftNumber > 0) {
      return `<pageMargins${upsertXmlAttribute(attrs, "right", left)}/>`;
    }

    if (right && left !== undefined && leftNumber === 0 && rightNumber !== undefined && rightNumber > 0) {
      return `<pageMargins${upsertXmlAttribute(attrs, "left", right)}/>`;
    }

    return match;
  });
}

function upsertSelfClosingElement(xml: string, tagName: string, elementXml: string, insertBefore: string): string {
  const existing = new RegExp(`<${tagName}\\b[^>]*/>`);
  if (existing.test(xml)) {
    return xml.replace(existing, elementXml);
  }

  const insertIndex = xml.indexOf(insertBefore);
  if (insertIndex === -1) {
    return xml;
  }
  return `${xml.slice(0, insertIndex)}${elementXml}${xml.slice(insertIndex)}`;
}

function upsertXmlAttribute(attrs: string, name: string, value: string): string {
  const pattern = new RegExp(`\\s${name}="[^"]*"`);
  if (pattern.test(attrs)) {
    return attrs.replace(pattern, ` ${name}="${value}"`);
  }
  return `${attrs} ${name}="${value}"`;
}

function xmlAttribute(attrs: string, name: string): string | undefined {
  return attrs.match(new RegExp(`\\s${name}="([^"]*)"`))?.[1];
}

function preserveProcessWorkbookPrintArea(zip: PizZip, templateZip: PizZip) {
  const workbook = zip.file("xl/workbook.xml");
  if (!workbook) {
    return;
  }

  const workbookXml = workbook.asText();
  const templateWorkbookXml = templateZip.file("xl/workbook.xml")?.asText() ?? "";
  const printAreas = printAreaDefinedNames(templateWorkbookXml);
  zip.file("xl/workbook.xml", insertPrintAreas(removePrintAreas(workbookXml), printAreas));
}

function removePrintAreas(workbookXml: string): string {
  return workbookXml
    .replace(/<definedName\b(?=[^>]*\bname="_xlnm\.Print_Area")[^>]*>[\s\S]*?<\/definedName>/g, "")
    .replace(/<definedNames>\s*<\/definedNames>/g, "");
}

function printAreaDefinedNames(workbookXml: string): string[] {
  return Array.from(
    workbookXml.matchAll(/<definedName\b(?=[^>]*\bname="_xlnm\.Print_Area")[^>]*>[\s\S]*?<\/definedName>/g),
    (match) => match[0],
  );
}

function insertPrintAreas(workbookXml: string, printAreas: string[]): string {
  if (printAreas.length === 0) {
    return workbookXml;
  }

  const printAreasXml = printAreas.join("");
  if (/<definedNames\b[^>]*>[\s\S]*?<\/definedNames>/.test(workbookXml)) {
    return workbookXml.replace(/<\/definedNames>/, `${printAreasXml}</definedNames>`);
  }
  if (/<definedNames\/>/.test(workbookXml)) {
    return workbookXml.replace(/<definedNames\/>/, `<definedNames>${printAreasXml}</definedNames>`);
  }
  return workbookXml.replace("</workbook>", `<definedNames>${printAreasXml}</definedNames></workbook>`);
}
