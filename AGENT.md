# AGENT.md

This file is for future maintainers and coding agents working on this project.

## Project Overview

This is a Tauri v2 desktop app for generating archive-related DOCX/XLSX files from a source Excel workbook such as `预立卷档案总目录（高明）(5).xlsx`.

Frontend stack:

- React
- TypeScript
- Vite
- Tauri v2 plugins: dialog, fs, opener, process, updater

Document stack:

- `xlsx` for source Excel parsing
- `exceljs` for normal XLSX generation
- `pizzip` and direct Open XML edits for DOCX/XLSX templates where style preservation matters

## Important Commands

Run these before handing off code changes when feasible:

```bash
npm test
npm run build
```

For Tauri/Rust changes:

```bash
cd src-tauri
cargo check
```

For desktop dev:

```bash
npm run tauri:dev
```

## Core Modules

- `src/lib/excel.ts`
  - Parses the archive source workbook.
  - Header row is row 2; data starts at row 4.
  - Non-empty archive code starts a new archive record; blank archive code rows belong to the previous record.

- `src/lib/docx.ts`
  - Generates archive cover, backup note, and spine DOCX files.

- `src/lib/catalog.ts`
  - Generates `1.2、案卷目录、卷内目录-著录台账（及打印模板）.xlsx`.

- `src/lib/process-docs/generator.ts`
  - Main process-doc generation entry.
  - Do not reintroduce sequence fallback matching.

- `src/lib/process-docs/templates.ts`
  - Template discovery and title matching rules.

- `src/lib/process-docs/docxRenderer.ts`
  - DOCX XML text replacement.
  - Be careful not to rewrite whole paragraphs when the template uses underlined runs for handwritten fields.

- `src/lib/process-docs/summaryWorkbookRenderer.ts`
  - Special renderer for summary quality acceptance workbooks.
  - Uses direct XML edits to preserve the original workbook style.

- `src/lib/process-docs/workbookRenderer.ts`
  - Normal XLSX template rendering with ExcelJS.

## Process Docs Rules

Process documents are generated only by explicit title/template rules. They must not fall back to `ArchiveItem.sequence`.

Current explicit rules:

- `开工报审.docx`
  - Match item title containing `开工报审` or `开工报审表`.

- `子单位工程报验申请单.docx`
  - Match item title containing `子单位` or `子单位工程`, and also `质量报验申请及验收记录` or `质量报审表及验收记录`.
  - In the body sentence after `根据施工承包合同的规定，`, write only the subject without the trailing `子单位工程` because the template text already contains `子单位工程现已施工完毕`.
  - In the attachment line, keep the full name with `子单位工程`.
  - Do not auto-fill `项目经理` or `日期`; those are handwritten fields and their underlines must remain.

- `子单位工程质量验收记录.xlsx`
  - Uses the same item-title rule as the subunit inspection application.
  - Must preserve summary workbook style through XML edits, not ExcelJS rewrites.
  - Print setup is important: A4, landscape, scale 95, fit to 1 page wide and 1 page tall, print area `光伏方阵安装!$A$1:$AL$30`.

## Template Files

Template root:

```text
public/templates/process-docs/
```

Manifest:

```text
public/templates/process-docs/manifest.json
```

The original reference files are under:

```text
template/
```

Do not edit reference templates unless explicitly requested. Built-in app templates are copied/converted under `public/templates`.

## Known Pitfalls

- Do not use sequence fallback for process docs.
  - The same卷内序号 can mean different document types in different archive records.

- Do not rewrite whole DOCX paragraphs when only one underlined field needs replacement.
  - Whole-paragraph replacement collapses Word runs and removes underlines.

- For `子单位工程质量验收记录.xlsx`, XML element order matters.
  - `pageMargins` must appear before `pageSetup`.
  - Missing or misplaced print setup can make Excel/WPS ignore page fitting and show a very wide print preview.

- For Tauri `openPath`, permissions are controlled in:

```text
src-tauri/capabilities/default.json
```

The app currently allows opening paths under `$HOME`.

## Release Notes

GitHub Actions workflow:

```text
.github/workflows/desktop-build.yml
```

Current behavior:

- Push branches: run tests only.
- Push `v*` tags: build Windows and macOS bundles and publish GitHub Release.

Use:

```bash
sh scripts/release-tag.sh v0.1.1
```

Updater notes are in `UPDATER.md`.
