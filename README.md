# 档案文档生成器

基于 Tauri v2、React 和 TypeScript 的桌面客户端，用于从档案总目录 Excel 中读取案卷和卷内目录数据，生成固定格式的 DOCX/XLSX 文件。目标平台为 Windows 和 macOS。

## 功能

- 档案文档生成
  - 案卷大封面 DOCX
  - 备考表 DOCX
  - 案卷脊背 DOCX
  - 案卷目录、卷内目录台账 XLSX
- 过程资料生成
  - 按用户勾选的模板类型生成
  - 当前已接入：开工报审、子单位工程报验申请单、子单位工程质量验收记录
  - 过程资料生成信息会保存在本机，下次打开自动带入
- 应用内更新
  - 发布 GitHub tag 后由 Actions 构建安装包和 `latest.json`
  - 已安装用户可在应用内点击更新

## 开发环境

需要安装：

- Node.js
- Rust
- Tauri 所需系统依赖

安装依赖：

```bash
npm ci
```

本地 Web 开发：

```bash
npm run dev
```

Tauri 桌面调试：

```bash
npm run tauri:dev
```

## 常用命令

```bash
npm test
npm run build
npm run tauri:build
```

只构建 Windows NSIS 包可在 Windows 环境运行：

```bash
npm run tauri -- build --bundles nsis
```

## 项目结构

```text
src/
  App.tsx                    主界面和 Tauri 交互
  lib/
    excel.ts                 解析预立卷档案总目录
    docx.ts                  案卷大封面、备考表、脊背生成
    catalog.ts               案卷目录台账 XLSX 生成
    process-docs/            过程资料生成
public/templates/
  process-docs/              过程资料内置模板
template/                    示例源文件和模板来源
src-tauri/                   Tauri/Rust 配置
scripts/release-tag.sh       创建并推送发布 tag
```

## 模板说明

过程资料模板清单在：

```text
public/templates/process-docs/manifest.json
```

第一版过程资料已经移除“按卷内序号兜底匹配”。模板必须通过明确题名规则匹配，避免不同案卷相同序号误生成。

## 发布

推送普通代码只跑测试。推送 `v*` tag 后 GitHub Actions 会构建 Windows 和 macOS 包并发布 Release。

推荐使用脚本：

```bash
sh scripts/release-tag.sh v0.1.1
```

更新配置和 GitHub Secrets 见 [UPDATER.md](./UPDATER.md)。
