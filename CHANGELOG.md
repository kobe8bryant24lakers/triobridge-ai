# Changelog

## 0.3.0 — 2026-11-23

Word can now insert structured reports with sections and native tables, Excel can create formatted data tables with explicit formulas, and PowerPoint outlines can be edited and reordered before creating slides. Previews retain drafts through settings changes, identify outdated results, and protect write targets from intervening changes.

The local shared edition stores multiple model profiles in one catalog for Word, Excel, and PowerPoint on the same computer and OS account. It uses macOS Keychain or Windows Credential Manager, synchronizes model switching, detects conflicting edits, and preserves shared state while the local service is unavailable. Migration is explicit and transactional; profile exports exclude keys. Setup requires a running local HTTPS service and a private installation manifest.

The hosted ZIP installers continue to install the online edition with separate per-application model storage. The website and README explain both editions in English and Chinese. The Office manifest remains at `1.1.0.0`; existing hosted identity, URL, and permissions are unchanged.

Validation includes native Mac Word, Excel, and PowerPoint with real configured models; structured document generation, editing, writing, saving, reopening, and Excel recalculation; shared-model switching, conflict handling, and service recovery. The pre-release regression suite passed 284 checks, with 810 additional saved-file assertions. The complex Word and PowerPoint PDFs were reviewed on all 16 pages. See [test scope](docs/TESTING.md). Windows native acceptance, other Office versions, and protected or custom-template documents remain separate validation work.

### 简体中文

- Word 新增章节与原生表格报告；Excel 新增带类型、公式及列格式的数据表；PowerPoint 支持逐页编辑与调整顺序后建页。
- 本机共享版支持同一电脑、同一系统账号配置一次、多模型切换、三应用同步，密钥使用系统安全存储；迁移、并发冲突及服务离线均有明确处理。
- 中英文 README 与网站说明在线安装包和本机共享版的区别。在线 ZIP 仍使用各应用独立配置；共享版需运行本机服务并安装私有清单。
- 已完成真实模型的 Mac 原生复杂文档验收；284 项回归及 810 项保存文件检查通过，16 页复杂文档 PDF 全部逐页复核。Windows 原生验收仍待完成。

## 0.2.0

Provider presets, simplified connection setup, and hosted Mac/Windows sideloading packages.

## 0.1.0

Initial Word, Excel, and PowerPoint task-pane workflows and hosted installation guide.
