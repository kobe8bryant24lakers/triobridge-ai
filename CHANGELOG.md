# Changelog

## 0.12.0 — 2026-11-26

- Compare three themes on a slide from your own deck. Choose a built-in design, ask the current model for three custom options, or adjust fonts and colors. Apply a design to the full deck, save a favorite, or undo the change.
- Review a generated deck for missing sources and notes, long or repeated titles, and chart units. Refine one slide with a before-and-after text review and accept individual edits.
- Preview alternative layouts on generated slides. On existing PPTX working copies, compare layouts side by side, move a picture and its text column, and review alignment, overflow and minor overlaps before exporting.
- The illustrated Chinese guide adds four pages with fictional examples. This release was tested with a 12-slide editable presentation in PowerPoint on Apple silicon.

### 简体中文

- 在真实页面上比较三套主题，可按描述设计，或手动微调标题、正文和配色。一键应用整稿，也可收藏主题和撤销换色。
- 在初稿中检查标题、来源、图表单位与演讲备注。逐条比较并采用单页修改建议。
- 为图文、要点和步骤页预览新布局。对已有文件，可指定页码并排查看改前、改后布局，勾选图文换位、对齐、越界与轻微重叠修复，再导出工作副本。
- 中文图解指南增加四页虚构教学示例。

## 0.11.0 — 2026-11-26

- Set native slide backgrounds across a deck, on specific pages or on a range with exceptions. Choose a solid color, a two-color gradient, a picture with transparency, or the slide master. New decks support a default background and per-slide overrides.
- Preview batch changes individually, save editing recipes, match a reference slide, insert a slide after an anchor, or split a text box across two slides.
- Update a supported native chart and its embedded Excel workbook from a table. Review comments, text density, overflow, small type and contrast in a PPTX working copy.
- Recognized background instructions run locally. Model-assisted edits send the explicitly read slides and selected materials. A cross-slide background preview can be refined with the model after explicitly reading the full deck.

### 简体中文

- 支持整稿、单页、页码范围及例外页的原生背景。可选纯色、双色渐变、带透明度的图片或恢复母版。新稿支持默认背景与单页覆盖。
- 批量改动可逐项勾选，常用操作可保存为配方。可套用参考页样式、在指定位置插页，或把一个长文本框拆成两页。
- 可用表格刷新支持的原生图表及其内嵌 Excel 工作簿。PPTX 工作副本支持按批注修改，并检查文字过密、越界、字号和对比度。
- 可识别的背景指令在本机处理。模型辅助编辑使用明确读取的页面和选中素材。跨页背景预览后，明确读取整稿即可让模型继续调整。

## 0.10.0 — 2026-11-26

- Add **Conversation** for multi-turn questions about the add-in, presentation planning, and general topics. This mode does not read or change the current document.
- Refine Word and Excel results, PowerPoint decks, existing-slide proposals, and chart recommendations across multiple turns. Editing uses the current preview, including manual changes, and recent history. Conversation and editing keep separate histories; document writes still require confirmation.
- Keep the previous preview and pending request when a revision fails or is cancelled. Clear editing history without discarding the current preview.

### 简体中文

- 新增「对话模式」，支持连续询问插件用法、整理 PPT 思路及一般问题，不读取或修改当前文档。
- Word、Excel 结果及 PowerPoint 整稿、现有页面修改方案、图表建议均支持多轮调整，结合当前预览、手动修改和近期历史继续编辑。两种模式分别保留历史，写入文档仍需确认。
- 修改失败或取消时保留上一版预览与待发送要求；清空编辑记录会保留当前预览。

## 0.9.1 — 2026-11-26

- PowerPoint’s **Create from Files → Generate deck** can now wait beyond the model profile’s request timeout, which defaults to 60 seconds. Other tasks and connection tests continue to use the configured timeout. Choose **Stop** to cancel a generation and keep the previous preview.
- A regression test used the production workbench, a simulated Windows task-pane bridge in Chromium on macOS, and a local Claude-compatible service that waited 65 seconds to respond. The two-slide preview appeared after 65.314 seconds. Stopping a second generation closed the request and preserved the first preview. Testing in desktop PowerPoint on physical Windows x64 and ARM64 machines, in both the task pane and standalone window, is still pending.

## 0.9.0 — 2026-11-25

- Switch between English and Simplified Chinese below the TrioBridge AI logo in the standalone workbench and Office task pane. Simplified Chinese is the default. The installed workbench shares the interface preference across Word, Excel and PowerPoint.
- Read task instructions, model settings, previews and write confirmations in the selected language. PowerPoint also localizes the source picker, chart controls, object edits and export prompts.

## 0.8.0 — 2026-11-25

- Publish a 40-slide illustrated Chinese user guide with step-by-step instructions for Word, Excel and PowerPoint. [Download the PowerPoint guide](https://kobe8bryant24lakers.github.io/triobridge-ai/assets/TrioBridge-AI-User-Guide-zh-CN.pptx)

- Add a dockable Windows workbench to the Word, Excel, and PowerPoint ribbon alongside the standalone window. Each pane is bound to its Office window and uses the existing local model profiles. Windows packages include the offline WebView2 runtime installer.
- Include an optional Mac Office task pane and its enable, restart, and removal commands in the desktop package. The pane uses local HTTPS and loads Office.js from Microsoft.
- Fix formula writes to a single cell in Excel for Mac and preserve adjacent paragraph styles when writing Markdown from the Word task pane.
- Convert Markdown in Word results to editable headings, lists, tables, and emphasis when inserting or replacing text. The workbench previews the same formatting before confirmation.
- Remove the fixed 4,096-token budget from Claude-compatible requests. Services that require `max_tokens` use the maximum reported by the selected model. Word can preview and write complete long text replies, and reports when the model runs out of context.
- Save model connections when a valid test response reaches the server's output limit, with a warning to raise that limit before generating content. Incomplete document generation is still rejected.
- Fix Windows PowerPoint document identification using its supported window caption instead of the Word/Excel window handle property, which could incorrectly report a selection error.
- Append generated PowerPoint decks and chart pages without reading the current selection or existing shape contents. Insertion still verifies the target document, page order and dimensions.

## 0.7.1 — 2026-11-25

- Fix Windows ZIP extraction in WinRAR. Both Windows packages now store the installer files and the Workbench directory directly at the archive root.
- Check every archived path against the staged installer before running installation tests. Manually dispatched Windows acceptance runs use WinRAR to extract the package.

## 0.7.0 — 2026-11-25

- Custom and local model profiles can use OpenAI Chat Completions or Claude (Anthropic Messages). Word, Excel and PowerPoint share the selected format. Existing profiles default to OpenAI.
- Model settings show the request URL before connecting. Word rewrites now ask the model to preserve the source date precision.

## 0.6.0

- Ship a single desktop workbench for Word, Excel and PowerPoint. Windows uses an Office ribbon button and Mac uses an application Services menu.
- Open the packaged page directly from disk. Model calls use the selected endpoint and credentials shared across the current user's three applications.
- Distribute installers by OS and processor. The download site hosts the guide and links to release assets.


## 0.5.2 — 2026-11-24

Windows shared setup verifies existing private permissions before changing them. A temporary file lock can clear during startup. After an unreadable first catalog is corrected, the same service can read it again. UTF-8 BOM catalogs retain their original bytes on read. The installer records a specific storage reason and the standalone diagnostic checks file metadata.

Windows 本机共享版改进了第 3 步的模型配置读取。安装窗口和只读诊断能区分文件占用、访问权限、JSON 解析和配置结构错误。安装器保留原文件和系统凭据。

## 0.5.1 — 2026-11-23

Windows local setup reports the active startup check and elapsed time. The installer connects directly to loopback with the installation certificate, waits for a cold shared configuration read, and checks both endpoints on restart. The read-only diagnostic includes the last startup phase and scheduler status even after a failed upgrade rolls back.

Windows 本机共享版修复配置读取超过两秒就提前回退的问题。安装窗口显示当前等待项。若失败，可用包内或独立的 `Diagnose.cmd` 查看端口、登录任务及最近一次启动检查记录。

## 0.5.0 — 2026-11-23

PowerPoint now accepts a local material folder or selected files as context for a complete editable deck or targeted changes to an existing presentation. DOCX, XLSX, PPTX, selectable-text PDF, text and images are parsed locally. Users review the included sources before submitting a model request; optional image understanding is explicit.

Full decks support nine layouts, 2–20 pages, source references, speaker notes, editable tables, images and native data charts. Existing slides support object-level before/after review, selective application, backup, stale-target rejection and post-write verification. Deterministic tools add cross-page text replacement, style unification, alignment, spacing and page ordering. Chart values come from source tables, with zero, negative and missing values preserved in both chart caches and embedded workbooks.

Mac PowerPoint acceptance used the saved real model for a ten-page deck, source-based local edits and a chart recommendation: three requests returned 6,719 tokens. Native files, both aspect ratios, embedded chart data and stale-edit rejection were checked separately from browser demos. Font substitution and unsupported grouped content remain explicit limits. See [PowerPoint guide](docs/POWERPOINT.md) and [validation scope](docs/TESTING.md). The manifest retains version `1.1.0.0` and its existing identity so saved Mac sideloaded documents keep resolving their add-in reference; application version is `0.5.0`.

### 简体中文

- 选择本地素材目录，生成可编辑的整套 PPT；支持文字、图片、表格和原生图表，并保留来源及备注。
- 参考素材局部修改现有页面，逐项预览与确认；支持原稿备份、过期方案拒绝和应用后核验。
- 新增跨页替换、统一样式、对象对齐 / 等距和页面重排；图表保留真实零值、负数和空值。
- 已用保存的真实模型完成 Mac PowerPoint 端到端验收；Windows 原生 Office 仍由用户发布后人工验收。详见测试记录。

## 0.4.0 — 2026-11-23

Self-contained local shared installers are provided for Windows and macOS, each in x64 and ARM64 builds. They bundle a pinned Node.js runtime and the native credential-store binding, create a private localhost certificate and Office manifest on the target computer, and start the service at sign-in. No separate Node.js or npm setup is required. Upgrade, repair, restart, and uninstall preserve existing model profiles and OS-stored keys.

Local HTTP APIs such as `http://127.0.0.1:8001/v1` work through the shared service without requiring model API CORS. Connection diagnostics distinguish a stopped local service, refused API connection, DNS, timeout, and certificate failures. A missing shared installation credential does not silently fall back to per-app keys.

The English and Chinese website and README now prioritize local shared downloads and identify the hosted edition separately. Native Windows Office acceptance remains pending; installer automation and synthetic transport checks are distinct from real-model Office acceptance.

Windows packages check for running Office apps before changing the catalog and include a read-only `Diagnose.cmd` for per-app discovery problems. Reports contain application versions and catalog/manifest status, without document content or model credentials.

PowerPoint generation instructions now describe append-only behavior, so preserving existing slides does not request placeholder copies in the generated outline. Generated Excel percentage columns display two decimal places. Mac setup and model settings also explain first-use Keychain prompts.

The installed Mac build passed real configured-model generation, writing, saving and reopening in all three native Office apps. Saved outputs passed 337 field assertions, including 96 native Excel formula results and unchanged original PowerPoint slides; Word pages and new slides were visually reviewed. See [validation scope and remaining Windows work](docs/TESTING.md).

### 简体中文

- 新增 Windows / Mac、x64 / ARM64 本机共享安装包，自带运行环境，双击安装并按系统提示完成配置。
- 自动配置当前账号的证书、后台服务和三应用入口；升级、修复及卸载保留模型配置和系统密钥。
- 本地 HTTP API 通过本机服务调用；连接报错明确区分插件到本机服务、服务到模型 API 的故障。
- README 与 GitHub Pages 优先提供本机共享下载，在线版独立说明。Windows 原生 Office 验收仍由使用者后续完成。
- Windows 安装前检查三应用是否关闭，并提供只读 `Diagnose.cmd`，定位目录注册、可达性及宿主声明，不读取密钥或文档。
- 明确 PowerPoint 只生成待新增页面，避免把保留原页误生成为占位页；Excel 百分比列默认显示两位小数；Mac 安装及模型设置补充钥匙串访问提示。

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
