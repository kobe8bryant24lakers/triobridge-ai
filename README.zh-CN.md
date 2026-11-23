# TrioBridge AI · 三桥 AI

[English](README.md) | **简体中文**

让 Word、Excel、PowerPoint 连接你自己的 AI 模型。

**v0.4.0** 提供自带运行环境的本机共享安装包。同一电脑、同一系统账号配置一次，Word、Excel、PowerPoint 复用多模型及默认选择；支持 `http://127.0.0.1:8001/v1` 等本地 API。[更新记录](CHANGELOG.md#040--2026-11-23) · [共享与迁移](docs/LOCAL-MODELS.md)

## 下载与安装

| 电脑 | 本机共享版安装包 |
| --- | --- |
| Windows Intel / AMD | [Windows x64](https://github.com/kobe8bryant24lakers/triobridge-ai/releases/download/v0.4.0/TrioBridge-AI-Local-Windows-x64.zip) |
| Windows ARM / Snapdragon | [Windows ARM64](https://github.com/kobe8bryant24lakers/triobridge-ai/releases/download/v0.4.0/TrioBridge-AI-Local-Windows-arm64.zip) |
| Mac Apple M | [Mac Apple silicon](https://github.com/kobe8bryant24lakers/triobridge-ai/releases/download/v0.4.0/TrioBridge-AI-Local-Mac-arm64.zip) |
| Mac Intel | [Mac Intel](https://github.com/kobe8bryant24lakers/triobridge-ai/releases/download/v0.4.0/TrioBridge-AI-Local-Mac-x64.zip) |

1. 先保存文档并完全退出 Word、Excel、PowerPoint，再解压完整文件夹，双击 Windows 的 `Install.cmd` 或 Mac 的 `Install.command`。无需另外安装 Node.js 或输入命令。安装器配置本机服务、HTTPS 证书和三应用清单；Windows 创建当前账号专用共享目录时请求管理员授权。
2. 按系统提示完成安装，重开 Office。在每个应用的“主页 → 加载项”分别添加 **三桥 AI（本机共享）**；Windows 从“高级 / 更多加载项 → 共享文件夹”添加。模型设置页应显示 **本机共享**。
3. 在任一应用新增模型并“连接并保存”。本地 API 无需鉴权时选择“不使用密钥”。其余应用可复用，无需重复填写，密钥保存在系统安全存储。

服务随当前账号登录启动，安装完成后可以关闭安装窗口。再次运行新版 `Install` 可升级或修复；`Restart` 重启服务；`Uninstall` 移除本次安装的服务、证书和入口，**保留模型配置与密钥**。需要删除模型时在插件内操作。

Windows 只有部分应用显示加载项时，运行 `Diagnose.cmd`，只读检查应用版本、目录注册 / 可达性和三应用声明，不读取模型密钥或文档。[按应用排查入口](docs/LOCAL-MODELS.md#windows-只有部分应用显示加载项)。

安装包尚未签名，首次运行可能需要系统确认。Mac 可按“系统设置 → 隐私与安全性”提示允许；Office 数据访问和证书信任需按系统提示处理。组织策略可能限制旁加载。Windows 原生 Office 验收仍待完成。[完整安装与故障处理](docs/LOCAL-MODELS.md)。

[浏览器演示（无需密钥）](https://kobe8bryant24lakers.github.io/triobridge-ai/app/) · [功能与下载页面](https://kobe8bryant24lakers.github.io/triobridge-ai/zh-CN.html#install)

**在线版仍可选用**：下载页单独提供在线 ZIP 与 XML 清单，从网站加载工作台，各应用独立保存配置，模型 API 需允许 CORS。Windows 的“共享文件夹”只是安装入口，不代表配置共享。访问本地 HTTP API 或复用模型配置，请选择上面的本机共享版。

## 已实现功能

| 应用 | 功能 | 写回方式 |
| --- | --- | --- |
| Word | 选区改写、摘要、翻译、内容起草、结构化报告 | 确认替换纯文字选区，或插入带章节、要点和原生表格的报告 |
| Excel | 公式生成与解释、表格问答、分析建议、创建数据表 | 写入单个公式，或在独立新工作表创建带公式、列格式的数据表 |
| PowerPoint | 演示大纲、文案改写、要点整理、逐页编辑 | 修改标题、要点和页序后确认建页；支持选中文字替换 |

- 多模型配置与切换、连接测试、5–300 秒超时、取消请求、中文错误提示。
- OpenAI 兼容 Chat Completions 接口，支持无 API Key 的本地模型。
- 本机共享版在同一系统账号下统一管理 Word、Excel、PowerPoint 的模型列表与默认模型，密钥保存在 Mac 钥匙串 / Windows 系统安全存储。见[同机共享配置](docs/LOCAL-MODELS.md)。
- 在线静态版仍使用各应用独立存储：密钥默认仅当前会话有效，明确勾选后才明文保存在 WebView 本地存储。配置导入导出不包含密钥。
- 读取、生成、写入分别由用户触发；写入前再次确认位置和内容。
- Word 跟踪原选区，Excel 固定工作表 ID 与单元格地址，覆盖前检查内容变化。
- Word 检查选区结构；包含表格、图片、域、公式等对象时保留读取能力，并限制纯文本替换。结构化报告只插入，不覆盖原文。
- Excel 同时读取原始值、格式化显示值和公式，避免使用四舍五入后的显示值计算。生成数据表使用独立工作表，名称冲突在写入前重新检查，公式错误在写入后报告。
- 报告、数据表和大纲均有可视化编辑及 JSON 编辑。长结果可滚动审阅；取消本轮编辑可恢复进入编辑前的内容。
- 进入模型设置或使用指南保留当前工作稿。修改参考内容或要求时保留旧预览并阻止写入；取消和失败不会清空已有结果。超长输入会完整保留并提示分段处理。
- 应用不保存文档、生成历史或遥测；真实请求只携带可见参考内容、任务与要求。
- 浏览器演示无需 Office 或密钥，固定输出明确标识；也能接入真实模型处理模拟文档。

## 兼容范围

订阅类型与 API 能力是两回事。本项目使用 XML manifest，并按运行时能力启用功能。

| 环境 | 目标支持范围 |
| --- | --- |
| Windows Microsoft 365 / Office 2021、2024、LTSC | 现代 WebView2，按实际 API 版本启用功能 |
| Mac Microsoft 365 / Office 2021、2024 | 现代 WKWebView，按实际 API 版本启用功能 |
| Office 2016、2019 / 旧构建 | 尽力兼容；必须具备现代 WebView 与所需 API，不承诺完整支持 |
| Office 网页版 | 架构兼容；需可信 HTTPS 和允许加载项的账号 / 组织策略，尚待宿主验证 |
| IE11、EdgeHTML、WPS、LibreOffice | 不支持 |

基础引擎目标为 ES2020 / Safari 15.4+，Windows 需现代 WebView2。Office 产品名称不能单独保证内置浏览器与 API 版本；企业策略也可能禁用加载项。

- 通用文本读取使用 `TextCoercion 1.1`；Word 文本与报告写回使用 `WordApi 1.1`；Excel 单公式写入使用 `ExcelApi 1.1`，新建格式化数据表需要 `ExcelApi 1.2`。
- PowerPoint 文案改写使用通用文本 API；无法监听选区变化时仅提供复制，避免误覆盖。
- 自动创建幻灯片需要 `PowerPointApi 1.4`；旧版本仍可生成和复制大纲。微软的版本表将 Windows 批量授权此能力列为 Office 2024；Mac 对应最低版本 16.62。
- 大纲最多 12 页，创建基础标题和要点文本框，不生成主题、图片或图表。可在 Office 中继续排版。
- PowerPointApi 1.10 可读取实际页面尺寸并适配文字框；较旧版本使用 720×405 点的基础区域。创建时检查页序变化，填充时按页面 ID 绑定目标。

**实机验证**覆盖 Mac 版 Word、Excel 和 PowerPoint 的选区读取、生成预览和写入。三类 Mac 应用均已使用配置的真实 OpenAI 模型完成复杂内容生成、原生写入、保存文件核对与重开；早期固定回包检查单独记录。Windows 和不同永久授权版本仍待实机验证。见[版本与测试记录](docs/TESTING.md)。

2026-11-21 至 11-23 完成 Mac 原生验收：真实模型生成六节两表的 Word 报告、含 41 个公式且经过原生重算的 Excel 工作簿，以及在三页原稿后追加十页的 PowerPoint。另有固定回包压力案例覆盖 Excel 48 行、96 个公式和 PPT 12 页密集中文。共享版也验证了跨应用切换、并发冲突、离线恢复与正式构建重开。记录包含 284 项回归检查、810 项保存文件检查，复杂 Word/PPT 的 PDF 全部逐页复核。详见[验证范围与证据](docs/TESTING.md)。Windows 原生验收仍待完成；其他版本、受保护文件和模板需独立验证。

## 复杂文档工作流

- **Word 报告**：选择「结构化报告」，描述读者、章节与表格要求；生成后直接编辑标题、正文、要点和表格单元格，核对插入位置后确认。每份最多 12 节，每节表格最多 8 列、60 行。
- **Excel 数据表**：选择「创建数据表」，描述字段、单位和计算项；预览中区分文本、数字、布尔值、空值与公式。选择单元格后编辑并应用修改，再确认新工作表名称和写入范围。最多 200 行数据、12 列，含表头不超过 2,000 格。公式预检不代替 Excel 计算。
- **PPT 大纲**：选择「演示大纲」，逐页修改标题和要点，可增删或调整页序。密集页面会提示拆分；确认框显示可读大纲。原有幻灯片保留，新页面仍需按实际演示需求排版。
- **免密钥体验**：浏览器工作台切换到「复杂文档示例」，使用 6 节含 2 张表的报告、48 行含 96 个公式的数据表、12 页大纲检验操作。全部为固定合成数据，界面持续标明未调用 AI。

生成结果仅在当前工作台内存中保留，刷新或切换演示应用 / 示例会重置。可复制内容或下载 JSON / 文本备份；JSON 下载不是 DOCX、XLSX 或 PPTX 导出。

官方依据：[加载项清单](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/add-in-manifests)、[Word API](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)、[Excel API](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-requirement-sets)、[PowerPoint API](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/powerpoint/powerpoint-api-requirement-sets)。

## 本地运行浏览器演示

开发环境使用 Node.js **22.13 及以上的 22.x，或 24 及以上版本**。

```sh
npm ci
npm run dev
```

打开 <http://localhost:3000>，选择应用 → 读取示例选区 → 选择任务 → 生成预览 → 确认写入。模拟文档仅在内存中保存，刷新后重置。HTTP 演示入口不能用于 Office 网页版。

## 在 Office 中使用本地开发版

仓库根目录的 `manifest.xml` 指向本机服务。以下步骤适用于开发和本机模型代理。在线版使用上方发布入口提供的独立清单。

### 1. 启动本机 HTTPS 服务

```sh
# 首次安装开发证书；系统可能要求确认信任或输入密码
npm run certs
npm run dev:office
```

运行构建产物：

```sh
npm run build
npm start
```

两者都监听 `https://localhost:3000`，只绑定本机回环接口。先打开该地址验证证书，再安装清单。`npm start` 服务 `dist/` 并支持下文的本机代理。`npm run preview` 仅用于 HTTP 预览，不包含代理。

证书默认读取 `~/.office-addin-dev-certs/localhost.crt` 与 `localhost.key`；也可在 `.env` 中配置 `TLS_CERT_PATH`、`TLS_KEY_PATH`。项目不会自动修改系统证书信任设置。

### 2A. Mac 安装

```sh
npm run sideload:mac
```

脚本为当前系统账号生成本机共享版专用清单，并以私有权限保存到下列目录中的 `triobridge-ai.xml`：

```text
~/Library/Containers/com.microsoft.Word/Data/Documents/wef/
~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/
```

重新打开 Office，打开文档，在 **主页 → 加载项**（部分版本为 **插入 → 我的加载项**）选择 **TrioBridge AI · 三桥 AI（本机共享）**。服务需保持运行。

移除：`npm run sideload:mac:remove`。若 Office 仍缓存入口，按微软文档清理加载项缓存。

官方步骤：[Mac 旁加载](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac)。

### 2B. Windows 安装（开发旁加载）

1. 运行 `npm run manifest:local`，将命令生成的私有清单放入专用目录，例如 `C:\OfficeAddins\TrioBridge`。这份文件只用于本机当前账号，不应分发；见[同机共享配置](docs/LOCAL-MODELS.md)。
2. 在文件夹属性中设置网络共享，仅向自己的账号授予所需读写权限，取得 UNC 地址，例如 `\\YOUR-PC\TrioBridge`。
3. 在 Office 的 **文件 → 选项 → 信任中心 → 信任中心设置 → 受信任的加载项目录** 中添加 UNC 地址，勾选“在菜单中显示”。
4. 保存文档并完全退出、重开 Office，在 **插入 → 我的加载项 → 共享文件夹** 中选择 **TrioBridge AI · 三桥 AI（本机共享）**。模型设置页应显示 **本机共享**。

Windows 的“共享文件夹”只是加载项安装目录，不等于模型配置共享。在线版 ZIP 也使用此入口。请安装上一步生成的 `triobridge-ai.xml`，不要使用仓库根目录的原始 `manifest.xml` 或在线版清单。若连接本地 API 时仍出现包含 CORS 的直连报错，请检查是否打开了旧入口；见[本地 API 排查](docs/LOCAL-MODELS.md#本地-http-api-与连接排查)。

清单中的 `localhost` 指运行 Office 的电脑，HTTPS 服务和证书必须配置在同一台电脑。共享文件夹目录用于开发测试，不是正式生产分发渠道。

官方步骤：[Windows 共享文件夹旁加载](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)。

### 3. 配置模型

在模型设置中选择 DeepSeek、OpenAI 或通义千问（北京），填写自己的 API Key，点击「连接并保存」。成功后自动保存；失败或取消不会保存本次修改。预设只提供默认参数，账户需具有对应模型权限。在线版要求服务允许浏览器跨域访问；本机共享版由本机服务调用，不依赖模型端点的 CORS。测试只发送一句简短文字，可能产生少量费用。

其他兼容服务选择「自定义 / 本地模型」，填写服务名称、Base URL 和模型 ID；共享版无需鉴权时请选择「不使用密钥」。已有密钥可选择保留或替换，不会读回到页面。高级设置可修改地址、模型、超时，也可选择「仅保存配置」而不发起请求。更换服务地址后必须填写新密钥或明确选择不使用密钥，不会把旧地址的密钥发送到新服务。

预设参数依据：[DeepSeek](https://api-docs.deepseek.com/)、[OpenAI GPT-4.1 mini](https://developers.openai.com/api/docs/models/gpt-4.1-mini)、[通义千问兼容接口](https://help.aliyun.com/zh/model-studio/compatibility-of-openai-with-dashscope)。预设不会覆盖已有自定义配置。

- Base URL 包含服务所需路径，例如 `https://api.example.com/v1`；客户端追加 `/chat/completions`，也接受完整 Chat Completions URL。
- 不自动补 `/v1`；不同服务的路径不同。无需鉴权时，在线版可留空密钥，共享版选择「不使用密钥」。
- 暂不支持流式响应，以及仅提供 Responses / Anthropic 原生等非 Chat Completions 协议的服务。

## 本地模型与 CORS

在线版由任务窗格直连模型。远程地址需 HTTPS，服务需允许插件来源（在线版为 `https://kobe8bryant24lakers.github.io`，本机开发版通常为 `https://localhost:3000`）的 CORS，包括 `Authorization`、`Content-Type` 和 OPTIONS 预检。GitHub Pages 提供静态文件；在线版使用的模型服务需允许直接访问。本机共享版由服务端读取系统密钥并请求模型，无需浏览器跨域授权，也支持 HTTP 回环模型地址。以下固定目标代理步骤同样适用于本机服务。

HTTPS 页面访问 HTTP 本机模型可能被混合内容 / 私有网络策略阻止。可用固定目标本机代理，也适用于不支持 CORS 的远程 HTTPS 服务：

```sh
cp .env.example .env
```

编辑 `.env`，填写上游根地址，不含 `/v1`：

```dotenv
MODEL_PROXY_TARGET=http://127.0.0.1:11434
```

重启 `npm run dev:office` 或 `npm start`，模型设置中 Base URL 填 **`/model-proxy/v1`**，模型 ID 填本地实际模型名。

代理只接受同源 JSON POST 和 Chat Completions 路径，仅访问本机配置的固定目标，不允许客户端指定任意目标、不跟随重定向、不记录正文和密钥。更换目标需修改 `.env` 并重启。不要将本机开发代理直接暴露到公共网络。

## 检查与项目结构

```sh
npm run typecheck
npm test
npm run build
npx playwright install chromium webkit
npm run test:e2e
npm run test:pages        # 构建并验证静态站点、在线工作台及清单
npm run manifest:validate  # 微软验证器，需要网络
```

GitHub Actions 使用 Node 22 运行类型检查、测试、构建和 Chromium / WebKit 工作流。Office 清单版本为 `1.1.0.0`，应用版本为 `0.4.0`。

```text
src/core/          任务、模型客户端、结果校验、配置存储
src/office/        宿主适配、版本检测、模拟文档
src/components/    工作台、模型设置、写入确认
scripts/           HTTPS 服务、固定代理、安装包构建与平台安装脚本
tests/             核心、Office 模拟、代理集成、浏览器测试
manifest.xml       三应用共用 XML 清单
```

## 部署与限制

本机共享包通过 GitHub Releases 分发。在对应操作系统上运行 `npm run build:local`，在 `dist-local/` 生成当前架构的包。CI 分别构建 Mac / Windows 的 x64 / ARM64；运行时来自带固定校验值的 Node.js 官方下载，原生凭据模块按锁文件完整性校验。发布包不含用户证书、密钥、私有清单或 `.env`，这些数据在目标电脑生成。

在线版通过本仓库的 GitHub Pages 分发。运行 `npm run build:pages` 可在 `dist-pages/` 生成安装指南、`app/` 工作台、在线清单，以及 `install/TrioBridge-AI-Mac.zip` 和 `install/TrioBridge-AI-Windows.zip`。安装包构建需在配有 `zip` 命令的 macOS / Linux 上运行（CI 已具备）；终端用户无需这些构建工具。每个包仅包含安装 / 卸载脚本、说明和固定在线清单，不含密钥、证书或 `.env`。合并到 `main` 后，检查通过的构建由 GitHub Actions 部署到 Pages。版本清单见 [GitHub Releases](https://github.com/kobe8bryant24lakers/triobridge-ai/releases)。

团队可将 `dist/` 托管到可信 HTTPS 域名，更新清单中的 SourceLocation、图标和 AppDomain，通过组织已有的加载项渠道分发。静态托管要求模型允许 CORS，也可自行部署经过鉴权的网关。

Word 纯文字替换不保证保留选区内部全部富文本样式；结构化报告插入标题、列表和表格，最终分页由 Word 决定。Excel 新表支持批量数据和显式公式，不自动创建图表；PowerPoint 创建基础文字页，不生成主题、图片或图表。Office 写入并非通用事务，失败可能留下部分新增页或工作表；界面会说明检查和恢复方法，不自动重复写入。

真实 API Key、文档内容、`.env`、证书和私有配置不得提交仓库。

GitHub：<https://github.com/kobe8bryant24lakers/triobridge-ai>

## License

[MIT License](LICENSE)。
