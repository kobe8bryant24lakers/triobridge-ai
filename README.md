# TrioBridge AI · 三桥 AI

让 Word、Excel、PowerPoint 连接你自己的 AI 模型。

TrioBridge AI 是面向 **Mac 和 Windows** 的 Office 任务窗格加载项，使用 React、TypeScript 和 Office.js。具备所需 WebView 和 API 的 Office 永久授权版也可以使用，无需 Microsoft 365 订阅。

[功能介绍与 Mac / Windows 安装指南](https://kobe8bryant24lakers.github.io/triobridge-ai/) · [免密钥浏览器演示](https://kobe8bryant24lakers.github.io/triobridge-ai/app/) · [下载在线版清单](https://github.com/kobe8bryant24lakers/triobridge-ai/releases/latest/download/manifest.xml)

**快速安装在线版**（个人测试旁加载，无需 Node.js、本机服务或开发证书）：

1. 在[安装页](https://kobe8bryant24lakers.github.io/triobridge-ai/#install)下载对应系统的 ZIP，解压后双击 Mac 的 `Install.command` 或 Windows 的 `Install.cmd`。Windows 需解压整个文件夹，按系统提示批准创建专用共享目录。
2. 保存文档并完全退出 Office。重新打开后，在「主页 → 加载项」选择「TrioBridge AI · 三桥 AI（在线版）」；Windows 从「高级 / 更多加载项 → 共享文件夹」添加。
3. 在「模型设置」选择服务、填入 API Key，点击「连接并保存」。常用服务自动填写地址与模型，其他参数在高级设置中。

安装包尚未签名，首次运行可能被系统拦截。Mac 可按「系统设置 → 隐私与安全性」中的提示允许；受组织策略限制时使用[手动安装指南](https://kobe8bryant24lakers.github.io/triobridge-ai/#install)或联系管理员。Windows 自动安装仍待实机验证。共享目录旁加载是微软的测试方式，正式分发应采用 Office 商店或组织部署。

移除安装文件：运行包内的 `Uninstall.command` / `Uninstall.cmd`。只移除三桥在线版自己的清单和目录登记，不清理其他加载项、Office 缓存或模型数据。需要清除密钥时，请先在插件内删除配置。开发版和在线版使用不同清单文件名，可并存。

## 已实现功能

| 应用 | 功能 | 写回方式 |
| --- | --- | --- |
| Word | 选区改写、摘要、翻译、内容起草 | 预览和编辑后确认替换原选区，或插入到光标处 |
| Excel | 公式生成与解释、表格问答、分析建议 | 核对目标位置后写入单个公式；分析和解释可复制 |
| PowerPoint | 演示大纲、文案改写、要点整理 | 确认替换选中文字；支持的版本可从大纲创建基础文字页 |

- 多模型配置与切换、连接测试、5–300 秒超时、取消请求、中文错误提示。
- OpenAI 兼容 Chat Completions 接口，支持无 API Key 的本地模型。
- 密钥默认仅保留到当前会话结束；明确勾选后才明文保存在当前 WebView 的本地存储。
- 配置导入导出不包含密钥。各 Office 应用可能隔离存储，可手动导入复用配置。
- 读取、生成、写入分别由用户触发；写入前再次确认位置和内容。
- Word 跟踪原选区，Excel 固定工作表 ID 与单元格地址，覆盖前检查内容变化。
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

- 通用文本读取使用 `TextCoercion 1.1`；Word 安全写回使用 `WordApi 1.1`；Excel 公式写入使用 `ExcelApi 1.1`。
- PowerPoint 文案改写使用通用文本 API；无法监听选区变化时仅提供复制，避免误覆盖。
- 自动创建幻灯片需要 `PowerPointApi 1.4`；旧版本仍可生成和复制大纲。微软的版本表将 Windows 批量授权此能力列为 Office 2024；Mac 对应最低版本 16.62。
- 大纲最多 12 页，创建基础标题和要点文本框，不生成主题、图片或图表。可在 Office 中继续排版。

**实机验证**覆盖 Mac 版 Word、Excel 和 PowerPoint 的选区读取、生成预览和写入。PowerPoint 还连接过 OpenAI 官方 API，Word 和 Excel 使用本机固定回包验证读写。Windows 和不同永久授权版本仍待实机验证。见[版本与测试记录](docs/TESTING.md)。

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

脚本把本项目清单复制为下列目录中的 `triobridge-ai.xml`：

```text
~/Library/Containers/com.microsoft.Word/Data/Documents/wef/
~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/
```

重新打开 Office，打开文档，在 **主页 → 加载项**（部分版本为 **插入 → 我的加载项**）选择 **TrioBridge AI · 三桥 AI**。服务需保持运行。

移除：`npm run sideload:mac:remove`。若 Office 仍缓存入口，按微软文档清理加载项缓存。

官方步骤：[Mac 旁加载](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac)。

### 2B. Windows 安装（开发旁加载）

1. 创建专用目录，例如 `C:\OfficeAddins\TrioBridge`，复制 `manifest.xml` 到其中。
2. 在文件夹属性中设置网络共享，仅向自己的账号授予所需读写权限，取得 UNC 地址，例如 `\\YOUR-PC\TrioBridge`。
3. 在 Office 的 **文件 → 选项 → 信任中心 → 信任中心设置 → 受信任的加载项目录** 中添加 UNC 地址，勾选“在菜单中显示”。
4. 重启 Office，在 **插入 → 我的加载项 → 共享文件夹** 中选择 TrioBridge AI。入口文字可能因版本而异。

清单中的 `localhost` 指运行 Office 的电脑，HTTPS 服务和证书必须配置在同一台电脑。共享文件夹目录用于开发测试，不是正式生产分发渠道。

官方步骤：[Windows 共享文件夹旁加载](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)。

### 3. 配置模型

在模型设置中选择 DeepSeek、OpenAI 或通义千问（北京），填写自己的 API Key，点击「连接并保存」。成功后自动保存；失败或取消不会保存本次修改。预设只提供默认参数，服务仍需允许浏览器跨域访问，账户也需具有对应模型权限。测试只发送一句简短文字，可能产生少量费用。

其他兼容服务选择「自定义 / 本地模型」，填写服务名称、Base URL 和模型 ID；无需鉴权时可不填 API Key。高级设置可修改地址、模型、超时，也可选择「仅保存配置」而不发起请求。切换服务或更换地址会清空原密钥，避免把密钥发送到另一服务。

预设参数依据：[DeepSeek](https://api-docs.deepseek.com/)、[OpenAI GPT-4.1 mini](https://developers.openai.com/api/docs/models/gpt-4.1-mini)、[通义千问兼容接口](https://help.aliyun.com/zh/model-studio/compatibility-of-openai-with-dashscope)。预设不会覆盖已有自定义配置。

- Base URL 包含服务所需路径，例如 `https://api.example.com/v1`；客户端追加 `/chat/completions`，也接受完整 Chat Completions URL。
- 不自动补 `/v1`；不同服务的路径不同。密钥可以留空。
- 暂不支持流式响应，以及仅提供 Responses / Anthropic 原生等非 Chat Completions 协议的服务。

## 本地模型与 CORS

默认由任务窗格直连模型。远程地址需 HTTPS，服务需允许插件来源（在线版为 `https://kobe8bryant24lakers.github.io`，本机开发版通常为 `https://localhost:3000`）的 CORS，包括 `Authorization`、`Content-Type` 和 OPTIONS 预检。GitHub Pages 提供静态文件；在线版使用的模型服务需允许直接访问。以下代理步骤适用于本机版。

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

GitHub Actions 使用 Node 22 运行类型检查、测试、构建和 Chromium / WebKit 工作流。Office 清单版本为 `1.1.0.0`，应用版本为 `0.2.0`。

```text
src/core/          任务、模型客户端、结果校验、配置存储
src/office/        宿主适配、版本检测、模拟文档
src/components/    工作台、模型设置、写入确认
scripts/           HTTPS 服务、固定代理、安装包构建与平台安装脚本
tests/             核心、Office 模拟、代理集成、浏览器测试
manifest.xml       三应用共用 XML 清单
```

## 部署与限制

在线版通过本仓库的 GitHub Pages 分发。运行 `npm run build:pages` 可在 `dist-pages/` 生成安装指南、`app/` 工作台、在线清单，以及 `install/TrioBridge-AI-Mac.zip` 和 `install/TrioBridge-AI-Windows.zip`。安装包构建需在配有 `zip` 命令的 macOS / Linux 上运行（CI 已具备）；终端用户无需这些构建工具。每个包仅包含安装 / 卸载脚本、说明和固定在线清单，不含密钥、证书或 `.env`。合并到 `main` 后，检查通过的构建由 GitHub Actions 部署到 Pages。版本清单见 [GitHub Releases](https://github.com/kobe8bryant24lakers/triobridge-ai/releases)。

团队可将 `dist/` 托管到可信 HTTPS 域名，更新清单中的 SourceLocation、图标和 AppDomain，通过组织已有的加载项渠道分发。静态托管要求模型允许 CORS，也可自行部署经过鉴权的网关。

写回主要是纯文本，Word 不保证保留选区内部全部富文本样式；PowerPoint 基础页面可继续排版；Excel 不自动批量写公式或创建图表。Office 写入并非通用事务，建页出错可能留下部分新增页，请检查后撤销。

真实 API Key、文档内容、`.env`、证书和私有配置不得提交仓库。

GitHub：<https://github.com/kobe8bryant24lakers/triobridge-ai>

## License

[MIT License](LICENSE)。
