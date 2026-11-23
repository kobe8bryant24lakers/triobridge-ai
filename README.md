# TrioBridge AI · 三桥 AI

**English** | [简体中文](README.zh-CN.md)

Connect Word, Excel, and PowerPoint to your own AI models.

**v0.3.0** adds structured Word reports, Excel data tables, editable PowerPoint outlines, and multiple models shared across local Office apps. [Release notes](CHANGELOG.md#030--2026-11-23) · [Configure once on this computer](docs/LOCAL-MODELS.md)

TrioBridge AI is an Office task pane add-in for **Mac and Windows**, built with React, TypeScript, and Office.js. Perpetually licensed Office editions with the required WebView and APIs can also use it; a Microsoft 365 subscription is not required.

[Features and Mac / Windows installation guide](https://kobe8bryant24lakers.github.io/triobridge-ai/) · [Browser demo without an API key](https://kobe8bryant24lakers.github.io/triobridge-ai/app/) · [Download the hosted manifest](https://github.com/kobe8bryant24lakers/triobridge-ai/releases/latest/download/manifest.xml)

**Quick installation of the hosted version** (sideloading for personal testing; no Node.js, local server, or development certificate required):

1. Download the ZIP for your operating system from the [installation page](https://kobe8bryant24lakers.github.io/triobridge-ai/#install), extract it, and double-click `Install.command` on Mac or `Install.cmd` on Windows. On Windows, extract the entire folder and approve the system prompt to create a dedicated shared directory.
2. Save your documents and quit Office completely. Reopen Office and select **TrioBridge AI · 三桥 AI（在线版）** under **Home → Add-ins**. On Windows, add it through **Advanced / More Add-ins → Shared Folder**.
3. In **Model settings (模型设置)**, select a provider, enter your API key, and click **Connect and save (连接并保存)**. Common providers have preset URLs and models; other options are available in advanced settings.

The installers are unsigned and may be blocked on first launch. On Mac, follow the prompt in **System Settings → Privacy & Security** to allow execution. If organizational policies restrict installation, use the [manual installation guide](https://kobe8bryant24lakers.github.io/triobridge-ai/#install) or contact your administrator. The Windows automatic installer still requires verification on a physical machine. Shared-folder sideloading is a Microsoft testing method; production distribution should use the Office Store or organizational deployment.

To remove the installation, run `Uninstall.command` / `Uninstall.cmd` from the package. This removes only TrioBridge's hosted manifest and directory registration; it does not clear other add-ins, Office caches, or model data. To remove saved keys, delete the configuration inside the add-in first. The development and hosted versions use different manifest filenames and can coexist.

## Features

| Application | Features | Write-back behavior |
| --- | --- | --- |
| Word | Rewrite selections, summarize, translate, draft content, and create structured reports | Confirm replacement of a plain-text selection, or insert a report with sections, bullet points, and native tables |
| Excel | Generate and explain formulas, answer questions about tables, suggest analyses, and create data tables | Write a single formula, or create a data table with formulas and column formatting in a separate new worksheet |
| PowerPoint | Create presentation outlines, rewrite copy, organize key points, and edit individual slides | Edit titles, bullet points, and slide order before confirming slide creation; selected-text replacement is also supported |

- Multiple model profiles, profile switching, connection tests, 5–300 second timeouts, request cancellation, and Chinese error messages.
- OpenAI-compatible Chat Completions APIs, including local models that do not require an API key.
- The local shared edition uses one model catalog across Word, Excel, and PowerPoint under the same OS account, with keys stored in macOS Keychain or the Windows credential store. See [local shared setup](docs/LOCAL-MODELS.md).
- The hosted static edition retains per-application storage: keys are session-only by default, or plaintext WebView storage when explicitly enabled. Imports and exports exclude keys.
- Reading, generation, and writing are each triggered by the user. The destination and content are confirmed again before writing.
- Word tracks the original selection. Excel pins the worksheet ID and cell address and checks for content changes before overwriting.
- Word inspects the selection structure. Selections containing tables, images, fields, equations, or other objects remain readable, but plain-text replacement is restricted. Structured reports are inserted without replacing the original content.
- Excel reads raw values, formatted display values, and formulas together to avoid calculations based on rounded display values. Generated tables use separate worksheets; name conflicts are checked again before writing, and formula errors are reported afterward.
- Reports, data tables, and outlines support both visual and JSON editing. Long results can be scrolled for review; canceling the current edit restores the content from before editing began.
- Opening model settings or the user guide preserves the current draft. Changing reference content or requirements keeps the previous preview and blocks writing until a new result is generated. Cancellation and failures do not clear existing results. Oversized input is preserved in full, with a prompt to process it in smaller sections.
- The application does not save documents, generation history, or telemetry. Real requests contain only the visible reference content, task, and requirements.
- The browser demo requires neither Office nor an API key and clearly labels its fixed outputs. It can also connect to a real model to work with simulated documents.

## Compatibility

Subscription type and API capabilities are separate concerns. This project uses an XML manifest and enables features according to runtime capabilities.

| Environment | Target support |
| --- | --- |
| Windows Microsoft 365 / Office 2021, 2024, LTSC | Modern WebView2; features enabled according to available API versions |
| Mac Microsoft 365 / Office 2021, 2024 | Modern WKWebView; features enabled according to available API versions |
| Office 2016, 2019 / older builds | Best effort; a modern WebView and the required APIs are mandatory, and full support is not guaranteed |
| Office on the web | Architecturally compatible; requires trusted HTTPS and an account / organizational policy that allows add-ins; host verification is still pending |
| IE11, EdgeHTML, WPS, LibreOffice | Not supported |

The baseline engine target is ES2020 / Safari 15.4+, with modern WebView2 required on Windows. An Office product name alone does not guarantee a particular embedded browser or API version. Enterprise policies may also disable add-ins.

- Common text reading uses `TextCoercion 1.1`. Word text and report write-back use `WordApi 1.1`. Excel single-formula writing uses `ExcelApi 1.1`; creating formatted data tables requires `ExcelApi 1.2`.
- PowerPoint copy rewriting uses the common text API. When selection changes cannot be monitored, only copying is offered to avoid overwriting the wrong content.
- Automatic slide creation requires `PowerPointApi 1.4`; older versions can still generate and copy outlines. Microsoft's version table lists Office 2024 for this capability in volume-licensed Windows editions, and version 16.62 as the minimum on Mac.
- Outlines support up to 12 slides and create basic title and bullet-point text boxes, without themes, images, or charts. Layout can be refined in Office.
- `PowerPointApi 1.10` can read the actual slide dimensions and adapt text boxes. Older versions use a basic 720×405 point area. Slide-order changes are checked during creation, and content is written to slides identified by their IDs.

**Testing in native Office applications** covers selection reading, generated previews, and writing in Mac versions of Word, Excel, and PowerPoint. All three Mac applications have now been tested end to end with the configured real OpenAI model, including structured output, native write-back, saved-file verification, and reopening. The earlier fixed-response checks remain separately identified. Windows and different perpetual-license editions still require testing on physical machines. See the [version and testing records](docs/TESTING.md).

Native Mac acceptance completed on November 21–23, 2026. Real-model cases include a six-section Word report with two tables, an Excel workbook with 41 formulas and native recalculation, and ten PowerPoint slides appended to a three-slide original. Separate fixed-response stress cases cover 48 Excel rows with 96 formulas and twelve dense PowerPoint slides. The shared edition was also checked for cross-app switching, conflicting edits, offline recovery, and production-build reopening. The recorded regression suite has 284 passing checks, with 810 additional saved-file assertions and every page of the complex Word/PPT PDFs reviewed. See the [test scope and evidence](docs/TESTING.md). Windows native acceptance remains pending; other versions, protected files, and templates require separate verification.

## Complex document workflows

- **Word reports**: Select **Structured report (结构化报告)** and describe the audience, sections, and table requirements. After generation, edit the title, body, bullet points, and table cells directly, then check the insertion point and confirm. Each report supports up to 12 sections, with tables limited to 8 columns and 60 rows per section.
- **Excel data tables**: Select **Create data table (创建数据表)** and describe the fields, units, and calculations. The preview distinguishes text, numbers, booleans, empty values, and formulas. Select a cell, edit it, apply the change, and then confirm the new worksheet name and destination range. Tables support up to 200 data rows and 12 columns, with no more than 2,000 cells including headers. Formula preflight checks do not replace Excel calculation.
- **PowerPoint outlines**: Select **Presentation outline (演示大纲)** to edit titles and bullet points slide by slide, add or delete slides, and change their order. Dense slides prompt you to split the content; the confirmation dialog shows a readable outline. Existing slides are preserved, and new slides still need layout refinement for the actual presentation.
- **Try without an API key**: Switch the browser workbench to **Complex document examples (复杂文档示例)** to try a six-section report with two tables, a 48-row table with 96 formulas, and a 12-slide outline. All examples use fixed synthetic data, and the interface continuously indicates that no AI has been called.

Generated results remain only in the current workbench's memory. Refreshing or switching the demo application / example resets them. Copy the content or download JSON / text for backup; JSON downloads are not DOCX, XLSX, or PPTX exports.

Official references: [Add-in manifests](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/add-in-manifests), [Word APIs](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets), [Excel APIs](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-requirement-sets), and [PowerPoint APIs](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/powerpoint/powerpoint-api-requirement-sets).

## Run the browser demo locally

Development requires Node.js **22.x starting at 22.13, or version 24 and later**.

```sh
npm ci
npm run dev
```

Open <http://localhost:3000>, then select an application → read the sample selection → choose a task → generate a preview → confirm writing. Simulated documents are stored only in memory and reset on refresh. The HTTP demo URL cannot be used in Office on the web.

## Use the local development version in Office

The root `manifest.xml` points to the local server. The following steps are for development and local model proxying. The hosted version uses the separate manifest available from the release links above.

### 1. Start the local HTTPS server

```sh
# Install the development certificate on first use; the system may ask you to confirm trust or enter a password
npm run certs
npm run dev:office
```

To serve the production build locally:

```sh
npm run build
npm start
```

Both listen at `https://localhost:3000` and bind only to the local loopback interface. Open that address first to verify the certificate, then install the manifest. `npm start` serves `dist/` and supports the local proxy described below. `npm run preview` provides only an HTTP preview and does not include the proxy.

Certificates are read from `~/.office-addin-dev-certs/localhost.crt` and `localhost.key` by default. You can also set `TLS_CERT_PATH` and `TLS_KEY_PATH` in `.env`. The project does not automatically change system certificate trust settings.

### 2A. Install on Mac

```sh
npm run sideload:mac
```

The script generates a private, per-installation shared-model manifest as `triobridge-ai.xml` in these directories:

```text
~/Library/Containers/com.microsoft.Word/Data/Documents/wef/
~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/
```

Reopen Office, open a document, and select **TrioBridge AI · 三桥 AI（本机共享）** under **Home → Add-ins** (or **Insert → My Add-ins** in some versions). Keep the server running.

To remove it, run `npm run sideload:mac:remove`. If Office still caches the entry, follow Microsoft's documentation to clear the add-in cache.

Official instructions: [Sideload on Mac](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac).

### 2B. Install on Windows (development sideloading)

1. Run `npm run manifest:local`, then copy the generated private manifest into a dedicated directory such as `C:\OfficeAddins\TrioBridge`. Keep this installation-specific file private; see [shared model setup](docs/LOCAL-MODELS.md).
2. Enable network sharing in the folder's properties, grant only your own account the required read/write permissions, and obtain the UNC address, such as `\\YOUR-PC\TrioBridge`.
3. In Office, go to **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**, add the UNC address, and enable **Show in Menu**.
4. Restart Office and select TrioBridge AI under **Insert → My Add-ins → Shared Folder**. Labels may vary by version.

The manifest's `localhost` refers to the computer running Office. The HTTPS server and certificate must be configured on that same computer. Shared-folder catalogs are for development testing, not production distribution.

Official instructions: [Windows shared-folder sideloading](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins).

### 3. Configure a model

In model settings, select DeepSeek, OpenAI, or Qwen (Beijing), enter your API key, and click **Connect and save (连接并保存)**. Settings are saved automatically on success; failures or cancellation do not save the current changes. Presets provide default parameters only. Your account must have access to the selected model. The hosted edition requires browser CORS support; the local shared edition calls the provider through its local service and does not require provider CORS support. The connection test sends only a short message and may incur a small charge.

For other compatible services, select **Custom / local model (自定义 / 本地模型)** and enter a provider name, Base URL, and model ID. In shared mode, choose **Do not use a key (不使用密钥)** when authentication is not required. Existing keys can be kept or replaced without returning them to the page. Advanced settings let you change the URL, model, and timeout, or choose **Save configuration only (仅保存配置)** without sending a request. Changing the service URL requires a new key or an explicit choice to use no key; the previous service’s key is never carried to the new address.

Preset references: [DeepSeek](https://api-docs.deepseek.com/), [OpenAI GPT-4.1 mini](https://developers.openai.com/api/docs/models/gpt-4.1-mini), and [Qwen's compatible API](https://help.aliyun.com/zh/model-studio/compatibility-of-openai-with-dashscope). Presets do not overwrite existing custom configurations.

- Include the provider's required path in the Base URL, such as `https://api.example.com/v1`. The client appends `/chat/completions` and also accepts a complete Chat Completions URL.
- `/v1` is not added automatically because providers use different paths. Without authentication, leave the key blank in the hosted edition or choose **Do not use a key** in shared mode.
- Streaming responses and providers that only expose non-Chat Completions protocols, such as Responses or native Anthropic APIs, are not currently supported.

## Local models and CORS

In the hosted edition, the task pane connects directly to the model. Remote endpoints require HTTPS, and the provider must allow CORS from the add-in's origin (`https://kobe8bryant24lakers.github.io` for the hosted version, usually `https://localhost:3000` for local development), including `Authorization`, `Content-Type`, and OPTIONS preflight requests. GitHub Pages serves static files, so model providers used with the hosted version must allow direct access. The local shared edition reads keys from the OS vault and calls models server-side, without provider CORS requirements; it also supports HTTP loopback model endpoints. The fixed-target proxy below is available in the local service.

Requests from an HTTPS page to an HTTP local model may be blocked by mixed-content or private-network policies. A local proxy with a fixed target can address this and can also be used for remote HTTPS services that do not support CORS:

```sh
cp .env.example .env
```

Edit `.env` and set the upstream root URL without `/v1`:

```dotenv
MODEL_PROXY_TARGET=http://127.0.0.1:11434
```

Restart `npm run dev:office` or `npm start`. In model settings, set the Base URL to **`/model-proxy/v1`** and the model ID to the actual name of your local model.

The proxy accepts only same-origin JSON POST requests to Chat Completions paths. It accesses only the locally configured fixed target, does not allow clients to specify arbitrary targets, does not follow redirects, and does not log request bodies or keys. To change the target, edit `.env` and restart. Do not expose the local development proxy directly to the public internet.

## Checks and project structure

```sh
npm run typecheck
npm test
npm run build
npx playwright install chromium webkit
npm run test:e2e
npm run test:pages         # Build and verify the static site, hosted workbench, and manifest
npm run manifest:validate # Microsoft validator; requires network access
```

GitHub Actions uses Node 22 to run type checking, tests, builds, and Chromium / WebKit workflows. The Office manifest version is `1.1.0.0`; the application version is `0.3.0`.

```text
src/core/          Tasks, model client, result validation, configuration storage
src/office/        Host adapters, version detection, simulated documents
src/components/    Workbench, model settings, write confirmation
scripts/           HTTPS server, fixed-target proxy, installer builds, platform install scripts
tests/             Core, Office mocks, proxy integration, browser tests
manifest.xml       Shared XML manifest for all three applications
```

## Deployment and limitations

The hosted version is distributed through this repository's GitHub Pages site. Run `npm run build:pages` to generate the installation guide, `app/` workbench, hosted manifest, `install/TrioBridge-AI-Mac.zip`, and `install/TrioBridge-AI-Windows.zip` in `dist-pages/`. Installer builds require macOS / Linux with the `zip` command available (as configured in CI); end users do not need these build tools. Each package contains only installation / uninstallation scripts, instructions, and the fixed hosted manifest, without keys, certificates, or `.env`. After merging into `main`, GitHub Actions deploys builds that pass the checks to Pages. See [GitHub Releases](https://github.com/kobe8bryant24lakers/triobridge-ai/releases) for release manifests.

Teams can host `dist/` on a trusted HTTPS domain, update the manifest's SourceLocation, icons, and AppDomain, and distribute it through their existing organizational add-in channels. Static hosting requires a model provider that allows CORS; teams can also deploy their own authenticated gateway.

Word plain-text replacement does not guarantee preservation of all rich-text formatting within the selection. Structured reports insert headings, lists, and tables; Word determines the final pagination. New Excel worksheets support bulk data and explicit formulas but do not automatically create charts. PowerPoint creates basic text slides without themes, images, or charts. Office writes are not general-purpose transactions: failures may leave partially created slides or worksheets. The interface explains how to inspect and recover from failures and does not automatically repeat writes.

Never commit real API keys, document content, `.env`, certificates, or private configurations to the repository.

GitHub: <https://github.com/kobe8bryant24lakers/triobridge-ai>

## License

[MIT License](LICENSE).
