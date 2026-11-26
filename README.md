# TrioBridge AI · 三桥 AI

**English** | [简体中文](README.zh-CN.md)

Use your own models in Word, Excel, and PowerPoint. Open the workbench in a separate desktop window or beside your document in an Office task pane.

Learn how to connect a model, generate content, and confirm changes in each Office app. [Download the illustrated user guide in English (PowerPoint, 40 slides)](https://kobe8bryant24lakers.github.io/triobridge-ai/assets/TrioBridge-AI-User-Guide-en.pptx)

The workbench opens in Simplified Chinese by default. Use the language selector near the TrioBridge AI logo to choose **English** or **简体中文** in the standalone window or Office task pane. The installed workbench shares this preference across Word, Excel and PowerPoint.

## Download and install

| System and processor | Installer |
| --- | --- |
| Windows · Intel/AMD x64 | [TrioBridge-AI-Windows-x64.zip](https://github.com/kobe8bryant24lakers/triobridge-ai/releases/download/v0.9.1/TrioBridge-AI-Windows-x64.zip) |
| Windows · ARM64 | [TrioBridge-AI-Windows-arm64.zip](https://github.com/kobe8bryant24lakers/triobridge-ai/releases/download/v0.9.1/TrioBridge-AI-Windows-arm64.zip) |
| Mac · Apple silicon | [TrioBridge-AI-Mac-arm64.zip](https://github.com/kobe8bryant24lakers/triobridge-ai/releases/download/v0.9.1/TrioBridge-AI-Mac-arm64.zip) |
| Mac · Intel | [TrioBridge-AI-Mac-x64.zip](https://github.com/kobe8bryant24lakers/triobridge-ai/releases/download/v0.9.1/TrioBridge-AI-Mac-x64.zip) |

**Windows** Save your documents, close all three desktop Office apps, extract the ZIP, then double-click **Install.cmd**. Open Office normally under the same Windows account and choose **打开侧边栏** (task pane) or **打开独立窗口** (separate window) on the **三桥 AI** ribbon.

**Mac** Check your processor in About This Mac, extract the ZIP, and double-click **Install.command**. In Word, Excel, or PowerPoint, choose the application menu → **Services → 打开三桥 AI**. macOS may ask for Automation or Keychain access.

Close the workbench and Office before installing an upgrade. Upgrading or uninstalling the desktop workbench preserves saved model profiles and keys.

The Mac package uses an ad hoc signature. The Windows package has no Authenticode signature. Your system may show an unverified-publisher prompt. Managed machines follow your organization's installation policy.

### Using the task pane

On Windows, open **三桥 AI → 打开侧边栏** in Word, Excel, or PowerPoint to use the workbench beside the current document. Choose **打开独立窗口** to use a separate window at the same time. Both use the same model profiles. The pane loads its files from the installation and works offline with a local model. If WebView2 is missing, close Office and run the Microsoft Edge WebView2 offline installer included at the root of the ZIP.

On Mac, quit Office and run **Enable-Embedded-Taskpane.command** to install the optional local HTTPS service, trust its certificate, and register the task pane. Open the local shared add-in under **My Add-ins**. This pane requires network access to load Office.js. The standalone workbench remains available offline from **Services**. When upgrading the Mac task pane, run **Enable-Embedded-Taskpane.command** from the new ZIP again. To remove both Mac entry points, run **Disable-Embedded-Taskpane.command** before **Uninstall.command**.

## Model settings

In model settings, add a **Local API** or **Custom** profile. Enter your server’s base URL (for example `http://127.0.0.1:8001/v1`) and model ID, then select the **API format** your server supports. With this base URL, OpenAI Chat Completions calls `/v1/chat/completions` and Claude (Anthropic Messages) calls `/v1/messages`. Existing profiles default to OpenAI.

Enter a key if your server requires one. **Connect and save** tests the connection with a real request. Word, Excel and PowerPoint share the saved profiles and the active selection. Keys are stored in macOS Keychain or Windows Credential Manager.

Cloud inference requires a connection to your provider. Local inference uses your existing model API. The installer contains the workbench and Office integration. Model weights are supplied by your inference runtime.

## Working with documents

- **Word** Rewrite, summarize, or translate a selection, or generate a structured report. Preview Markdown results, then insert or replace text with editable Word headings, lists, tables, and emphasis.
- **Excel** Read values and formulas from a contiguous selection, preview a formula for a target cell, or create data on a new worksheet.
- **PowerPoint** Generate editable PPTX slides, charts, and tables from a local materials folder. Read existing slides and review changes to text, styling, and positions. To append a generated deck to the current presentation, use the Windows workbench or the Mac task pane. The Mac standalone window can open the deck as a new presentation.

The integration targets desktop Office for Windows and Mac. Interactive Windows Office acceptance is manual. [Verification notes](docs/TESTING.md) · [Download and installation guide](https://kobe8bryant24lakers.github.io/triobridge-ai/#install)

## Development and verification

```sh
npm ci
npm ci --prefix offline/tooling
npm run check
npm run build:installer
npm run verify:installer
```

Build on the target OS and architecture. Installers are written to `dist-install/`. GitHub Pages publishes the bilingual guide and download links.

## License

[MIT License](LICENSE).
