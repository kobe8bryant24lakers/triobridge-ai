# TrioBridge AI · 三桥 AI

**English** | [简体中文](README.zh-CN.md)

Use your own models in Word, Excel, and PowerPoint. The workbench opens compiled files from the installer. Running the plugin requires neither a local web server nor a localhost certificate.

## Download and install

| System and processor | Installer |
| --- | --- |
| Windows · Intel/AMD x64 | [TrioBridge-AI-Windows-x64.zip](https://github.com/kobe8bryant24lakers/triobridge-ai/releases/download/v0.6.0/TrioBridge-AI-Windows-x64.zip) |
| Windows · ARM64 | [TrioBridge-AI-Windows-arm64.zip](https://github.com/kobe8bryant24lakers/triobridge-ai/releases/download/v0.6.0/TrioBridge-AI-Windows-arm64.zip) |
| Mac · Apple silicon | [TrioBridge-AI-Mac-arm64.zip](https://github.com/kobe8bryant24lakers/triobridge-ai/releases/download/v0.6.0/TrioBridge-AI-Mac-arm64.zip) |
| Mac · Intel | [TrioBridge-AI-Mac-x64.zip](https://github.com/kobe8bryant24lakers/triobridge-ai/releases/download/v0.6.0/TrioBridge-AI-Mac-x64.zip) |

**Windows** Save your documents, close all three desktop Office apps, extract the ZIP, then double-click **Install.cmd**. Reopen a document and choose **三桥 AI → 打开工作台** in the ribbon.

**Mac** Check your processor in About This Mac, extract the ZIP, and double-click **Install.command**. In Word, Excel, or PowerPoint, choose the application menu → **Services → 打开三桥 AI**. macOS may ask for Automation or Keychain access.

Close the workbench and Office before installing an upgrade. Run the included Uninstall to remove the workbench. Both operations preserve saved model profiles and keys. Use the previous installer’s Uninstall to remove a legacy sign-in HTTPS service.

The Mac package uses an ad hoc signature. The Windows package has no Authenticode signature. Your system may show an unverified-publisher prompt. Managed machines follow your organization's installation policy.

## Model settings

Add a profile in model settings or select an existing provider. Enter your existing OpenAI-compatible API base URL (for example `http://127.0.0.1:8001/v1`), actual model ID, and a key if required. Connect and save sends a real request. The three Office apps share profiles and the active selection. Keys are held in macOS Keychain or Windows Credential Manager.

Cloud inference requires a connection to your provider. Local inference uses your existing model API. The installer contains the workbench and Office integration. Model weights are supplied by your inference runtime.

## Working with documents

- **Word** Read a selection, rewrite, summarize, translate, or insert a report with headings, bullets, and tables. Preview before replacing or inserting.
- **Excel** Read values and formulas from a contiguous selection, preview a formula for a target cell, or create data on a new worksheet.
- **PowerPoint** Generate editable PPTX slides, charts, and tables from a local materials folder. Read existing slides and review changes to text, styling, and positions. The Windows adapter also supports full-slide insertion and reordering.

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
