# TrioBridge AI · 三桥 AI

[English](README.md) | **简体中文**

在 Word、Excel 和 PowerPoint 里使用自己的模型。可在独立窗口中打开工作台，也可在 Office 侧边栏中边编辑边预览结果。

从连接模型到确认写入，按图操作 Word、Excel 和 PowerPoint。[下载中文图解使用指南（PowerPoint，40 页）](https://kobe8bryant24lakers.github.io/triobridge-ai/assets/TrioBridge-AI-User-Guide-zh-CN.pptx)

工作台默认使用简体中文。在独立窗口或 Office 侧边栏中，点三桥 AI 标志旁的语言菜单，选择 **English** 或 **简体中文**。安装版的 Word、Excel 和 PowerPoint 共用此偏好。

## 下载与安装

| 系统和芯片 | 安装包 |
| --- | --- |
| Windows · Intel/AMD x64 | [TrioBridge-AI-Windows-x64.zip](https://github.com/kobe8bryant24lakers/triobridge-ai/releases/download/v0.11.0/TrioBridge-AI-Windows-x64.zip) |
| Windows · ARM64 | [TrioBridge-AI-Windows-arm64.zip](https://github.com/kobe8bryant24lakers/triobridge-ai/releases/download/v0.11.0/TrioBridge-AI-Windows-arm64.zip) |
| Mac · Apple silicon | [TrioBridge-AI-Mac-arm64.zip](https://github.com/kobe8bryant24lakers/triobridge-ai/releases/download/v0.11.0/TrioBridge-AI-Mac-arm64.zip) |
| Mac · Intel | [TrioBridge-AI-Mac-x64.zip](https://github.com/kobe8bryant24lakers/triobridge-ai/releases/download/v0.11.0/TrioBridge-AI-Mac-x64.zip) |

**Windows** 保存文档，完全退出三款 Office，解压 ZIP，再双击 **Install.cmd**。安装后用同一 Windows 账户以普通权限打开文档，在“三桥 AI”功能区选择“打开侧边栏”或“打开独立窗口”。Intel / AMD 电脑选 x64，ARM 电脑选 ARM64。

**Mac** 在“关于本机”确认芯片，解压 ZIP，双击 **Install.command**。在目标应用的菜单栏选择 **Word / Excel / PowerPoint → 服务 → 打开三桥 AI**。首次使用可能出现 macOS 自动化或钥匙串访问提示。

升级时先关闭工作台和 Office，再次运行安装。卸载请运行包内的 Uninstall。两项操作均保留当前账户下的模型配置与密钥。

### 使用内嵌侧边栏

Windows 安装后，在 Word、Excel 或 PowerPoint 的 **三桥 AI → 打开侧边栏** 打开当前文档的工作台。选择 **打开独立窗口** 可同时使用大窗口。两个入口共用模型设置。侧边栏从安装目录加载页面，连接本地模型时可以断网使用。如果电脑缺少 WebView2，请关闭 Office，运行同一 ZIP 根目录中的 Microsoft Edge WebView2 离线安装程序，再打开 Office。

Mac 的内嵌窗格可按需启用。退出三款 Office，双击 **Enable-Embedded-Taskpane.command**，按提示信任本机 HTTPS 证书。重新打开 Office，在“我的加载项”中选择名称含“三桥 AI”和“本机共享”的加载项。该窗格通过本机服务访问模型，加载 Office.js 需要网络。独立窗口仍可从“服务”菜单打开并离线使用。升级 Mac 窗格时，再运行新包中的 **Enable-Embedded-Taskpane.command**。完整卸载两个入口时，先运行 **Disable-Embedded-Taskpane.command**，再运行 **Uninstall.command**。

Mac 包使用临时签名，Windows 包未签署 Authenticode。系统可能提示无法验证发布者，受管设备需符合组织的软件安装策略。

## 模型设置

在“模型设置 → 新增”中选择“本机 API”或“自定义 / 本地模型”。填写服务地址（例如 `http://127.0.0.1:8001/v1`）和模型 ID，按推理服务支持的类型选择 **接口格式**。以上述地址为例，OpenAI（Chat Completions）请求 `/v1/chat/completions`，Claude（Anthropic Messages）请求 `/v1/messages`。已有配置默认使用 OpenAI。

按服务要求填写密钥，或选择“不使用密钥”。点击 **连接并保存** 会发起一次真实请求。Word、Excel 和 PowerPoint 共用模型列表和当前选择，密钥保存在 macOS 钥匙串或 Windows 凭据管理器。

云端推理需要联网。本机推理使用你已运行的模型 API，安装包包含工作台与桌面集成，模型权重由你自己的推理程序提供。

## 在文档中工作

- **Word** 读取选区后可以改写、摘要、翻译，或生成结构化报告。模型返回 Markdown 时，先预览排版，再将标题、列表、表格和加粗等格式写入 Word。
- **Excel** 读取连续区域的数值和公式，预览公式后写入目标单元格，或在新工作表中生成结构化数据。
- **PowerPoint** 选择本地素材目录，生成带可编辑文字、数据图表和表格的 PPTX。也能读取现有页面，预览局部文字、样式和位置改动后应用。要把生成的整套 PPT 追加到当前演示稿，可用 Windows 工作台或 Mac 内嵌侧栏。Mac 独立窗口可将整稿打开为新演示文稿。

可以统一整稿背景，也可以单独指定某几页。打开 PPTX 工作副本，还能套用参考页样式、插页和拆页、刷新原图表及其内嵌工作簿，并按批注和版面问题逐项修改。[PowerPoint 编辑说明](docs/POWERPOINT.md)

桌面集成面向 Windows 和 Mac 版 Office。Windows 桌面 Office 交互按人工验收。 [验证记录](docs/TESTING.md) · [下载与安装说明](https://kobe8bryant24lakers.github.io/triobridge-ai/zh-CN.html#install)

## 开发与验证

```sh
npm ci
npm ci --prefix offline/tooling
npm run check
npm run build:installer
npm run verify:installer
```

在与目标系统和芯片相符的机器上打包。安装文件生成在 `dist-install/`。GitHub Pages 发布双语说明和下载链接。

## License

[MIT License](LICENSE).
