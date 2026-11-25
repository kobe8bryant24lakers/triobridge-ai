# 同机共享模型

## 桌面工作台

打开“模型设置 → 新增 → 本机 API”或“自定义 / 本地模型”。填写推理程序显示的地址和模型 ID。按推理服务支持的接口选择格式。同一模型 ID 可分别保存两种格式，升级前已有的配置默认使用 OpenAI。

| 接口格式 | 基础地址示例 | 工作台发起请求的路径 |
| --- | --- | --- |
| OpenAI（Chat Completions） | `http://127.0.0.1:8001/v1` | `/v1/chat/completions` |
| Claude（Anthropic Messages） | `http://127.0.0.1:8001/v1` | `/v1/messages` |

可根据推理服务文档的请求示例识别格式。OpenAI 响应的正文位于 `choices[0].message.content`，Claude 响应使用 `content` 内容块。[OpenAI 接口文档](https://developers.openai.com/api/reference/resources/chat/subresources/completions/methods/create) · [Claude 接口文档](https://platform.claude.com/docs/en/api/messages/create)

设置页会显示实际请求地址，也可以直接粘贴该格式的完整接口地址。推理服务允许匿名访问时，选择“不使用密钥”。点击“连接并保存”会用所选格式发送一条测试消息。在任一工作台切换当前模型，三个应用会共用该配置。原有配置默认使用 OpenAI。

## v0.5.x 及更早版本

以下安装和诊断步骤适用于使用本机 HTTPS 服务的旧版加载项。

本机共享版在同一台电脑、同一个系统账号下，让 Word、Excel、PowerPoint 共用模型列表和当前默认模型。可以新增多个模型，并在任一应用切换；其他应用会在刷新、重新聚焦或下一次同步时更新。

在线静态版仍使用各应用自己的存储，不能保证跨 Office 共享。微软明确不保证 cookies 或 localStorage 在不同 Office 应用之间通用，见[加载项状态与设置](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/persisting-add-in-state-and-settings)。

## 安装本机共享版

从[下载页面](https://kobe8bryant24lakers.github.io/triobridge-ai/zh-CN.html#install)选择与你的电脑和芯片匹配的 **Local / 本机共享版** ZIP，解压整个文件夹。先保存文档并完全退出 Word、Excel、PowerPoint，再双击 Windows 的 `Install.cmd` 或 Mac 的 `Install.command`，按系统提示完成安装。无需另外安装 Node.js、npm 或开发工具。

安装程序为当前账号配置本机 HTTPS 服务、证书和三应用清单；Windows 为当前账号建立专用共享目录，只有创建 / 移除共享目录这一步请求管理员授权。服务随当前账号登录启动，安装完成后可以关闭窗口。包内不含个人密钥或证书，每台电脑安装时分别生成。

重开 Office，在每个应用内分别添加 **三桥 AI（本机共享）**。Windows 在“主页 → 加载项 → 高级 / 更多加载项 → 共享文件夹”添加。模型设置必须显示 **本机共享**；在线版入口仍然存在时，请核对名称。

- **升级 / 修复**：下载新版 ZIP，解压后再次运行 `Install`。保留已有模型和密钥；失败会尝试恢复原服务。
- **重启**：运行包内 `Restart`。同一电脑的开发服务与安装版都使用 3000 端口，切换前停止旧服务。
- **卸载**：运行 `Uninstall`，移除本次安装的后台服务、入口与证书。模型目录和系统密钥保留；删除模型请在插件内操作。
- **系统提示**：安装包尚未签名。Mac 如阻止运行，在“系统设置 → 隐私与安全性”按提示允许。若系统请求访问 Word / Excel / PowerPoint 数据目录或信任 localhost 证书，需在本机处理。首次使用已存密钥或升级后，还可能请求本机服务访问钥匙串；在系统窗口确认后重试，无需重新填写密钥。不要在聊天或日志中提供系统密码。
- **目录**：程序位于 Mac 的 `~/Library/Application Support/TrioBridge AI/local-app/` 或 Windows 的 `%LOCALAPPDATA%\TrioBridgeAI\LocalApp\`。`logs/` 保存服务启动日志，不保存模型密钥或文档正文。

安装完成后参考下方“配置与迁移”。以下源码方式只供开发和自行构建使用。

### Windows 安装停在第 3 步

从本机共享版 v0.5.1 起，安装窗口会显示正在检查的项目及等待秒数。第 3 步先释放 3000 端口，然后准备本账号的凭据、启动登录任务、验证服务的 HTTPS 证书和安装标识，最后读取共享配置。连接模型 API 是安装完成后在插件里单独执行的操作。

如果窗口报错，先记下方括号中的阶段和代码。双击包内 `Diagnose.cmd`，或解压并运行[独立诊断工具](https://kobe8bryant24lakers.github.io/triobridge-ai/install/TrioBridge-Diagnose-Windows.zip)。工具可以在安装期间和回退后运行。

| 诊断字段或代码 | 如何使用 |
| --- | --- |
| `localService.lastStartup` | 最近一次完成的启动检查。`phase` 与 `code` 记录结果。失败时尽可能保留当时的计划任务状态和 Windows 错误码。旧安装包没有这条记录。 |
| `localService.task` 和 `tcp3000` | 运行诊断时计划任务的状态、最近返回值及端口能否建立 TCP 连接。升级回退后，当前运行的可能是已恢复的旧版本。 |
| `PORT_BUSY` / `PORT_DENIED` | 前者请检查当前使用 3000 端口的程序。后者请让管理员检查回环端口保留范围和本机策略。 |
| `TASK_START` / `NO_LISTENER` | 查看记录中的任务状态、返回值或 HRESULT，确认当前账号可启动该登录任务。 |
| `TLS_CERTIFICATE` / `SERVICE_IDENTITY` | 核对端口上的程序与本次安装的证书，再用同一账号运行 Install 修复。 |
| `LOCAL_AUTH` / `CAPABILITY_ACCESS` | 核对安装和打开 Office 使用的系统账号。保留原凭据，使用同一账号运行 Install 修复。 |
| `MODEL_STORAGE` | 本机 HTTPS 已通过验证，读取共享配置时收到 503。v0.5.2 会显示第二个原因码，回退后可在 `localService.lastStartup.storageDetail` 查看。 |
| `CATALOG_BUSY` | 安装器会重试临时占用。如果最终仍被占用，关闭正在独占访问该配置的程序，再运行 Install。 |
| `ACL_OWNER` / `ACL_APPLY` / `ACL_VERIFY` / `ACL_INSPECT` | 在诊断报告的 `localService.storage` 查看目录和文件的状态、所有者是否匹配及是否设置了专用权限。需要协助时可提供这些字段。 |
| `CATALOG_JSON` / `CATALOG_SCHEMA` | 已有配置的 JSON 或结构未通过校验。保留原文件，核对可信备份。 |
| `CATALOG_ACCESS` / `DIRECTORY_ACCESS` | 检查当前系统账号是否可读取对应文件或目录。首次安装时 `profiles.json` 可以尚未创建。 |

`Running / 0x00041301` 表示登录任务正在运行。遇到 `models/MODEL_STORAGE` 时应先查看配置存储原因码。v0.5.1 只记录总类，请使用新版安装包复测。独立诊断仅检查文件类型、所有者匹配和访问权限标志，报告不会包含模型配置值、用户名或密钥。[微软任务计划程序状态码](https://learn.microsoft.com/en-us/windows/win32/taskschd/task-scheduler-error-and-success-constants)

## Windows 只有部分应用显示加载项

共享目录与模型配置可复用，但首次添加加载项需要在 Word、Excel、PowerPoint 中分别操作。请打开普通空白文档，在“主页 → 加载项 → 高级 / 更多加载项”查看“共享文件夹”；“文件 → 选项 → 加载项”中的 COM 加载项列表不是此入口。

若某个应用仍没有“共享文件夹”，在该应用打开“文件 → 选项 → 信任中心 → 信任中心设置 → 受信任的加载项目录”，检查是否存在与可用应用相同的目录，并勾选“显示在菜单中”。注册目录前须保存文档并关闭所有 Office 应用；新版安装器会检查当前会话中仍运行的三个宿主，不强制关闭文档。`Restart` 只重启本机服务，可在 Office 开着时使用。参见[微软共享目录旁加载步骤](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)。

包内 `Diagnose.cmd` 可在 Office 打开时运行，不需要管理员授权，也不改动目录、缓存、策略或模型配置。报告包含上方的本机服务状态、检测到的 Office 版本、运行中的应用名称和下表的加载项目录状态。

| 字段 | 含义 |
| --- | --- |
| `registered` / `urlMatches` | 当前账号的目录 ID、路径是否与本安装匹配 |
| `showInMenu` | 注册项是否启用了目录菜单显示 |
| `reachable` | 当前账号能否读取本机共享目录中的清单 |
| `manifest.Word` / `Excel` / `PowerPoint` | 实际安装的清单是否声明了对应应用 |
| `detectedVersions` | 从应用注册信息读取的文件版本；`not detected` 表示未能自动识别，不等于没有安装 |

目录字段均通过，仍不能证明 Office 界面已加载成功。若重开后仍不可见，保留这份诊断结果及应用内“关于”的产品版本，继续核对对应宿主的目录页面与策略；不要把该情况记录为验收通过。

### 从源码运行（开发者）

需要 Node.js 22.13+ 或 24+，并安装项目依赖。Mac 与 Windows 的密钥存储使用平台原生依赖；请在实际运行服务的操作系统上执行 `npm ci`，不要跨系统复制 `node_modules`。

```sh
npm ci
# 首次配置本机 HTTPS 证书，按系统提示完成信任
npm run certs
npm run build
npm start
```

保持本机服务运行。开发时可用 `npm run dev:office` 替代 `npm start`。两者使用同一份共享模型目录；不能同时占用 3000 端口。

Mac 在另一个终端运行：

```sh
npm run sideload:mac
```

保存文档并完全退出、重开 Office，选择「TrioBridge AI · 三桥 AI（本机共享）」。安装脚本为三个应用生成同一账号专用的清单，不改动在线版清单。

Windows 先生成本机专用清单：

```sh
npm run manifest:local
```

命令只打印清单文件路径。按 README 的 Windows 受信任共享目录步骤安装这份生成文件，并把目录权限限制为当前账号；不要使用仓库中未引导共享连接的原始 `manifest.xml`。Windows 原生 Office 行为仍需在实际 Office 中验收。

生成的清单含本机安装凭证，必须保留在本机私有目录，不要上传、提交仓库或分发给其他人。模型 API Key 不写入清单。不同电脑和系统账号各自安装，不同步账户之间的密钥。

## 配置与迁移

已使用本机源码共享版的账号会复用同一份模型目录及系统密钥。在线版位于另一个网页来源，浏览器不允许本机共享版读取它的密钥；切换版本后需在共享版配置一次，之后三个应用复用。

1. 在任一应用进入「模型设置」，确认页面显示本机共享状态。
2. 新增模型，填写服务地址、模型 ID 和 API Key，连接并保存。密钥进入 macOS 钥匙串或 Windows 系统凭据存储。
3. 若已有旧配置，使用迁移入口选择要迁移的模型。整批成功后才移除相应的旧凭证；失败或并发修改时保留来源。未选择的旧配置不被删除。
4. 在工作台切换模型，三类应用同步当前默认模型。模型更新会检查版本冲突；遇到冲突先重新读取并复核草稿，不自动覆盖另一应用的新配置。

共享模式只向页面返回模型元数据和“已保存密钥”的状态，不读回已保存的密钥。保留、替换、删除密钥是不同操作；更换服务地址不会把旧密钥带到新地址。配置导出仍不包含密钥，也不包含本机安装凭证。

服务离线后，已进入共享模式的页面保持共享模型信息并提示重新连接，不会悄悄使用旧的本应用密钥继续调用。缺少本机安装凭证时，重新执行安装清单步骤并重开 Office。

## 本地 HTTP API 与连接排查

支持 `http://127.0.0.1:8001/v1` 这样的本地 OpenAI 兼容 API。插件追加 `/chat/completions`；模型 ID 填写 API 实际提供的名称。服务无需鉴权时，明确选择“不使用密钥”。共享版由本机 Node.js 服务调用 API，不要求模型 API 提供浏览器 CORS。

**Windows 的“共享文件夹”不等于“本机共享模型”。** 它只是加载项安装目录，在线版也能通过它安装。判断依据是模型设置页显示“本机共享”，而不是“仅此应用”。

若连接时报“无法连接模型服务……CORS……本机代理”，表示当前使用浏览器直连流程。请在运行 Office 的同一电脑、同一系统账号下安装上方本机共享 ZIP，并重新选择“本机共享”入口。源码开发者可执行 `npm run manifest:local`（Windows）或 `npm run sideload:mac`（Mac）。不要安装仓库原始清单、在线版清单或另一台电脑生成的私有清单。保存文档并完全退出、重开 Office，选择名称含“本机共享”的加载项；安装版服务自动随登录启动，源码开发者需保持 `npm start` 运行。

确认共享模式后，可在运行 TrioBridge 本机服务的电脑上用 Node.js 检查 API 是否可达（不发送密钥，不修改配置）：

```sh
node -e "fetch('http://127.0.0.1:8001/v1/models',{signal:AbortSignal.timeout(5000)}).then(async r=>{console.log('HTTP',r.status);await r.body?.cancel()}).catch(e=>console.error(e.cause?.code||e.name))"
```

- `HTTP 200` 表示该模型列表接口可访问；`HTTP 401/403` 表示已连到服务，但需要鉴权；`HTTP 404` 表示已收到 HTTP 响应，需要确认服务是否提供此路径。这些结果都不能代替 Chat Completions 的“连接并保存”测试。
- `ECONNREFUSED` 表示目标拒绝连接，请检查 API 是否启动，以及监听地址和端口。`127.0.0.1` 指运行本机 Node.js 服务的电脑；API 在容器或虚拟机内时，需要将端口映射到这台电脑。
- 更新后的共享服务会区分拒绝连接、域名解析失败、连接中断、超时和证书错误，不展示原始错误正文或密钥。不需要为 HTTP 回环 API 关闭证书校验或添加通配 CORS。

## 存储与调用范围

- Mac 元数据位于 `~/Library/Application Support/TrioBridge AI/models/`，Windows 位于 `%LOCALAPPDATA%\TrioBridgeAI\models\`。元数据不含 API Key；本机文件限定当前账号访问。
- 密钥使用 [`@napi-rs/keyring`](https://github.com/Brooooooklyn/keyring-node) 调用系统安全存储，不经过命令行参数，也不落入明文 JSON 文件。
- 每台安装有独立随机凭证；共享接口同时检查本机凭证、来源与主机地址，只监听回环接口。凭证在加载项启动时从安装清单引导，随后从页面 URL 移除。
- 页面按已保存的模型 ID 发起调用，由本机服务读取系统密钥并访问对应地址。远程地址必须 HTTPS；HTTP 仅允许本机回环地址。固定 `/model-proxy` 的目标在保存时绑定，随后修改环境配置不会把原密钥转发到新目标。
- 本机服务不持久化文档或生成历史。开发验收的 `VITE_REAL_MODEL_AUDIT=1` 是独立的显式开关，只应使用合成数据；正式构建不包含该审计客户端。

移除 Office 清单不会删除模型和系统密钥。需要删除模型时，在任一共享设置页执行删除，三应用同步。系统安全存储暂时不可用时会显示清理未完成状态，并在后续服务启动时重试，不把已删除模型恢复到列表。

## 本机实测记录

2026-11-23 在 Mac Word、Excel、PowerPoint 中完成配置迁移、两个不同真实模型连接与切换、三应用同步、原生写回、并发冲突、服务离线和重启恢复。正式构建重新打开三个文档后仍可复用模型。验收使用两个独立模型配置，具体模型名称和默认项由使用者自行设置。

详细范围、审计用量与未覆盖平台见[验证说明](TESTING.md#真实模型与本机共享最终验收2026-11-21-至-2026-11-23)。Windows 原生验收仍待完成。
