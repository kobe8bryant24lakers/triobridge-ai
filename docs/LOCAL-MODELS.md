# 同机共享模型

本机共享版在同一台电脑、同一个系统账号下，让 Word、Excel、PowerPoint 共用模型列表和当前默认模型。可以新增多个模型，并在任一应用切换；其他应用会在刷新、重新聚焦或下一次同步时更新。

在线静态版仍使用各应用自己的存储，不能保证跨 Office 共享。微软明确不保证 cookies 或 localStorage 在不同 Office 应用之间通用，见[加载项状态与设置](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/persisting-add-in-state-and-settings)。

## 安装本机共享版

从[下载页面](https://kobe8bryant24lakers.github.io/triobridge-ai/zh-CN.html#install)选择与你的电脑和芯片匹配的 **Local / 本机共享版** ZIP，解压整个文件夹。Windows 双击 `Install.cmd`，Mac 双击 `Install.command`，按系统提示完成安装。无需另外安装 Node.js、npm 或开发工具。

安装程序为当前账号配置本机 HTTPS 服务、证书和三应用清单；Windows 为当前账号建立专用共享目录，只有创建 / 移除共享目录这一步请求管理员授权。服务随当前账号登录启动，安装完成后可以关闭窗口。包内不含个人密钥或证书，每台电脑安装时分别生成。

保存文档并完全退出、重开 Office，选择 **三桥 AI（本机共享）**。Windows 在“高级 / 更多加载项 → 共享文件夹”添加。模型设置必须显示 **本机共享**；在线版入口仍然存在时，请核对名称。

- **升级 / 修复**：下载新版 ZIP，解压后再次运行 `Install`。保留已有模型和密钥；失败会尝试恢复原服务。
- **重启**：运行包内 `Restart`。同一电脑的开发服务与安装版都使用 3000 端口，切换前停止旧服务。
- **卸载**：运行 `Uninstall`，移除本次安装的后台服务、入口与证书。模型目录和系统密钥保留；删除模型请在插件内操作。
- **系统提示**：安装包尚未签名。Mac 如阻止运行，在“系统设置 → 隐私与安全性”按提示允许。若系统请求访问 Word / Excel / PowerPoint 数据目录或信任 localhost 证书，需在本机处理。不要在聊天或日志中提供系统密码。
- **目录**：程序位于 Mac 的 `~/Library/Application Support/TrioBridge AI/local-app/` 或 Windows 的 `%LOCALAPPDATA%\TrioBridgeAI\LocalApp\`。`logs/` 保存服务启动日志，不保存模型密钥或文档正文。

安装完成后参考下方“配置与迁移”。以下源码方式只供开发和自行构建使用。

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

1. 在任一应用进入「模型设置」，确认页面显示本机共享状态。
2. 新增模型，填写服务地址、模型 ID 和 API Key，连接并保存。密钥进入 macOS 钥匙串或 Windows 系统凭据存储。
3. 若已有旧配置，使用迁移入口选择要迁移的模型。整批成功后才移除相应的旧凭证；失败或并发修改时保留来源。未选择的旧配置不被删除。
4. 在工作台切换模型，三类应用同步当前默认模型。模型更新会检查版本冲突；遇到冲突先重新读取并复核草稿，不自动覆盖另一应用的新配置。

共享模式只向页面返回模型元数据和“已保存密钥”的状态，不读回已保存的密钥。保留、替换、删除密钥是不同操作；更换服务地址不会把旧密钥带到新地址。配置导出仍不包含密钥，也不包含本机安装凭证。

服务离线后，已进入共享模式的页面保持共享模型信息并提示重新连接，不会悄悄使用旧的本应用密钥继续调用。缺少本机安装凭证时，重新执行安装清单步骤并重开 Office。

## 本地 HTTP API 与连接排查

支持 `http://127.0.0.1:8001/v1` 这样的本地 OpenAI 兼容 API。插件追加 `/chat/completions`；模型 ID 填写 API 实际提供的名称。服务无需鉴权时，明确选择“不使用密钥”。共享版由本机 Node.js 服务调用 API，不要求模型 API 提供浏览器 CORS。

**Windows 的“共享文件夹”不等于“本机共享模型”。** 它只是加载项安装目录，在线版也能通过它安装。判断依据是模型设置页显示“本机共享”，而不是“仅此应用”。

若连接时报“无法连接模型服务……CORS……本机代理”，表示当前使用浏览器直连流程。请在运行 Office 的同一电脑、同一系统账号下安装上方本机共享 ZIP，并重新选择“本机共享”入口。源码开发者可执行 `npm run manifest:local`（Windows）或 `npm run sideload:mac`（Mac）。不要安装仓库原始清单、在线版清单或另一台电脑生成的私有清单。保存文档并完全退出、重开 Office，选择名称含“本机共享”的加载项，；安装版服务自动随登录启动，源码开发者需保持 `npm start` 运行。

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
