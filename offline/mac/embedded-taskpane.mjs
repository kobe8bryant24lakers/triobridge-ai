import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { mkdir, mkdtemp, rm, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { buildLocalInstaller } from '../../scripts/build-local-installers.mjs';

const execute = promisify(execFile);
export const EMBEDDED_DIRECTORY = 'Embedded-Taskpane';
export const EMBEDDED_COMMANDS = {
  'Enable-Embedded-Taskpane.command': 'install',
  'Disable-Embedded-Taskpane.command': 'remove',
  'Restart-Embedded-Taskpane.command': 'restart',
};

export function embeddedTaskpaneLauncher(action) {
  if (!Object.values(EMBEDDED_COMMANDS).includes(action))
    throw new Error('Unknown embedded taskpane action.');
  const introduction =
    action === 'install'
      ? '即将启用 Office 内嵌窗格：安装本机 HTTPS 证书和随登录启动的服务。\n内嵌窗格需要联网加载微软 Office.js；独立离线工作台仍可使用。\n请先保存文档并完全退出 Word、Excel、PowerPoint。'
      : action === 'remove'
        ? '即将移除内嵌窗格、本机服务及其证书；独立工作台、模型配置和钥匙串凭据保留。\n请先保存文档并完全退出 Word、Excel、PowerPoint。'
        : '即将重启内嵌窗格的本机服务，保留两种入口的共享模型配置。';
  return `#!/bin/bash
set -euo pipefail
package_root=$(cd "$(dirname "$0")" && pwd)
cat <<'TRIOBRIDGE_NOTICE'
${introduction}
TRIOBRIDGE_NOTICE
set +e
"$package_root/${EMBEDDED_DIRECTORY}/runtime/node" "$package_root/${EMBEDDED_DIRECTORY}/app/scripts/install-local.mjs" ${action}
result=$?
set -e
if [ -t 0 ]; then read -r -p '按回车关闭此窗口…' || true; fi
exit "$result"
`;
}

export const EMBEDDED_README = `三桥 AI · Mac 可选 Office 内嵌窗格

同一安装包包含两个可独立使用的入口：
• Install.command 安装独立离线工作台，不安装 HTTPS 证书或本机 Web 服务。
• Enable-Embedded-Taskpane.command 启用 Word、Excel、PowerPoint 内的任务窗格。

启用内嵌窗格
1. 解压整个 ZIP，保存文档并完全退出 Word、Excel、PowerPoint。
2. 双击 Enable-Embedded-Taskpane.command，按系统提示信任本次安装的 localhost HTTPS 证书。
3. 重新打开 Office，在“主页 → 加载项”选择“三桥 AI（本机共享）”。
   不同 Office 版本可能将入口放在“插入 → 加载项”。
4. 模型设置应显示“本机共享”。两种入口复用当前账号的模型目录和钥匙串，不必重新填写模型。

网络与运行条件
内嵌窗格页面由包内文件通过 https://localhost:3000 提供，不依赖远程网页托管。
微软 Office.js 仍从 https://appsforoffice.microsoft.com/lib/1/hosted/office.js 加载，需要网络连接。
本包没有捆绑或修改 Office.js，浏览器缓存不保证断网可用。需要完全断网时，请使用独立工作台和本地模型。
嵌入服务只监听 127.0.0.1，会随当前账号登录启动；无需安装系统 Node 或保持安装窗口开启。
3000 端口不能同时运行另一个服务。首次使用钥匙串可能出现系统授权。
Office 或组织策略可能禁止旁加载；本入口使用微软公开的 Mac 旁加载机制，不是 Marketplace 发布。

维护
更新安装包后，再运行 Enable-Embedded-Taskpane.command 升级内嵌窗格；它可修复失败的安装。
Restart-Embedded-Taskpane.command 仅重启本机服务。
Disable-Embedded-Taskpane.command 移除内嵌窗格、登录服务和本次安装的证书，不移除独立工作台。
Uninstall.command 只移除独立工作台。完全卸载两个入口时，先运行 Disable-Embedded-Taskpane.command，再运行 Uninstall.command。
以上操作均保留模型配置和钥匙串凭据。
本包可接管以前的 TrioBridge 本机 HTTPS 共享版；不要同时启用开发服务。
生成的每账号私有清单含本机访问凭证，不能提交、上传或发给其他人。

Optional embedded Office taskpane for Mac
Enable-Embedded-Taskpane.command installs a trusted localhost HTTPS certificate and a per-user login service.
Microsoft Office.js still requires Internet access. Use the independent workbench for offline startup.
Both entries share model profiles and Keychain credentials. Disable removes only the embedded component;
Uninstall removes only the independent workbench. Run both to remove both entries. Model credentials are preserved.

Microsoft documentation:
https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac
https://learn.microsoft.com/en-us/office/dev/add-ins/develop/referencing-the-javascript-api-for-office-library-from-its-cdn
`;

/** Build-time only: never installs a certificate, LaunchAgent or Office manifest. */
export async function addMacEmbeddedTaskpane({ root, payload, arch }) {
  if (!['arm64', 'x64'].includes(arch)) throw new Error('Unsupported Mac architecture.');
  await mkdir(join(root, 'build'), { recursive: true });
  const temporary = await mkdtemp(join(root, 'build', '.mac-embedded-'));
  try {
    const webRoot = join(temporary, 'ui');
    await execute(
      process.execPath,
      [
        join(root, 'node_modules/vite/bin/vite.js'),
        'build',
        '--config',
        join(root, 'vite.config.ts'),
        '--outDir',
        webRoot,
        '--emptyOutDir',
      ],
      { cwd: root, maxBuffer: 4 * 1024 * 1024 },
    );
    await buildLocalInstaller('darwin', arch, temporary, {
      webRoot,
      unpackedDirectory: join(payload, EMBEDDED_DIRECTORY),
    });
    for (const [file, action] of Object.entries(EMBEDDED_COMMANDS))
      await writeFile(join(payload, file), embeddedTaskpaneLauncher(action), { mode: 0o755 });
    await writeFile(join(payload, 'Embedded-Taskpane-README.txt'), EMBEDDED_README);
  } finally {
    await rm(temporary, { recursive: true, force: true });
  }
}
