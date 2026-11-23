import { execFile } from 'node:child_process';
import { createHash } from 'node:crypto';
import { chmod, mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { promisify } from 'node:util';

const exec = promisify(execFile);
export const INSTALLER_FILES = ['TrioBridge-AI-Mac.zip', 'TrioBridge-AI-Windows.zip'];

export async function buildInstallers(root, destination, manifest, manifestId) {
  const staging = await mkdtemp(join(tmpdir(), 'triobridge-installers-'));
  const hash = createHash('sha256').update(manifest).digest('hex');
  const render = (text) =>
    text.replaceAll('__MANIFEST_SHA256__', hash).replaceAll('__HOSTED_MANIFEST_ID__', manifestId);
  try {
    await mkdir(destination, { recursive: true });
    for (const platform of ['Mac', 'Windows']) {
      const name = `TrioBridge-AI-${platform}`;
      const directory = join(staging, name);
      await mkdir(directory);
      await writeFile(join(directory, 'manifest.xml'), manifest);
      await writeFile(
        join(directory, 'README.txt'),
        `TrioBridge AI · 三桥 AI（在线版）\n\n` +
          (platform === 'Mac'
            ? '解压后双击 Install.command。若系统阻止打开，请在系统设置的「隐私与安全性」中按提示允许，或使用网页中的手动安装步骤。\n卸载安装文件：双击 Uninstall.command。\n'
            : '先保存文档并完全退出 Word、Excel、PowerPoint，再解压整个文件夹，双击 Install.cmd。按系统提示批准创建仅当前账号可访问的共享目录。请正常双击运行，不要用另一个账号运行整个安装程序。\n卸载安装文件和共享目录：双击 Uninstall.cmd，同样可能要求管理员确认。\n只有部分应用显示入口时，运行 Diagnose.cmd 只读检查目录和三应用声明；报告不含模型密钥或文档。\n') +
          '\n保存文档并完全退出 Office，然后重新打开。在「主页 → 加载项」中添加三桥 AI 在线版；Windows 需进入「高级 / 更多加载项 → 共享文件夹」。\n' +
          '在模型设置中选择服务、填入 API Key、点击「连接并保存」。无需 Node.js、本机服务或开发证书。\n' +
          '此安装包是在线版，三个 Office 应用各自保存模型配置。同机配置一次与跨应用共享需另行安装本机共享版并保持本机服务运行。\n' +
          '本机共享版说明：https://github.com/kobe8bryant24lakers/triobridge-ai/blob/main/docs/LOCAL-MODELS.md\n' +
          '这是未签名的个人测试旁加载包，组织策略可能禁止安装。Windows 需要可用的文件共享服务；不会自动修改防火墙或组织策略。\n' +
          '卸载不会清空 Office 缓存、模型配置或密钥；需要清除模型配置时，请先在插件内删除。\n' +
          '安装与故障处理：https://kobe8bryant24lakers.github.io/triobridge-ai/#install\n',
      );
      if (platform === 'Mac') {
        const template = await readFile(join(root, 'scripts/installers/mac/manage.sh'), 'utf8');
        await writeFile(join(directory, 'manage.sh'), render(template));
        for (const [file, action] of [
          ['Install.command', 'install'],
          ['Uninstall.command', 'remove'],
        ]) {
          await writeFile(
            join(directory, file),
            `#!/bin/bash\npackage_root=$(cd "$(dirname "$0")" && pwd)\n/bin/bash "$package_root/manage.sh" ${action}\nresult=$?\nif [ "$result" -ne 0 ]; then echo '操作未完成，请检查上方错误后重试；已完成的部分可用卸载脚本移除。'; fi\nif [ -t 0 ]; then read -r -p '按回车关闭此窗口…'; fi\nexit "$result"\n`,
          );
          await chmod(join(directory, file), 0o755);
        }
      } else {
        for (const file of ['Manage.ps1', 'Share.ps1']) {
          const template = await readFile(join(root, 'scripts/installers/windows', file), 'utf8');
          // BOM allows Windows PowerShell 5.1 to read UTF-8 on non-English systems.
          await writeFile(
            join(directory, file),
            '\uFEFF' + render(template).replaceAll('\n', '\r\n'),
          );
        }
        await writeFile(
          join(directory, 'windows-office-discovery.ps1'),
          '\uFEFF' +
            (await readFile(join(root, 'scripts/windows-office-discovery.ps1'), 'utf8'))
              .replaceAll('\r\n', '\n')
              .replaceAll('\n', '\r\n'),
        );
        await writeFile(
          join(directory, 'Diagnose.cmd'),
          '@echo off\r\nsetlocal\r\nset "PSModulePath="\r\nchcp 65001 >nul\r\necho Read-only Office discovery check. No model settings or keys are read.\r\npowershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0windows-office-discovery.ps1" -Diagnose\r\nset "result=%ERRORLEVEL%"\r\npause\r\nexit /b %result%\r\n',
        );
        for (const [file, action] of [
          ['Install.cmd', 'Install'],
          ['Uninstall.cmd', 'Remove'],
        ])
          await writeFile(
            join(directory, file),
            `@echo off\r\nsetlocal\r\nset "PSModulePath="\r\nchcp 65001 >nul\r\npowershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Manage.ps1" -Action ${action}\r\nset "result=%ERRORLEVEL%"\r\npause\r\nexit /b %result%\r\n`,
          );
      }
      const output = resolve(destination, `${name}.zip`);
      await rm(output, { force: true });
      // zip preserves executable bits on the Mac launchers; only staging's allowlisted files enter it.
      await exec('zip', ['-q', '-X', '-r', output, name], { cwd: staging });
    }
  } finally {
    await rm(staging, { recursive: true, force: true });
  }
}
