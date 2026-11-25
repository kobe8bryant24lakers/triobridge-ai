import { createHash } from 'node:crypto';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import {
  chmod,
  copyFile,
  cp,
  lstat,
  mkdir,
  mkdtemp,
  readFile,
  readdir,
  rm,
  writeFile,
} from 'node:fs/promises';
import { join, resolve, dirname } from 'node:path';
import { tmpdir } from 'node:os';
import { fileURLToPath } from 'node:url';

const exec = promisify(execFile);
const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const runtimeScripts = [
  'serve.mjs',
  'serve-installed.mjs',
  'server-utils.mjs',
  'shared-model-runtime.mjs',
  'shared-model-store.mjs',
  'shared-model-http.mjs',
  'shared-model-vault.mjs',
  'local-request-guard.mjs',
  'local-capability.mjs',
  'private-local-permissions.mjs',
  'local-manifest.mjs',
  'model-connection-error.mjs',
  'install-local.mjs',
  'local-service-readiness.mjs',
  'model-storage-diagnostic.mjs',
  'local-install-files.mjs',
  'local-install-platform.mjs',
  'native-local.ps1',
  'local-catalog-share.ps1',
  'uninstall-local.ps1',
  'start-installed.ps1',
  'windows-office-discovery.ps1',
];

export const localInstallerName = (platform, arch) =>
  `TrioBridge-AI-Local-${platform === 'darwin' ? 'Mac' : 'Windows'}-${arch}`;

async function download(url, integrity, cache) {
  const [algorithm, expected] = integrity.split('-', 2);
  const matches = (data) => createHash(algorithm).update(data).digest('base64') === expected;
  try {
    const saved = await readFile(cache);
    if (matches(saved)) return saved;
  } catch (error) {
    if (error.code !== 'ENOENT') throw error;
  }
  const response = await fetch(url, { signal: AbortSignal.timeout(180_000) });
  if (!response.ok) throw new Error(`Download failed: HTTP ${response.status}`);
  const data = Buffer.from(await response.arrayBuffer());
  if (!matches(data)) throw new Error('Downloaded runtime dependency failed its pinned checksum.');
  await mkdir(dirname(cache), { recursive: true });
  await writeFile(cache, data);
  return data;
}

async function extract(archive, target) {
  await mkdir(target, { recursive: true });
  if (archive.endsWith('.zip') && process.platform !== 'win32')
    await exec('unzip', ['-q', archive, '-d', target]);
  else await exec('tar', ['-xf', archive, '-C', target]);
}

export async function packageFiles(directory, prefix = '') {
  const files = [];
  for (const name of (await readdir(join(directory, prefix))).sort()) {
    const relative = prefix ? `${prefix}/${name}` : name;
    const info = await lstat(join(directory, relative));
    if (info.isSymbolicLink()) throw new Error('Installer payload cannot contain symbolic links.');
    if (info.isDirectory()) files.push(...(await packageFiles(directory, relative)));
    else if (info.isFile()) {
      const data = await readFile(join(directory, relative));
      files.push({
        path: relative,
        sha256: createHash('sha256').update(data).digest('hex'),
        size: data.length,
        executable: Boolean(info.mode & 0o111),
      });
    } else throw new Error('Installer payload must contain only regular files.');
  }
  return files;
}

export async function buildLocalInstaller(
  platform = process.platform,
  arch = process.arch,
  output = resolve(root, 'dist-local'),
) {
  if (!['darwin', 'win32'].includes(platform) || !['arm64', 'x64'].includes(arch))
    throw new Error('Supported local packages: macOS or Windows, arm64 or x64.');
  const lock = JSON.parse(await readFile(join(root, 'scripts/local-runtime-lock.json'), 'utf8'));
  const packages = JSON.parse(await readFile(join(root, 'package-lock.json'), 'utf8')).packages;
  const { version } = JSON.parse(await readFile(join(root, 'package.json'), 'utf8'));
  const stage = await mkdtemp(join(tmpdir(), 'triobridge-local-package-'));
  const name = localInstallerName(platform, arch);
  const directory = join(stage, name);
  const cache = join(root, '.cache', 'local-runtime');
  try {
    await mkdir(join(directory, 'app/scripts'), { recursive: true });
    await mkdir(join(directory, 'runtime'), { recursive: true });
    await cp(join(root, 'dist'), join(directory, 'app/dist'), { recursive: true });
    for (const script of runtimeScripts) {
      const target = join(directory, 'app/scripts', script);
      if (script.endsWith('.ps1')) {
        const text = await readFile(join(root, 'scripts', script), 'utf8');
        await writeFile(
          target,
          '\uFEFF' +
            text
              .replace(/^\uFEFF/, '')
              .replaceAll('\r\n', '\n')
              .replaceAll('\n', '\r\n'),
        );
      } else await copyFile(join(root, 'scripts', script), target);
    }
    for (const file of ['manifest.xml', 'LICENSE'])
      await copyFile(join(root, file), join(directory, 'app', file));
    await writeFile(
      join(directory, 'app/package.json'),
      JSON.stringify(
        { name: 'triobridge-ai-local', version, type: 'module', private: true },
        null,
        2,
      ) + '\n',
    );
    const dependencyNames = [
      '@napi-rs/keyring',
      `@napi-rs/keyring-${platform}-${arch}${platform === 'win32' ? '-msvc' : ''}`,
    ];
    for (const dependency of dependencyNames) {
      const metadata = packages[`node_modules/${dependency}`];
      if (
        !metadata ||
        !/^\d+\.\d+\.\d+$/.test(metadata.version) ||
        !metadata.integrity?.startsWith('sha512-')
      )
        throw new Error('Missing locked native credential-store dependency.');
      const archive = join(cache, `${dependency.replaceAll('/', '-')}-${metadata.version}.tgz`);
      const url = `https://registry.npmjs.org/${dependency}/-/${dependency.split('/')[1]}-${metadata.version}.tgz`;
      await download(url, metadata.integrity, archive);
      const unpacked = join(stage, dependency.replaceAll('/', '-'));
      await extract(archive, unpacked);
      await cp(join(unpacked, 'package'), join(directory, 'app/node_modules', dependency), {
        recursive: true,
      });
    }
    const runtime = lock.artifacts[`${platform}-${arch}`];
    const archive = join(cache, runtime.file);
    await download(
      lock.baseUrl + runtime.file,
      `sha256-${Buffer.from(runtime.sha256, 'hex').toString('base64')}`,
      archive,
    );
    const unpacked = join(stage, 'node');
    await extract(archive, unpacked);
    const nodeRoot = join(unpacked, runtime.file.replace(/\.(tar\.gz|zip)$/, ''));
    const nodeName = platform === 'win32' ? 'node.exe' : 'node';
    await copyFile(
      join(nodeRoot, platform === 'win32' ? 'node.exe' : 'bin/node'),
      join(directory, 'runtime', nodeName),
    );
    await chmod(join(directory, 'runtime', nodeName), 0o755);
    await copyFile(join(nodeRoot, 'LICENSE'), join(directory, 'runtime/LICENSE'));
    for (const [name, action] of [
      ['Install', 'install'],
      ['Restart', 'restart'],
      ['Uninstall', 'remove'],
    ]) {
      if (platform === 'darwin') {
        const launcher = `#!/bin/bash\npackage_root=$(cd "$(dirname "$0")" && pwd)\n"$package_root/runtime/node" "$package_root/app/scripts/install-local.mjs" ${action}\nresult=$?\nif [ -t 0 ]; then read -r -p '按回车关闭此窗口…'; fi\nexit "$result"\n`;
        await writeFile(join(directory, `${name}.command`), launcher, { mode: 0o755 });
      } else {
        const launcher =
          action === 'remove'
            ? 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0app\\scripts\\uninstall-local.ps1"'
            : `"%~dp0runtime\\node.exe" "%~dp0app\\scripts\\install-local.mjs" ${action}`;
        await writeFile(
          join(directory, `${name}.cmd`),
          `@echo off\r\nsetlocal\r\nset "PSModulePath="\r\nchcp 65001 >nul\r\n${launcher}\r\nset "result=%ERRORLEVEL%"\r\npause\r\nexit /b %result%\r\n`,
        );
      }
    }
    if (platform === 'win32')
      await writeFile(
        join(directory, 'Diagnose.cmd'),
        '@echo off\r\nsetlocal\r\nset "PSModulePath="\r\nchcp 65001 >nul\r\necho Read-only local service and Office catalog check.\r\npowershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0app\\scripts\\windows-office-discovery.ps1" -Diagnose\r\nset "result=%ERRORLEVEL%"\r\npause\r\nexit /b %result%\r\n',
      );
    await writeFile(
      join(directory, 'README.txt'),
      `TrioBridge AI ${version} · 本机共享版 / Local shared edition\n\n先保存文档并完全退出 Word、Excel、PowerPoint，再解压整个文件夹，双击 Install.${platform === 'darwin' ? 'command' : 'cmd'}。无需另外安装 Node.js 或运行 npm。\n安装会请求系统证书信任${platform === 'win32' ? '和当前账号专用共享目录的管理员授权' : ''}。按系统提示批准。\n${platform === 'darwin' ? '首次使用已存密钥或升级后，macOS 可能请求本机服务访问钥匙串；在系统窗口确认后重试，无需重新填写密钥。\n' : ''}保存文档并完全退出、重开 Office，选择“三桥 AI（本机共享）”。模型设置应显示“本机共享”。\n本机服务会随当前账号登录启动；模型列表与密钥在 Word、Excel、PowerPoint 之间复用。\n本地 API 可使用 http://127.0.0.1:8001/v1；服务无需鉴权时选择“不使用密钥”。\n${platform === 'win32' ? '第 3 步失败或部分应用看不到入口时，双击 Diagnose.cmd 查看本机服务和加载项目录状态。\n' : ''}再次双击 Install 可升级或修复，Restart 可重启服务，Uninstall 可卸载服务及入口。卸载保留模型配置和系统密钥。\n安装包未签名；组织策略可能限制安装，系统可能要求确认。不会修改防火墙或关闭 TLS 验证。\n\nSave documents and completely close Word, Excel and PowerPoint, then extract the entire folder and open Install. No separate Node.js or npm setup is required. Approve the OS certificate/catalog prompts, then restart Office and select the local shared add-in. Uninstall preserves model profiles and OS-stored keys.\n\nhttps://kobe8bryant24lakers.github.io/triobridge-ai/zh-CN.html\n`,
    );
    const files = await packageFiles(directory);
    await writeFile(
      join(directory, 'package-manifest.json'),
      JSON.stringify(
        {
          format: 1,
          product: 'triobridge-ai-local',
          version,
          platform,
          arch,
          nodeVersion: lock.version,
          files,
        },
        null,
        2,
      ) + '\n',
    );
    await mkdir(output, { recursive: true });
    const zip = join(output, `${name}.zip`);
    await rm(zip, { force: true });
    if (process.platform === 'win32') {
      await exec(
        'powershell.exe',
        [
          '-NoProfile',
          '-NonInteractive',
          '-Command',
          'Compress-Archive -LiteralPath $env:TRIOBRIDGE_PACKAGE_SOURCE -DestinationPath $env:TRIOBRIDGE_PACKAGE_ZIP -CompressionLevel Optimal',
        ],
        {
          env: {
            ...process.env,
            TRIOBRIDGE_PACKAGE_SOURCE: directory,
            TRIOBRIDGE_PACKAGE_ZIP: zip,
          },
          maxBuffer: 1024 * 1024,
        },
      );
    } else await exec('zip', ['-q', '-X', '-r', zip, name], { cwd: stage });
    return zip;
  } finally {
    await rm(stage, { recursive: true, force: true });
  }
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url))
  console.log(await buildLocalInstaller(process.argv[2], process.argv[3]));
