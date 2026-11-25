import { cp, mkdir, readFile, readdir, rm, writeFile, chmod, stat } from 'node:fs/promises';
import { execFileSync } from 'node:child_process';
import { basename, dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { packager } from './tooling/node_modules/@electron/packager/dist/index.js';
const root = resolve(dirname(fileURLToPath(import.meta.url)), '..'),
  platform = process.platform,
  arch = process.arch;
if (!['win32', 'darwin'].includes(platform) || !['arm64', 'x64'].includes(arch))
  throw Error('Build on the matching Windows or Mac system.');
const work = join(root, 'build', 'native-' + platform + '-' + arch),
  stage = join(work, 'stage'),
  output = join(work, 'package'),
  payload = join(work, 'bundle');
await rm(work, { recursive: true, force: true });
for (const d of [stage, payload, output]) await mkdir(d, { recursive: true });
const run = (name, args, opts = {}) => execFileSync(name, args, { stdio: 'inherit', ...opts });
const copy = (from, to) => cp(from, to, { recursive: true, dereference: false });
run(
  process.execPath,
  [
    join(root, 'node_modules', 'vite', 'bin', 'vite.js'),
    'build',
    '--config',
    'offline/vite.config.mjs',
  ],
  { cwd: root },
);
await copy(join(root, 'offline/app'), stage);
await copy(join(root, 'build/native-ui'), join(stage, 'ui'));
await copy(join(root, 'public/icon.svg'), join(stage, 'ui/icon.svg'));
await mkdir(join(stage, 'shared'));
await mkdir(join(stage, 'native'));
for (const f of [
  'shared-model-store',
  'shared-model-vault',
  'local-capability',
  'private-local-permissions',
  'model-connection-error',
  'model-storage-diagnostic',
])
  await copy(join(root, 'scripts', f + '.mjs'), join(stage, 'shared', f + '.mjs'));
const binding =
  platform === 'darwin' ? 'keyring-darwin-' + arch : 'keyring-win32-' + arch + '-msvc';
for (const pkg of ['keyring', binding])
  await copy(
    join(root, 'node_modules', '@napi-rs', pkg),
    join(stage, 'node_modules', '@napi-rs', pkg),
  );
const info = JSON.parse(await readFile(join(stage, 'package.json'), 'utf8'));
const name = 'TrioBridge AI';
const version = JSON.parse(await readFile(join(root, 'offline/tooling/package.json'), 'utf8'))
  .devDependencies.electron;
const assembly = (target, out, src, refs = []) => {
  const framework = join(
    process.env.SystemRoot,
    'Microsoft.NET',
    'Framework',
    'v4.0.30319',
    'csc.exe',
  );
  run(framework, [
    '/nologo',
    '/utf8output',
    '/codepage:65001',
    '/optimize+',
    '/platform:anycpu',
    '/target:' + target,
    '/out:' + out,
    ...refs.map((r) => '/reference:' + r + '.dll'),
    join(root, 'offline/windows', src),
  ]);
};
if (platform === 'darwin')
  await copy(join(root, 'offline/mac/office-mac.js'), join(stage, 'native/office-mac.js'));
else
  assembly('exe', join(stage, 'native/Office.exe'), 'Office.cs', [
    'Microsoft.CSharp',
    'System.Web.Extensions',
    'System.Core',
  ]);
const [packed] = await packager({
  dir: stage,
  out: output,
  name,
  executableName: name,
  appBundleId: 'ai.triobridge.workbench',
  appVersion: info.version,
  electronVersion: version,
  platform,
  arch,
  prune: false,
  overwrite: true,
  asar: false,
  extendInfo: {
    NSAppleEventsUsageDescription: '读取你选择的 Office 内容，并在确认后写入当前文档。',
  },
  win32metadata: {
    CompanyName: 'TrioBridge AI',
    FileDescription: 'TrioBridge AI Office Workbench',
    ProductName: name,
  },
  ...(platform === 'darwin'
    ? {
        osxSign: {
          identity: '-',
          identityValidation: false,
          optionsForFile: () => ({ hardenedRuntime: false }),
          gatekeeperAssess: false,
        },
      }
    : {}),
});
if (platform === 'win32') {
  const dest = join(payload, 'Workbench');
  await cp(packed, dest, { recursive: true });
  // Desktop Office uses its own bitness, independent of the installed Electron architecture.
  for (const [target, archName] of [
    ['Win32', 'x86'],
    ['x64', 'x64'],
  ]) {
    const build = join(work, 'ribbon-' + archName);
    run('cmake', ['-S', join(root, 'offline/windows'), '-B', build, '-A', target]);
    run('cmake', ['--build', build, '--config', 'Release']);
    await copy(
      join(build, 'Release', 'Ribbon.dll'),
      join(dest, 'TrioBridge.Ribbon.' + archName + '.dll'),
    );
  }
  assembly('exe', join(payload, 'Setup.exe'), 'Setup.cs', ['System.Core']);
  await writeFile(join(dest, 'workbench.id'), 'ai.triobridge.workbench\n');
  for (const [file, args] of [
    ['Install.cmd', ''],
    ['Uninstall.cmd', ' --uninstall'],
  ])
    await writeFile(
      join(payload, file),
      '@echo off\r\n"%~dp0Setup.exe"' +
        args +
        '\r\nset trio_result=%ERRORLEVEL%\r\npause\r\nexit /b %trio_result%\r\n',
    );
} else {
  const app = join(payload, name + '.app');
  run('/usr/bin/ditto', [join(packed, name + '.app'), app]);
  const service = join(payload, name + '.service', 'Contents');
  await mkdir(join(service, 'MacOS'), { recursive: true });
  const ids = {
    word: 'com.microsoft.Word',
    excel: 'com.microsoft.Excel',
    powerpoint: 'com.microsoft.Powerpoint',
  };
  const services = Object.entries(ids)
    .map(
      ([host, id]) =>
        '<dict><key>NSMenuItem</key><dict><key>default</key><string>打开三桥 AI（' +
        { word: 'Word', excel: 'Excel', powerpoint: 'PowerPoint' }[host] +
        '）</string></dict><key>NSMessage</key><string>openWorkbench</string><key>NSPortName</key><string>TrioBridge Office Services</string><key>NSUserData</key><string>' +
        host +
        '</string><key>NSRequiredContext</key><dict><key>NSApplicationIdentifier</key><string>' +
        id +
        '</string></dict></dict>',
    )
    .join('');
  await writeFile(
    join(service, 'Info.plist'),
    '<?xml version="1.0" encoding="UTF-8"?><plist version="1.0"><dict><key>CFBundleIdentifier</key><string>ai.triobridge.workbench.service</string><key>CFBundleExecutable</key><string>TrioBridgeService</string><key>CFBundleName</key><string>TrioBridge Office Services</string><key>CFBundlePackageType</key><string>APPL</string><key>CFBundleVersion</key><string>' +
      info.version +
      '</string><key>CFBundleShortVersionString</key><string>' +
      info.version +
      '</string><key>LSBackgroundOnly</key><true/><key>NSServices</key><array>' +
      services +
      '</array></dict></plist>',
  );
  run('xcrun', [
    'swiftc',
    '-O',
    join(root, 'offline/mac/Service.swift'),
    '-o',
    join(service, 'MacOS/TrioBridgeService'),
  ]);
  run('/usr/bin/codesign', ['--force', '--sign', '-', dirname(service)]);
  for (const f of ['Install.command', 'Uninstall.command']) {
    await copy(join(root, 'offline/mac', f), join(payload, f));
    await chmod(join(payload, f), 0o755);
  }
}
const dist = join(root, 'dist-install');
await mkdir(dist, { recursive: true });
const zip = join(
  dist,
  'TrioBridge-AI-' + (platform === 'darwin' ? 'Mac' : 'Windows') + '-' + arch + '.zip',
);
await rm(zip, { force: true });
if (platform === 'darwin')
  run('/usr/bin/ditto', ['-c', '-k', '--keepParent', '--norsrc', basename(payload), zip], {
    cwd: work,
  });
else run('tar.exe', ['-a', '-cf', zip, '.'], { cwd: payload });
console.log(basename(zip), Math.round((await stat(zip)).size / 1048576) + ' MiB');
