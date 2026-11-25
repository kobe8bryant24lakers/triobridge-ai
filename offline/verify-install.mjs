import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { copyFile, mkdir, rm, stat, readFile, readdir } from 'node:fs/promises';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { verifyPackage } from '../scripts/local-install-files.mjs';
import { verifyPaneWorker } from './verify-pane-worker.mjs';
const mac = process.platform === 'darwin';
const work = resolve('build', 'native-' + process.platform + '-' + process.arch);
const extracted = join(work, 'verified-archive');
const bundle = mac ? join(extracted, 'bundle') : extracted;
const archive = resolve(
  'dist-install',
  'TrioBridge-AI-' + (mac ? 'Mac' : 'Windows') + '-' + process.arch + '.zip',
);
await rm(extracted, { recursive: true, force: true });
await mkdir(extracted, { recursive: true });
if (mac) execFileSync('/usr/bin/ditto', ['-x', '-k', archive, extracted]);
else {
  const listing = execFileSync('tar.exe', ['-tf', archive], { encoding: 'utf8' });
  const archivePaths = listing
    .trimEnd()
    .split(/\r?\n/)
    .map((entry) => entry.replace(/\/$/, ''));
  const stagedPaths = await readdir(join(work, 'bundle'), { recursive: true });
  assert.deepEqual(
    archivePaths.sort(),
    stagedPaths.map((entry) => entry.replaceAll('\\', '/')).sort(),
    'ZIP entries must use the exact relative names of the staged installer files and directories.',
  );
  const winrar = process.env.TRIOBRIDGE_TEST_WINRAR;
  if (winrar) {
    execFileSync(winrar, ['x', '-inul', '-y', archive, extracted + '\\'], { timeout: 120000 });
    console.log('WinRAR extracted the Windows installer ZIP.');
  } else {
    execFileSync('tar.exe', ['-xf', archive, '-C', extracted]);
  }
}
const app = join(bundle, mac ? 'TrioBridge AI.app' : 'Workbench');
const resources = mac ? join(app, 'Contents/Resources/app') : join(app, 'resources/app');
const { wordMarkdownHtml } = await import(pathToFileURL(join(resources, 'word-markdown.mjs')).href);
assert.match(wordMarkdownHtml('## 标题'), /<h2\b/);
assert.equal(JSON.parse(await readFile(join(resources, 'package.json'))).main, 'main.mjs');
const page = await readFile(join(resources, 'ui/offline.html'), 'utf8');
for (const path of page.matchAll(/(?:src|href)="([^"]+)"/g)) {
  assert.ok(path[1].startsWith('./'));
  await stat(join(resources, 'ui', path[1]));
}
assert.ok(!page.includes('office.js'));
assert.ok(!(await readdir(resources)).includes('.env'));
await stat(join(resources, 'LICENSE'));
await stat(join(resources, 'ui/THIRD-PARTY-NOTICES.md'));
await verifyPaneWorker(
  mac ? join(app, 'Contents/MacOS/TrioBridge AI') : join(app, 'TrioBridge AI.exe'),
  join(resources, 'pane-worker.mjs'),
);
if (mac) {
  execFileSync('/usr/bin/codesign', ['--verify', '--deep', '--strict', app]);
  execFileSync('/usr/bin/codesign', [
    '--verify',
    '--strict',
    join(bundle, 'TrioBridge AI.service'),
  ]);
  await stat(join(resources, 'native/office-mac.js'));
  await verifyPackage(join(bundle, 'Embedded-Taskpane'), 'darwin', process.arch);
  for (const operation of ['Enable', 'Disable', 'Restart']) {
    const script = await stat(join(bundle, operation + '-Embedded-Taskpane.command'));
    assert.ok(script.mode & 0o100);
  }
} else {
  const run = (cmd, args) => execFileSync(cmd, args, { stdio: 'pipe', encoding: 'utf8' });
  const setup = join(bundle, 'Setup.exe');
  const paneProgId = 'HKCU\\Software\\Classes\\TrioBridge.TaskPane';
  const paneClass = 'HKCU\\Software\\Classes\\CLSID\\{640D4723-A40B-46E5-A862-30C4254606D4}';
  await stat(
    join(bundle, 'MicrosoftEdgeWebView2RuntimeInstaller' + process.arch.toUpperCase() + '.exe'),
  );
  for (const arch of ['x64', 'x86', 'arm64'])
    await stat(join(app, 'pane/runtimes/win-' + arch + '/native/WebView2Loader.dll'));
  try {
    run(setup, []);
    run(setup, []);
    for (const host of ['Word', 'Excel', 'PowerPoint'])
      for (const view of ['32', ...(process.arch === 'ia32' ? [] : ['64'])]) {
        const val = run('reg.exe', [
          'query',
          'HKCU\\Software\\Microsoft\\Office\\' + host + '\\Addins\\TrioBridge.Workbench',
          '/v',
          'LoadBehavior',
          '/reg:' + view,
        ]);
        assert.match(val, /0x3/);
      }
    for (const arch of ['x86', 'x64'])
      run(join(work, 'ribbon-' + arch, 'Release', 'RibbonCheck.exe'), []);
    for (const view of ['32', '64']) {
      const cls = run('reg.exe', ['query', paneProgId + '\\CLSID', '/ve', '/reg:' + view]);
      assert.match(cls, /640D4723-A40B-46E5-A862-30C4254606D4/i);
    }
    const smokeDir = join(work, 'pane-check');
    await mkdir(smokeDir, { recursive: true });
    const sdk = ['Microsoft.Web.WebView2.Core.dll', 'Microsoft.Web.WebView2.WinForms.dll'];
    for (const dll of sdk) await copyFile(join(app, 'pane', dll), join(smokeDir, dll));
    for (const bitness of ['x86', 'x64']) {
      const smoke = join(smokeDir, 'PaneCheck-' + bitness + '.exe');
      execFileSync(
        join(process.env.SystemRoot, 'Microsoft.NET/Framework/v4.0.30319/csc.exe'),
        [
          '/nologo',
          '/codepage:65001',
          '/platform:' + bitness,
          '/target:exe',
          '/main:TrioBridge.TaskPaneTests',
          '/out:' + smoke,
          '/reference:Microsoft.CSharp.dll',
          '/reference:System.Web.Extensions.dll',
          '/reference:System.Windows.Forms.dll',
          '/reference:System.Drawing.dll',
          ...sdk.map((dll) => '/reference:' + join(smokeDir, dll)),
          resolve('offline/windows/Office.cs'),
          resolve('offline/windows/TaskPane.cs'),
          resolve('tests/native-taskpane-windows.cs'),
        ],
        { stdio: 'inherit' },
      );
      execFileSync(smoke, [], { stdio: 'inherit' });
      execFileSync(smoke, ['--registered', bitness, ...(bitness === 'x64' ? ['--webview'] : [])], {
        stdio: 'inherit',
        timeout: 90000,
      });
    }
    const result = JSON.parse(
      execFileSync(join(resources, 'native/Office.exe'), {
        input: JSON.stringify({ host: 'word', op: 'selection' }),
        encoding: 'utf8',
      }),
    );
    assert.equal(result.ok, false);
    assert.ok(result.message);
  } finally {
    run(setup, ['--uninstall']);
  }
  for (const view of ['32', '64'])
    for (const key of [paneClass, paneProgId])
      assert.throws(() => run('reg.exe', ['query', key, '/reg:' + view]), { status: 1 });
}
console.log('Verified packaged page, platform integration, and install payload.');
