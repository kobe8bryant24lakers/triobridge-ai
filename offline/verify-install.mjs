import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { stat, readFile, readdir } from 'node:fs/promises';
import { join, resolve } from 'node:path';
const mac = process.platform === 'darwin';
const bundle = resolve('build', 'native-' + process.platform + '-' + process.arch, 'bundle');
const app = join(bundle, mac ? 'TrioBridge AI.app' : 'Workbench');
const resources = mac ? join(app, 'Contents/Resources/app') : join(app, 'resources/app');
assert.equal(JSON.parse(await readFile(join(resources, 'package.json'))).main, 'main.mjs');
const page = await readFile(join(resources, 'ui/offline.html'), 'utf8');
for (const path of page.matchAll(/(?:src|href)="([^"]+)"/g)) {
  assert.ok(path[1].startsWith('./'));
  await stat(join(resources, 'ui', path[1]));
}
assert.ok(!page.includes('office.js'));
assert.ok(!(await readdir(resources)).includes('.env'));
if (mac) {
  execFileSync('/usr/bin/codesign', ['--verify', '--deep', '--strict', app]);
  execFileSync('/usr/bin/codesign', [
    '--verify',
    '--strict',
    join(bundle, 'TrioBridge AI.service'),
  ]);
  await stat(join(resources, 'native/office-mac.js'));
} else {
  const run = (cmd, args) => execFileSync(cmd, args, { stdio: 'pipe', encoding: 'utf8' });
  const setup = join(bundle, 'Setup.exe');
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
    const ps = join(process.env.SystemRoot, 'System32/WindowsPowerShell/v1.0/powershell.exe');
    const verify =
      "$ErrorActionPreference='Stop'; $t=[type]::GetTypeFromProgID('TrioBridge.Workbench',$true); $o=[Activator]::CreateInstance($t); $a=[Reflection.Assembly]::LoadFrom((Join-Path $env:LOCALAPPDATA 'TrioBridgeAI\\Workbench\\TrioBridge.Ribbon.dll')); $i=$a.GetType('TrioBridge.IRibbon'); $x=[xml]$i.GetMethod('GetCustomUI').Invoke($o,@('Microsoft.Word.Document')); if($x.customUI.ribbon.tabs.tab.group.button.onAction -ne 'OpenWorkbench'){throw 'Invalid ribbon'}";
    run(ps, ['-NoProfile', '-NonInteractive', '-Command', verify]);
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
}
console.log('Verified packaged page, platform integration, and install payload.');
