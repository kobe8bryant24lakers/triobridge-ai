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
    for (const arch of ['x86', 'x64'])
      run(join(bundle, '..', 'ribbon-' + arch, 'Release', 'RibbonCheck.exe'), []);
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
