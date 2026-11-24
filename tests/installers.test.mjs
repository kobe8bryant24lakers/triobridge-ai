import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, mkdir, readFile, readdir, rm, stat, symlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { promisify } from 'node:util';
import { test } from 'node:test';
import { buildInstallers, INSTALLER_FILES } from '../scripts/build-installers.mjs';
import { hostedManifest, HOSTED_MANIFEST_ID } from '../scripts/build-pages.mjs';

const exec = promisify(execFile);
const root = resolve(import.meta.dirname, '..');

test('downloadable packages install, update and remove only the hosted Mac add-in', async () => {
  const temporary = await mkdtemp(join(tmpdir(), 'triobridge-install-test-'));
  try {
    const output = join(temporary, 'packages');
    const manifest = hostedManifest(await readFile(join(root, 'manifest.xml'), 'utf8'));
    await buildInstallers(root, output, manifest, HOSTED_MANIFEST_ID);
    assert.deepEqual((await readdir(output)).sort(), INSTALLER_FILES);
    for (const file of INSTALLER_FILES)
      await exec('unzip', ['-q', join(output, file), '-d', temporary]);
    const mac = join(temporary, 'TrioBridge-AI-Mac');
    assert.deepEqual((await readdir(mac)).sort(), [
      'Install.command',
      'README.txt',
      'Uninstall.command',
      'manage.sh',
      'manifest.xml',
    ]);
    assert.ok((await stat(join(mac, 'Install.command'))).mode & 0o111);
    assert.ok((await stat(join(mac, 'Uninstall.command'))).mode & 0o111);
    const windows = join(temporary, 'TrioBridge-AI-Windows');
    const diagnostic = join(temporary, 'TrioBridge-Diagnose-Windows');
    assert.deepEqual((await readdir(diagnostic)).sort(), [
      'Diagnose.cmd',
      'LICENSE',
      'README.txt',
      'windows-office-discovery.ps1',
    ]);
    for (const script of ['Diagnose.cmd', 'windows-office-discovery.ps1'])
      assert.deepEqual(
        await readFile(join(diagnostic, script)),
        await readFile(join(windows, script)),
      );
    assert.deepEqual((await readdir(windows)).sort(), [
      'Diagnose.cmd',
      'Install.cmd',
      'Manage.ps1',
      'README.txt',
      'Share.ps1',
      'Uninstall.cmd',
      'manifest.xml',
      'windows-office-discovery.ps1',
    ]);
    assert.equal(await readFile(join(windows, 'manifest.xml'), 'utf8'), manifest);
    assert.doesNotMatch(
      await readFile(join(windows, 'Manage.ps1'), 'utf8'),
      /__MANIFEST_|__HOSTED_/,
    );

    // Pass a temporary home to the sourced library; never touch real Office containers.
    const userRoot = join(temporary, '用户 folder with spaces');
    const destination = (app) =>
      join(
        userRoot,
        'Library/Containers',
        `com.microsoft.${app}`,
        'Data/Documents/wef/triobridge-ai-online.xml',
      );
    await mkdir(dirname(destination('Word')), { recursive: true });
    const local = join(dirname(destination('Word')), 'triobridge-ai.xml');
    await writeFile(local, 'unrelated local development manifest');
    const manage = (action) =>
      exec('bash', [
        '-c',
        'source "$1/manage.sh"; triobridge_manage "$2" "$3" "$1"',
        'installer-test',
        mac,
        action,
        userRoot,
      ]);
    await manage('install');
    await manage('install');
    for (const app of ['Word', 'Excel', 'Powerpoint'])
      assert.equal(await readFile(destination(app), 'utf8'), manifest);
    await manage('remove');
    await manage('remove');
    for (const app of ['Word', 'Excel', 'Powerpoint'])
      await assert.rejects(stat(destination(app)), { code: 'ENOENT' });
    assert.equal(await readFile(local, 'utf8'), 'unrelated local development manifest');

    // A conflicting file in the last app must be found before writing the first app.
    await writeFile(destination('Powerpoint'), '<Id>another-add-in</Id>');
    await assert.rejects(manage('install'), /同名文件不属于/);
    await assert.rejects(stat(destination('Word')), { code: 'ENOENT' });
    await assert.rejects(manage('remove'), /同名文件不属于/);
    await rm(destination('Powerpoint'));
    await symlink(local, destination('Word'));
    await assert.rejects(manage('install'), /符号链接/);
    await rm(destination('Word'));
    await writeFile(join(mac, 'manifest.xml'), 'corrupted package');
    await assert.rejects(manage('install'), /清单缺失或损坏/);
    await assert.rejects(stat(destination('Word')), { code: 'ENOENT' });
  } finally {
    await rm(temporary, { recursive: true, force: true });
  }
});
