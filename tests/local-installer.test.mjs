import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  mkdtemp,
  mkdir,
  readFile,
  readdir,
  realpath,
  rm,
  symlink,
  writeFile,
} from 'node:fs/promises';
import { join, dirname } from 'node:path';
import { tmpdir } from 'node:os';
import { packageFiles } from '../scripts/build-local-installers.mjs';
import {
  initializeInstallation,
  readInstallation,
  writeInstallation,
  verifyPackage,
  stagePackage,
} from '../scripts/local-install-files.mjs';
import { recoverInstallation } from '../scripts/install-local.mjs';

async function fixture(t) {
  const directory = await mkdtemp(join(await realpath(tmpdir()), 'triobridge-local-package-test-'));
  t.after(() => rm(directory, { recursive: true, force: true }));
  const source = join(directory, 'source');
  const root = join(directory, 'installed');
  for (const name of [
    'app/dist/index.html',
    'app/scripts/serve-installed.mjs',
    'app/scripts/install-local.mjs',
    'app/manifest.xml',
    process.platform === 'win32' ? 'runtime/node.exe' : 'runtime/node',
    ...(process.platform === 'win32'
      ? ['app/scripts/start-installed.ps1', 'app/scripts/windows-office-discovery.ps1']
      : []),
  ]) {
    await mkdir(dirname(join(source, name)), { recursive: true });
    await writeFile(join(source, name), 'synthetic payload');
  }
  const metadata = {
    format: 1,
    product: 'triobridge-ai-local',
    version: '0.4.0',
    platform: process.platform,
    arch: process.arch,
    files: await packageFiles(source),
  };
  await writeFile(join(source, 'package-manifest.json'), JSON.stringify(metadata));
  return { directory, source, root, metadata };
}

test('local package staging verifies bytes and does not copy adjacent private files', async (t) => {
  const { source, root } = await fixture(t);
  await writeFile(join(source, '.env'), 'private-test-value');
  const state = await initializeInstallation(root);
  const metadata = await verifyPackage(source);
  const staged = await stagePackage(source, root, metadata);
  assert.equal(staged, await stagePackage(source, root, metadata));
  assert.equal((await verifyPackage(join(root, 'versions', staged))).digest, metadata.digest);
  assert.equal((await readdir(join(root, 'versions', staged))).includes('.env'), false);
  await writeInstallation(root, { ...state, active: staged });
  assert.equal((await readInstallation(root)).active, staged);
});

test('damaged, incompatible, linked and traversal packages fail before installation', async (t) => {
  const { source, root, metadata } = await fixture(t);
  await assert.rejects(verifyPackage(source, 'other-platform'), /不适用/);
  await writeFile(join(source, 'app/dist/index.html'), 'damaged');
  await assert.rejects(verifyPackage(source), /损坏/);
  await rm(join(source, 'app/dist/index.html'));
  await symlink(join(source, 'app/manifest.xml'), join(source, 'app/dist/index.html'));
  await assert.rejects(verifyPackage(source), /普通文件/);
  await writeFile(
    join(source, 'package-manifest.json'),
    JSON.stringify({
      ...metadata,
      files: [{ path: '../outside', size: 0, sha256: '0'.repeat(64) }],
    }),
  );
  await assert.rejects(verifyPackage(source), /清单无效/);
  await assert.rejects(readFile(join(root, 'installation.json')), { code: 'ENOENT' });
});

test('installer never adopts unrelated directories or linked installation roots', async (t) => {
  const { root, directory } = await fixture(t);
  await mkdir(root);
  await writeFile(join(root, 'user-document.txt'), 'keep');
  await assert.rejects(initializeInstallation(root), /已有其他内容/);
  assert.equal(await readFile(join(root, 'user-document.txt'), 'utf8'), 'keep');
  const link = join(directory, 'link');
  await symlink(root, link, 'dir');
  await assert.rejects(initializeInstallation(link), /链接/);
});

test('an interrupted upgrade restores the old service and certificate without touching model storage', async (t) => {
  const { root, directory } = await fixture(t);
  const state = await initializeInstallation(root);
  const previous = '0.3.0-aaaaaaaaaaaa';
  const active = '0.4.0-bbbbbbbbbbbb';
  await writeInstallation(root, { ...state, active });
  const backup = '.cert-before-11111111-1111-1111-1111-111111111111';
  const next = '.cert-next-22222222-2222-2222-2222-222222222222';
  await mkdir(join(root, backup));
  await writeFile(join(root, backup, 'server.crt'), 'old-certificate');
  await mkdir(join(root, 'certificates'));
  await writeFile(join(root, 'certificates/server.crt'), 'new-certificate');
  await writeFile(join(directory, 'profiles.json'), 'private-model-metadata');
  await writeFile(
    join(root, 'transaction.json'),
    JSON.stringify({
      id: state.id,
      previous,
      hadCertificate: true,
      nextCertificate: next,
      previousCertificate: backup,
    }),
  );
  const events = [];
  const driver = {
    async stop() {
      events.push('stop');
    },
    async removeCertificate(path, id) {
      assert.equal(id, state.id);
      assert.equal(await readFile(join(path, 'server.crt'), 'utf8'), 'new-certificate');
      events.push('remove-new-certificate');
    },
    async start(restored) {
      assert.equal(restored.active, previous);
      events.push('start-old');
    },
    async registerUninstaller() {
      events.push('restore-uninstaller');
    },
  };
  await recoverInstallation(root, driver, async () => {
    events.push('verify-old');
  });
  assert.deepEqual(events, [
    'stop',
    'remove-new-certificate',
    'start-old',
    'verify-old',
    'restore-uninstaller',
  ]);
  assert.equal((await readInstallation(root)).active, previous);
  assert.equal(await readFile(join(root, 'certificates/server.crt'), 'utf8'), 'old-certificate');
  assert.equal(await readFile(join(directory, 'profiles.json'), 'utf8'), 'private-model-metadata');
  await assert.rejects(readFile(join(root, 'transaction.json')), { code: 'ENOENT' });
});

test('invalid recovery paths do not stop services or touch files', async (t) => {
  const { root } = await fixture(t);
  const state = await initializeInstallation(root);
  await writeFile(
    join(root, 'transaction.json'),
    JSON.stringify({
      id: state.id,
      previous: null,
      nextCertificate: '../models',
      previousCertificate: '../documents',
    }),
  );
  await assert.rejects(
    recoverInstallation(root, {
      stop() {
        assert.fail('must not stop');
      },
    }),
    /恢复记录无效/,
  );
});
