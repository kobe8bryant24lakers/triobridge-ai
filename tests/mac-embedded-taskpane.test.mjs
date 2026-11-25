import { test } from 'node:test';
import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { mkdtemp, mkdir, readFile, readdir, realpath, rm, stat, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import {
  addMacEmbeddedTaskpane,
  EMBEDDED_COMMANDS,
  EMBEDDED_DIRECTORY,
  EMBEDDED_README,
  embeddedTaskpaneLauncher,
} from '../offline/mac/embedded-taskpane.mjs';
import { verifyPackage } from '../scripts/local-install-files.mjs';
import { sharedModelsDirectory } from '../scripts/shared-model-store.mjs';

const execute = promisify(execFile);
const repository = dirname(dirname(fileURLToPath(import.meta.url)));
async function temporary(t) {
  const directory = await mkdtemp(join(await realpath(tmpdir()), 'triobridge embedded test '));
  t.after(() => rm(directory, { recursive: true, force: true }));
  return directory;
}

test(
  'embedded control commands invoke only their packaged installer, preserve exit status and quote paths',
  {
    skip: process.platform === 'win32',
  },
  async (t) => {
    const directory = await temporary(t);
    const runtime = join(directory, EMBEDDED_DIRECTORY, 'runtime');
    await mkdir(runtime, { recursive: true });
    // Execute the real generated shell wrappers with an inert recording executable.
    // No installer module, user directory, Keychain or LaunchAgent is accessed.
    await writeFile(
      join(runtime, 'node'),
      '#!/bin/bash\nprintf "%s\\n" "$@"\nexit "${TRIO_TEST_EXIT:-0}"\n',
      { mode: 0o755 },
    );
    for (const [file, action] of Object.entries(EMBEDDED_COMMANDS)) {
      const launcher = join(directory, file);
      await writeFile(launcher, embeddedTaskpaneLauncher(action), { mode: 0o755 });
      const { stdout } = await execute('/bin/bash', [launcher], { cwd: tmpdir() });
      assert.match(stdout, /共享|独立/);
      assert.ok(
        stdout.includes(join(directory, EMBEDDED_DIRECTORY, 'app/scripts/install-local.mjs')),
      );
      assert.match(stdout, new RegExp(`\\n${action}\\n$`));
      await assert.rejects(
        execute('/bin/bash', [launcher], {
          env: { ...process.env, TRIO_TEST_EXIT: '17' },
        }),
        (error) => error.code === 17,
      );
    }
    assert.throws(() => embeddedTaskpaneLauncher('install; unexpected'), /Unknown/);
  },
);

test('standalone installation stays certificate-free and the optional component discloses its network requirement', async () => {
  const standalone = await readFile(join(repository, 'offline/mac/Install.command'), 'utf8');
  assert.doesNotMatch(standalone, /install-local\.mjs|add-trusted-cert|launchctl|runtime\/node/);
  assert.match(embeddedTaskpaneLauncher('install'), /HTTPS 证书/);
  assert.match(embeddedTaskpaneLauncher('install'), /需要联网加载微软 Office\.js/);
  assert.match(EMBEDDED_README, /浏览器缓存不保证断网可用/);
  assert.match(
    EMBEDDED_README,
    /先运行 Disable-Embedded-Taskpane\.command，再运行 Uninstall\.command/,
  );
});

test(
  'Mac desktop ZIP component contains the current taskpane, verified runtime and shared credential implementation',
  {
    skip: process.platform !== 'darwin',
    timeout: 120_000,
  },
  async (t) => {
    const payload = await temporary(t);
    await writeFile(join(payload, 'standalone-sentinel'), 'independent workbench');
    await addMacEmbeddedTaskpane({ root: repository, payload, arch: process.arch });
    assert.equal(
      await readFile(join(payload, 'standalone-sentinel'), 'utf8'),
      'independent workbench',
    );
    const component = join(payload, EMBEDDED_DIRECTORY);
    const metadata = await verifyPackage(component, 'darwin', process.arch);
    assert.equal(
      metadata.version,
      JSON.parse(await readFile(join(repository, 'package.json'), 'utf8')).version,
    );
    const manifest = await readFile(join(component, 'app/manifest.xml'), 'utf8');
    assert.match(manifest, /https:\/\/localhost:3000\/\?office=1/);
    for (const host of ['Document', 'Workbook', 'Presentation'])
      assert.ok(manifest.includes(`Name="${host}"`));
    assert.doesNotMatch(manifest, /local-token=/);
    for (const script of [
      'shared-model-store.mjs',
      'shared-model-vault.mjs',
      'model-protocol.mjs',
      'install-local.mjs',
    ]) {
      assert.deepEqual(
        await readFile(join(component, 'app/scripts', script)),
        await readFile(join(repository, 'scripts', script)),
      );
    }
    const assets = join(component, 'app/dist/assets');
    const javascript = (
      await Promise.all(
        (await readdir(assets))
          .filter((file) => file.endsWith('.js'))
          .map((file) => readFile(join(assets, file), 'utf8')),
      )
    ).join('\n');
    assert.ok(javascript.includes('https://appsforoffice.microsoft.com/lib/1/hosted/office.js'));
    assert.equal(
      await readFile(join(payload, 'Embedded-Taskpane-README.txt'), 'utf8'),
      EMBEDDED_README,
    );
    for (const file of Object.keys(EMBEDDED_COMMANDS))
      assert.ok((await stat(join(payload, file))).mode & 0o100);

    // Load modules and the native binding from the packaged Node without touching secrets.
    const storeUrl = pathToFileURL(join(component, 'app/scripts/shared-model-store.mjs')).href;
    const vaultUrl = pathToFileURL(join(component, 'app/scripts/shared-model-vault.mjs')).href;
    const { stdout } = await execute(
      join(component, 'runtime/node'),
      [
        '--input-type=module',
        '-e',
        `
    import { sharedModelsDirectory } from ${JSON.stringify(storeUrl)};
    import { createNativeModelVault } from ${JSON.stringify(vaultUrl)};
    const { AsyncEntry } = await import('@napi-rs/keyring');
    console.log(JSON.stringify({ directory: sharedModelsDirectory(), vault: createNativeModelVault().kind, binding: typeof AsyncEntry }));
  `,
      ],
      { cwd: join(component, 'app') },
    );
    assert.deepEqual(JSON.parse(stdout), {
      directory: sharedModelsDirectory('darwin'),
      vault: 'macos-keychain',
      binding: 'function',
    });
  },
);
