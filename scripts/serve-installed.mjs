import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { Console } from 'node:console';
import { randomUUID } from 'node:crypto';
import { createWriteStream } from 'node:fs';
import { mkdir, rename, stat, writeFile } from 'node:fs/promises';
import { installedRoot, readInstallation, assertRegularOwned } from './local-install-files.mjs';
import { enforcePrivatePath } from './private-local-permissions.mjs';

const root = installedRoot();
const state = await readInstallation(root);
await mkdir(join(root, 'logs'), { recursive: true, mode: 0o700 });
if (process.platform === 'win32') {
  for (const name of ['service.log', 'service-error.log']) {
    const path = join(root, 'logs', name);
    try {
      if ((await stat(path)).size > 1_000_000) await rename(path, path + '.previous');
    } catch (error) {
      if (error.code !== 'ENOENT') throw error;
    }
  }
  globalThis.console = new Console({
    stdout: createWriteStream(join(root, 'logs/service.log'), { flags: 'a', mode: 0o600 }),
    stderr: createWriteStream(join(root, 'logs/service-error.log'), { flags: 'a', mode: 0o600 }),
  });
}
const app = join(root, 'versions', state.active || '', 'app');
if (!state.active || resolve(fileURLToPath(new URL('..', import.meta.url))) !== resolve(app))
  throw new Error('This is not the active installed TrioBridge runtime.');
if (process.platform === 'win32') {
  const destination = join(root, 'service-process.json');
  try {
    await assertRegularOwned(destination);
  } catch (error) {
    if (error.code !== 'ENOENT') throw error;
  }
  const temporary = join(root, `.service-process-${randomUUID()}.tmp`);
  await writeFile(
    temporary,
    JSON.stringify({ installationId: state.id, version: state.active, pid: process.pid }) + '\n',
    { mode: 0o600, flag: 'wx' },
  );
  await enforcePrivatePath(temporary, 'file', { created: true });
  await rename(temporary, destination);
}
await assertRegularOwned(join(root, 'certificates'), 'directory');
if (process.platform === 'win32') {
  process.env.TLS_PFX_PATH = join(root, 'certificates/server.pfx');
  process.env.TLS_PFX_PASSWORD_PATH = join(root, 'certificates/server.password');
  await assertRegularOwned(process.env.TLS_PFX_PATH);
  await assertRegularOwned(process.env.TLS_PFX_PASSWORD_PATH);
} else {
  process.env.TLS_CERT_PATH = join(root, 'certificates/server.crt');
  process.env.TLS_KEY_PATH = join(root, 'certificates/server.key');
  await assertRegularOwned(process.env.TLS_CERT_PATH);
  await assertRegularOwned(process.env.TLS_KEY_PATH);
}
process.env.TRIOBRIDGE_INSTALLATION_ID = state.id;
process.env.TRIOBRIDGE_SHARED_MODELS = '1';
process.chdir(app);
try {
  await import('./serve.mjs');
} catch (error) {
  console.error(
    'Installed service startup failed:',
    /^[A-Z0-9_]+$/.test(error?.code || '') ? error.code : 'STARTUP_FAILED',
  );
  process.exitCode = 1;
}
